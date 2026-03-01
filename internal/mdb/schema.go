package mdb

import (
	"fmt"
	"math"
	"sort"
	"strings"
)

// Schema describes the complete schema of a Jet4 database.
type Schema struct {
	Tables        []TableSchema
	Relationships []Relationship
	Queries       []QueryDef
}

// TableSchema describes one user table.
type TableSchema struct {
	Name    string
	Columns []ColumnDef
}

// ColumnDef describes one column in a table.
type ColumnDef struct {
	Name          string
	JetType       byte
	SQLType       string // mapped SQL type string, e.g. "VARCHAR(50)", "INTEGER"
	Size          int    // character count for TEXT/BINARY; 0 for fixed-size types
	Required      bool
	AutoIncrement bool
}

// Relationship describes a foreign-key link between two tables,
// read from MSysRelationships.
type Relationship struct {
	Name          string
	FromTable     string
	FromColumns   []string
	ToTable       string
	ToColumns     []string
	CascadeUpdate bool
	CascadeDelete bool
}

// SQLStatus describes why a query's SQL text is present or absent.
type SQLStatus string

const (
	// SQLStatusFound means the SQL text was successfully read from MSysQueries.
	SQLStatusFound SQLStatus = "found"
	// SQLStatusTableMissing means MSysQueries could not be opened at all
	// (e.g. the table does not exist or the layout is unsupported).
	SQLStatusTableMissing SQLStatus = "table-missing"
	// SQLStatusNotInTable means MSysQueries was opened but contained no
	// Attribute=0 row matching this query name.
	SQLStatusNotInTable SQLStatus = "not-in-table"
)

// QueryDef is a named saved query.
type QueryDef struct {
	Name      string
	SQL       string
	SQLStatus SQLStatus
}

// MSysRelationships cascade flags.
const (
	relFlagCascadeUpdate = 0x0100
	relFlagCascadeDelete = 0x1000
)

// colFlagAutoIncrement is the Jet4 column-flag bit for AutoNumber columns.
const colFlagAutoIncrement = 0x04

// ReadSchema reads all user tables, relationships, and saved queries from the
// database. Unreadable tables and system tables (MSys*) are silently skipped.
func (db *Database) ReadSchema() (*Schema, error) {
	entries, err := db.Catalog()
	if err != nil {
		return nil, fmt.Errorf("mdb: ReadSchema: catalog: %w", err)
	}

	s := &Schema{}
	queryNames := make(map[string]struct{})

	for _, e := range entries {
		switch e.Type {
		case ObjTypeLocalTable:
			if strings.HasPrefix(e.Name, "MSys") {
				continue
			}

			ts, tErr := db.readTableSchema(int64(e.ID), e.Name)
			if tErr != nil {
				continue
			}

			s.Tables = append(s.Tables, ts)

		case ObjTypeQuery:
			queryNames[e.Name] = struct{}{}
		}
	}

	sort.Slice(s.Tables, func(i, j int) bool {
		return s.Tables[i].Name < s.Tables[j].Name
	})

	s.Relationships = db.readRelationships()
	s.Queries = db.readQueries(queryNames)

	return s, nil
}

func (db *Database) readTableSchema(tdefPage int64, name string) (TableSchema, error) {
	td, err := db.ReadTableDef(tdefPage)
	if err != nil {
		return TableSchema{}, err
	}

	// Sort columns by ColNum for correct field order.
	cols := make([]*Column, len(td.Columns))
	copy(cols, td.Columns)
	sort.Slice(cols, func(i, j int) bool {
		return cols[i].ColNum < cols[j].ColNum
	})

	ts := TableSchema{Name: name}

	for _, col := range cols {
		size := 0
		if col.Type == ColTypeText || col.Type == ColTypeBinary {
			// Jet4 stores text length in bytes (UCS-2); divide by 2 for char count.
			size = int(col.Length) / 2
		}

		ts.Columns = append(ts.Columns, ColumnDef{
			Name:          col.Name,
			JetType:       col.Type,
			SQLType:       jetTypeToSQL(col.Type, int(col.Length), col.Scale, col.Precision),
			Size:          size,
			Required:      !col.IsNullable(),
			AutoIncrement: col.Flags&colFlagAutoIncrement != 0,
		})
	}

	return ts, nil
}

// readRelationships reads MSysRelationships. Each row represents one column
// pair; rows are grouped by relationship name.
func (db *Database) readRelationships() []Relationship {
	td, err := db.FindTable("MSysRelationships")
	if err != nil {
		return nil
	}

	rows, err := td.ReadRows()
	if err != nil {
		return nil
	}

	byName := make(map[string]*Relationship)
	var order []string

	for _, row := range rows {
		relName := stringField(row, "szRelationship")
		if relName == "" {
			continue
		}

		rel, exists := byName[relName]
		if !exists {
			grbit := intField(row, "grbit")
			rel = &Relationship{
				Name:          relName,
				FromTable:     stringField(row, "szObject"),
				ToTable:       stringField(row, "szReferencedObject"),
				CascadeUpdate: grbit&relFlagCascadeUpdate != 0,
				CascadeDelete: grbit&relFlagCascadeDelete != 0,
			}
			byName[relName] = rel
			order = append(order, relName)
		}

		if col := stringField(row, "szColumn"); col != "" {
			rel.FromColumns = append(rel.FromColumns, col)
		}

		if col := stringField(row, "szReferencedColumn"); col != "" {
			rel.ToColumns = append(rel.ToColumns, col)
		}
	}

	rels := make([]Relationship, 0, len(order))
	for _, name := range order {
		rels = append(rels, *byName[name])
	}

	return rels
}

// readQueries fetches SQL text from MSysQueries for each known query name.
//
// MSysQueries layout:
//   - Rows are keyed by ObjectId, which matches the ID column in MSysObjects (the catalog).
//   - The SQL text is in the "Expression" column on the row with Attribute == 0.
//   - Expression is a Memo/LVAL column: ReadRows returns raw reference bytes that must be
//     resolved via ResolveMemo and then decoded as UCS-2.
//   - Some query types (append, update, delete, crosstab) store their definition in
//     structured Attribute=5/6/7 rows instead of a single SQL string; for those,
//     Attribute=0 Expression is empty and SQLStatusNotInTable is returned.
//
// Each returned QueryDef carries a SQLStatus explaining why SQL is present or absent.
func (db *Database) readQueries(names map[string]struct{}) []QueryDef {
	td, err := db.FindTable("MSysQueries")
	if err != nil {
		return buildQueryDefs(names, nil, SQLStatusTableMissing)
	}

	rows, err := td.ReadRows()
	if err != nil {
		return buildQueryDefs(names, nil, SQLStatusTableMissing)
	}

	// Build ObjectId → catalog query name map from the catalog entries.
	// MSysObjects stores the name; MSysQueries rows reference it via ObjectId = MSysObjects.ID.
	entries, err := db.Catalog()
	if err != nil {
		return buildQueryDefs(names, nil, SQLStatusTableMissing)
	}

	nameByObjID := make(map[int32]string)
	for _, e := range entries {
		if e.Type != ObjTypeQuery {
			continue
		}
		if _, wanted := names[e.Name]; wanted {
			nameByObjID[int32(e.ID)] = e.Name
		}
	}

	// Collect SQL from Attribute=0 rows, joined to catalog names via ObjectId.
	sqlByName := make(map[string]string)
	for _, row := range rows {
		if intField(row, "Attribute") != 0 {
			continue
		}
		oid, _ := row["ObjectId"].(int32)

		name, ok := nameByObjID[oid]
		if !ok {
			continue
		}

		// Expression is a Memo/LVAL: raw bytes returned by ReadRows need ResolveMemo + UCS-2 decode.
		exprRaw, _ := row["Expression"].([]byte)
		if len(exprRaw) == 0 {
			continue
		}

		var resolved []byte
		resolved, err = db.ResolveMemo(exprRaw)
		if err != nil || len(resolved) == 0 {
			continue
		}

		sqlByName[name] = decodeUCS2(resolved)
	}

	return buildQueryDefs(names, sqlByName, SQLStatusNotInTable)
}

// buildQueryDefs assembles a sorted []QueryDef from a name set and a SQL map.
// missingStatus is assigned to any query whose name is absent from sqlByName.
func buildQueryDefs(names map[string]struct{}, sqlByName map[string]string, missingStatus SQLStatus) []QueryDef {
	queries := make([]QueryDef, 0, len(names))

	for name := range names {
		sql := sqlByName[name]
		status := missingStatus

		if sql != "" {
			status = SQLStatusFound
		}

		queries = append(queries, QueryDef{Name: name, SQL: sql, SQLStatus: status})
	}

	sort.Slice(queries, func(i, j int) bool {
		return queries[i].Name < queries[j].Name
	})

	return queries
}

// stringField extracts a string value from a Row, returning "" if absent or wrong type.
func stringField(row Row, key string) string {
	v, _ := row[key].(string)
	return v
}

// intField extracts an integer value from a Row, handling int16/int32.
func intField(row Row, key string) int32 {
	switch v := row[key].(type) {
	case int32:
		return v
	case int16:
		return int32(v)
	case int8:
		return int32(v)
	case uint8:
		return int32(v)
	case uint32:
		if v > math.MaxInt32 {
			return 0
		}

		return int32(v)
	}

	return 0
}

// jetTypeToSQL maps a Jet4 column type to a SQL type string.
func jetTypeToSQL(t byte, length int, scale, precision byte) string {
	switch t {
	case ColTypeBool:
		return "BOOLEAN"
	case ColTypeByte:
		return "TINYINT"
	case ColTypeInt:
		return "SMALLINT"
	case ColTypeLong:
		return "INTEGER"
	case ColTypeMoney:
		return "DECIMAL(19,4)"
	case ColTypeFloat:
		return "REAL"
	case ColTypeDouble:
		return "DOUBLE PRECISION"
	case ColTypeDatetime:
		return "DATETIME"
	case ColTypeText:
		chars := length / 2
		if chars <= 0 {
			chars = 255
		}

		return fmt.Sprintf("VARCHAR(%d)", chars)
	case ColTypeMemo:
		return "TEXT"
	case ColTypeBinary:
		chars := length / 2
		if chars <= 0 {
			chars = 255
		}

		return fmt.Sprintf("BINARY(%d)", chars)
	case ColTypeOLE:
		return "OLE"
	case ColTypeGUID:
		return "CHAR(38)"
	case ColTypeNumeric:
		if precision == 0 {
			return "DECIMAL"
		}

		return fmt.Sprintf("DECIMAL(%d,%d)", precision, scale)
	default:
		return fmt.Sprintf("TYPE_0x%02X", t)
	}
}
