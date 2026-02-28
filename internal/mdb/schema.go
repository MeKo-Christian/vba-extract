package mdb

import (
	"fmt"
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

// QueryDef is a named saved query.
// SQL may be empty if it could not be read from MSysQueries.
type QueryDef struct {
	Name string
	SQL  string
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
// Rows with Attribute == 0 carry the SQL in the Expression column.
func (db *Database) readQueries(names map[string]struct{}) []QueryDef {
	sqlByName := make(map[string]string)

	td, err := db.FindTable("MSysQueries")
	if err == nil {
		rows, rErr := td.ReadRows()
		if rErr == nil {
			for _, row := range rows {
				if intField(row, "Attribute") != 0 {
					continue
				}

				name := stringField(row, "Name")
				if _, wanted := names[name]; !wanted || name == "" {
					continue
				}

				if expr := stringField(row, "Expression"); expr != "" {
					sqlByName[name] = expr
				}
			}
		}
	}

	queries := make([]QueryDef, 0, len(names))
	for name := range names {
		queries = append(queries, QueryDef{Name: name, SQL: sqlByName[name]})
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
	case uint32:
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
