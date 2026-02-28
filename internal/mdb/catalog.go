package mdb

import (
	"fmt"
)

// ObjectType values from MSysObjects.Type column.
const (
	ObjTypeLocalTable = 1
	ObjTypeLinkedAcc  = 2
	ObjTypeLinkedODBC = 4
	ObjTypeQuery      = 5
	ObjTypeForm       = -32768
	ObjTypeReport     = -32764
	ObjTypeMacro      = -32766
	ObjTypeModule     = -32761
)

// CatalogEntry represents an entry in MSysObjects.
type CatalogEntry struct {
	ID       int32
	ParentID int32
	Name     string
	Type     int16
	Flags    int32
}

// Catalog reads the MSysObjects table and returns all catalog entries.
func (db *Database) Catalog() ([]CatalogEntry, error) {
	td, err := db.ReadTableDef(2) // MSysObjects always at page 2
	if err != nil {
		return nil, fmt.Errorf("mdb: catalog: %w", err)
	}

	rows, err := td.ReadRows()
	if err != nil {
		return nil, fmt.Errorf("mdb: catalog: %w", err)
	}

	var entries []CatalogEntry

	for _, row := range rows {
		e := CatalogEntry{}

		if v, ok := row["Id"].(int32); ok {
			e.ID = v
		}

		if v, ok := row["ParentId"].(int32); ok {
			e.ParentID = v
		}

		if v, ok := row["Name"].(string); ok {
			e.Name = v
		}

		if v, ok := row["Type"].(int16); ok {
			e.Type = v
		}

		if v, ok := row["Flags"].(int32); ok {
			e.Flags = v
		}

		if e.Name != "" {
			entries = append(entries, e)
		}
	}

	return entries, nil
}

// FindTable looks up a table by name in the catalog and returns its TDEF page number.
// The ID in MSysObjects IS the TDEF page number for local tables.
func (db *Database) FindTable(name string) (*TableDef, error) {
	entries, err := db.Catalog()
	if err != nil {
		return nil, err
	}

	for _, e := range entries {
		if e.Name == name && e.Type == ObjTypeLocalTable {
			return db.ReadTableDef(int64(e.ID))
		}
	}

	return nil, fmt.Errorf("mdb: table %q not found", name)
}

// TableNames returns a list of all user table names.
func (db *Database) TableNames() ([]string, error) {
	entries, err := db.Catalog()
	if err != nil {
		return nil, err
	}

	var names []string

	for _, e := range entries {
		if e.Type == ObjTypeLocalTable {
			names = append(names, e.Name)
		}
	}

	return names, nil
}
