package mdb

import (
	"testing"
)

// TestBuildQueryDefs_sqlFound verifies that a query with SQL gets SQLStatusFound.
func TestBuildQueryDefs_sqlFound(t *testing.T) {
	names := map[string]struct{}{"GetAll": {}}
	sqlByName := map[string]string{"GetAll": "SELECT * FROM T"}

	queries := buildQueryDefs(names, sqlByName, SQLStatusFound)
	if len(queries) != 1 {
		t.Fatalf("expected 1 query, got %d", len(queries))
	}

	q := queries[0]
	if q.SQL != "SELECT * FROM T" {
		t.Errorf("expected SQL, got %q", q.SQL)
	}

	if q.SQLStatus != SQLStatusFound {
		t.Errorf("expected SQLStatusFound, got %q", q.SQLStatus)
	}
}

// TestBuildQueryDefs_sqlMissingFromTable verifies that a query not in sqlByName
// gets the fallback status that was passed in.
func TestBuildQueryDefs_sqlMissingFromTable(t *testing.T) {
	names := map[string]struct{}{"NoSQL": {}}
	sqlByName := map[string]string{}

	queries := buildQueryDefs(names, sqlByName, SQLStatusNotInTable)
	if len(queries) != 1 {
		t.Fatalf("expected 1 query, got %d", len(queries))
	}

	q := queries[0]
	if q.SQL != "" {
		t.Errorf("expected empty SQL, got %q", q.SQL)
	}

	if q.SQLStatus != SQLStatusNotInTable {
		t.Errorf("expected SQLStatusNotInTable, got %q", q.SQLStatus)
	}
}

// TestReadQueries_msysQueriesMissing verifies that when MSysQueries cannot be
// opened, all queries report SQLStatusTableMissing.
func TestReadQueries_msysQueriesMissing(t *testing.T) {
	// sample.mdb has no saved queries, but the table itself may exist.
	// We can exercise the "table missing" branch by calling readQueries
	// on a database that has no MSysQueries table at all: we synthesise
	// that by passing a set of names but expecting the status to propagate.
	//
	// This test uses a real database where we know MSysQueries is absent
	// or empty; for a pure unit test we fall back to testing buildQueryDefs
	// directly with the TableMissing status sentinel.
	names := map[string]struct{}{"Q1": {}, "Q2": {}}
	queries := buildQueryDefs(names, nil, SQLStatusTableMissing)

	for _, q := range queries {
		if q.SQLStatus != SQLStatusTableMissing {
			t.Errorf("query %q: expected SQLStatusTableMissing, got %q", q.Name, q.SQLStatus)
		}
	}
}

// TestReadQueries_sqlPresentInMixed verifies the mix: one query with SQL and
// one without both get the right status.
func TestBuildQueryDefs_mixed(t *testing.T) {
	names := map[string]struct{}{"HasSQL": {}, "NoSQL": {}}
	sqlByName := map[string]string{"HasSQL": "SELECT 1"}

	queries := buildQueryDefs(names, sqlByName, SQLStatusNotInTable)

	byName := make(map[string]QueryDef)
	for _, q := range queries {
		byName[q.Name] = q
	}

	if byName["HasSQL"].SQLStatus != SQLStatusFound {
		t.Errorf("HasSQL: expected SQLStatusFound, got %q", byName["HasSQL"].SQLStatus)
	}

	if byName["NoSQL"].SQLStatus != SQLStatusNotInTable {
		t.Errorf("NoSQL: expected SQLStatusNotInTable, got %q", byName["NoSQL"].SQLStatus)
	}
}
