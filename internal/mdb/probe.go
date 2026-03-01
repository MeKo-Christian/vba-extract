package mdb

import (
	"fmt"
)

const (
	LayoutJet32K  = "jet3-2k"
	LayoutJet44K  = "jet4-4k"
	LayoutLegacy4 = "legacy-4k"
	LayoutUnknown = "unknown"
)

// ProbeResult summarizes storage-layout and parser viability signals for one DB.
type ProbeResult struct {
	Path                string
	JetVersion          uint32
	PageSize            int64
	LayoutClass         string
	Page1Type           byte
	Page2Type           byte
	MSysObjectsTDEF     bool
	MSysObjectsReadable bool
	MSysObjectsError    string
	HasDataPageInSample bool
	SampleScannedPages  int
}

// Probe inspects DB layout and parser viability without extracting data.
func Probe(path string) (*ProbeResult, error) {
	db, err := Open(path)
	if err != nil {
		return nil, fmt.Errorf("mdb: probe: open %q: %w", path, err)
	}
	defer db.Close()

	result := &ProbeResult{
		Path:       path,
		JetVersion: db.Header.JetVersion,
		PageSize:   db.PageSize(),
	}

	pg1, err := db.ReadPage(1)
	if err == nil && len(pg1) > 0 {
		result.Page1Type = pg1[0]
	}

	pg2, err := db.ReadPage(2)
	if err == nil && len(pg2) > 0 {
		result.Page2Type = pg2[0]
		result.MSysObjectsTDEF = PageType(pg2) == PageTypeTDEF
	}

	switch {
	case db.PageSize() == PageSizeJet3:
		result.LayoutClass = LayoutJet32K
	case db.Header.JetVersion == JetVersion3 || db.Header.JetVersion == 259:
		result.LayoutClass = LayoutLegacy4
	case db.PageSize() == PageSizeJet4:
		result.LayoutClass = LayoutJet44K
	default:
		result.LayoutClass = LayoutUnknown
	}

	td, err := db.ReadTableDef(2)
	if err != nil {
		result.MSysObjectsError = err.Error()
	} else {
		result.MSysObjectsReadable = td != nil && len(td.Columns) > 0
	}

	const maxSamplePages = 64

	limit := min(int(db.PageCount()), maxSamplePages)

	result.SampleScannedPages = limit
	for i := 1; i < limit; i++ {
		page, pErr := db.ReadPage(int64(i))
		if pErr != nil || len(page) == 0 {
			continue
		}

		if PageType(page) == PageTypeData {
			result.HasDataPageInSample = true
			break
		}
	}

	return result, nil
}
