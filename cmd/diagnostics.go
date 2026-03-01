package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
	"github.com/MeKo-Christian/accessdump/internal/vba"
)

func layoutClass(pageSize int64, jetVersion uint32) string {
	switch {
	case pageSize == mdb.PageSizeJet3:
		return mdb.LayoutJet32K
	case jetVersion == mdb.JetVersion3 || jetVersion == 259:
		return mdb.LayoutLegacy4
	case pageSize == mdb.PageSizeJet4:
		return mdb.LayoutJet44K
	default:
		return mdb.LayoutUnknown
	}
}

func layoutHint(class string) string {
	switch class {
	case mdb.LayoutJet32K:
		return "Jet 3.x 2K layout detected; parser support is improving and some edge cases may still fail."
	case mdb.LayoutLegacy4:
		return "Legacy/non-standard 4K header detected; layout probing is used instead of header version alone."
	case mdb.LayoutUnknown:
		return "Unknown layout; schema/VBA extraction may fail on unsupported or corrupted files."
	default:
		return ""
	}
}

func commandErrorHint(err error) string {
	if err == nil {
		return ""
	}

	lower := strings.ToLower(err.Error())

	switch {
	case strings.Contains(lower, "encrypted"), strings.Contains(lower, "password"):
		return "encrypted/password-protected databases are not currently supported"
	case strings.Contains(lower, "invalid magic bytes"):
		return "file is not a valid Access database header"
	case strings.Contains(lower, "file too small"),
		strings.Contains(lower, "out of range"),
		strings.Contains(lower, "too short"),
		strings.Contains(lower, "invalid bounds"),
		strings.Contains(lower, "truncated"):
		return "file appears truncated or structurally corrupt"
	case strings.Contains(lower, "unsupported"), strings.Contains(lower, "not implemented"):
		return "database layout/path is not supported yet"
	default:
		return ""
	}
}

func formatCommandError(path string, err error) string {
	if err == nil {
		return ""
	}

	details := []string{err.Error()}

	if hint := commandErrorHint(err); hint != "" {
		details = append(details, "hint: "+hint)
	}

	// Add layout context when possible to make troubleshooting actionable.
	details = append(details, probeLayoutDetails(path)...)

	if len(details) == 1 {
		return details[0]
	}

	return fmt.Sprintf("%s (%s)", details[0], strings.Join(details[1:], "; "))
}

func probeLayoutDetails(path string) []string {
	if path == "" {
		return nil
	}

	_, statErr := os.Stat(path)
	if statErr != nil {
		return nil
	}

	p, probeErr := mdb.Probe(path)
	if probeErr != nil {
		return nil
	}

	details := []string{fmt.Sprintf("layout=%s pageSize=%d jetVersion=%d", p.LayoutClass, p.PageSize, p.JetVersion)}
	if !p.MSysObjectsReadable && p.MSysObjectsError != "" {
		details = append(details, "catalog="+p.MSysObjectsError)
	}

	return details
}

func extractionStats(modules []vba.ExtractedModule) (partialModules int, warningCount int) {
	for _, m := range modules {
		if m.Partial {
			partialModules++
		}

		warningCount += len(m.Warnings)
	}

	return partialModules, warningCount
}
