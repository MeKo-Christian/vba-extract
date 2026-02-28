package vba

import (
	"fmt"
	"sort"
	"strings"
)

type ForensicKind string

const (
	ForensicProjectText   ForensicKind = "project-text"
	ForensicDirRecords    ForensicKind = "dir-records"
	ForensicVBASourceText ForensicKind = "vba-source-text"
	ForensicCompressedVBA ForensicKind = "compressed-vba"
	ForensicAccessArtifact ForensicKind = "access-artifact"
)

type ForensicHit struct {
	NodeID   int32
	NodeName string
	NodeType int32
	DataSize int
	Kind     ForensicKind
	Score    int
	Summary  string
}

type ForensicReport struct {
	Hits               []ForensicHit
	ProjectCandidates  int
	DirCandidates      int
	SourceCandidates   int
	CompressedCandidates int
	ArtifactCandidates int
}

func ForensicScanStorage(st *StorageTree) ForensicReport {
	report := ForensicReport{}

	for _, node := range st.Nodes {
		if node == nil || len(node.Data) == 0 {
			continue
		}

		data := node.Data

		if project, err := ParseProjectStream(data); err == nil && project != nil && len(project.Modules) > 0 {
			report.ProjectCandidates++
			report.Hits = append(report.Hits, ForensicHit{
				NodeID:   node.ID,
				NodeName: node.Name,
				NodeType: node.Type,
				DataSize: len(data),
				Kind:     ForensicProjectText,
				Score:    100 + len(project.Modules),
				Summary:  fmt.Sprintf("PROJECT-like stream with %d module declarations", len(project.Modules)),
			})
		}

		if dir, err := ParseDirStream(data, func(in []byte) ([]byte, error) {
			out, _, derr := DecompressContainerWithFallback(in, false, nil)
			return out, derr
		}); err == nil && dir != nil && len(dir.Modules) > 0 {
			report.DirCandidates++
			report.Hits = append(report.Hits, ForensicHit{
				NodeID:   node.ID,
				NodeName: node.Name,
				NodeType: node.Type,
				DataSize: len(data),
				Kind:     ForensicDirRecords,
				Score:    90 + len(dir.Modules),
				Summary:  fmt.Sprintf("dir-like stream with %d module records", len(dir.Modules)),
			})
		}

		decoded := cleanupVBA(decodeBestText(data))
		if looksLikeVBASource(decoded) {
			report.SourceCandidates++
			report.Hits = append(report.Hits, ForensicHit{
				NodeID:   node.ID,
				NodeName: node.Name,
				NodeType: node.Type,
				DataSize: len(data),
				Kind:     ForensicVBASourceText,
				Score:    70 + scoreVBAText(decoded),
				Summary:  summarizeText(decoded),
			})
		}

		if len(data) > 0 && data[0] == compressedContainerSig {
			if dec, _, err := DecompressContainerWithFallback(data, false, nil); err == nil {
				text := cleanupVBA(decodeBestText(dec))
				if scoreVBAText(text) >= 3 {
					report.CompressedCandidates++
					report.Hits = append(report.Hits, ForensicHit{
						NodeID:   node.ID,
						NodeName: node.Name,
						NodeType: node.Type,
						DataSize: len(data),
						Kind:     ForensicCompressedVBA,
						Score:    60 + scoreVBAText(text),
						Summary:  summarizeText(text),
					})
				}
			}
		}

		if isLikelyAccessArtifactNode(node.Name) {
			artifactText := decodeBestText(data)
			score := scoreAccessArtifactText(artifactText)
			if score >= 5 {
				report.ArtifactCandidates++
				report.Hits = append(report.Hits, ForensicHit{
					NodeID:   node.ID,
					NodeName: node.Name,
					NodeType: node.Type,
					DataSize: len(data),
					Kind:     ForensicAccessArtifact,
					Score:    40 + score,
					Summary:  summarizeText(cleanupVBA(artifactText)),
				})
			}
		}
	}

	sort.Slice(report.Hits, func(i, j int) bool {
		if report.Hits[i].Score == report.Hits[j].Score {
			if report.Hits[i].NodeID == report.Hits[j].NodeID {
				return report.Hits[i].Kind < report.Hits[j].Kind
			}
			return report.Hits[i].NodeID < report.Hits[j].NodeID
		}
		return report.Hits[i].Score > report.Hits[j].Score
	})

	return report
}

func summarizeText(text string) string {
	text = strings.TrimSpace(text)
	if text == "" {
		return ""
	}
	line := text
	if idx := strings.Index(line, "\n"); idx >= 0 {
		line = line[:idx]
	}
	if len(line) > 100 {
		line = line[:100] + "..."
	}
	return line
}

func isLikelyAccessArtifactNode(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "BLOB" || upper == "PROPDATA" || upper == "DIRDATA" || upper == "TYPEINFO"
}

func scoreAccessArtifactText(text string) int {
	t := strings.ToLower(text)
	score := 0

	markers := []string{
		"form_", "report_", "caption", "recordsource", "controlsource", "onclick", "onopen",
		"sourceobject", "datasheet", "row source", "select ", "where ", "insert into", "update ",
	}
	for _, marker := range markers {
		if strings.Contains(t, marker) {
			score += 2
		}
	}

	printable := 0
	for _, r := range text {
		if r == '\n' || r == '\r' || r == '\t' || (r >= 32 && r <= 126) || (r >= 160 && r <= 255) {
			printable++
		}
	}
	if len(text) > 0 {
		ratio := float64(printable) / float64(len(text))
		if ratio > 0.90 {
			score += 2
		}
	}

	return score
}
