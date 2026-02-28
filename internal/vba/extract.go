package vba

import (
	"bytes"
	"fmt"
	"log/slog"
	"strings"
	"unicode/utf8"

	"golang.org/x/text/encoding/charmap"
)

// ExtractedModule is the final extraction result for a single module.
type ExtractedModule struct {
	Name      string
	Type      ProjectModuleType
	Text      string
	Partial   bool
	Stream    string
	StorageID int32
	Warnings  []string
}

// ExtractAllModules extracts VBA source text from all modules in a storage tree.
// It parses the PROJECT stream to discover module names, the dir stream for
// stream offsets, and decompresses each module's MS-OVBA compressed source.
// Debug messages (non-standard decompression strategies, dir parse failures)
// are written to log. Pass slog.New(slog.DiscardHandler) to suppress all output.
func ExtractAllModules(st *StorageTree, log *slog.Logger) ([]ExtractedModule, error) {
	required, err := st.RequiredStreams()
	if err != nil {
		return nil, err
	}

	projectNode := required["PROJECT"]
	if projectNode == nil || len(projectNode.Data) == 0 {
		return nil, fmt.Errorf("vba: PROJECT stream missing or empty")
	}

	projectInfo, err := ParseProjectStream(projectNode.Data)
	if err != nil {
		return nil, fmt.Errorf("vba: parse PROJECT stream: %w", err)
	}

	var dirInfo *DirInfo
	dirNode := required["dir"]
	if dirNode != nil && len(dirNode.Data) > 0 {
		dirInfo, err = ParseDirStream(dirNode.Data, func(in []byte) ([]byte, error) {
			out, _, derr := DecompressContainerWithFallback(in, log)
			return out, derr
		})
		if err != nil {
			log.Debug("vba: dir parsing failed; fallback mapping will be used", "err", err)
		}
	}

	moduleStreams, err := st.ModuleStreams()
	if err != nil {
		return nil, err
	}

	mappings, mappingWarns := BuildModuleMappings(projectInfo, dirInfo, moduleStreams)

	results := make([]ExtractedModule, 0, len(mappings))
	for _, mapping := range mappings {
		result := ExtractedModule{
			Name:      mapping.ModuleName,
			Type:      mapping.ModuleType,
			Stream:    mapping.StreamName,
			StorageID: mapping.StorageID,
		}

		for _, w := range mappingWarns {
			if strings.Contains(strings.ToLower(w), strings.ToLower(mapping.ModuleName)) {
				result.Warnings = append(result.Warnings, w)
			}
		}

		text, warns, partial := extractModuleSource(mapping, log)
		result.Text = text
		result.Partial = partial
		result.Warnings = append(result.Warnings, warns...)

		results = append(results, result)
	}

	return results, nil
}

func extractModuleSource(mapping ModuleMapping, log *slog.Logger) (string, []string, bool) {
	if len(mapping.Data) == 0 {
		return "", []string{"empty module stream data"}, true
	}

	var warnings []string

	// Module stream layout (MS-OVBA §2.3.4.3):
	//   bytes [0 .. SourceOffset-1] = compiled p-code (binary, not OVBA)
	//   bytes [SourceOffset ..]     = OVBA compressed VBA source
	// SourceOffset is in the RAW bytes — do NOT decompress the whole stream first.
	offset := int(mapping.SourceOffset)
	if offset > len(mapping.Data) {
		warnings = append(warnings, fmt.Sprintf("sourceOffset %d exceeds stream length %d; trying brute-force scan", offset, len(mapping.Data)))
		if text, ok := bruteForceOffsetScan(mapping.Data); ok {
			return cleanupVBA(text), warnings, false
		}
		return recoverPartialFromRaw(mapping.Data, warnings)
	}

	sourceCandidate := mapping.Data[offset:]
	finalRaw, srcStrategy, srcErr := DecompressContainerWithFallback(sourceCandidate, log)
	if srcErr != nil {
		warnings = append(warnings, fmt.Sprintf("source decompression failed at offset %d: %v; trying brute-force scan", offset, srcErr))
		if text, ok := bruteForceOffsetScan(mapping.Data); ok {
			return cleanupVBA(text), warnings, false
		}
		return recoverPartialFromRaw(mapping.Data, warnings)
	}
	if srcStrategy != StrategyStandard {
		warnings = append(warnings, fmt.Sprintf("source decompression used strategy %s", srcStrategy))
	}

	decoded := decodeBestText(finalRaw)
	decoded = cleanupVBA(decoded)
	if !looksLikeVBASource(decoded) {
		warnings = append(warnings, "decoded source does not contain strong VBA markers")
		if text, ok := bruteForceOffsetScan(mapping.Data); ok {
			return cleanupVBA(text), warnings, false
		}
		if partial, pOK := recoverPartialText(finalRaw); pOK {
			return partial, warnings, true
		}
	}

	return decoded, warnings, false
}

func bruteForceOffsetScan(streamDecompressed []byte) (string, bool) {
	bestScore := -1
	best := ""

	for i := 0; i < len(streamDecompressed); i++ {
		if streamDecompressed[i] != compressedContainerSig {
			continue
		}

		candidateRaw, err := DecompressContainer(streamDecompressed[i:])
		if err != nil {
			continue
		}

		candidate := cleanupVBA(decodeBestText(candidateRaw))
		score := scoreVBAText(candidate)
		if score > bestScore {
			bestScore = score
			best = candidate
		}
	}

	if bestScore < 2 || strings.TrimSpace(best) == "" {
		return "", false
	}

	return best, true
}

func recoverPartialFromRaw(raw []byte, warnings []string) (string, []string, bool) {
	partial, ok := recoverPartialText(raw)
	if !ok {
		return "", append(warnings, "no recoverable VBA fragments found"), true
	}
	return partial, warnings, true
}

func recoverPartialText(raw []byte) (string, bool) {
	keywords := []string{"Attribute VB_Name", "Option Compare", "Option Explicit", "Sub ", "Function ", "MsgBox", "DoCmd"}

	decoded := decodeBestText(raw)
	lines := strings.Split(decoded, "\n")
	var kept []string

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		for _, keyword := range keywords {
			if strings.Contains(trimmed, keyword) {
				kept = append(kept, trimmed)
				break
			}
		}
	}

	if len(kept) == 0 {
		return "", false
	}

	text := "[PARTIAL - reconstructed from p-code tokens]\n" + strings.Join(kept, "\n") + "\n"
	return text, true
}

func decodeBestText(raw []byte) string {
	if len(raw) == 0 {
		return ""
	}

	candidates := []string{}
	if utf8.Valid(raw) {
		candidates = append(candidates, string(raw))
	}

	if decoded, err := charmap.Windows1252.NewDecoder().Bytes(raw); err == nil {
		candidates = append(candidates, string(decoded))
	}
	if decoded, err := charmap.ISO8859_1.NewDecoder().Bytes(raw); err == nil {
		candidates = append(candidates, string(decoded))
	}

	best := ""
	bestScore := -1
	for _, candidate := range candidates {
		score := scoreVBAText(candidate)
		if score > bestScore {
			best = candidate
			bestScore = score
		}
	}

	if best == "" {
		return string(raw)
	}

	return best
}

func scoreVBAText(text string) int {
	t := strings.ToLower(text)
	score := 0

	if strings.Contains(t, "attribute vb_name") {
		score += 5
	}
	if strings.Contains(t, "option explicit") {
		score += 3
	}
	if strings.Contains(t, "sub ") {
		score += 2
	}
	if strings.Contains(t, "function ") {
		score += 2
	}
	if strings.Contains(t, "docmd") || strings.Contains(t, "msgbox") {
		score += 1
	}

	printable := 0
	for _, r := range text {
		if r == '\n' || r == '\r' || r == '\t' || (r >= 32 && r <= 126) || (r >= 160 && r <= 255) {
			printable++
		}
	}
	if len(text) > 0 {
		ratio := float64(printable) / float64(len(text))
		if ratio > 0.95 {
			score += 2
		}
	}

	return score
}

func cleanupVBA(text string) string {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	text = strings.TrimRight(text, "\x00")
	text = strings.TrimSpace(text)
	if text != "" {
		text += "\n"
	}
	return text
}

func looksLikeVBASource(text string) bool {
	t := strings.ToLower(text)
	return strings.Contains(t, "attribute vb_name") || strings.Contains(t, "option explicit") || strings.Contains(t, "sub ") || strings.Contains(t, "function ")
}

// ExtractModuleMap is a convenience wrapper around ExtractAllModules that returns
// the results as a map keyed by module name.
func ExtractModuleMap(st *StorageTree, log *slog.Logger) (map[string]ExtractedModule, error) {
	modules, err := ExtractAllModules(st, log)
	if err != nil {
		return nil, err
	}

	out := make(map[string]ExtractedModule, len(modules))
	for _, module := range modules {
		if strings.TrimSpace(module.Name) == "" {
			continue
		}
		out[module.Name] = module
	}

	return out, nil
}

func joinWarnings(a []string, b ...string) []string {
	if len(b) == 0 {
		return a
	}
	out := make([]string, 0, len(a)+len(b))
	out = append(out, a...)
	out = append(out, b...)
	return out
}

func containsAny(text string, needles []string) bool {
	for _, n := range needles {
		if strings.Contains(text, n) {
			return true
		}
	}
	return false
}

func splitNonEmptyLines(s string) []string {
	parts := strings.Split(s, "\n")
	out := make([]string, 0, len(parts))
	for _, line := range parts {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		out = append(out, line)
	}
	return out
}

func compactWhitespace(s string) string {
	return strings.Join(splitNonEmptyLines(strings.ReplaceAll(s, "\r", "\n")), "\n")
}

func normalizeForMatch(s string) string {
	s = cleanupVBA(s)
	s = compactWhitespace(s)
	return strings.TrimSpace(s)
}

func similarVBA(a, b string) bool {
	na := normalizeForMatch(a)
	nb := normalizeForMatch(b)
	if na == "" || nb == "" {
		return false
	}
	if na == nb {
		return true
	}
	return bytes.Contains([]byte(na), []byte(nb)) || bytes.Contains([]byte(nb), []byte(na))
}
