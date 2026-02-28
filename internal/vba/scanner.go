package vba

import (
	"encoding/binary"
	"regexp"
	"strings"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
)

const (
	pageTypeData   = 0x01 // LVAL page type byte
	lvalRowTable   = 0x0E
	lvalNumRows    = 0x0C
	lvalRowMask    = 0x1FFF
	maxOrphanChain = 256 * 1024 // 256 KB: largest observed VBA module is ~154 KB compressed
)

var moduleNameRe = regexp.MustCompile(`(?i)Attribute VB_Name\s*=\s*"([^"]{1,80})"`)

// ScanOrphanedLvalModules scans all LVAL page chains in the database for VBA
// module streams that are not referenced by MSysAccessStorage. This is a
// forensic fallback for databases where the VBA structure has been stripped
// from the system tables but the underlying page data remains on disk.
//
// The scan is O(pages) for the first pass and then only reads chains that look
// like they could be module streams.
func ScanOrphanedLvalModules(db *mdb.Database) ([]ExtractedModule, error) {
	pageCount := int(db.PageCount())

	// First pass: collect all next-pointer targets so we can identify chain starts.
	targets := make(map[uint64]struct{}, 8192)

	for pg := range pageCount {
		page, err := db.ReadPage(int64(pg))
		if err != nil || page[0] != pageTypeData {
			continue
		}

		numRows := int(binary.LittleEndian.Uint16(page[lvalNumRows:]))
		for row := range numRows {
			rowOff := lvalRowTable + row*2
			if rowOff+2 > len(page) {
				break
			}

			offVal := binary.LittleEndian.Uint16(page[rowOff:])
			if offVal&0x4000 != 0 { // deleted flag
				continue
			}

			start := int(offVal & lvalRowMask)
			if start+4 > len(page) {
				continue
			}

			nextPtr := binary.LittleEndian.Uint32(page[start:])
			if nextPtr != 0 {
				targets[uint64(nextPtr)] = struct{}{}
			}
		}
	}

	// Second pass: for each LVAL page row that is a chain start, try to recover VBA.
	seen := make(map[string]struct{})
	var results []ExtractedModule

	for pg := range pageCount {
		page, err := db.ReadPage(int64(pg))
		if err != nil || page[0] != pageTypeData {
			continue
		}

		numRows := int(binary.LittleEndian.Uint16(page[lvalNumRows:]))
		for row := range numRows {
			key := uint64(pg)<<8 | uint64(row)
			if _, isTarget := targets[key]; isTarget {
				continue // not a chain start
			}

			rowOff := lvalRowTable + row*2
			if rowOff+2 > len(page) {
				break
			}

			offVal := binary.LittleEndian.Uint16(page[rowOff:])
			if offVal&0x4000 != 0 {
				continue
			}

			start := int(offVal & lvalRowMask)
			if start+4 > len(page) {
				continue
			}
			// Peek at the first content byte (after the 4-byte next-ptr).
			firstContent := start + 4
			if firstContent >= len(page) {
				continue
			}
			// Filter to chains that are likely module streams.
			// VBA module streams start with p-code binary followed by a
			// CompressedContainer beginning with 0x01. Two heuristics:
			//   a) first content byte is 0x01 AND the next 2 bytes form a valid
			//      CompressedChunkHeader (bits 14-12 == 0x3) -- rejects ~93% of
			//      false-positive 0x01 bytes with a single cheap check.
			//   b) row size >= 1000 bytes, covering p-code-prefixed streams where
			//      the container signature is not at content offset 0.
			rowEnd := len(page)
			if row > 0 {
				prevOff := binary.LittleEndian.Uint16(page[lvalRowTable+(row-1)*2:])
				rowEnd = int(prevOff & lvalRowMask)
			}

			rowSize := rowEnd - start
			firstByte := page[firstContent]
			// Keep chains where the first content byte is the container signature,
			// or where the row is large enough to contain p-code + container.
			// Note: firstByte==0x01 may be a coincidental byte (not the container
			// start); tryExtractModuleFromChain will scan for the real container and
			// validate each candidate with the nibble check internally.
			if firstByte != compressedContainerSig && rowSize < 1000 {
				continue
			}

			chainData, chainErr := db.ReadLvalChain(int64(pg), row, maxOrphanChain)
			if chainErr != nil || len(chainData) < 100 {
				continue
			}

			module := tryExtractModuleFromChain(chainData)
			if module == nil {
				continue
			}

			if _, dup := seen[module.Name]; dup {
				continue
			}

			seen[module.Name] = struct{}{}
			results = append(results, *module)
		}
	}

	return results, nil
}

// tryExtractModuleFromChain tries to decompress and extract a VBA module from
// raw LVAL chain data. It scans all positions looking for a CompressedContainer
// that decompresses to valid VBA source.
func tryExtractModuleFromChain(data []byte) *ExtractedModule {
	for i := range len(data) - 4 {
		if data[i] != compressedContainerSig {
			continue
		}
		// Quick-validate the CompressedChunkHeader (bytes i+1..i+2).
		// Bits 14-12 of the header must be 0x3; anything else is not a valid
		// CompressedContainer start. This rejects ~93% of 0x01 positions
		// without any decompression work.
		if i+3 > len(data) {
			break
		}

		nibble := data[i+2] >> 4
		if nibble != 0x3 && nibble != 0xB {
			continue
		}
		// We are already positioned at a 0x01 sig byte, so call
		// DecompressContainer directly instead of DecompressContainerWithFallback
		// (which would redundantly re-scan the whole slice for another 0x01).
		dec, err := DecompressContainer(data[i:])
		if err != nil || len(dec) < 40 {
			continue
		}

		text := cleanupVBA(decodeBestText(dec))
		if !strings.Contains(strings.ToLower(text), "attribute vb_name") {
			continue
		}

		name := extractModuleNameFromText(text)
		if name == "" {
			continue
		}

		moduleType := ProjectModuleStandard
		if strings.Contains(strings.ToLower(text), "attribute vb_base") {
			moduleType = ProjectModuleDocument
		}

		return &ExtractedModule{
			Name:   name,
			Type:   moduleType,
			Text:   text,
			Stream: "orphaned-lval",
		}
	}

	return nil
}

func extractModuleNameFromText(text string) string {
	m := moduleNameRe.FindStringSubmatch(text)
	if m == nil {
		return ""
	}

	name := strings.TrimSpace(m[1])
	// Remove null bytes that can appear in p-code embedded text
	name = strings.ReplaceAll(name, "\x00", "")

	return name
}
