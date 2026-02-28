package vba

import (
	"encoding/binary"
	"fmt"
	"strings"
	"testing"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
)

func TestDiagAbteilungen(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/Abteilungen_Zuordnen.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	pageCount := int(db.PageCount())
	t.Logf("PageCount: %d", pageCount)

	// Count pages by type and find pages with VBA-like signatures
	typeCounts := map[byte]int{}
	vbaPages := 0
	for pg := 0; pg < pageCount; pg++ {
		page, err := db.ReadPage(int64(pg))
		if err != nil { continue }
		typeCounts[page[0]]++
		// Check if page contains "Attribute VB_Name" anywhere
		if strings.Contains(string(page), "Attribute VB_Name") ||
			strings.Contains(string(page), "Attribute VB_Base") {
			vbaPages++
			t.Logf("VBA text found on page %d (type=0x%02X)", pg, page[0])
		}
	}
	t.Logf("Page type distribution: %v", typeCounts)
	t.Logf("Pages with raw VBA text: %d", vbaPages)

	// Try to find any 0x01 bytes that look like a CompressedContainer
	containerCount := 0
	for pg := 0; pg < pageCount; pg++ {
		page, err := db.ReadPage(int64(pg))
		if err != nil || page[0] != 0x01 { continue }
		numRows := int(binary.LittleEndian.Uint16(page[lvalNumRows:]))
		for row := 0; row < numRows; row++ {
			rowOff := lvalRowTable + row*2
			if rowOff+2 > len(page) { break }
			offVal := binary.LittleEndian.Uint16(page[rowOff:])
			if offVal&0x4000 != 0 { continue }
			start := int(offVal & lvalRowMask)
			firstContent := start + 4
			if firstContent+3 > len(page) { continue }
			// scan for 0x01 with valid nibble
			for i := firstContent; i < len(page)-3; i++ {
				if page[i] != 0x01 { continue }
				nibble := page[i+2] >> 4
				if nibble == 0x3 || nibble == 0xB {
					containerCount++
					if containerCount <= 5 {
						t.Logf("Candidate container at pg=%d row=%d offset=%d nibble=0x%X", pg, row, i-start, nibble)
						// Try decompression
						chainData, err := db.ReadLvalChain(int64(pg), row, 256*1024)
						if err != nil { t.Logf("  chain read error: %v", err); continue }
						dec, decErr := DecompressContainer(chainData[i-start:])
						t.Logf("  chain=%d dec=%d err=%v", len(chainData), len(dec), decErr)
						if decErr == nil {
							t.Logf("  text preview: %q", fmt.Sprintf("%.100s", decodeBestText(dec)))
						}
					}
					break
				}
			}
		}
	}
	t.Logf("Total container candidates: %d", containerCount)
}
