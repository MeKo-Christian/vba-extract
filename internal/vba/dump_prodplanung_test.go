package vba

import (
	"encoding/hex"
	"fmt"
	"strings"
	"testing"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
)

func TestDumpProdPlanungStorage(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	t.Logf("Total nodes: %d", len(st.Nodes))
	for _, node := range st.Nodes {
		errStr := ""
		if node.ResolveErr != nil {
			errStr = " ERR:" + node.ResolveErr.Error()
		}
		// Show first 40 chars of data
		dataPreview := ""
		if len(node.Data) > 0 {
			preview := string(node.Data)
			if len(preview) > 80 {
				preview = preview[:80]
			}
			dataPreview = fmt.Sprintf(" data=%q", preview)
		}
		t.Logf("  id=%-5d parent=%-5d type=%-3d name=%-40q dataLen=%d%s%s",
			node.ID, node.ParentID, node.Type, node.Name, len(node.Data), errStr, dataPreview)
	}

	// Check PROJECT stream specifically
	required, err := st.RequiredStreams()
	t.Logf("RequiredStreams err: %v", err)
	for k, v := range required {
		if v != nil {
			t.Logf("  stream %q: id=%d dataLen=%d", k, v.ID, len(v.Data))
		} else {
			t.Logf("  stream %q: nil", k)
		}
	}
}

func TestDumpProdPlanungVBANode(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	// Check what tables exist
	catalog, err := db.Catalog()
	if err != nil {
		t.Fatalf("Catalog: %v", err)
	}
	t.Logf("Catalog entries (first 30):")
	for i, e := range catalog {
		if i >= 30 {
			t.Logf("  ... and %d more", len(catalog)-30)
			break
		}
		t.Logf("  type=%d id=%-8d name=%q", e.Type, e.ID, e.Name)
	}

	// Check VBAProject node children
	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	vbaNode, _ := st.VBAProjectNode()
	if vbaNode != nil {
		t.Logf("VBAProject node: id=%d parent=%d dataLen=%d", vbaNode.ID, vbaNode.ParentID, len(vbaNode.Data))
		t.Logf("  Children of VBAProject (id=%d):", vbaNode.ID)
		for _, child := range st.Children[vbaNode.ID] {
			t.Logf("    id=%-5d type=%-3d name=%-40q dataLen=%d", child.ID, child.Type, child.Name, len(child.Data))
			for _, grandchild := range st.Children[child.ID] {
				preview := ""
				if len(grandchild.Data) > 0 {
					p := string(grandchild.Data)
					if len(p) > 60 { p = p[:60] }
					preview = fmt.Sprintf(" data=%q", p)
				}
				t.Logf("      id=%-5d type=%-3d name=%-40q dataLen=%d%s", grandchild.ID, grandchild.Type, grandchild.Name, len(grandchild.Data), preview)
			}
		}
	}
}

func TestDumpProdPlanungFindProject(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	// Check all nodes for PROJECT-like content
	t.Logf("Scanning all nodes for PROJECT content:")
	for _, node := range st.Nodes {
		if node == nil || len(node.Data) < 10 {
			continue
		}
		// Check for Module= or DocClass= anywhere
		data := string(node.Data)
		if strings.Contains(data, "Module=") || strings.Contains(data, "DocClass=") {
			preview := data
			if len(preview) > 200 { preview = preview[:200] }
			t.Logf("  FOUND! id=%-5d parent=%-5d type=%-3d name=%-30q dataLen=%d data=%q",
				node.ID, node.ParentID, node.Type, node.Name, len(node.Data), preview)
		}
		// Check for OLE compound document header
		if len(node.Data) >= 8 && node.Data[0] == 0xD0 && node.Data[1] == 0xCF {
			t.Logf("  OLE compound doc! id=%-5d name=%-30q dataLen=%d", node.ID, node.Name, len(node.Data))
		}
		// Check for MS-OVBA compressed container 
		if node.Data[0] == 0x01 {
			t.Logf("  OVBA? id=%-5d name=%-30q dataLen=%d", node.ID, node.Name, len(node.Data))
		}
	}
}

func TestDumpProdPlanungMSysObjects(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	// Read MSysObjects table
	td, err := db.FindTable("MSysObjects")
	if err != nil {
		t.Fatalf("FindTable MSysObjects: %v", err)
	}
	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}

	t.Logf("MSysObjects rows: %d", len(rows))
	for _, row := range rows {
		lv, hasLv := row["Lv"].([]byte)
		if !hasLv || len(lv) < 10 {
			continue
		}
		name, _ := row["Name"].(string)
		typ, _ := row["Type"].(int32)
		
		// Resolve the Lv memo
		resolved, err := db.ResolveMemo(lv)
		if err != nil {
			t.Logf("  row name=%-40q type=%d lvLen=%d resolveerr=%v", name, typ, len(lv), err)
			continue
		}
		
		// Check if it contains VBA-related content
		data := string(resolved)
		if strings.Contains(data, "Module=") || strings.Contains(data, "DocClass=") || strings.Contains(data, "Attribute VB_") {
			preview := data
			if len(preview) > 300 { preview = preview[:300] }
			t.Logf("  VBA! name=%-40q type=%d resolved=%d preview=%q", name, typ, len(resolved), preview)
		}
		if resolved[0] == 0xD0 && resolved[1] == 0xCF {
			t.Logf("  OLE compound doc! name=%-40q type=%d resolved=%d", name, typ, len(resolved))
		}
		if resolved[0] == 0x01 {
			dec, _, decErr := DecompressContainerWithFallback(resolved, false, nil)
			if decErr == nil && len(dec) > 0 {
				text := string(dec)
				if strings.Contains(text, "Module=") || strings.Contains(text, "Attribute VB_") {
					t.Logf("  OVBA VBA! name=%-40q type=%d rawLen=%d decLen=%d", name, typ, len(resolved), len(dec))
					if len(text) > 200 { text = text[:200] }
					t.Logf("    content=%q", text)
				}
			}
		}
	}
}

func TestDumpProdPlanungMSysDb(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	// Try MSysDb
	for _, tableName := range []string{"MSysDb", "MSysAccessObjects"} {
		td, err := db.FindTable(tableName)
		if err != nil {
			t.Logf("Table %q not found: %v", tableName, err)
			continue
		}
		rows, err := td.ReadRows()
		if err != nil {
			t.Logf("ReadRows %q: %v", tableName, err)
			continue
		}
		t.Logf("Table %q: %d rows", tableName, len(rows))
		for i, row := range rows {
			if i > 5 { t.Logf("  ... truncated"); break }
			t.Logf("  row %d: keys=%v", i, rowKeys(row))
			for k, v := range row {
				switch vv := v.(type) {
				case []byte:
					if len(vv) > 0 {
						t.Logf("    %q: []byte len=%d", k, len(vv))
						// Try to resolve
						resolved, err := db.ResolveMemo(vv)
						if err == nil && len(resolved) > 4 {
							t.Logf("    resolved len=%d first4=%x", len(resolved), resolved[:4])
							if resolved[0] == 0xD0 && resolved[1] == 0xCF {
								t.Logf("    *** OLE COMPOUND DOC! ***")
							}
							if strings.Contains(string(resolved), "Module=") {
								t.Logf("    *** CONTAINS Module= ***")
							}
						}
					}
				}
			}
		}
	}
}

func rowKeys(row map[string]interface{}) []string {
	keys := make([]string, 0, len(row))
	for k := range row {
		keys = append(keys, k)
	}
	return keys
}

func TestDumpProdPlanungLvRaw(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	t.Logf("JetVersion: %d, IsJet4: %v", db.Header.JetVersion, db.IsJet4())
	t.Logf("PageCount: %d", db.PageCount())

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	// Show LvRaw for nodes related to VBA 
	for _, node := range st.Nodes {
		if node.LvRaw != nil || node.ResolveErr != nil {
			t.Logf("  id=%-5d name=%-30q lvRawLen=%d dataLen=%d err=%v lvRaw=%x",
				node.ID, node.Name, len(node.LvRaw), len(node.Data), node.ResolveErr, node.LvRaw)
		}
	}
}

func TestDumpProdPlanungRawTableRows(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	td, err := db.FindTable("MSysAccessStorage")
	if err != nil {
		t.Fatalf("FindTable: %v", err)
	}
	rows, err := td.ReadRows()
	if err != nil {
		t.Fatalf("ReadRows: %v", err)
	}
	
	t.Logf("MSysAccessStorage raw row count: %d", len(rows))
	
	// Show all IDs and ParentIDs
	for _, row := range rows {
		id := row["Id"]
		parentId := row["ParentId"]
		name := row["Name"]
		typ := row["Type"]
		lv := row["Lv"]
		t.Logf("  id=%-10v parent=%-10v type=%-5v name=%-30v hasLv=%v",
			id, parentId, typ, name, lv != nil)
	}
}

func TestDumpProdPlanungBlob166(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	t.Logf("PageCount: %d", db.PageCount())

	// Manually resolve Blob 166 LvRaw = 37fc00000041090017bf8400
	lvRaw, _ := hex.DecodeString("37fc00000041090017bf8400")
	t.Logf("LvRaw: %x (len=%d)", lvRaw, len(lvRaw))

	resolved, err := db.ResolveMemo(lvRaw)
	t.Logf("Resolve result: len=%d err=%v", len(resolved), err)
	if len(resolved) > 20 {
		t.Logf("First 20 bytes: %x", resolved[:20])
		t.Logf("Last 20 bytes: %x", resolved[len(resolved)-20:])
	}
}

// TestDumpProdPlanungScanAllModules scans all LVAL chain starts looking for recoverable VBA.
// This test checks page groups identified by raw scan as containing module stream data.
func TestDumpProdPlanungScanAllModules(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	// Chain starts identified by reverse-index trace from VBA-containing LVAL pages.
	// Sizes and page numbers confirmed by Python scan.
	candidates := []struct {
		page     int32
		totalLen int32
	}{
		{1132, 94376},  // large module: compressed at offset ~69927
		{5702, 154403}, // Zeitkalkulation: compressed at offset ~129954
		{27913, 153992}, // large module: compressed at offset ~129543
		// All chain starts from first-chunk pattern scan (38 candidates):
		{48, 1508},
		{59, 4543},
		{157, 3201},
		{218, 4099},
		{245, 2047},
		{247, 3429},
		{278, 3073},
		{282, 4201},
		{301, 2387},
		{987, 3809},
		{999, 1641},
		{1055, 4200},
		{1074, 1933},
		{1171, 4343},
		{1175, 4927},
		{1204, 2807},
		{1985, 1440},
		{27574, 4120},
		{27576, 4111},
		{27578, 4110},
		{27580, 4121},
		{27582, 4118},
		{27587, 4088},
		{27590, 4364},
		{27711, 4551},
		{27762, 4095},
		{27764, 4137},
		{27829, 4083},
		{27831, 3985},
		{27832, 4131},
		{27835, 4139},
		{27978, 4138},
		{27991, 4067},
		{27993, 4299},
		{27997, 3982},
		{28272, 4095},
		{28274, 4092},
	}

	recovered := 0
	for _, c := range candidates {
		tl := c.totalLen
		lvalPage := c.page << 8
		lvRaw := []byte{
			byte(tl), byte(tl >> 8), byte(tl >> 16), 0x00,
			byte(lvalPage), byte(lvalPage >> 8), byte(lvalPage >> 16), byte(lvalPage >> 24),
			0, 0, 0, 0,
		}
		resolved, err := db.ResolveMemo(lvRaw)
		if err != nil || len(resolved) < 100 {
			t.Logf("pg=%d: chain len=%d err=%v", c.page, len(resolved), err)
			continue
		}
		// Brute-force offset scan for CompressedContainer (up to full chain length)
		found := false
		for offset := 0; offset < len(resolved)-10; offset++ {
			if resolved[offset] != compressedContainerSig {
				continue
			}
			dec, _, decErr := DecompressContainerWithFallback(resolved[offset:], false, nil)
			if decErr != nil || len(dec) < 50 {
				continue
			}
			text := cleanupVBA(decodeBestText(dec))
			if !strings.Contains(text, "Attribute VB_Name") {
				continue
			}
			lines := strings.Count(text, "\n")
			preview := text
			if len(preview) > 80 {
				preview = preview[:80]
			}
			t.Logf("FOUND pg=%d chain=%d dec=%d lines=%d offset=%d: %q",
				c.page, len(resolved), len(dec), lines, offset, preview)
			recovered++
			found = true
			break
		}
		if !found {
			t.Logf("MISS  pg=%d chain=%d - no VBA source found", c.page, len(resolved))
		}
	}
	t.Logf("Recovered %d/%d modules", recovered, len(candidates))
}

// TestDumpProdPlanungOrphanedModules investigates orphaned LVAL chains
// that contain VBA module data not referenced by any MSysAccessStorage row.
// Page 5733 was found to contain "Attribute VB_Name = Zeitkalkulation" text.
func TestDumpProdPlanungOrphanedModules(t *testing.T) {
	path := "/mnt/projekte/Code/MeKo/IPOffice-VBA/mdb-files/ProdPlanung.mdb"
	db, err := mdb.Open(path)
	if err != nil {
		t.Skipf("open: %v", err)
	}
	defer db.Close()

	// Page 5733 row 0 is an orphaned LVAL chain containing VBA module stream.
	// LvRaw for page 5733 row 0: type=0x00 (multi-page), lvalPage = (5733<<8)|0 = 0x00166500
	// Python tracing showed this chain is 28171 bytes (0x006E0B)
	// bytes 0-2: 0B 6E 00 (totalLen 28171 LE 3-byte), byte 3: 0x00 (LvalMultiPage)
	// bytes 4-7: 00 65 16 00 (lvalPage = 0x00166500 = page 5733 row 0)
	lvRaw, _ := hex.DecodeString("0B6E0000" + "00651600" + "00000000")
	resolved, err := db.ResolveMemo(lvRaw)
	t.Logf("Page 5733 chain: len=%d err=%v", len(resolved), err)
	if len(resolved) < 100 {
		t.Logf("  data: %x", resolved)
		return
	}
	t.Logf("  First 40 bytes: %x", resolved[:40])

	// Try to decompress as OVBA CompressedContainer at various offsets
	t.Logf("  Trying brute-force decompression...")
	for offset := 0; offset < len(resolved) && offset < len(resolved)-10; offset++ {
		if resolved[offset] != compressedContainerSig {
			continue
		}
		dec, _, decErr := DecompressContainerWithFallback(resolved[offset:], false, nil)
		if decErr != nil {
			continue
		}
		text := cleanupVBA(decodeBestText(dec))
		if strings.Contains(text, "Attribute VB_Name") || strings.Contains(text, "Sub ") || strings.Contains(text, "Function ") {
			t.Logf("  SUCCESS at offset %d: decompressed %d bytes", offset, len(dec))
			if len(text) > 300 {
				text = text[:300]
			}
			t.Logf("  text=%q", text)
			break
		}
	}
}
