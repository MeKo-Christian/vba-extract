package vba

import (
	"bytes"
	"encoding/binary"
	"errors"
	"testing"
)

func buildDirRecord(id uint16, payload []byte) []byte {
	var b bytes.Buffer
	_ = binary.Write(&b, binary.LittleEndian, id)
	_ = binary.Write(&b, binary.LittleEndian, uint32(len(payload)))
	b.Write(payload)
	return b.Bytes()
}

func TestParseProjectStream(t *testing.T) {
	input := []byte("ID=\r\nDocument=ThisDocument/&H00000000\r\nName=SampleProject\r\nModule=Mod1\r\nClass=Class1\r\nDocClass=Form_Main/&H00000000\r\n[Host Extender Info]\r\n&H00000001={GUID}\r\n")

	info, err := ParseProjectStream(input)
	if err != nil {
		t.Fatalf("ParseProjectStream: %v", err)
	}

	if info.Name != "SampleProject" {
		t.Fatalf("project name = %q, want %q", info.Name, "SampleProject")
	}

	if len(info.Modules) != 3 {
		t.Fatalf("modules = %d, want 3", len(info.Modules))
	}

	if info.Modules[0].Name != "Mod1" || info.Modules[0].Type != ProjectModuleStandard {
		t.Fatalf("module[0] = %+v, want Mod1/standard", info.Modules[0])
	}
	if info.Modules[1].Name != "Class1" || info.Modules[1].Type != ProjectModuleClass {
		t.Fatalf("module[1] = %+v, want Class1/class", info.Modules[1])
	}
	if info.Modules[2].Name != "Form_Main" || info.Modules[2].Type != ProjectModuleDocument {
		t.Fatalf("module[2] = %+v, want Form_Main/document", info.Modules[2])
	}
}

func TestParseDirStreamNeedsDecompressor(t *testing.T) {
	_, err := ParseDirStream([]byte{0x01, 0x02, 0x03}, nil)
	if !errors.Is(err, ErrDirNeedsDecompression) {
		t.Fatalf("error = %v, want ErrDirNeedsDecompression", err)
	}
}

func TestParseDirStreamRecords(t *testing.T) {
	var data []byte
	data = append(data, buildDirRecord(0x0019, []byte("Mod1"))...)
	data = append(data, buildDirRecord(0x001A, []byte("STREAM_ABC"))...)
	data = append(data, buildDirRecord(0x0032, []byte("S\x00T\x00R\x00"))...)
	off := make([]byte, 4)
	binary.LittleEndian.PutUint32(off, 123)
	data = append(data, buildDirRecord(0x0031, off)...)
	data = append(data, buildDirRecord(0x0021, nil)...)
	data = append(data, buildDirRecord(0x002B, nil)...)

	info, err := ParseDirStream(data, nil)
	if err != nil {
		t.Fatalf("ParseDirStream: %v", err)
	}

	if len(info.Modules) != 1 {
		t.Fatalf("modules = %d, want 1", len(info.Modules))
	}

	m := info.Modules[0]
	if m.ModuleName != "Mod1" {
		t.Fatalf("ModuleName = %q, want %q", m.ModuleName, "Mod1")
	}
	if m.StreamName != "STREAM_ABC" {
		t.Fatalf("StreamName = %q, want %q", m.StreamName, "STREAM_ABC")
	}
	if m.SourceOff != 123 {
		t.Fatalf("SourceOff = %d, want 123", m.SourceOff)
	}
	if m.Type != DirModuleProcedural {
		t.Fatalf("Type = %q, want %q", m.Type, DirModuleProcedural)
	}
}

func TestParseDirStreamWithDecompressor(t *testing.T) {
	var records []byte
	records = append(records, buildDirRecord(0x0019, []byte("Mod1"))...)
	records = append(records, buildDirRecord(0x001A, []byte("STREAM_ABC"))...)
	off := make([]byte, 4)
	binary.LittleEndian.PutUint32(off, 7)
	records = append(records, buildDirRecord(0x0031, off)...)
	records = append(records, buildDirRecord(0x0021, nil)...)
	records = append(records, buildDirRecord(0x002B, nil)...)

	chunkSize := len(records) + 2
	header := uint16(0x3000 | uint16(chunkSize-3)) // uncompressed chunk

	container := []byte{0x01}
	h := make([]byte, 2)
	binary.LittleEndian.PutUint16(h, header)
	container = append(container, h...)
	container = append(container, records...)

	info, err := ParseDirStream(container, DecompressContainer)
	if err != nil {
		t.Fatalf("ParseDirStream: %v", err)
	}

	if len(info.Modules) != 1 {
		t.Fatalf("modules = %d, want 1", len(info.Modules))
	}
	if info.Modules[0].SourceOff != 7 {
		t.Fatalf("SourceOff = %d, want 7", info.Modules[0].SourceOff)
	}
}

func TestBuildModuleMappingsFallback(t *testing.T) {
	project := &ProjectInfo{
		Modules: []ProjectModule{
			{Name: "Mod1", Type: ProjectModuleStandard},
			{Name: "Mod2", Type: ProjectModuleClass},
		},
	}

	streams := []*StorageNode{
		{ID: 20, Name: "ZSTREAM"},
		{ID: 10, Name: "ASTREAM"},
		{ID: 30, Name: "XSTREAM"},
	}

	mappings, warns := BuildModuleMappings(project, nil, streams)
	if len(warns) == 0 {
		t.Fatal("expected fallback warning")
	}
	if len(mappings) != 2 {
		t.Fatalf("mappings = %d, want 2", len(mappings))
	}

	if mappings[0].ModuleName != "Mod1" || mappings[0].StreamName != "ASTREAM" {
		t.Fatalf("mapping[0] = %+v, want Mod1->ASTREAM", mappings[0])
	}
	if mappings[1].ModuleName != "Mod2" || mappings[1].StreamName != "ZSTREAM" {
		t.Fatalf("mapping[1] = %+v, want Mod2->ZSTREAM", mappings[1])
	}
}

func TestParseProjectFromStartMDB(t *testing.T) {
	db := testDB(t)

	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}

	projectNode := required["PROJECT"]
	if projectNode == nil || len(projectNode.Data) == 0 {
		t.Fatal("PROJECT stream missing or empty")
	}

	info, err := ParseProjectStream(projectNode.Data)
	if err != nil {
		t.Fatalf("ParseProjectStream: %v", err)
	}

	if len(info.Modules) != 15 {
		t.Errorf("modules = %d, want 15", len(info.Modules))
	}

	docClasses, standards := 0, 0
	for _, m := range info.Modules {
		switch m.Type {
		case ProjectModuleDocument:
			docClasses++
		case ProjectModuleStandard:
			standards++
		}
	}
	if docClasses != 6 {
		t.Errorf("DocClass modules = %d, want 6", docClasses)
	}
	if standards != 9 {
		t.Errorf("Standard modules = %d, want 9", standards)
	}
	if info.Name != "Start" {
		t.Errorf("project name = %q, want %q", info.Name, "Start")
	}
}

func TestBuildModuleMappingsStartMDB(t *testing.T) {
	db := testDB(t)
	st, err := LoadStorageTree(db)
	if err != nil {
		t.Fatalf("LoadStorageTree: %v", err)
	}

	required, err := st.RequiredStreams()
	if err != nil {
		t.Fatalf("RequiredStreams: %v", err)
	}
	moduleStreams, err := st.ModuleStreams()
	if err != nil {
		t.Fatalf("ModuleStreams: %v", err)
	}

	projectInfo, err := ParseProjectStream(required["PROJECT"].Data)
	if err != nil {
		t.Fatalf("ParseProjectStream: %v", err)
	}

	dirInfo, err := ParseDirStream(required["dir"].Data, DecompressContainer)
	if err != nil {
		t.Fatalf("ParseDirStream: %v", err)
	}

	mappings, warns := BuildModuleMappings(projectInfo, dirInfo, moduleStreams)
	for _, w := range warns {
		t.Logf("warn: %s", w)
	}

	if len(mappings) != 15 {
		t.Errorf("mappings = %d, want 15", len(mappings))
	}

	for _, m := range mappings {
		if m.StreamName == "" {
			t.Errorf("module %q has empty StreamName", m.ModuleName)
		}
		if m.SourceOffset == 0 {
			t.Logf("module %q has sourceOffset=0 (may be valid)", m.ModuleName)
		}
	}
}
