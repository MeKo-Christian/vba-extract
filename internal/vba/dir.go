package vba

import (
	"encoding/binary"
	"errors"
	"fmt"
	"sort"
	"strings"

	"golang.org/x/text/encoding/charmap"
)

var ErrDirNeedsDecompression = errors.New("vba: dir stream appears compressed and needs decompression")

// DirModuleType classifies module type from dir stream.
type DirModuleType string

const (
	DirModuleUnknown    DirModuleType = "unknown"
	DirModuleProcedural DirModuleType = "procedural"
	DirModuleClass      DirModuleType = "class"
)

// DirModule is one module entry decoded from dir stream records.
type DirModule struct {
	ModuleName string
	StreamName string
	SourceOff  uint32
	Type       DirModuleType
}

// DirInfo contains parsed dir stream metadata.
type DirInfo struct {
	Modules []DirModule
}

// DirDecompressor decodes raw dir stream bytes when compression is used.
type DirDecompressor func([]byte) ([]byte, error)

// MS-OVBA §2.3.4.2 dir stream record type IDs.
const (
	dirRecModuleName            = 0x0019 // MODULENAME record
	dirRecModuleStreamName      = 0x001A // MODULESTREAMNAME (MBCS)
	dirRecModuleStreamNameUTF16 = 0x0032 // MODULESTREAMNAME (UTF-16)
	dirRecModuleOffset          = 0x0031 // MODULEOFFSET: source start in raw stream
	dirRecModuleTypeProcedural  = 0x0021 // MODULETYPE: procedural (.bas)
	dirRecModuleTypeClass       = 0x0022 // MODULETYPE: class/document (.cls)
	dirRecModuleTerminator      = 0x002B // MODULE_TERMINATOR
	dirRecProjectVersion        = 0x0009 // PROJECTVERSION (has extra 2-byte MinorVersion)
)

// ParseDirStream parses a dir stream. If the stream appears compressed and no
// decompressor is provided, ErrDirNeedsDecompression is returned.
func ParseDirStream(data []byte, decompressor DirDecompressor) (*DirInfo, error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("vba: dir stream is empty")
	}

	work := data
	if data[0] == 0x01 {
		if decompressor == nil {
			return nil, ErrDirNeedsDecompression
		}

		decoded, err := decompressor(data)
		if err != nil {
			return nil, fmt.Errorf("vba: decompress dir stream: %w", err)
		}
		work = decoded
	}

	return parseDirRecords(work)
}

func parseDirRecords(data []byte) (*DirInfo, error) {
	result := &DirInfo{}
	var current *DirModule

	appendCurrent := func() {
		if current == nil {
			return
		}
		if current.ModuleName == "" && current.StreamName == "" {
			return
		}
		result.Modules = append(result.Modules, *current)
		current = nil
	}

	for i := 0; i < len(data); {
		if i+6 > len(data) {
			break
		}
		recID := binary.LittleEndian.Uint16(data[i : i+2])
		recSize := int(binary.LittleEndian.Uint32(data[i+2 : i+6]))
		if i+6+recSize > len(data) {
			break
		}

		payload := data[i+6 : i+6+recSize]

		switch recID {
		case dirRecModuleName:
			appendCurrent()
			current = &DirModule{
				ModuleName: trimText(payload),
				Type:       DirModuleUnknown,
			}
		case dirRecModuleStreamName:
			if current == nil {
				current = &DirModule{Type: DirModuleUnknown}
			}
			current.StreamName = trimText(payload)
		case dirRecModuleStreamNameUTF16:
			// Keep ANSI stream name from 0x001A as canonical; unicode variant is optional.
			if current == nil {
				current = &DirModule{Type: DirModuleUnknown}
			}
		case dirRecModuleOffset:
			if current == nil {
				current = &DirModule{Type: DirModuleUnknown}
			}
			if len(payload) >= 4 {
				current.SourceOff = binary.LittleEndian.Uint32(payload[:4])
			}
		case dirRecModuleTypeProcedural:
			if current == nil {
				current = &DirModule{}
			}
			current.Type = DirModuleProcedural
		case dirRecModuleTypeClass:
			if current == nil {
				current = &DirModule{}
			}
			current.Type = DirModuleClass
		case dirRecModuleTerminator:
			appendCurrent()
		}

		i += 6 + recSize

		// PROJECTVERSION (0x0009) has an undocumented extra MinorVersion WORD
		// appended outside the stated Size field (MS-OVBA §2.3.4.2.1.9).
		// Without this +2 the parser would be 2 bytes misaligned for all
		// subsequent records.
		if recID == dirRecProjectVersion && i+2 <= len(data) {
			i += 2
		}
	}

	appendCurrent()

	if len(result.Modules) == 0 {
		return nil, fmt.Errorf("vba: dir stream parser found no module records")
	}

	return result, nil
}

// trimText decodes a byte slice as Windows-1252/Latin-1 and strips null
// bytes and surrounding whitespace. Module names in the dir stream are MBCS
// (Windows codepage), which for Western European characters is Windows-1252.
// Using raw string(b) would corrupt non-ASCII bytes like Ü (0xDC) since Go
// strings are UTF-8, causing name mismatches with the PROJECT stream.
func trimText(b []byte) string {
	decoded, err := charmap.ISO8859_1.NewDecoder().Bytes(b)
	if err != nil {
		decoded = b
	}
	text := strings.TrimRight(string(decoded), "\x00")
	return strings.TrimSpace(text)
}

// ModuleMapping resolves PROJECT + dir information to physical stream entries.
type ModuleMapping struct {
	ModuleName   string
	ModuleType   ProjectModuleType
	StreamName   string
	SourceOffset uint32
	StorageID    int32
	Data         []byte // raw stream bytes from MSysAccessStorage Lv field
}

// BuildModuleMappings builds module→stream mappings.
// If dirInfo is missing/empty, it falls back to stream-order mapping.
func BuildModuleMappings(project *ProjectInfo, dirInfo *DirInfo, moduleStreams []*StorageNode) ([]ModuleMapping, []string) {
	sortedStreams := make([]*StorageNode, len(moduleStreams))
	copy(sortedStreams, moduleStreams)
	sort.Slice(sortedStreams, func(i, j int) bool {
		if sortedStreams[i].ID == sortedStreams[j].ID {
			return sortedStreams[i].Name < sortedStreams[j].Name
		}
		return sortedStreams[i].ID < sortedStreams[j].ID
	})

	if project == nil {
		return nil, []string{"project info is nil"}
	}

	if dirInfo == nil || len(dirInfo.Modules) == 0 {
		return fallbackModuleMappings(project, sortedStreams), []string{"dir mapping unavailable, used stream-order fallback"}
	}

	dirByModule := make(map[string]DirModule, len(dirInfo.Modules))
	for _, module := range dirInfo.Modules {
		if module.ModuleName == "" {
			continue
		}
		dirByModule[strings.ToLower(module.ModuleName)] = module
	}

	streamByName := make(map[string]*StorageNode, len(sortedStreams))
	for _, stream := range sortedStreams {
		streamByName[strings.ToLower(stream.Name)] = stream
	}

	var (
		mappings []ModuleMapping
		warns    []string
	)

	for _, projectModule := range project.Modules {
		dirModule, ok := dirByModule[strings.ToLower(projectModule.Name)]
		if !ok {
			warns = append(warns, fmt.Sprintf("module %q missing in dir mapping", projectModule.Name))
			continue
		}

		stream := streamByName[strings.ToLower(dirModule.StreamName)]
		if stream == nil {
			warns = append(warns, fmt.Sprintf("module %q mapped to unknown stream %q", projectModule.Name, dirModule.StreamName))
			continue
		}

		mappings = append(mappings, ModuleMapping{
			ModuleName:   projectModule.Name,
			ModuleType:   projectModule.Type,
			StreamName:   stream.Name,
			SourceOffset: dirModule.SourceOff,
			StorageID:    stream.ID,
			Data:         stream.Data,
		})
	}

	if len(mappings) == 0 {
		warns = append(warns, "dir mapping produced 0 usable modules, used stream-order fallback")
		return fallbackModuleMappings(project, sortedStreams), warns
	}

	return mappings, warns
}

func fallbackModuleMappings(project *ProjectInfo, sortedStreams []*StorageNode) []ModuleMapping {
	n := len(project.Modules)
	if len(sortedStreams) < n {
		n = len(sortedStreams)
	}

	mappings := make([]ModuleMapping, 0, n)
	for i := 0; i < n; i++ {
		module := project.Modules[i]
		stream := sortedStreams[i]
		mappings = append(mappings, ModuleMapping{
			ModuleName: module.Name,
			ModuleType: module.Type,
			StreamName: stream.Name,
			StorageID:  stream.ID,
			Data:       stream.Data,
		})
	}

	return mappings
}
