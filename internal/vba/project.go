package vba

import (
	"bufio"
	"bytes"
	"errors"
	"fmt"
	"strings"

	"golang.org/x/text/encoding/charmap"
)

// ProjectModuleType classifies module kind parsed from PROJECT stream.
type ProjectModuleType string

const (
	ProjectModuleStandard ProjectModuleType = "standard"
	ProjectModuleClass    ProjectModuleType = "class"
	ProjectModuleDocument ProjectModuleType = "document"
)

// ProjectModule is one module declaration from PROJECT stream text.
type ProjectModule struct {
	Name string
	Type ProjectModuleType
}

// ProjectInfo contains parsed data from PROJECT stream.
type ProjectInfo struct {
	Name         string
	Modules      []ProjectModule
	Metadata     map[string]string
	HostExtender map[string]string
}

// ParseProjectStream parses Access PROJECT stream text (latin-1 encoded).
func ParseProjectStream(data []byte) (*ProjectInfo, error) {
	if len(data) == 0 {
		return nil, errors.New("vba: PROJECT stream is empty")
	}

	decoded, err := charmap.ISO8859_1.NewDecoder().Bytes(data)
	if err != nil {
		return nil, fmt.Errorf("vba: decode PROJECT stream: %w", err)
	}

	info := &ProjectInfo{
		Metadata:     make(map[string]string),
		HostExtender: make(map[string]string),
	}

	scanner := bufio.NewScanner(bytes.NewReader(decoded))
	section := ""

	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" || strings.HasPrefix(line, "'") {
			continue
		}

		if strings.HasPrefix(line, "[") && strings.HasSuffix(line, "]") {
			section = strings.TrimSpace(line[1 : len(line)-1])
			continue
		}

		key, value, ok := splitKV(line)
		if !ok {
			continue
		}

		switch key {
		case "Module":
			info.Modules = append(info.Modules, ProjectModule{Name: value, Type: ProjectModuleStandard})
		case "Class":
			info.Modules = append(info.Modules, ProjectModule{Name: value, Type: ProjectModuleClass})
		case "DocClass":
			name := value
			if idx := strings.Index(name, "/"); idx >= 0 {
				name = name[:idx]
			}

			info.Modules = append(info.Modules, ProjectModule{Name: strings.TrimSpace(name), Type: ProjectModuleDocument})
		default:
			if strings.EqualFold(section, "Host Extender Info") {
				info.HostExtender[key] = value
			} else {
				info.Metadata[key] = value
			}

			if strings.EqualFold(key, "Name") && info.Name == "" {
				info.Name = stripQuotes(value)
			}
		}
	}

	if err := scanner.Err(); err != nil {
		return nil, fmt.Errorf("vba: scan PROJECT stream: %w", err)
	}

	return info, nil
}

// stripQuotes removes a single layer of surrounding double-quotes.
// The PROJECT stream stores Name as e.g. `"Start"` (VBA string literal syntax).
func stripQuotes(s string) string {
	if len(s) >= 2 && s[0] == '"' && s[len(s)-1] == '"' {
		return s[1 : len(s)-1]
	}

	return s
}

func splitKV(line string) (key string, value string, ok bool) {
	idx := strings.Index(line, "=")
	if idx <= 0 {
		return "", "", false
	}

	key = strings.TrimSpace(line[:idx])
	value = strings.TrimSpace(line[idx+1:])

	if key == "" {
		return "", "", false
	}

	return key, value, true
}
