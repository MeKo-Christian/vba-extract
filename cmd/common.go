package cmd

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"log/slog"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
	"github.com/MeKo-Tech/vba-extract/internal/vba"
)

type listEntry struct {
	Name      string `json:"name"`
	Type      string `json:"type"`
	Stream    string `json:"stream"`
	SizeBytes int    `json:"sizeBytes"`
	Partial   bool   `json:"partial"`
}

func discoverInputFiles(args []string, recursive bool) ([]string, error) {
	seen := map[string]struct{}{}
	var files []string

	for _, arg := range args {
		expanded, err := expandArg(arg)
		if err != nil {
			return nil, err
		}

		for _, candidate := range expanded {
			info, err := os.Stat(candidate)
			if err != nil {
				return nil, fmt.Errorf("stat %q: %w", candidate, err)
			}

			if info.IsDir() {
				if !recursive {
					continue
				}

				err = filepath.WalkDir(candidate, func(path string, d os.DirEntry, walkErr error) error {
					if walkErr != nil {
						return walkErr
					}

					if d.IsDir() {
						return nil
					}

					if !isAccessFile(path) {
						return nil
					}

					addUniquePath(path, seen, &files)

					return nil
				})
				if err != nil {
					return nil, fmt.Errorf("walk %q: %w", candidate, err)
				}

				continue
			}

			if isAccessFile(candidate) {
				addUniquePath(candidate, seen, &files)
			}
		}
	}

	sort.Strings(files)

	return files, nil
}

func expandArg(arg string) ([]string, error) {
	if hasGlob(arg) {
		matches, err := filepath.Glob(arg)
		if err != nil {
			return nil, fmt.Errorf("glob %q: %w", arg, err)
		}

		if len(matches) == 0 {
			return nil, fmt.Errorf("glob %q matched no files", arg)
		}

		return matches, nil
	}

	return []string{arg}, nil
}

func hasGlob(s string) bool {
	return strings.ContainsAny(s, "*?[")
}

func addUniquePath(path string, seen map[string]struct{}, files *[]string) {
	abs, err := filepath.Abs(path)
	if err != nil {
		abs = path
	}

	if _, ok := seen[abs]; ok {
		return
	}

	seen[abs] = struct{}{}
	*files = append(*files, abs)
}

func isAccessFile(path string) bool {
	ext := strings.ToLower(filepath.Ext(path))
	return ext == ".mdb" || ext == ".accdb"
}

func loadModules(path string) ([]vba.ExtractedModule, error) {
	db, err := mdb.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open %q: %w", path, err)
	}
	defer db.Close()

	log := slog.Default()
	st, stErr := vba.LoadStorageTree(db)

	var modules []vba.ExtractedModule

	var extractErr error
	if stErr == nil {
		modules, extractErr = vba.ExtractAllModules(st, log)
		if extractErr == nil && len(modules) > 0 {
			return modules, nil
		}
	} else {
		extractErr = stErr
	}

	// Fallback: scan raw LVAL chains for orphaned module streams.
	// This handles databases where MSysAccessStorage is missing entirely or
	// where the VBA structure was stripped from it but page data remains on disk.
	scanned, scanErr := vba.ScanOrphanedLvalModules(db)
	if scanErr == nil && len(scanned) > 0 {
		log.Debug("vba: standard extraction failed; recovered modules via raw LVAL scan",
			"err", extractErr, "count", len(scanned))

		return scanned, nil
	}

	// Both standard extraction and LVAL scan found nothing. If the only failure
	// was "no VBA structure found" (MSysAccessStorage missing or empty), treat
	// that as a database with no VBA rather than a hard error — return empty.
	// Only propagate errors that indicate a real read failure.
	if extractErr != nil && scanErr == nil {
		log.Debug("vba: no VBA found", "path", path, "err", extractErr)
		return nil, nil
	}

	if extractErr != nil {
		return nil, fmt.Errorf("extract modules %q: %w", path, extractErr)
	}

	return modules, nil
}

func moduleExt(moduleType vba.ProjectModuleType) string {
	if moduleType == vba.ProjectModuleClass || moduleType == vba.ProjectModuleDocument {
		return ".cls"
	}

	return ".bas"
}

func safeModuleName(name string) string {
	name = strings.TrimSpace(name)
	if name == "" {
		return "unnamed"
	}

	replacer := strings.NewReplacer("/", "_", "\\", "_", ":", "_", "*", "_", "?", "_", "\"", "_", "<", "_", ">", "_", "|", "_")

	return replacer.Replace(name)
}

func writeModules(baseOutDir, dbPath string, modules []vba.ExtractedModule, flat bool) (int, int, error) {
	dbName := strings.TrimSuffix(filepath.Base(dbPath), filepath.Ext(dbPath))

	targetDir := baseOutDir
	if !flat {
		targetDir = filepath.Join(baseOutDir, dbName)
	}

	err := os.MkdirAll(targetDir, 0o755)
	if err != nil {
		return 0, 0, fmt.Errorf("create output dir %q: %w", targetDir, err)
	}

	written := 0
	totalLines := 0

	for _, module := range modules {
		filename := safeModuleName(module.Name) + moduleExt(module.Type)

		outPath := filepath.Join(targetDir, filename)

		err := os.WriteFile(outPath, []byte(module.Text), 0o600)
		if err != nil {
			return written, totalLines, fmt.Errorf("write %q: %w", outPath, err)
		}

		written++
		totalLines += strings.Count(module.Text, "\n")
	}

	return written, totalLines, nil
}

func computeFileHash(path string) (string, error) {
	f, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer f.Close()

	h := sha256.New()

	_, err = io.Copy(h, f)
	if err != nil {
		return "", err
	}

	return hex.EncodeToString(h.Sum(nil)), nil
}

func defaultOutputDir() string {
	if outputDir != "" {
		return outputDir
	}

	return filepath.Join(".", "vba-output")
}

func printListTable(entries []listEntry) {
	fmt.Printf("%-30s %-10s %-30s %-8s %-7s\n", "MODULE", "TYPE", "STREAM", "BYTES", "PARTIAL")

	for _, e := range entries {
		fmt.Printf("%-30s %-10s %-30s %-8d %-7v\n", e.Name, e.Type, e.Stream, e.SizeBytes, e.Partial)
	}
}

func printListJSON(entries []listEntry) error {
	b, err := json.MarshalIndent(entries, "", "  ")
	if err != nil {
		return err
	}

	fmt.Println(string(b))

	return nil
}

func colorEnabled() bool {
	fi, err := os.Stdout.Stat()
	if err != nil {
		return false
	}

	return (fi.Mode() & os.ModeCharDevice) != 0
}

func colorize(code, text string) string {
	if !colorEnabled() {
		return text
	}

	return "\x1b[" + code + "m" + text + "\x1b[0m"
}
