package cmd

import (
	"errors"
	"fmt"
	"log/slog"
	"path/filepath"

	"github.com/spf13/cobra"
)

var (
	extractRecursive       bool
	extractFlat            bool
	extractStrict          bool
	extractDedupe          bool
	extractOverwriteReadme bool
)

var extractCmd = &cobra.Command{
	Use:   "extract [files...]",
	Short: "Extract VBA modules from one or more Access files",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		inputs, err := discoverInputFiles(args, extractRecursive)
		if err != nil {
			return err
		}

		out := cmd.OutOrStdout()

		if len(inputs) == 0 {
			return errors.New("no .mdb/.accdb files found in inputs")
		}

		baseOut := defaultOutputDir()

		seenHashes := map[string]string{}
		processed := 0
		writtenModules := 0
		totalLines := 0
		failed := 0

		for _, file := range inputs {
			if extractDedupe {
				hash, hashErr := computeFileHash(file)
				if hashErr == nil {
					if prev, ok := seenHashes[hash]; ok {
						slog.Default().Debug("skip duplicate", "file", file, "sameAs", prev)
						continue
					}

					seenHashes[hash] = file
				}
			}

			processed++

			modules, loadErr := loadModules(file)
			if loadErr != nil {
				failed++

				fmt.Fprintf(out, "%s %s: %s\n", colorize("31", "ERROR"), filepath.Base(file), formatCommandError(file, loadErr))

				if extractStrict {
					return loadErr
				}

				continue
			}

			count, lines, writeErr := writeModules(baseOut, file, modules, extractFlat || format == outputFormatFlat, extractOverwriteReadme)
			if writeErr != nil {
				failed++

				fmt.Fprintf(out, "%s %s: %s\n", colorize("31", "ERROR"), filepath.Base(file), formatCommandError(file, writeErr))

				if extractStrict {
					return writeErr
				}

				continue
			}

			writtenModules += count
			totalLines += lines
			fmt.Fprintf(out, "%s %s -> modules=%d lines=%d\n", colorize("32", "OK"), filepath.Base(file), count, lines)

			partialModules, warningCount := extractionStats(modules)
			if partialModules > 0 || warningCount > 0 {
				fmt.Fprintf(out, "%s %s -> partial=%d warnings=%d\n",
					colorize("33", "WARN"), filepath.Base(file), partialModules, warningCount)
			}
		}

		fmt.Fprintf(out, "summary: processed=%d modules=%d lines=%d failed=%d output=%s\n", processed, writtenModules, totalLines, failed, baseOut)

		if extractStrict && failed > 0 {
			return fmt.Errorf("strict mode: %d file(s) failed", failed)
		}

		return nil
	},
}

func init() {
	rootCmd.AddCommand(extractCmd)
	extractCmd.Flags().BoolVar(&extractRecursive, "recursive", false, "Recursively process directories")
	extractCmd.Flags().BoolVar(&extractFlat, "flat", false, "Write all modules into one directory")
	extractCmd.Flags().BoolVar(&extractStrict, "strict", false, "Fail on first file error")
	extractCmd.Flags().BoolVar(&extractDedupe, "dedupe", false, "Skip duplicate files by content hash")
	extractCmd.Flags().BoolVar(&extractOverwriteReadme, "overwrite-readme", false, "Overwrite existing README.md files")
}
