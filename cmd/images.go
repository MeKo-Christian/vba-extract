package cmd

import (
	"errors"
	"fmt"
	"log/slog"
	"os"
	"path/filepath"
	"strings"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
	"github.com/spf13/cobra"
)

var (
	imagesRecursive bool
	imagesFlat      bool
	imagesStrict    bool
	imagesDedupe    bool
)

var imagesCmd = &cobra.Command{
	Use:   "images [files...]",
	Short: "Extract embedded images from Access form definitions",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		inputs, err := discoverInputFiles(args, imagesRecursive)
		if err != nil {
			return err
		}

		out := cmd.OutOrStdout()

		if len(inputs) == 0 {
			return errors.New("no .mdb/.accdb files found in inputs")
		}

		baseOut := defaultOutputDir()
		seenHashes := map[string]string{}
		processed, totalImages, failed := 0, 0, 0

		for _, file := range inputs {
			if imagesDedupe {
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

			images, loadErr := loadImages(file)
			if loadErr != nil {
				failed++

				fmt.Fprintf(out, "%s %s: %s\n", colorize("31", "ERROR"), filepath.Base(file), formatCommandError(file, loadErr))

				if imagesStrict {
					return loadErr
				}

				continue
			}

			if len(images) == 0 {
				fmt.Fprintf(out, "%s %s -> no images found\n", colorize("33", "SKIP"), filepath.Base(file))
				continue
			}

			count, writeErr := writeImages(baseOut, file, images, imagesFlat || format == outputFormatFlat)
			if writeErr != nil {
				failed++

				fmt.Fprintf(out, "%s %s: %s\n", colorize("31", "ERROR"), filepath.Base(file), formatCommandError(file, writeErr))

				if imagesStrict {
					return writeErr
				}

				continue
			}

			totalImages += count
			fmt.Fprintf(out, "%s %s -> images=%d\n", colorize("32", "OK"), filepath.Base(file), count)
		}

		fmt.Fprintf(out, "summary: processed=%d images=%d failed=%d output=%s\n", processed, totalImages, failed, baseOut)

		if imagesStrict && failed > 0 {
			return fmt.Errorf("strict mode: %d file(s) failed", failed)
		}

		return nil
	},
}

func loadImages(path string) ([]mdb.EmbeddedImage, error) {
	db, err := mdb.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open %q: %w", path, err)
	}
	defer db.Close()

	return mdb.ExtractImages(db)
}

func writeImages(baseOutDir, dbPath string, images []mdb.EmbeddedImage, flat bool) (int, error) {
	dbName := strings.TrimSuffix(filepath.Base(dbPath), filepath.Ext(dbPath))

	targetDir := baseOutDir
	if !flat {
		targetDir = filepath.Join(baseOutDir, dbName, "images")
	}

	err := os.MkdirAll(targetDir, 0o755)
	if err != nil {
		return 0, fmt.Errorf("create output dir %q: %w", targetDir, err)
	}

	written := 0
	usedNames := map[string]int{}

	for _, img := range images {
		filename := imageFilename(img, usedNames)
		outPath := filepath.Join(targetDir, filename)

		err := os.WriteFile(outPath, img.Data, 0o600)
		if err != nil {
			return written, fmt.Errorf("write %q: %w", outPath, err)
		}

		written++
	}

	return written, nil
}

// imageFilename generates a unique filename for an extracted image.
func imageFilename(img mdb.EmbeddedImage, usedNames map[string]int) string {
	ext := "." + img.Format
	if img.Format == "jpeg" {
		ext = ".jpg"
	}

	var base string
	if img.FileName != "" {
		base = safeModuleName(img.FileName)
		// Remove double extension if the original filename already has the right one.
		for _, suffix := range []string{".jpg", ".jpeg", ".png", ".bmp", ".gif"} {
			if strings.HasSuffix(strings.ToLower(base), suffix) {
				base = base[:len(base)-len(suffix)]
				break
			}
		}
	} else {
		base = safeModuleName(img.FormName)
	}

	name := base + ext
	if n, ok := usedNames[strings.ToLower(name)]; ok {
		usedNames[strings.ToLower(name)] = n + 1
		name = fmt.Sprintf("%s_%d%s", base, n+1, ext)
	}

	usedNames[strings.ToLower(name)] = 1

	return name
}

func init() {
	rootCmd.AddCommand(imagesCmd)
	imagesCmd.Flags().BoolVar(&imagesRecursive, "recursive", false, "Recursively process directories")
	imagesCmd.Flags().BoolVar(&imagesFlat, "flat", false, "Write all images into one directory")
	imagesCmd.Flags().BoolVar(&imagesStrict, "strict", false, "Fail on first file error")
	imagesCmd.Flags().BoolVar(&imagesDedupe, "dedupe", false, "Skip duplicate files by content hash")
}
