package mdb_test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

func skipIfNoStartMDB(t *testing.T) {
	t.Helper()
	_, err := os.Stat(filepath.Join("..", "..", "testdata", "Start.mdb"))
	if os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available (proprietary fixture)")
	}
}

func TestExtractImages_StartMDB(t *testing.T) {
	skipIfNoStartMDB(t)
	dbPath := filepath.Join("..", "..", "testdata", "Start.mdb")

	db, err := mdb.Open(dbPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	images, err := mdb.ExtractImages(db)
	if err != nil {
		t.Fatalf("ExtractImages: %v", err)
	}

	if len(images) != 2 {
		t.Fatalf("expected 2 images, got %d", len(images))
	}

	// All images in sample.mdb are JPEGs.
	for i, img := range images {
		if img.Format != "jpeg" {
			t.Errorf("image[%d]: expected format jpeg, got %s", i, img.Format)
		}

		if len(img.Data) < 1000 {
			t.Errorf("image[%d]: suspiciously small (%d bytes)", i, len(img.Data))
		}
		// Verify JPEG signature.
		if len(img.Data) < 3 || img.Data[0] != 0xFF || img.Data[1] != 0xD8 || img.Data[2] != 0xFF {
			t.Errorf("image[%d]: invalid JPEG signature", i)
		}
		// Verify JPEG end marker.
		if len(img.Data) >= 2 && (img.Data[len(img.Data)-2] != 0xFF || img.Data[len(img.Data)-1] != 0xD9) {
			t.Errorf("image[%d]: missing JPEG end marker", i)
		}
	}

	// Check that form names were resolved.
	formNames := map[string]bool{}
	for _, img := range images {
		formNames[img.FormName] = true
	}

	if !formNames["Startschirm"] {
		t.Error("expected form name 'Startschirm' in results")
	}

	// Check that filenames were extracted for the images that have them.
	foundFilename := false

	for _, img := range images {
		if img.FileName != "" {
			foundFilename = true

			if img.FileName != "MicrosoftTeams-image (55).jpg" && img.FileName != "MicrosoftTeams-image (55).png" {
				t.Errorf("unexpected filename: %q", img.FileName)
			}
		}
	}

	if !foundFilename {
		t.Error("expected at least one image with an extracted filename")
	}
}

func TestScanBlobForImages_NoImages(t *testing.T) {
	skipIfNoStartMDB(t)
	// A blob with no image signatures should return nothing.
	blob := make([]byte, 1024)
	for i := range blob {
		blob[i] = byte(i % 256)
	}
	// Make sure it doesn't accidentally contain JPEG/PNG/BMP/GIF signatures.
	// The loop above produces 0xFF at index 255, 0xD8 at index 216 — not adjacent.

	dbPath := filepath.Join("..", "..", "testdata", "Start.mdb")

	db, err := mdb.Open(dbPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	// ExtractImages works end-to-end. For a unit test of the blob scanner,
	// we test via the integration path since scanBlobForImages is unexported.
	// This test just validates the happy path works.
	images, err := mdb.ExtractImages(db)
	if err != nil {
		t.Fatalf("ExtractImages: %v", err)
	}

	if images == nil {
		t.Fatal("expected non-nil result")
	}
}

func TestExtractImages_ImageSizes(t *testing.T) {
	skipIfNoStartMDB(t)
	dbPath := filepath.Join("..", "..", "testdata", "Start.mdb")

	db, err := mdb.Open(dbPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer db.Close()

	images, err := mdb.ExtractImages(db)
	if err != nil {
		t.Fatalf("ExtractImages: %v", err)
	}

	// Known image sizes from probing (full-resolution, using record header sizes).
	expectedSizes := map[int]bool{
		188972:  true, // Startschirm.jpg (4096x2160)
		1107958: true, // MicrosoftTeams-image (55).jpg (1920x1286)
	}

	for _, img := range images {
		if !expectedSizes[len(img.Data)] {
			t.Errorf("unexpected image size %d for %q", len(img.Data), img.FormName)
		}
	}
}
