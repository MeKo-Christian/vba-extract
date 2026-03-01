package cmd

import (
	"os"
	"testing"
)

const testMDB = "../testdata/sample.mdb"

const startMDB = "../testdata/Start.mdb"

func skipIfNoFixture(t *testing.T) {
	t.Helper()

	_, err := os.Stat(testMDB)
	if os.IsNotExist(err) {
		t.Skip("testdata/sample.mdb not available")
	}
}

func skipIfNoStartFixture(t *testing.T) {
	t.Helper()

	_, err := os.Stat(startMDB)
	if os.IsNotExist(err) {
		t.Skip("testdata/Start.mdb not available (proprietary fixture)")
	}
}

// loadModules

func TestLoadModules_returnsModules(t *testing.T) {
	skipIfNoStartFixture(t)

	modules, err := loadModules(startMDB)
	if err != nil {
		t.Fatalf("loadModules: %v", err)
	}

	if len(modules) == 0 {
		t.Fatal("expected at least one module from Start.mdb")
	}
}

func TestLoadModules_missingFile(t *testing.T) {
	_, err := loadModules("/nonexistent/path.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

// loadSchema

func TestLoadSchema_returnsSchema(t *testing.T) {
	skipIfNoFixture(t)

	schema, err := loadSchema(testMDB)
	if err != nil {
		t.Fatalf("loadSchema: %v", err)
	}

	if schema == nil {
		t.Fatal("expected non-nil schema")
	}

	if len(schema.Tables) == 0 {
		t.Fatal("expected at least one table in schema")
	}
}

func TestLoadSchema_missingFile(t *testing.T) {
	_, err := loadSchema("/nonexistent/path.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

// loadImages

func TestLoadImages_returnsImages(t *testing.T) {
	skipIfNoStartFixture(t)

	images, err := loadImages(startMDB)
	if err != nil {
		t.Fatalf("loadImages: %v", err)
	}

	if len(images) != 2 {
		t.Fatalf("expected 2 images, got %d", len(images))
	}

	for i, img := range images {
		if img.Format != "jpeg" {
			t.Errorf("image[%d]: expected jpeg, got %s", i, img.Format)
		}

		if len(img.Data) < 1000 {
			t.Errorf("image[%d]: suspiciously small (%d bytes)", i, len(img.Data))
		}
	}
}

func TestLoadImages_missingFile(t *testing.T) {
	_, err := loadImages("/nonexistent/path.mdb")
	if err == nil {
		t.Error("expected error for missing file")
	}
}

func TestWriteImages_createsFiles(t *testing.T) {
	skipIfNoStartFixture(t)

	images, err := loadImages(startMDB)
	if err != nil {
		t.Fatalf("loadImages: %v", err)
	}

	dir := t.TempDir()

	count, err := writeImages(dir, startMDB, images, false)
	if err != nil {
		t.Fatalf("writeImages: %v", err)
	}

	if count != len(images) {
		t.Errorf("expected %d written, got %d", len(images), count)
	}

	entries, err := os.ReadDir(dir + "/Start/images")
	if err != nil {
		t.Fatalf("ReadDir: %v", err)
	}

	if len(entries) != count {
		t.Errorf("expected %d files on disk, got %d", count, len(entries))
	}
}

func TestLoadSchema_legacyFixture(t *testing.T) {
	path := "../testdata/jet35/st990426.mdb"

	_, err := os.Stat(path)

	if os.IsNotExist(err) {
		t.Skip("legacy fixture not available")
	}

	schema, err := loadSchema(path)
	if err != nil {
		t.Fatalf("loadSchema legacy fixture: %v", err)
	}

	if schema == nil {
		t.Fatal("expected non-nil schema from legacy fixture")
	}

	if len(schema.Tables) == 0 {
		t.Fatal("expected at least one table in legacy fixture schema")
	}
}
