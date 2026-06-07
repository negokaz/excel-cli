package excel

import (
	"os"
	"path/filepath"
	"testing"
)

func TestParseRangeParsesSingleCell(t *testing.T) {
	t.Parallel()

	rangeRef := "B3"

	startCol, startRow, endCol, endRow, err := ParseRange(rangeRef)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if startCol != 2 || startRow != 3 || endCol != 2 || endRow != 3 {
		t.Fatalf("unexpected coordinates: %d,%d,%d,%d", startCol, startRow, endCol, endRow)
	}
}

func TestParseRangeParsesAbsoluteRange(t *testing.T) {
	t.Parallel()

	rangeRef := "$A$1:$C$3"

	startCol, startRow, endCol, endRow, err := ParseRange(rangeRef)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if startCol != 1 || startRow != 1 || endCol != 3 || endRow != 3 {
		t.Fatalf("unexpected coordinates: %d,%d,%d,%d", startCol, startRow, endCol, endRow)
	}
}

func TestParseRangeReturnsErrorForInvalidFormat(t *testing.T) {
	t.Parallel()

	rangeRef := "A1:"

	_, _, _, _, err := ParseRange(rangeRef)

	if err == nil {
		t.Fatal("expected invalid range error")
	}
}

func TestNormalizeRangeReturnsCanonicalRange(t *testing.T) {
	t.Parallel()

	rangeRef := "$A$1:$C$3"

	normalized := NormalizeRange(rangeRef)

	if normalized != "A1:C3" {
		t.Fatalf("expected A1:C3, got %s", normalized)
	}
}

func TestFileIsNotWritableReflectsFilesystemState(t *testing.T) {
	t.Parallel()

	tempDir := t.TempDir()
	writablePath := filepath.Join(tempDir, "writable.txt")
	if err := os.WriteFile(writablePath, []byte("ok"), 0o644); err != nil {
		t.Fatalf("failed to create file: %v", err)
	}
	missingPath := filepath.Join(tempDir, "missing.txt")

	writable := FileIsNotWritable(writablePath)
	missing := FileIsNotWritable(missingPath)

	if writable {
		t.Fatalf("expected writable file to be writable: %s", writablePath)
	}
	if !missing {
		t.Fatalf("expected missing file to be treated as not writable: %s", missingPath)
	}
}
