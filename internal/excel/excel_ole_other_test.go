//go:build !windows

package excel

import (
	"strings"
	"testing"
)

func TestNewExcelOleReturnsUnsupportedErrorOnNonWindows(t *testing.T) {
	t.Parallel()

	path := "/tmp/test.xlsx"

	workbook, release, err := NewExcelOle(path)
	defer release()

	if err == nil {
		t.Fatal("expected unsupported error")
	}
	if workbook != nil {
		t.Fatalf("expected nil workbook, got %#v", workbook)
	}
	if !strings.Contains(err.Error(), "only supported on Windows") {
		t.Fatalf("expected Windows support error, got %v", err)
	}
}
