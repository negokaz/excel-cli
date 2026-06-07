package cli

import "testing"

func TestParseTargetPath(t *testing.T) {
	t.Parallel()

	t.Run("parses workbook root", func(t *testing.T) {
		t.Parallel()
		target, err := parseTargetPath("/")
		if err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
		if target.Kind != pathKindWorkbook || target.Canonical() != "/" {
			t.Fatalf("unexpected target: %+v", target)
		}
	})

	t.Run("parses encoded sheet path", func(t *testing.T) {
		t.Parallel()
		target, err := parseTargetPath("/Hidden%20Sheet")
		if err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
		if target.Kind != pathKindSheet || target.SheetName != "Hidden Sheet" {
			t.Fatalf("unexpected target: %+v", target)
		}
		if target.Canonical() != "/Hidden%20Sheet" {
			t.Fatalf("unexpected canonical path: %s", target.Canonical())
		}
	})

	t.Run("preserves unicode in canonical path output", func(t *testing.T) {
		t.Parallel()
		target, err := parseTargetPath("/テスト2/A1:B1")
		if err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
		if target.Canonical() != "/テスト2/A1:B1" {
			t.Fatalf("unexpected canonical path: %s", target.Canonical())
		}
	})

	t.Run("keeps ascii escapes in canonical path output", func(t *testing.T) {
		t.Parallel()
		target, err := parseTargetPath("/100%25")
		if err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
		if target.Canonical() != "/100%25" {
			t.Fatalf("unexpected canonical path: %s", target.Canonical())
		}
	})

	t.Run("parses range path and normalizes address", func(t *testing.T) {
		t.Parallel()
		target, err := parseTargetPath("/Data/$a$1:$c$2")
		if err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
		if target.Kind != pathKindRange || target.RangeRef != "A1:C2" {
			t.Fatalf("unexpected target: %+v", target)
		}
	})

	t.Run("rejects invalid syntax", func(t *testing.T) {
		t.Parallel()
		if _, err := parseTargetPath("Data/A1"); err == nil {
			t.Fatal("expected invalid path error")
		}
	})
}
