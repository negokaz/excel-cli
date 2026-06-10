package cli

import (
	"fmt"
	"net/url"
	"strings"
	"unicode/utf8"

	"github.com/negokaz/excel-cli/internal/excel"
)

type pathKind string

const (
	pathKindWorkbook pathKind = "workbook"
	pathKindSheet    pathKind = "sheet"
	pathKindRange    pathKind = "range"
)

type targetPath struct {
	Kind      pathKind
	SheetName string
	RangeRef  string
}

func parseTargetPath(raw string) (targetPath, error) {
	if strings.HasPrefix(raw, "/") {
		return targetPath{}, fmt.Errorf("invalid path syntax: path must not begin with /")
	}
	if raw == "" {
		return targetPath{Kind: pathKindWorkbook}, nil
	}

	parts := strings.Split(raw, "/")
	if len(parts) > 2 {
		return targetPath{}, fmt.Errorf("unsupported path kind: %s", raw)
	}

	sheetName, err := url.PathUnescape(parts[0])
	if err != nil {
		return targetPath{}, fmt.Errorf("invalid path syntax: %w", err)
	}
	if sheetName == "" {
		return targetPath{}, fmt.Errorf("invalid path syntax: empty sheet segment")
	}
	if len(parts) == 1 {
		return targetPath{Kind: pathKindSheet, SheetName: sheetName}, nil
	}

	if strings.Contains(parts[1], "!") {
		return targetPath{}, fmt.Errorf("invalid path syntax: Excel-style references are not accepted")
	}
	if _, _, _, _, err := excel.ParseRange(parts[1]); err != nil {
		return targetPath{}, fmt.Errorf("invalid path syntax: %w", err)
	}
	return targetPath{
		Kind:      pathKindRange,
		SheetName: sheetName,
		RangeRef:  excel.NormalizeRange(parts[1]),
	}, nil
}

func canonicalSheetPath(sheetName string) string {
	return canonicalPathSegment(sheetName)
}

// extractPathArg separates the optional path argument from remaining args.
// If args is non-empty and the first element does not start with '-', it is
// treated as the path argument and removed from the remaining slice.
func extractPathArg(args []string) (path string, remaining []string) {
	if len(args) > 0 && !strings.HasPrefix(args[0], "-") {
		return args[0], args[1:]
	}
	return "", args
}

func canonicalRangePath(sheetName, rangeRef string) string {
	return canonicalSheetPath(sheetName) + "/" + excel.NormalizeRange(rangeRef)
}

func canonicalPathSegment(value string) string {
	escaped := url.PathEscape(value)
	if !strings.Contains(escaped, "%") {
		return escaped
	}

	var builder strings.Builder
	for i := 0; i < len(escaped); {
		if escaped[i] != '%' || i+2 >= len(escaped) {
			builder.WriteByte(escaped[i])
			i++
			continue
		}

		b, ok := decodeHexByte(escaped[i+1], escaped[i+2])
		if !ok || b < utf8.RuneSelf {
			builder.WriteString(escaped[i : i+3])
			i += 3
			continue
		}

		start := i
		buf := make([]byte, 0, 4)
		for i+2 < len(escaped) && escaped[i] == '%' {
			next, valid := decodeHexByte(escaped[i+1], escaped[i+2])
			if !valid || next < utf8.RuneSelf {
				break
			}
			buf = append(buf, next)
			i += 3
		}
		if utf8.Valid(buf) {
			builder.Write(buf)
			continue
		}
		builder.WriteString(escaped[start:i])
	}

	return builder.String()
}

func decodeHexByte(hi, lo byte) (byte, bool) {
	high, ok := fromHexDigit(hi)
	if !ok {
		return 0, false
	}
	low, ok := fromHexDigit(lo)
	if !ok {
		return 0, false
	}
	return high<<4 | low, true
}

func fromHexDigit(value byte) (byte, bool) {
	switch {
	case value >= '0' && value <= '9':
		return value - '0', true
	case value >= 'A' && value <= 'F':
		return value - 'A' + 10, true
	case value >= 'a' && value <= 'f':
		return value - 'a' + 10, true
	default:
		return 0, false
	}
}

func (p targetPath) Canonical() string {
	switch p.Kind {
	case pathKindWorkbook:
		return ""
	case pathKindSheet:
		return canonicalSheetPath(p.SheetName)
	case pathKindRange:
		return canonicalRangePath(p.SheetName, p.RangeRef)
	default:
		return ""
	}
}
