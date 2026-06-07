package excel

import (
	"fmt"
	"os"
	"path"
	"regexp"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// ParseRange parses an Excel range string (e.g. "A1:C10" or "A1").
func ParseRange(rangeStr string) (int, int, int, int, error) {
	re := regexp.MustCompile(`^(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?$`)
	matches := re.FindStringSubmatch(rangeStr)
	if matches == nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format: %s", rangeStr)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(matches[1])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	if matches[2] == "" {
		return startCol, startRow, startCol, startRow, nil
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(matches[2])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	return startCol, startRow, endCol, endRow, nil
}

func NormalizeRange(rangeStr string) string {
	startCol, startRow, endCol, endRow, _ := ParseRange(rangeStr)
	startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, _ := excelize.CoordinatesToCellName(endCol, endRow)
	return fmt.Sprintf("%s:%s", startCell, endCell)
}

func anyMatrixToStringMatrix(values [][]any) [][]string {
	result := make([][]string, len(values))
	for rowIdx, row := range values {
		result[rowIdx] = make([]string, len(row))
		for colIdx, value := range row {
			result[rowIdx][colIdx] = stringifyCellValue(value)
		}
	}
	return result
}

func stringifyCellValue(value any) string {
	switch v := value.(type) {
	case nil:
		return ""
	case string:
		return v
	case bool:
		if v {
			return "TRUE"
		}
		return "FALSE"
	case int:
		return strconv.FormatInt(int64(v), 10)
	case int8:
		return strconv.FormatInt(int64(v), 10)
	case int16:
		return strconv.FormatInt(int64(v), 10)
	case int32:
		return strconv.FormatInt(int64(v), 10)
	case int64:
		return strconv.FormatInt(v, 10)
	case uint:
		return strconv.FormatUint(uint64(v), 10)
	case uint8:
		return strconv.FormatUint(uint64(v), 10)
	case uint16:
		return strconv.FormatUint(uint64(v), 10)
	case uint32:
		return strconv.FormatUint(uint64(v), 10)
	case uint64:
		return strconv.FormatUint(v, 10)
	case float32:
		return strconv.FormatFloat(float64(v), 'f', -1, 64)
	case float64:
		return strconv.FormatFloat(v, 'f', -1, 64)
	default:
		return fmt.Sprintf("%v", v)
	}
}

func parseNumericString(raw string) (any, bool) {
	if raw == "" {
		return "", false
	}
	f, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return nil, false
	}
	if float64(int64(f)) == f {
		return int64(f), true
	}
	return f, true
}

func FileIsNotWritable(absolutePath string) bool {
	f, err := os.OpenFile(path.Clean(absolutePath), os.O_WRONLY, os.ModePerm)
	if err != nil {
		return true
	}
	defer f.Close()
	return false
}
