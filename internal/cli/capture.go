package cli

import (
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runCapture(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("usage: excel-cli capture <file> <sheet> [range]")
	}

	filePath := args[0]
	sheetName := args[1]
	var rangeStr string
	if len(args) >= 3 {
		rangeStr = args[2]
	}

	absPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("failed to resolve path: %w", err)
	}

	workbook, release, err := excel.OpenFile(absPath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer release()

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to find sheet %q: %w", sheetName, err)
	}
	defer worksheet.Release()

	if rangeStr == "" {
		rangeStr, err = worksheet.GetDimention()
		if err != nil {
			return fmt.Errorf("failed to get sheet dimension: %w", err)
		}
	}

	base64image, err := worksheet.CapturePicture(rangeStr)
	if err != nil {
		return fmt.Errorf("failed to capture picture: %w", err)
	}

	imgBytes, err := base64.StdEncoding.DecodeString(base64image)
	if err != nil {
		return fmt.Errorf("failed to decode image data: %w", err)
	}

	outPath, err := writeCaptureOutput(imgBytes)
	if err != nil {
		return fmt.Errorf("failed to write output: %w", err)
	}

	fmt.Fprintln(os.Stdout, outPath)
	return nil
}

func captureOutputFileName() string {
	t := time.Now().UTC()
	ms := t.Nanosecond() / 1_000_000
	return fmt.Sprintf("capture-%s-%03dZ.png", t.Format("2006-01-02T15-04-05"), ms)
}

func writeCaptureOutput(data []byte) (string, error) {
	dir := ".excel-cli"
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", fmt.Errorf("failed to create output directory: %w", err)
	}
	filename := captureOutputFileName()
	outPath := filepath.Join(dir, filename)
	if err := os.WriteFile(outPath, data, 0o644); err != nil {
		return "", fmt.Errorf("failed to write file: %w", err)
	}
	absPath, err := filepath.Abs(outPath)
	if err != nil {
		return outPath, nil
	}
	return absPath, nil
}
