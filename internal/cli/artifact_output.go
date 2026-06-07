package cli

import (
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"
	"time"
)

func writeOutput(content string) (string, error) {
	dir := ".excel-cli"
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", fmt.Errorf("failed to create output directory: %w", err)
	}
	filename := outputFileName()
	outPath := filepath.Join(dir, filename)
	if err := os.WriteFile(outPath, []byte(content), 0o644); err != nil {
		return "", fmt.Errorf("failed to write file: %w", err)
	}
	absPath, err := filepath.Abs(outPath)
	if err != nil {
		return outPath, nil
	}
	return absPath, nil
}

func outputFileName() string {
	t := time.Now().UTC()
	ms := t.Nanosecond() / 1_000_000
	return fmt.Sprintf("sheet-%s-%03dZ.html", t.Format("2006-01-02T15-04-05"), ms)
}

func decodeBase64Image(base64image string) ([]byte, error) {
	imgBytes, err := base64.StdEncoding.DecodeString(base64image)
	if err != nil {
		return nil, fmt.Errorf("failed to decode image data: %w", err)
	}
	return imgBytes, nil
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

func captureOutputFileName() string {
	t := time.Now().UTC()
	ms := t.Nanosecond() / 1_000_000
	return fmt.Sprintf("capture-%s-%03dZ.png", t.Format("2006-01-02T15-04-05"), ms)
}
