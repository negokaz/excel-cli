package cli

import (
	"encoding/json"
	"fmt"
)

func parseValuesJSON(raw string) ([][]any, error) {
	var parsed any
	if err := json.Unmarshal([]byte(raw), &parsed); err != nil {
		return nil, fmt.Errorf("failed to parse values as JSON: %w", err)
	}

	outer, ok := parsed.([]any)
	if !ok {
		return nil, fmt.Errorf("values must be a JSON 2-dimensional array")
	}

	result := make([][]any, len(outer))
	for i, item := range outer {
		row, ok := item.([]any)
		if !ok {
			return nil, fmt.Errorf("values must be a JSON 2-dimensional array: row %d is not an array", i)
		}
		result[i] = row
	}

	return result, nil
}
