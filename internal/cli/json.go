package cli

import (
	"encoding/json"
	"fmt"
	"os"
)

func writeJSON(value any) error {
	payload, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to encode JSON response: %w", err)
	}
	fmt.Fprintln(os.Stdout, string(payload))
	return nil
}
