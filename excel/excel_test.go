//go:build windows && excel_integration

// Package-level smoke test: workbook + sheet + range round trip through the
// typed wrappers. Build with `-tags=excel_integration`.

package excel_test

import (
	"testing"

	"github.com/xll-gen/sugar/excel"
)

func TestExcel_Package(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")
		rng.SetValue("Sugar Excel")
		if err := rng.Err(); err != nil {
			t.Fatalf("failed to set value: %v", err)
		}

		// Test Cells on Worksheet
		cell := sheet.Cells(2, 2) // B2
		cell.SetValue("Cell B2")
		if err := cell.Err(); err != nil {
			t.Fatalf("failed to set value via Cells: %v", err)
		}

		val, err := cell.Value()
		if err != nil {
			t.Fatalf("failed to get value from B2: %v", err)
		}
		if val != "Cell B2" {
			t.Errorf("expected 'Cell B2', got %v", val)
		}
	})
}
