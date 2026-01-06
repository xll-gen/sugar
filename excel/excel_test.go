//go:build windows

package excel_test

import (
	"testing"
	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

func TestExcel_Package(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		
		// Ensure cleanup
		defer app.Put("DisplayAlerts", false).Call("Quit")

		wb := app.Workbooks().Add()
		if err := wb.Err(); err != nil {
			t.Fatalf("failed to add workbook: %v", err)
		}

		sheet := wb.ActiveSheet()
		if err := sheet.Err(); err != nil {
			t.Fatalf("failed to get active sheet: %v", err)
		}

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

		val, err := cell.Get("Value").Value()
		if err != nil {
			t.Fatalf("failed to get value from B2: %v", err)
		}
		if val != "Cell B2" {
			t.Errorf("expected 'Cell B2', got %v", val)
		}

		return nil
	})
}
