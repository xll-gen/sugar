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

		val, err := rng.Get("Value").Value()
		if err != nil {
			t.Fatalf("failed to get value: %v", err)
		}

		if val != "Sugar Excel" {
			t.Errorf("expected 'Sugar Excel', got %v", val)
		}

		return nil
	})
}
