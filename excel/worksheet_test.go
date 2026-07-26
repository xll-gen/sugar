//go:build windows && excel_integration

// Integration tests for excel.Worksheet.
//
// These run against a live Excel instance and the build tag keeps them off
// CI hosts without Office. Run them with:
//
//	go test -tags=excel_integration ./excel/...

package excel_test

import (
	"strings"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestWorksheet_Clear pins the Clear/ClearContents fix: Worksheet.Cells is a
// COM *property* (propget), so it must be read with Get. Reading it with Call
// (DISPATCH_METHOD) made both methods fail 100% of the time with a COM
// member-not-found error, so neither cleared anything.
func TestWorksheet_Clear(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		seed := func() {
			if err := sheet.Range("A1", "B2").SetValue([][]interface{}{
				{"alpha", "beta"},
				{"gamma", "delta"},
			}).Err(); err != nil {
				t.Fatalf("seed: %v", err)
			}
			if err := sheet.Range("A1").SetColor(excel.RGB(255, 255, 0)).Err(); err != nil {
				t.Fatalf("seed color: %v", err)
			}
		}

		// --- ClearContents: values go, formatting stays. ---
		seed()
		if err := sheet.ClearContents(); err != nil {
			t.Fatalf("ClearContents: %v", err)
		}
		if v, err := sheet.Range("A1").Value(); err != nil || v != nil {
			t.Errorf("after ClearContents A1 = %v (%T) err=%v; want nil", v, v, err)
		}
		if v, err := sheet.Range("B2").Value(); err != nil || v != nil {
			t.Errorf("after ClearContents B2 = %v (%T) err=%v; want nil", v, v, err)
		}
		if c, err := sheet.Range("A1").Color(); err != nil || c != excel.RGB(255, 255, 0) {
			t.Errorf("after ClearContents A1 color = %d err=%v; want the fill preserved", c, err)
		}

		// --- Clear: values AND formatting go. ---
		seed()
		if err := sheet.Clear(); err != nil {
			t.Fatalf("Clear: %v", err)
		}
		if v, err := sheet.Range("A1").Value(); err != nil || v != nil {
			t.Errorf("after Clear A1 = %v (%T) err=%v; want nil", v, v, err)
		}
		noFill, err := sheet.Range("A1").Color()
		if err != nil {
			t.Fatalf("Color after Clear: %v", err)
		}
		if noFill == excel.RGB(255, 255, 0) {
			t.Errorf("after Clear A1 still carries the yellow fill (%d); formatting was not cleared", noFill)
		}
	})
}

// TestWorksheet_ClearIsWholeSheet proves Clear/ClearContents act on the whole
// sheet (Cells), not just UsedRange as recorded when the call was made.
func TestWorksheet_ClearIsWholeSheet(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("near")
		sheet.Range("Z100").SetValue("far")

		if err := sheet.ClearContents(); err != nil {
			t.Fatalf("ClearContents: %v", err)
		}
		for _, addr := range []string{"A1", "Z100"} {
			if v, err := sheet.Range(addr).Value(); err != nil || v != nil {
				t.Errorf("after ClearContents %s = %v err=%v; want nil", addr, v, err)
			}
		}
	})
}

// TestWorksheet_NameIndexVisible covers the remaining scalar Worksheet
// getters/setters — the file had no integration coverage at all before the
// Clear fix (AGENTS.md §5 rule 9).
func TestWorksheet_NameIndexVisible(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		if err := sheet.SetName("SugarSheet").Err(); err != nil {
			t.Fatalf("SetName: %v", err)
		}
		if got, err := sheet.Name(); err != nil || got != "SugarSheet" {
			t.Errorf("Name: got %q err=%v; want SugarSheet", got, err)
		}
		if idx, err := sheet.Index(); err != nil || idx != 1 {
			t.Errorf("Index: got %d err=%v; want 1", idx, err)
		}

		if got, err := sheet.Visible(); err != nil || got != excel.SheetVisible {
			t.Errorf("Visible: got %d err=%v; want %d", got, err, excel.SheetVisible)
		}
	})
}

// TestWorksheet_SetVisibleRoundTrip exercises SetVisible, which the getter test
// above cannot: Excel refuses to hide the last visible sheet, so a second sheet
// has to exist first.
func TestWorksheet_SetVisibleRoundTrip(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		sheet := wb.ActiveSheet()
		if err := sheet.Err(); err != nil {
			t.Fatalf("ActiveSheet: %v", err)
		}
		// Keep a visible sheet behind us; hiding the only one fails in Excel.
		if err := wb.Worksheets().Add().Err(); err != nil {
			t.Fatalf("Worksheets().Add: %v", err)
		}

		for _, want := range []excel.SheetVisibility{excel.SheetHidden, excel.SheetVeryHidden, excel.SheetVisible} {
			if err := sheet.SetVisible(want).Err(); err != nil {
				t.Fatalf("SetVisible(%d): %v", want, err)
			}
			got, err := sheet.Visible()
			if err != nil {
				t.Fatalf("Visible after SetVisible(%d): %v", want, err)
			}
			if got != want {
				t.Errorf("Visible: got %d, want %d", got, want)
			}
		}
	})
}

// TestWorksheet_AutoFit widens a narrow column by auto-fitting long content.
func TestWorksheet_AutoFit(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		if err := sheet.Range("A1").SetColumnWidth(2).Err(); err != nil {
			t.Fatalf("SetColumnWidth: %v", err)
		}
		sheet.Range("A1").SetValue(strings.Repeat("wide ", 8))

		if err := sheet.AutoFit(); err != nil {
			t.Fatalf("AutoFit: %v", err)
		}
		w, err := sheet.Range("A1").ColumnWidth()
		if err != nil || w <= 2 {
			t.Errorf("after AutoFit ColumnWidth = %v err=%v; want > 2", w, err)
		}
	})
}
