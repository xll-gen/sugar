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

// TestWorksheet_Delete closes the last coverage gap on Worksheet: Delete had no
// test at all.
//
// Two details make or break this test:
//
//   - A fresh Workbooks.Add() workbook has ONE sheet on modern Excel
//     (SheetsInNewWorkbook defaults to 1), and Excel REFUSES to delete the last
//     visible sheet. Deleting wb.ActiveSheet() would therefore assert an error
//     path while claiming to assert deletion, so a second sheet is added first.
//   - Count alone is not enough: it would still pass if Delete removed a
//     DIFFERENT sheet. The name has to be gone too.
//
// The subtest at the end pins the refusal, which also proves the test can see
// Delete failing at all.
func TestWorksheet_Delete(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		// The harness suppresses the "permanently delete" modal. If that ever
		// regresses, this test would hang until the 5 s force-kill tier fires
		// and would look like a product failure — so assert it, do not trust it.
		if alerts, err := wb.App().DisplayAlerts(); err != nil || alerts {
			t.Fatalf("DisplayAlerts = %v err=%v; want false (Delete raises a modal otherwise)", alerts, err)
		}

		sheets := wb.Worksheets()
		before, err := sheets.Count()
		if err != nil {
			t.Fatalf("Worksheets().Count(): %v", err)
		}

		const name = "SugarDeleteMe"
		victim := sheets.Add(excel.AddName(name))
		if err := victim.Err(); err != nil {
			t.Fatalf("Worksheets().Add(AddName(%q)): %v", name, err)
		}
		if got, err := sheets.Count(); err != nil || got != before+1 {
			t.Fatalf("Count after Add = %d err=%v; want %d", got, err, before+1)
		}
		if got, err := victim.Name(); err != nil || got != name {
			t.Fatalf("new sheet Name = %q err=%v; want %q", got, err, name)
		}

		if err := victim.Delete(); err != nil {
			t.Fatalf("Delete: %v", err)
		}
		if got, err := sheets.Count(); err != nil || got != before {
			t.Errorf("Count after Delete = %d err=%v; want %d", got, err, before)
		}
		// The count would also be satisfied by deleting the wrong sheet.
		if err := sheets.Item(name).Err(); err == nil {
			t.Errorf("Worksheets().Item(%q) still resolves after Delete", name)
		}
		remaining, err := sheets.Count()
		if err != nil {
			t.Fatalf("Count: %v", err)
		}
		var names []string
		for i := int32(1); i <= remaining; i++ {
			n, err := sheets.Item(i).Name()
			if err != nil {
				t.Fatalf("Item(%d).Name(): %v", i, err)
			}
			names = append(names, n)
		}
		for _, n := range names {
			if n == name {
				t.Errorf("sheet %q survived Delete (sheets: %v)", name, names)
			}
		}

		// Deleting the last remaining sheet must fail, not silently succeed —
		// which also proves this test can observe Delete failing at all.
		//
		// Inline, NOT a t.Run subtest: t.Run runs the closure on a fresh
		// goroutine, which is outside this thread's COM apartment, and every
		// call there fails with "CoInitialize has not been called". Any COM work
		// must stay on the sugar.Do thread.
		if remaining != 1 {
			t.Fatalf("expected a single sheet left, got %d (%v)", remaining, names)
		}
		last := sheets.Item(1)
		if err := last.Err(); err != nil {
			t.Fatalf("Worksheets().Item(1): %v", err)
		}
		if err := last.Delete(); err == nil {
			t.Error("Delete on the workbook's only sheet returned nil; Excel must refuse it")
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
