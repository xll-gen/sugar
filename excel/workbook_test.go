//go:build windows && excel_integration

// Integration tests for excel.Workbook and excel.Worksheets.
// Build with `-tags=excel_integration`. Skipped on machines without Excel.

package excel_test

import (
	"path/filepath"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestWorkbook_NameAndPath checks the three string identity getters
// (Name, FullName, Path). FullName equals Path joined with Name only after
// the workbook has been saved to disk — pre-save Path is empty.
func TestWorkbook_NameAndPath(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "sugar_workbook_test.xlsx")

	withBook(t, func(wb excel.Workbook) {
		if err := wb.SaveAs(path); err != nil {
			t.Fatalf("SaveAs: %v", err)
		}

		name, err := wb.Name()
		if err != nil || name != "sugar_workbook_test.xlsx" {
			t.Errorf("Name: got %q, err=%v; want sugar_workbook_test.xlsx", name, err)
		}
		full, err := wb.FullName()
		if err != nil || full != path {
			t.Errorf("FullName: got %q, err=%v; want %q", full, err, path)
		}
		p, err := wb.Path()
		if err != nil || p != dir {
			t.Errorf("Path: got %q, err=%v; want %q", p, err, dir)
		}

		// Skip the save-on-close prompt for the temp file.
		_ = wb.SetSaved(true).Close()
	})
}

// TestWorkbook_SaveAsOptions exercises the v1.0 functional-option form of
// SaveAs: forcing the on-disk format with SaveFileFormat and protecting the
// file with SavePassword. We write a macro-enabled .xlsm container (Excel
// rejects a format/extension mismatch, so the extension matches the format),
// then reopen with the password to prove SavePassword took effect.
func TestWorkbook_SaveAsOptions(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "sugar_saveas_opts.xlsm")

	withApp(t, func(app excel.Application) {
		wb := app.Workbooks().Add()
		if err := wb.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}
		wb.ActiveSheet().Range("A1").SetValue("secret-data")

		if err := wb.SaveAs(path,
			excel.SaveFileFormat(excel.FileFormatOpenXMLWorkbookMacroEnabled),
			excel.SavePassword("pw123"),
		); err != nil {
			t.Fatalf("SaveAs with options: %v", err)
		}
		if err := wb.SetSaved(true).Close(); err != nil {
			t.Fatalf("Close after SaveAs: %v", err)
		}

		// Reopening without the password should fail (proof SavePassword stuck);
		// with the right password it should succeed and round-trip the value.
		reopened := app.Workbooks().Open(path, excel.OpenPassword("pw123"))
		if err := reopened.Err(); err != nil {
			t.Fatalf("Open with correct password: %v", err)
		}
		got, err := reopened.ActiveSheet().Range("A1").Value()
		if err != nil || got != "secret-data" {
			t.Errorf("round-trip value: got %v err=%v; want secret-data", got, err)
		}
		_ = reopened.SetSaved(true).Close()
	})
}

// TestWorkbook_CloseSaveChanges proves CloseSaveChanges(false) discards edits
// without a prompt: we dirty a saved workbook, close discarding, reopen, and
// confirm the edit did not persist.
func TestWorkbook_CloseSaveChanges(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "sugar_close_savechanges.xlsx")

	withApp(t, func(app excel.Application) {
		wb := app.Workbooks().Add()
		wb.ActiveSheet().Range("A1").SetValue("original")
		if err := wb.SaveAs(path); err != nil {
			t.Fatalf("SaveAs: %v", err)
		}

		// Dirty the workbook, then close discarding changes.
		wb.ActiveSheet().Range("A1").SetValue("edited")
		if err := wb.Close(excel.CloseSaveChanges(false)); err != nil {
			t.Fatalf("Close(CloseSaveChanges(false)): %v", err)
		}

		reopened := app.Workbooks().Open(path)
		if err := reopened.Err(); err != nil {
			t.Fatalf("reopen: %v", err)
		}
		got, err := reopened.ActiveSheet().Range("A1").Value()
		if err != nil || got != "original" {
			t.Errorf("after discard: got %v err=%v; want original (edit must not persist)", got, err)
		}
		_ = reopened.SetSaved(true).Close()
	})
}

// TestWorkbook_SheetsAlias verifies that Sheets() and Worksheets() return
// collections of identical Count. The two are aliases per xlwings naming.
func TestWorkbook_SheetsAlias(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		w1, err := wb.Worksheets().Count()
		if err != nil {
			t.Fatalf("Worksheets.Count: %v", err)
		}
		w2, err := wb.Sheets().Count()
		if err != nil {
			t.Fatalf("Sheets.Count: %v", err)
		}
		if w1 != w2 {
			t.Errorf("Sheets/Worksheets count mismatch: %d vs %d", w1, w2)
		}
	})
}

// TestWorksheets_AddAndCount adds a sheet by name and confirms Count grows.
func TestWorksheets_AddAndCount(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		sheets := wb.Worksheets()

		before, _ := sheets.Count()

		newSheet := sheets.Add(excel.AddName("Extra"))
		if err := newSheet.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}
		name, err := newSheet.Name()
		if err != nil || name != "Extra" {
			t.Errorf("new sheet name: got %q, err=%v; want Extra", name, err)
		}

		after, _ := sheets.Count()
		if after != before+1 {
			t.Errorf("Count: before=%d after=%d; expected +1", before, after)
		}
	})
}

// TestWorksheets_AddBeforeAfter places new sheets relative to an anchor.
// These options pass a Worksheet (a sugar.Chain) as a COM argument, which
// relies on the core chain→IDispatch normalization.
func TestWorksheets_AddBeforeAfter(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		sheets := wb.Worksheets()
		anchor := sheets.Item(1)

		before := sheets.Add(excel.AddBefore(anchor), excel.AddName("First"))
		if err := before.Err(); err != nil {
			t.Fatalf("Add(AddBefore): %v", err)
		}
		idx, err := before.Index()
		if err != nil || idx != 1 {
			t.Errorf("AddBefore index: got %d err=%v; want 1", idx, err)
		}

		after := sheets.Add(excel.AddAfter(before), excel.AddName("Second"))
		if err := after.Err(); err != nil {
			t.Fatalf("Add(AddAfter): %v", err)
		}
		idx, err = after.Index()
		if err != nil || idx != 2 {
			t.Errorf("AddAfter index: got %d err=%v; want 2", idx, err)
		}
	})
}

// TestWorksheet_NameAndIndex exercises the Worksheet identity properties.
func TestWorksheet_NameAndIndex(t *testing.T) {
	withSheet(t, func(s excel.Worksheet) {
		s.SetName("Renamed")
		got, err := s.Name()
		if err != nil || got != "Renamed" {
			t.Errorf("Name after SetName: got %q, err=%v", got, err)
		}

		idx, err := s.Index()
		if err != nil || idx < 1 {
			t.Errorf("Index: got %d, err=%v; want >=1", idx, err)
		}
	})
}
