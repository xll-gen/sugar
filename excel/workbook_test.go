//go:build windows && excel_integration

// Integration tests for excel.Workbook and excel.Worksheets.
// Build with `-tags=excel_integration`. Skipped on machines without Excel.

package excel_test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

// TestWorkbook_NameAndPath checks the three string identity getters
// (Name, FullName, Path). FullName equals Path joined with Name only after
// the workbook has been saved to disk — pre-save Path is empty.
func TestWorkbook_NameAndPath(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "sugar_workbook_test.xlsx")
	t.Cleanup(func() { _ = os.Remove(path) })

	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)
		defer app.Quit()

		wb := app.Workbooks().Add()

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
		return nil
	})
}

// TestWorkbook_SheetsAlias verifies that Sheets() and Worksheets() return
// collections of identical Count. The two are aliases per xlwings naming.
func TestWorkbook_SheetsAlias(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)
		defer app.Quit()

		wb := app.Workbooks().Add()

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
		return nil
	})
}

// TestWorksheets_AddAndCount adds a sheet by name and confirms Count grows.
// Excel defaults a workbook to 1 sheet; adding two more should yield 3.
func TestWorksheets_AddAndCount(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)
		defer app.Quit()

		wb := app.Workbooks().Add()
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
		return nil
	})
}

// TestWorksheet_NameAndIndex exercises the Worksheet identity properties.
func TestWorksheet_NameAndIndex(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)
		defer app.Quit()

		wb := app.Workbooks().Add()
		s := wb.ActiveSheet()

		s.SetName("Renamed")
		got, err := s.Name()
		if err != nil || got != "Renamed" {
			t.Errorf("Name after SetName: got %q, err=%v", got, err)
		}

		idx, err := s.Index()
		if err != nil || idx < 1 {
			t.Errorf("Index: got %d, err=%v; want >=1", idx, err)
		}
		return nil
	})
}
