//go:build windows && excel_integration

// Integration tests for excel.Range.
//
// These run against a live Excel instance and the build tag keeps them off
// CI hosts without Office. Run them with:
//
//	go test -tags=excel_integration ./excel/...

package excel_test

import (
	"reflect"
	"testing"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

// TestRange_ScalarRoundTrip writes a single cell and reads it back through
// the typed Range.Value() getter. This is the path xlwings calls
// `Range("A1").value = 42; Range("A1").value` and is the simplest possible
// proof that the value getter works at all.
func TestRange_ScalarRoundTrip(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("Hello")
		got, err := sheet.Range("A1").Value()
		if err != nil {
			t.Fatalf("Value() failed: %v", err)
		}
		if got != "Hello" {
			t.Errorf("expected \"Hello\", got %v (%T)", got, got)
		}

		sheet.Range("B1").SetValue(42.5)
		got, err = sheet.Range("B1").Value()
		if err != nil {
			t.Fatalf("Value() failed: %v", err)
		}
		if got != 42.5 {
			t.Errorf("expected 42.5, got %v (%T)", got, got)
		}
	})
}

// TestRange_Value2D exercises the SAFEARRAY decode path: a multi-cell range
// must return [][]interface{} shaped [rows][cols]. Before v0.7.0 this path
// silently returned nil because go-ole's VARIANT.Value() does not handle
// VT_ARRAY|VT_VARIANT.
func TestRange_Value2D(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		// Seed a 2×3 block via individual cells to keep this test focused on
		// the read path (the block-write path has its own test below).
		sheet.Cells(1, 1).SetValue(1.0)
		sheet.Cells(1, 2).SetValue(2.0)
		sheet.Cells(1, 3).SetValue("c")
		sheet.Cells(2, 1).SetValue(4.0)
		sheet.Cells(2, 2).SetValue(5.0)
		sheet.Cells(2, 3).SetValue("f")

		got, err := sheet.Range("A1", "C2").Value()
		if err != nil {
			t.Fatalf("Value() failed: %v", err)
		}
		grid, ok := got.([][]interface{})
		if !ok {
			t.Fatalf("expected [][]interface{}, got %T (%v)", got, got)
		}
		want := [][]interface{}{
			{1.0, 2.0, "c"},
			{4.0, 5.0, "f"},
		}
		if !reflect.DeepEqual(grid, want) {
			t.Errorf("grid mismatch:\n got  %v\n want %v", grid, want)
		}
	})
}

// TestRange_SetValue2D writes a whole block with one SetValue call — the
// write-direction mirror of TestRange_Value2D. xlwings equivalent:
// `sheet.range("A1").value = [[...], [...]]`.
func TestRange_SetValue2D(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		want := [][]interface{}{
			{"h1", "h2", "h3"},
			{1.0, 2.0, 3.0},
		}
		if err := sheet.Range("A1", "C2").SetValue(want).Err(); err != nil {
			t.Fatalf("SetValue 2-D failed: %v", err)
		}

		got, err := sheet.Range("A1", "C2").Value()
		if err != nil {
			t.Fatalf("Value() failed: %v", err)
		}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("grid mismatch:\n got  %v\n want %v", got, want)
		}
	})
}

// TestRange_End covers the Ctrl+Arrow navigation primitive in all four
// directions plus the invalid-direction error path.
func TestRange_End(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		if err := sheet.Range("A1", "A3").SetValue([][]interface{}{
			{1.0}, {2.0}, {3.0},
		}).Err(); err != nil {
			t.Fatalf("seed: %v", err)
		}

		addr, err := sheet.Range("A1").End("down").Address()
		if err != nil || addr != "$A$3" {
			t.Errorf("End(down): got %q err=%v; want $A$3", addr, err)
		}
		addr, err = sheet.Range("A3").End("up").Address()
		if err != nil || addr != "$A$1" {
			t.Errorf("End(up): got %q err=%v; want $A$1", addr, err)
		}

		if err := sheet.Range("A1").End("sideways").Err(); err == nil {
			t.Errorf("End(sideways) should error")
		}
	})
}

// TestRange_ColorRoundTrip writes and reads the Interior fill color.
func TestRange_ColorRoundTrip(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		yellow := excel.RGB(255, 255, 0)
		if err := sheet.Range("B2").SetColor(yellow).Err(); err != nil {
			t.Fatalf("SetColor: %v", err)
		}
		got, err := sheet.Range("B2").Color()
		if err != nil || got != yellow {
			t.Errorf("Color: got %d err=%v; want %d", got, err, yellow)
		}
	})
}

// TestRange_Dimensions covers Width/Height (points, read-only) and
// ColumnWidth/RowHeight (settable).
func TestRange_Dimensions(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")

		if err := rng.SetColumnWidth(20).SetRowHeight(30).Err(); err != nil {
			t.Fatalf("set dimensions: %v", err)
		}
		cw, err := rng.ColumnWidth()
		if err != nil || cw != 20 {
			t.Errorf("ColumnWidth: got %v err=%v; want 20", cw, err)
		}
		rh, err := rng.RowHeight()
		if err != nil || rh != 30 {
			t.Errorf("RowHeight: got %v err=%v; want 30", rh, err)
		}

		w, err := rng.Width()
		if err != nil || w <= 0 {
			t.Errorf("Width: got %v err=%v; want > 0", w, err)
		}
		h, err := rng.Height()
		if err != nil || h != 30 {
			t.Errorf("Height: got %v err=%v; want 30 (single row of height 30)", h, err)
		}
	})
}

// TestRange_Insert shifts cells down and verifies the displaced value.
func TestRange_Insert(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("original")

		if err := sheet.Range("A1").Insert("down"); err != nil {
			t.Fatalf("Insert(down): %v", err)
		}
		got, err := sheet.Range("A2").Value()
		if err != nil || got != "original" {
			t.Errorf("after Insert(down): A2 = %v err=%v; want original", got, err)
		}

		if err := sheet.Range("A1").Insert("diagonal"); err == nil {
			t.Errorf("Insert(diagonal) should error")
		}
	})
}

// TestRange_Find covers both the hit (returns the cell) and the miss
// (Excel's Nothing → found=false, not an error or a panic).
func TestRange_Find(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		if err := sheet.Range("A1", "B2").SetValue([][]interface{}{
			{"alpha", "beta"},
			{"gamma", "delta"},
		}).Err(); err != nil {
			t.Fatalf("seed: %v", err)
		}

		cell, found, err := sheet.Range("A1", "B2").Find("gamma")
		if err != nil || !found {
			t.Fatalf("Find(gamma): found=%v err=%v; want hit", found, err)
		}
		addr, err := cell.Address()
		if err != nil || addr != "$A$2" {
			t.Errorf("Find(gamma) address: got %q err=%v; want $A$2", addr, err)
		}

		_, found, err = sheet.Range("A1", "B2").Find("no_such_value")
		if err != nil || found {
			t.Errorf("Find(miss): found=%v err=%v; want clean miss", found, err)
		}
	})
}

// TestRange_Geometry checks the Row/Column/Count/Address quartet against a
// known-shape range. Excel uses 1-based indexing, mirrored in our API.
func TestRange_Geometry(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("B3", "D5") // 3 rows × 3 cols

		row, err := rng.Row()
		if err != nil || row != 3 {
			t.Errorf("Row: got %d, err=%v; want 3", row, err)
		}
		col, err := rng.Column()
		if err != nil || col != 2 {
			t.Errorf("Column: got %d, err=%v; want 2", col, err)
		}
		count, err := rng.Count()
		if err != nil || count != 9 {
			t.Errorf("Count: got %d, err=%v; want 9", count, err)
		}
		addr, err := rng.Address()
		if err != nil || addr != "$B$3:$D$5" {
			t.Errorf("Address: got %q, err=%v; want $B$3:$D$5", addr, err)
		}
	})
}

// TestRange_OffsetAndResize verifies the two grid-navigation primitives.
// xlwings' equivalents are `rng.offset(r, c)` and `rng.resize(rows, cols)`.
func TestRange_OffsetAndResize(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		// Offset(1, 1) on A1 should be B2.
		offset := sheet.Range("A1").Offset(1, 1)
		addr, err := offset.Address()
		if err != nil || addr != "$B$2" {
			t.Errorf("Offset: got %q, err=%v; want $B$2", addr, err)
		}

		// Resize(2, 3) on A1 should give A1:C2.
		resized := sheet.Range("A1").Resize(2, 3)
		addr, err = resized.Address()
		if err != nil || addr != "$A$1:$C$2" {
			t.Errorf("Resize: got %q, err=%v; want $A$1:$C$2", addr, err)
		}
	})
}

// TestRange_FormulaRoundTrip writes a formula via SetFormula and reads it
// back. NumberFormat is also exercised — together they cover the four
// string-typed property setters added in v0.7.0.
func TestRange_FormulaRoundTrip(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")
		rng.SetFormula("=1+2")
		f, err := rng.Formula()
		if err != nil || f != "=1+2" {
			t.Errorf("Formula: got %q, err=%v; want =1+2", f, err)
		}

		rng.SetNumberFormat("0.00")
		nf, err := rng.NumberFormat()
		if err != nil || nf != "0.00" {
			t.Errorf("NumberFormat: got %q, err=%v; want 0.00", nf, err)
		}
	})
}

// TestRange_ClearAndMerge covers the destructive helpers added in v0.7.0:
// ClearContents must drop values but leave the range usable; Merge/UnMerge
// must round-trip.
func TestRange_ClearAndMerge(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("x")
		if err := sheet.Range("A1").ClearContents(); err != nil {
			t.Fatalf("ClearContents: %v", err)
		}
		got, _ := sheet.Range("A1").Value()
		if got != nil {
			t.Errorf("after ClearContents, expected nil, got %v", got)
		}

		rng := sheet.Range("B2", "C3")
		if err := rng.Merge(); err != nil {
			t.Fatalf("Merge: %v", err)
		}
		merged, err := rng.MergeCells()
		if err != nil || !merged {
			t.Errorf("after Merge: MergeCells=%v, err=%v; want true", merged, err)
		}
		if err := rng.UnMerge(); err != nil {
			t.Fatalf("UnMerge: %v", err)
		}
	})
}

// withSheet is the standard integration-test harness for excel.Range tests:
// it launches Excel invisibly, opens a fresh workbook, hands the first sheet
// to the test, and guarantees Excel is closed even on panic.
func withSheet(t *testing.T, fn func(sheet excel.Worksheet)) {
	t.Helper()
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)
		defer app.Quit()

		wb := app.Workbooks().Add()
		if err := wb.Err(); err != nil {
			t.Fatalf("Add workbook failed: %v", err)
		}
		sheet := wb.ActiveSheet()
		fn(sheet)
		return nil
	})
}
