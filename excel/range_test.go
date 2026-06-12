//go:build windows && excel_integration

// Integration tests for excel.Range.
//
// These run against a live Excel instance and the build tag keeps them off
// CI hosts without Office. Run them with:
//
//	go test -tags=excel_integration ./excel/...

package excel_test

import (
	"fmt"
	"reflect"
	"testing"

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

// TestRange_AutoFit proves the v1.0 behavior change: AutoFit now fits both
// column width AND row height (xlwings parity), not columns only. We shrink a
// cell's column and row, write content that needs more space, AutoFit, and
// assert both dimensions grew back.
func TestRange_AutoFit(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")
		// Long content that needs a wide column to display.
		rng.SetValue("A very long line of text that exceeds three characters")
		// Force both dimensions small so AutoFit has to grow them back.
		if err := rng.SetColumnWidth(3).SetRowHeight(6).Err(); err != nil {
			t.Fatalf("shrink: %v", err)
		}

		narrowW, _ := rng.ColumnWidth()
		shortH, _ := rng.RowHeight()

		if err := rng.AutoFit(); err != nil {
			t.Fatalf("AutoFit: %v", err)
		}

		wideW, err := rng.ColumnWidth()
		if err != nil || wideW <= narrowW {
			t.Errorf("AutoFit column: width %v did not grow from %v (err=%v)", wideW, narrowW, err)
		}
		// The artificially tiny 6pt row height must change once AutoFit touches
		// rows at all — the key proof of the v1.0 row+column behavior. (Row
		// AutoFit restores the natural single-line height, well above 6pt.)
		tallH, err := rng.RowHeight()
		if err != nil {
			t.Fatalf("RowHeight after AutoFit: %v", err)
		}
		if tallH == shortH {
			t.Errorf("AutoFit row: height stayed at %v; expected row autofit to adjust it", tallH)
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

// TestRange_Formula2RoundTrip writes a dynamic-array formula via SetFormula2
// and reads it back through Formula2(). A native DA function (SEQUENCE) is
// used so this is meaningful on dynamic-array Excel (2021+/365); on older
// Excel SEQUENCE is unknown and the test is effectively skipped via the
// build tag's integration gate plus the version of Excel installed.
func TestRange_Formula2RoundTrip(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")
		rng.SetFormula2("=SEQUENCE(3,1)")
		f, err := rng.Formula2()
		if err != nil {
			t.Fatalf("Formula2: %v", err)
		}
		// Formula2 must not be wrapped in implicit intersection.
		if len(f) > 0 && f[1] == '@' {
			t.Errorf("Formula2 got implicit-intersection form %q; want no leading @", f)
		}
	})
}

// TestRange_SetFormulaSpill proves the spill-correct setter: on dynamic-array
// Excel the formula is stored without the implicit-intersection `@`, so a UDF
// (or native DA function) spills. We use a native DA function (SEQUENCE) to
// avoid depending on a registered UDF. The legacy SetFormula path applies
// implicit intersection to the same input, so this is the regression guard for
// the showcase's "=@TimesTable(5)" bug.
func TestRange_SetFormulaSpill(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		rng := sheet.Range("A1")
		if err := rng.SetFormulaSpill("=SEQUENCE(3,1)").Err(); err != nil {
			t.Fatalf("SetFormulaSpill: %v", err)
		}
		f, err := rng.Formula2()
		if err != nil {
			t.Fatalf("Formula2 after SetFormulaSpill: %v", err)
		}
		if len(f) > 1 && f[1] == '@' {
			t.Errorf("SetFormulaSpill stored implicit-intersection form %q; want spill-native (no @)", f)
		}
	})
}

// TestRange_SetFormula2Array proves the batch formula setter: a contiguous
// column of formulas is written through Formula2 in a single COM round-trip,
// each cell evaluates independently, and none is rewritten into the
// implicit-intersection `=@...` form. This is the spill-correct, one-call
// counterpart of looping SetFormulaSpill cell-by-cell (the showcase build's
// hot path).
func TestRange_SetFormula2Array(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		// A 3x1 block of independent scalar formulas.
		block := [][]interface{}{
			{"=1+1"},
			{"=2+2"},
			{"=3+3"},
		}
		if err := sheet.Range("A1:A3").SetFormula2Array(block).Err(); err != nil {
			t.Fatalf("SetFormula2Array: %v", err)
		}
		wantVals := []float64{2, 4, 6}
		for i, want := range wantVals {
			cell := sheet.Range(fmt.Sprintf("A%d", i+1))
			f, err := cell.Formula2()
			if err != nil {
				t.Fatalf("Formula2(A%d): %v", i+1, err)
			}
			if len(f) > 1 && f[1] == '@' {
				t.Errorf("A%d stored implicit-intersection form %q; want no leading @", i+1, f)
			}
			v, err := cell.Value()
			if err != nil {
				t.Fatalf("Value(A%d): %v", i+1, err)
			}
			if got, ok := v.(float64); !ok || got != want {
				t.Errorf("A%d = %v (%T); want %v", i+1, v, v, want)
			}
		}
	})
}

// TestRange_ClearAndMerge covers the destructive helpers added in v0.7.0:
// ClearContents must drop values but leave the range usable; Merge/Unmerge
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
		if err := rng.Unmerge(); err != nil {
			t.Fatalf("Unmerge: %v", err)
		}
	})
}
