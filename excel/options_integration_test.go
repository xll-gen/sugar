//go:build windows && excel_integration

// Integration tests for the excel.Options conversion framework.
//
// These tests cover the option flags that depend on real COM behaviour and
// cannot be faked: Expand (uses Range.End / Worksheet.Range) and the
// struct-by-header decode driven off Range.Value. Run with:
//
//	go test -tags=excel_integration ./excel/...

package excel_test

import (
	"reflect"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestOptions_ExpandDown grows a one-cell anchor down through a contiguous
// column of values and verifies the resulting range covers the whole block.
// xlwings parity: `rng.options(expand="down").value`.
func TestOptions_ExpandDown(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(1.0)
		sheet.Range("A2").SetValue(2.0)
		sheet.Range("A3").SetValue(3.0)
		// A4 deliberately empty — Expand stops at the gap.

		v, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Options(Expand=down).Value(): %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{} for 3x1 expand, got %T (%v)", v, v)
		}
		want := []interface{}{1.0, 2.0, 3.0}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_ExpandRight is the row analogue of ExpandDown.
func TestOptions_ExpandRight(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(10.0)
		sheet.Range("B1").SetValue(20.0)
		sheet.Range("C1").SetValue(30.0)

		v, err := sheet.Range("A1").Options(excel.Expand("right")).Value()
		if err != nil {
			t.Fatalf("Expand=right: %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{}, got %T (%v)", v, v)
		}
		want := []interface{}{10.0, 20.0, 30.0}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_ExpandTable seeds a 2x3 block and verifies expand="table"
// walks both directions from the anchor.
func TestOptions_ExpandTable(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(1.0)
		sheet.Range("B1").SetValue(2.0)
		sheet.Range("C1").SetValue(3.0)
		sheet.Range("A2").SetValue(4.0)
		sheet.Range("B2").SetValue(5.0)
		sheet.Range("C2").SetValue(6.0)

		v, err := sheet.Range("A1").Options(excel.Expand("table")).Value()
		if err != nil {
			t.Fatalf("Expand=table: %v", err)
		}
		got, ok := v.([][]interface{})
		if !ok {
			t.Fatalf("expected [][]interface{}, got %T (%v)", v, v)
		}
		want := [][]interface{}{
			{1.0, 2.0, 3.0},
			{4.0, 5.0, 6.0},
		}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_ExpandDown_SingleCell verifies the blank-neighbor guard: a lone
// value with an empty cell directly below stays a 1x1 read (scalar), instead of
// End(xlDown) overshooting to the sheet boundary.
func TestOptions_ExpandDown_SingleCell(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("solo")
		// A2 empty.
		v, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Expand=down single cell: %v", err)
		}
		if v != "solo" {
			t.Errorf("blank-neighbor guard: got %T %v, want scalar \"solo\"", v, v)
		}
	})
}

// TestOptions_ExpandDown_SeparateIsland is the key regression: a blank cell
// below the anchor sits above a *separate* data island further down. The guard
// must keep the expansion at the anchor rather than letting End(xlDown) jump
// across the gap to the island (which would read a block full of empty cells).
func TestOptions_ExpandDown_SeparateIsland(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("top")
		// A2..A4 empty; A5 begins an unrelated island.
		sheet.Range("A5").SetValue("island")
		sheet.Range("A6").SetValue("island2")

		v, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Expand=down separate island: %v", err)
		}
		if v != "top" {
			t.Errorf("guard should stop at anchor; got %T %v, want scalar \"top\"", v, v)
		}
	})
}

// TestOptions_ExpandTable_BlankCorner is the live-Excel regression for the
// endpoint ladder. The layout is the commonest table shape there is — an empty
// top-left corner, headers along row 1, labels down column A:
//
//	    A      B    C
//	1 (empty) Jan  Feb
//	2  North   1    2
//	3  South   3    4
//	4  East    5    6
//
// The old guard probed A2, found it non-empty, and then called End("down") from
// the *blank* A1. Excel's End() from an empty cell lands on the first non-empty
// cell rather than the end of the run, so the expansion reported A1:B2 and the
// read returned 4 cells out of 12 with err == nil. Only real Excel can attest
// that End()-from-blank semantics; the Excel-free fake asserts the same shape.
func TestOptions_ExpandTable_BlankCorner(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		// A1 deliberately left empty.
		seed := [][]interface{}{
			{nil, "Jan", "Feb"},
			{"North", 1.0, 2.0},
			{"South", 3.0, 4.0},
			{"East", 5.0, 6.0},
		}
		if err := sheet.Range("A1", "C4").SetValue(seed).Err(); err != nil {
			t.Fatalf("seed block: %v", err)
		}

		v, err := sheet.Range("A1").Options(excel.Expand("table")).Value()
		if err != nil {
			t.Fatalf("Expand=table with a blank corner: %v", err)
		}
		if !reflect.DeepEqual(v, seed) {
			t.Errorf("blank-corner table: got %v, want %v", v, seed)
		}

		down, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Expand=down with a blank corner: %v", err)
		}
		wantDown := []interface{}{nil, "North", "South", "East"}
		if !reflect.DeepEqual(down, wantDown) {
			t.Errorf("blank-corner down: got %v, want %v", down, wantDown)
		}
	})
}

// TestOptions_ExpandDown_TwoCellBlock pins the ladder's middle rung against
// real Excel. With exactly two filled cells the endpoint is the second one:
// End("down") from the neighbor would sail past the block to row 1048576 and
// drag in a million blanks.
func TestOptions_ExpandDown_TwoCellBlock(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("first")
		sheet.Range("A2").SetValue("second")
		// A3 onwards empty, nothing else on the sheet.

		v, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Expand=down on a 2-cell block: %v", err)
		}
		want := []interface{}{"first", "second"}
		if !reflect.DeepEqual(v, want) {
			t.Errorf("2-cell block: got %T %v, want %v", v, v, want)
		}
	})
}

// TestOptions_ExpandRight_SingleCell is the row analogue of the single-cell
// down guard.
func TestOptions_ExpandRight_SingleCell(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("solo")
		// B1 empty.
		v, err := sheet.Range("A1").Options(excel.Expand("right")).Value()
		if err != nil {
			t.Fatalf("Expand=right single cell: %v", err)
		}
		if v != "solo" {
			t.Errorf("blank-neighbor guard: got %T %v, want scalar \"solo\"", v, v)
		}
	})
}

// TestOptions_ExpandTable_SingleCell confirms the table guard fires on both
// dimensions: an isolated cell expands to just itself.
func TestOptions_ExpandTable_SingleCell(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("B2").SetValue("solo")
		// All four neighbors empty.
		v, err := sheet.Range("B2").Options(excel.Expand("table")).Value()
		if err != nil {
			t.Fatalf("Expand=table single cell: %v", err)
		}
		if v != "solo" {
			t.Errorf("table guard: got %T %v, want scalar \"solo\"", v, v)
		}
	})
}

// TestOptions_ExpandDown_MultiCellAnchor is the live-Excel regression for the
// cross-axis collapse: expanding a 1x3 anchor down must keep all three columns.
// Before the fix the rectangle was built from two addresses that both sat in
// column A, so the read returned a 10x1 column and B/C were silently dropped
// (with err == nil, so a caller could not detect the truncation).
func TestOptions_ExpandDown_MultiCellAnchor(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		block := make([][]interface{}, 10)
		for r := range block {
			block[r] = []interface{}{float64(r*10 + 1), float64(r*10 + 2), float64(r*10 + 3)}
		}
		if err := sheet.Range("A1", "C10").SetValue(block).Err(); err != nil {
			t.Fatalf("seed block: %v", err)
		}
		// A11 deliberately empty — the block ends at row 10.

		v, err := sheet.Range("A1", "C1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Options(Expand=down) on A1:C1: %v", err)
		}
		got, ok := v.([][]interface{})
		if !ok {
			t.Fatalf("expected [][]interface{} for a 10x3 expand, got %T (%v) — "+
				"a 1-D result means the anchor's columns were collapsed", v, v)
		}
		if !reflect.DeepEqual(got, block) {
			t.Errorf("got %v, want %v", got, block)
		}
	})
}

// TestOptions_ExpandRight_MultiCellAnchor is the row-wise mirror: a 3x1 anchor
// grown right must keep all three rows.
func TestOptions_ExpandRight_MultiCellAnchor(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		block := make([][]interface{}, 3)
		for r := range block {
			block[r] = make([]interface{}, 6)
			for c := range block[r] {
				block[r][c] = float64((r+1)*100 + c + 1)
			}
		}
		if err := sheet.Range("A1", "F3").SetValue(block).Err(); err != nil {
			t.Fatalf("seed block: %v", err)
		}

		v, err := sheet.Range("A1", "A3").Options(excel.Expand("right")).Value()
		if err != nil {
			t.Fatalf("Options(Expand=right) on A1:A3: %v", err)
		}
		got, ok := v.([][]interface{})
		if !ok {
			t.Fatalf("expected [][]interface{} for a 3x6 expand, got %T (%v) — "+
				"a 1-D result means the anchor's rows were collapsed", v, v)
		}
		if !reflect.DeepEqual(got, block) {
			t.Errorf("got %v, want %v", got, block)
		}
	})
}

// TestOptions_ExpandMultiCellMatchesTable cross-checks the fixed "down"/"right"
// branches against the "table" branch that was always two-dimensional: on a
// rectangular block anchored at its top-left corner all three directions must
// read the identical grid.
func TestOptions_ExpandMultiCellMatchesTable(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		block := [][]interface{}{
			{1.0, 2.0, 3.0},
			{4.0, 5.0, 6.0},
			{7.0, 8.0, 9.0},
		}
		if err := sheet.Range("A1", "C3").SetValue(block).Err(); err != nil {
			t.Fatalf("seed block: %v", err)
		}

		table, err := sheet.Range("A1").Options(excel.Expand("table")).Value()
		if err != nil {
			t.Fatalf("Expand=table: %v", err)
		}
		down, err := sheet.Range("A1", "C1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Expand=down: %v", err)
		}
		right, err := sheet.Range("A1", "A3").Options(excel.Expand("right")).Value()
		if err != nil {
			t.Fatalf("Expand=right: %v", err)
		}
		if !reflect.DeepEqual(down, table) {
			t.Errorf("down %v != table %v", down, table)
		}
		if !reflect.DeepEqual(right, table) {
			t.Errorf("right %v != table %v", right, table)
		}
		if !reflect.DeepEqual(table, block) {
			t.Errorf("table %v != seeded block %v", table, block)
		}
	})
}

// TestOptions_ExpandDownHeaderStructDecode exercises the most typical xlwings
// idiom on a multi-cell anchor: the header row is the anchor, Expand("down")
// grows it over the records. With the cross-axis collapse this returned structs
// whose 2nd and 3rd fields were all zero.
func TestOptions_ExpandDownHeaderStructDecode(t *testing.T) {
	type Row struct {
		Name string
		Age  float64
		City string
	}
	withSheet(t, func(sheet excel.Worksheet) {
		seed := [][]interface{}{
			{"Name", "Age", "City"},
			{"alice", 30.0, "seoul"},
			{"bob", 25.0, "busan"},
		}
		if err := sheet.Range("A1", "C3").SetValue(seed).Err(); err != nil {
			t.Fatalf("seed block: %v", err)
		}

		var rows []Row
		err := sheet.Range("A1", "C1").
			Options(excel.Header(true), excel.Expand("down")).
			Get(&rows)
		if err != nil {
			t.Fatalf("Header+Expand(down): %v", err)
		}
		want := []Row{
			{Name: "alice", Age: 30, City: "seoul"},
			{Name: "bob", Age: 25, City: "busan"},
		}
		if !reflect.DeepEqual(rows, want) {
			t.Errorf("got %+v, want %+v", rows, want)
		}
	})
}

// TestOptions_StructDecode is the end-to-end happy path for the struct-by-
// header decode driven directly from real Excel data.
func TestOptions_StructDecode(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("Name")
		sheet.Range("B1").SetValue("Age")
		sheet.Range("A2").SetValue("alice")
		sheet.Range("B2").SetValue(30.0)
		sheet.Range("A3").SetValue("bob")
		sheet.Range("B3").SetValue(25.0)

		var rows []Row
		err := sheet.Range("A1", "B3").Options(excel.Header(true)).Get(&rows)
		if err != nil {
			t.Fatalf("Get(&rows): %v", err)
		}
		want := []Row{{Name: "alice", Age: 30}, {Name: "bob", Age: 25}}
		if !reflect.DeepEqual(rows, want) {
			t.Errorf("got %+v, want %+v", rows, want)
		}
	})
}

// TestOptions_ExpandTableAndStruct chains Expand + Header(true) — the
// typical xlwings idiom `rng.options(MyType, header=1, expand="table")`.
func TestOptions_ExpandTableAndStruct(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("Name")
		sheet.Range("B1").SetValue("Age")
		sheet.Range("A2").SetValue("alice")
		sheet.Range("B2").SetValue(30.0)
		sheet.Range("A3").SetValue("bob")
		sheet.Range("B3").SetValue(25.0)
		// A4 empty -> expand=table stops here.

		var rows []Row
		err := sheet.Range("A1").Options(
			excel.Expand("table"),
			excel.Header(true),
		).Get(&rows)
		if err != nil {
			t.Fatalf("Expand+Header: %v", err)
		}
		want := []Row{{Name: "alice", Age: 30}, {Name: "bob", Age: 25}}
		if !reflect.DeepEqual(rows, want) {
			t.Errorf("got %+v, want %+v", rows, want)
		}
	})
}

// TestOptions_ScalarForcing exercises the .options(Scalar()) shape-forcing
// path against real Excel: a 1x1 read should hand back the bare scalar.
func TestOptions_ScalarForcing(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(123.5)
		v, err := sheet.Range("A1").Options(excel.Scalar()).Value()
		if err != nil {
			t.Fatalf("Scalar.Value: %v", err)
		}
		if v != 123.5 {
			t.Errorf("got %v (%T), want 123.5", v, v)
		}
	})
}

// TestOptions_SetGetStructRoundTrip writes a struct slice with a header row
// and reads it back through the header decode — the full write/read mirror.
func TestOptions_SetGetStructRoundTrip(t *testing.T) {
	type Person struct {
		Name string
		Age  float64
	}
	withSheet(t, func(sheet excel.Worksheet) {
		in := []Person{{"alice", 30}, {"bob", 25}}

		if err := sheet.Range("A1").Options(excel.Header(true)).Set(in); err != nil {
			t.Fatalf("Set: %v", err)
		}

		// The block now spans A1:B3 (header + 2 rows). Expand("table") from
		// the anchor must rediscover it.
		var out []Person
		err := sheet.Range("A1").
			Options(excel.Expand("table"), excel.Header(true)).
			Get(&out)
		if err != nil {
			t.Fatalf("Get: %v", err)
		}
		if !reflect.DeepEqual(out, in) {
			t.Errorf("round trip: got %+v, want %+v", out, in)
		}
	})
}

// TestOptions_ExpandDeferredReevaluation is the regression for the eager-Expand
// bug: xlwings evaluates options only on value access, so a stored
// OptionedRange must re-discover the current data block on every read. We
// capture an OptionedRange over a 3-row column, then grow the column, then read
// again — the second read must include the new rows. Before the fix, Options()
// snapshotted the expanded address at construction time and the second read
// returned only the original 3 rows.
func TestOptions_ExpandDeferredReevaluation(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(1.0)
		sheet.Range("A2").SetValue(2.0)
		sheet.Range("A3").SetValue(3.0)

		// Capture the OptionedRange BEFORE growing the data.
		opt := sheet.Range("A1").Options(excel.Expand("down"))

		// First read: 3 rows.
		v, err := opt.Value()
		if err != nil {
			t.Fatalf("first Value(): %v", err)
		}
		if got, want := v, []interface{}{1.0, 2.0, 3.0}; !reflect.DeepEqual(got, want) {
			t.Fatalf("first read: got %v, want %v", got, want)
		}

		// Grow the block.
		sheet.Range("A4").SetValue(4.0)
		sheet.Range("A5").SetValue(5.0)

		// Second read on the SAME OptionedRange must re-evaluate the expand and
		// include the new rows.
		v, err = opt.Value()
		if err != nil {
			t.Fatalf("second Value(): %v", err)
		}
		want := []interface{}{1.0, 2.0, 3.0, 4.0, 5.0}
		if !reflect.DeepEqual(v, want) {
			t.Errorf("second read did not re-evaluate expand: got %v, want %v", v, want)
		}
	})
}

// TestOptions_EmptyReplacement validates that Empty(value) substitutes nil
// cells in the raw read.
func TestOptions_EmptyReplacement(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("x")
		// B1 left blank.
		sheet.Range("C1").SetValue("z")

		v, err := sheet.Range("A1", "C1").
			Options(excel.Empty("N/A"), excel.Vector()).
			Value()
		if err != nil {
			t.Fatalf("Empty+Vector: %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{}, got %T", v)
		}
		want := []interface{}{"x", "N/A", "z"}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}
