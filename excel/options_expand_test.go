//go:build windows

// Excel-free regression tests for the Expand address arithmetic.
//
// The COM round trips Expand performs (Cells / Address / Rows / Columns /
// Offset / End / Worksheet.Range) are all driven through sugar.Chain, so a
// recording fake chain backed by an in-memory grid can exercise the *whole*
// expansion — including which rectangle is finally requested from
// Worksheet.Range — with no Excel installed. The live-Excel counterparts live
// in options_integration_test.go behind the excel_integration tag.
//
// The bug these pin: "down"/"right" used to build their rectangle from two
// addresses that both sat in the anchor's first column (down) or first row
// (right), collapsing the perpendicular axis of a multi-cell anchor to 1 and
// silently truncating the read.

package excel

import (
	"fmt"
	"strconv"
	"strings"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// fakeGrid is a minimal in-memory worksheet: a sparse set of non-empty cells
// plus a log of every Worksheet.Range(cell1, cell2) request the expansion made.
type fakeGrid struct {
	cells map[[2]int]interface{}
	asked [][2]string
}

func newFakeGrid() *fakeGrid {
	return &fakeGrid{cells: map[[2]int]interface{}{}}
}

// fill seeds a rows x cols block anchored at (row, col) with non-empty values.
func (g *fakeGrid) fill(row, col, rows, cols int) *fakeGrid {
	for r := row; r < row+rows; r++ {
		for c := col; c < col+cols; c++ {
			g.cells[[2]int{r, c}] = float64(r*100 + c)
		}
	}
	return g
}

// rangeAt returns a Range wrapper over the given rectangle. It is the entry
// point the tests hand to applyExpand / expandFromEnd.
func (g *fakeGrid) rangeAt(row, col, rows, cols int) Range {
	return wrapRange(&fakeCell{g: g, kind: "range", row: row, col: col, rows: rows, cols: cols})
}

func (g *fakeGrid) lastAsked(t *testing.T) [2]string {
	t.Helper()
	if len(g.asked) == 0 {
		t.Fatal("no Worksheet.Range(cell1, cell2) request was recorded")
	}
	return g.asked[len(g.asked)-1]
}

// fakeCell is a node in the faked COM object graph: a range rectangle, the
// parent sheet, a Rows/Columns collection (carrying only its Count), or a
// terminal value.
type fakeCell struct {
	g    *fakeGrid
	kind string // "range" | "sheet" | "count" | "value"

	row, col, rows, cols int         // kind == "range"
	n                    int32       // kind == "count"
	val                  interface{} // kind == "value"
}

func (f *fakeCell) value(v interface{}) sugar.Chain {
	return &fakeCell{g: f.g, kind: "value", val: v}
}

func (f *fakeCell) cellAt(row, col int) sugar.Chain {
	return &fakeCell{g: f.g, kind: "range", row: row, col: col, rows: 1, cols: 1}
}

// Get emulates the handful of Excel members the Expand path touches. Anything
// else becomes an error chain so an unexpected COM call fails the test loudly
// instead of silently returning a zero value.
func (f *fakeCell) Get(prop string, params ...interface{}) sugar.Chain {
	switch prop {
	case "Cells":
		if len(params) != 2 {
			return sugar.Error(fmt.Errorf("fake: Cells wants 2 args, got %d", len(params)))
		}
		return f.cellAt(f.row+fakeArgInt(params[0])-1, f.col+fakeArgInt(params[1])-1)
	case "Offset":
		if len(params) != 2 {
			return sugar.Error(fmt.Errorf("fake: Offset wants 2 args, got %d", len(params)))
		}
		return &fakeCell{
			g:    f.g,
			kind: "range",
			row:  f.row + fakeArgInt(params[0]),
			col:  f.col + fakeArgInt(params[1]),
			rows: f.rows,
			cols: f.cols,
		}
	case "Address":
		return f.value(f.address())
	case "Rows":
		return &fakeCell{g: f.g, kind: "count", n: int32(f.rows)}
	case "Columns":
		return &fakeCell{g: f.g, kind: "count", n: int32(f.cols)}
	case "Count":
		if f.kind == "count" {
			return f.value(f.n)
		}
		return f.value(int32(f.rows * f.cols))
	case "End":
		if len(params) != 1 {
			return sugar.Error(fmt.Errorf("fake: End wants 1 arg, got %d", len(params)))
		}
		return f.end(int32(fakeArgInt(params[0])))
	case "Value":
		if f.rows == 1 && f.cols == 1 {
			return f.value(f.g.cells[[2]int{f.row, f.col}])
		}
		return f.value(f.block())
	case "Worksheet":
		return &fakeCell{g: f.g, kind: "sheet"}
	case "Range":
		if len(params) != 2 {
			return sugar.Error(fmt.Errorf("fake: Range wants 2 args, got %d", len(params)))
		}
		a, aok := params[0].(string)
		b, bok := params[1].(string)
		if !aok || !bok {
			return sugar.Error(fmt.Errorf("fake: Range wants string corners, got %T and %T", params[0], params[1]))
		}
		f.g.asked = append(f.g.asked, [2]string{a, b})
		r1, c1, err := parseFakeA1(a)
		if err != nil {
			return sugar.Error(err)
		}
		r2, c2, err := parseFakeA1(b)
		if err != nil {
			return sugar.Error(err)
		}
		return &fakeCell{
			g:    f.g,
			kind: "range",
			row:  min(r1, r2),
			col:  min(c1, c2),
			rows: abs(r2-r1) + 1,
			cols: abs(c2-c1) + 1,
		}
	}
	return sugar.Error(fmt.Errorf("fake: unexpected Get(%q)", prop))
}

// end walks the contiguous run of non-empty cells, mirroring Ctrl+Arrow from a
// non-empty start cell (the only case the blank-neighbor guard lets through).
func (f *fakeCell) end(dir int32) sugar.Chain {
	var dr, dc int
	switch dir {
	case xlDown:
		dr = 1
	case xlUp:
		dr = -1
	case xlToRight:
		dc = 1
	case xlToLeft:
		dc = -1
	default:
		return sugar.Error(fmt.Errorf("fake: unknown End direction %d", dir))
	}
	r, c := f.row, f.col
	for {
		nr, nc := r+dr, c+dc
		if nr < 1 || nc < 1 {
			break
		}
		if _, ok := f.g.cells[[2]int{nr, nc}]; !ok {
			break
		}
		r, c = nr, nc
	}
	return f.cellAt(r, c)
}

func (f *fakeCell) block() [][]interface{} {
	out := make([][]interface{}, f.rows)
	for r := 0; r < f.rows; r++ {
		out[r] = make([]interface{}, f.cols)
		for c := 0; c < f.cols; c++ {
			out[r][c] = f.g.cells[[2]int{f.row + r, f.col + c}]
		}
	}
	return out
}

// address renders the A1 form Excel's Address property returns: "$A$1" for a
// single cell, "$A$1:$C$10" for a rectangle.
func (f *fakeCell) address() string {
	tl := "$" + fakeColName(f.col) + "$" + strconv.Itoa(f.row)
	if f.rows == 1 && f.cols == 1 {
		return tl
	}
	return tl + ":$" + fakeColName(f.col+f.cols-1) + "$" + strconv.Itoa(f.row+f.rows-1)
}

func (f *fakeCell) Call(method string, params ...interface{}) sugar.Chain {
	return sugar.Error(fmt.Errorf("fake: unexpected Call(%q)", method))
}
func (f *fakeCell) Put(prop string, params ...interface{}) sugar.Chain {
	return sugar.Error(fmt.Errorf("fake: unexpected Put(%q)", prop))
}
func (f *fakeCell) ForEach(cb func(item sugar.Chain) error) sugar.Chain { return f }
func (f *fakeCell) Fork() sugar.Chain                                   { return f }
func (f *fakeCell) Store() (*ole.IDispatch, error)                      { return nil, nil }
func (f *fakeCell) Release() error                                      { return nil }
func (f *fakeCell) IsDispatch() bool                                    { return f.kind != "value" }
func (f *fakeCell) Value() (interface{}, error) {
	if f.kind != "value" {
		return nil, fmt.Errorf("fake: Value() on a %s node", f.kind)
	}
	return f.val, nil
}
func (f *fakeCell) Err() error { return nil }

// fakeArgInt narrows the int/int32 arguments the typed Range wrappers pass.
func fakeArgInt(v interface{}) int {
	switch x := v.(type) {
	case int:
		return x
	case int32:
		return int(x)
	case int64:
		return int(x)
	}
	return 0
}

func fakeColName(col int) string {
	name := ""
	for col > 0 {
		col--
		name = string(rune('A'+col%26)) + name
		col /= 26
	}
	return name
}

func parseFakeA1(addr string) (row, col int, err error) {
	a := strings.ReplaceAll(addr, "$", "")
	if i := strings.Index(a, ":"); i >= 0 {
		a = a[:i]
	}
	i := 0
	for i < len(a) && a[i] >= 'A' && a[i] <= 'Z' {
		col = col*26 + int(a[i]-'A') + 1
		i++
	}
	if i == 0 || i == len(a) {
		return 0, 0, fmt.Errorf("fake: cannot parse address %q", addr)
	}
	row, err = strconv.Atoi(a[i:])
	if err != nil {
		return 0, 0, fmt.Errorf("fake: cannot parse address %q: %w", addr, err)
	}
	return row, col, nil
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func abs(a int) int {
	if a < 0 {
		return -a
	}
	return a
}

// expandedAddress runs applyExpand and returns the address of the resulting
// range plus the two corners handed to Worksheet.Range.
func expandedAddress(t *testing.T, g *fakeGrid, anchor Range, direction string) (string, [2]string) {
	t.Helper()
	got, err := applyExpand(anchor, direction)
	if err != nil {
		t.Fatalf("applyExpand(%q): %v", direction, err)
	}
	addr, err := got.Address()
	if err != nil {
		t.Fatalf("expanded Address(): %v", err)
	}
	return addr, g.lastAsked(t)
}

// TestExpand_MultiCellAnchorKeepsCrossAxis is the regression for the silent
// truncation: a 1x3 anchor expanded "down" must keep all three columns, and a
// 3x1 anchor expanded "right" must keep all three rows. Before the fix both
// corners came from the anchor's origin column/row, so the result collapsed to
// a single column ("$A$1:$A$10") / row.
func TestExpand_MultiCellAnchorKeepsCrossAxis(t *testing.T) {
	t.Run("down keeps columns", func(t *testing.T) {
		g := newFakeGrid().fill(1, 1, 10, 3) // A1:C10 populated
		addr, corners := expandedAddress(t, g, g.rangeAt(1, 1, 1, 3), "down")
		if want := "$A$1:$C$10"; addr != want {
			t.Errorf("Expand(down) on A1:C1: got %s, want %s", addr, want)
		}
		if want := [2]string{"$A$1", "$C$10"}; corners != want {
			t.Errorf("Worksheet.Range corners: got %v, want %v", corners, want)
		}
	})

	t.Run("right keeps rows", func(t *testing.T) {
		g := newFakeGrid().fill(1, 1, 3, 6) // A1:F3 populated
		addr, corners := expandedAddress(t, g, g.rangeAt(1, 1, 3, 1), "right")
		if want := "$A$1:$F$3"; addr != want {
			t.Errorf("Expand(right) on A1:A3: got %s, want %s", addr, want)
		}
		if want := [2]string{"$A$1", "$F$3"}; corners != want {
			t.Errorf("Worksheet.Range corners: got %v, want %v", corners, want)
		}
	})

	t.Run("down off a non-A1 anchor", func(t *testing.T) {
		g := newFakeGrid().fill(2, 2, 4, 2) // B2:C5 populated
		addr, corners := expandedAddress(t, g, g.rangeAt(2, 2, 1, 2), "down")
		if want := "$B$2:$C$5"; addr != want {
			t.Errorf("Expand(down) on B2:C2: got %s, want %s", addr, want)
		}
		if want := [2]string{"$B$2", "$C$5"}; corners != want {
			t.Errorf("Worksheet.Range corners: got %v, want %v", corners, want)
		}
	})
}

// TestExpand_DownMatchesTableOnRectangularBlock is the semantic cross-check
// against the branch that was already correct: for a fully rectangular block
// anchored at its top-left header row, "down" and "table" must describe the
// same rectangle. ("table" derives both extents from the block itself, so it
// never depended on the anchor span.)
//
// Only the resulting rectangle is compared, not the corner pair handed to
// Worksheet.Range: "down" passes top-left + bottom-right while "table" passes
// bottom-left + top-right. Both are opposite corners of the same box — the
// bounding rectangle is what Worksheet.Range(cell1, cell2) resolves to.
func TestExpand_DownMatchesTableOnRectangularBlock(t *testing.T) {
	g := newFakeGrid().fill(1, 1, 10, 3)

	downAddr, _ := expandedAddress(t, g, g.rangeAt(1, 1, 1, 3), "down")
	tableAddr, _ := expandedAddress(t, g, g.rangeAt(1, 1, 1, 3), "table")

	if downAddr != tableAddr {
		t.Errorf("down vs table rectangle: %s != %s", downAddr, tableAddr)
	}
	if want := "$A$1:$C$10"; downAddr != want {
		t.Errorf("both should cover the whole block: got %s, want %s", downAddr, want)
	}
}

// TestExpand_RightMatchesTableOnRectangularBlock is the row-wise mirror: a
// 3x1 anchor grown right over a rectangular block equals the table expansion.
func TestExpand_RightMatchesTableOnRectangularBlock(t *testing.T) {
	g := newFakeGrid().fill(1, 1, 3, 6)

	rightAddr, _ := expandedAddress(t, g, g.rangeAt(1, 1, 3, 1), "right")
	tableAddr, _ := expandedAddress(t, g, g.rangeAt(1, 1, 3, 1), "table")

	if rightAddr != tableAddr {
		t.Errorf("right vs table rectangle: %s != %s", rightAddr, tableAddr)
	}
	if want := "$A$1:$F$3"; rightAddr != want {
		t.Errorf("both should cover the whole block: got %s, want %s", rightAddr, want)
	}
}

// TestExpand_SingleCellAnchorUnchanged guards the pre-existing behavior of the
// common 1x1 anchor: the cross-axis span is 1, so no widening happens.
func TestExpand_SingleCellAnchorUnchanged(t *testing.T) {
	t.Run("down", func(t *testing.T) {
		g := newFakeGrid().fill(1, 1, 10, 3)
		addr, corners := expandedAddress(t, g, g.rangeAt(1, 1, 1, 1), "down")
		if want := "$A$1:$A$10"; addr != want {
			t.Errorf("Expand(down) on A1: got %s, want %s", addr, want)
		}
		if want := [2]string{"$A$1", "$A$10"}; corners != want {
			t.Errorf("corners: got %v, want %v", corners, want)
		}
	})
	t.Run("right", func(t *testing.T) {
		g := newFakeGrid().fill(1, 1, 3, 6)
		addr, _ := expandedAddress(t, g, g.rangeAt(1, 1, 1, 1), "right")
		if want := "$A$1:$F$1"; addr != want {
			t.Errorf("Expand(right) on A1: got %s, want %s", addr, want)
		}
	})
}

// TestExpand_BlankNeighborKeepsMultiCellAnchor pins the blank-neighbor guard
// for a multi-cell anchor: with nothing below the header row, Expand("down")
// must hand back the anchor rectangle itself (1x3), not its first cell.
func TestExpand_BlankNeighborKeepsMultiCellAnchor(t *testing.T) {
	g := newFakeGrid().fill(1, 1, 1, 3) // only A1:C1 populated
	addr, corners := expandedAddress(t, g, g.rangeAt(1, 1, 1, 3), "down")
	if want := "$A$1:$C$1"; addr != want {
		t.Errorf("Expand(down) with blank neighbor: got %s, want %s", addr, want)
	}
	if want := [2]string{"$A$1", "$C$1"}; corners != want {
		t.Errorf("corners: got %v, want %v", corners, want)
	}
}

// TestExpandCornerOffset is the pure-arithmetic unit test for the corner
// derivation: growing down widens the endpoint cell by the anchor's column
// span, growing right deepens it by the row span, and a 1-wide cross axis
// needs no shift at all.
func TestExpandCornerOffset(t *testing.T) {
	cases := []struct {
		name      string
		direction int32
		crossSpan int
		wantRow   int
		wantCol   int
	}{
		{"down, 3 columns", xlDown, 3, 0, 2},
		{"down, 1 column", xlDown, 1, 0, 0},
		{"down, zero span", xlDown, 0, 0, 0},
		{"right, 4 rows", xlToRight, 4, 3, 0},
		{"right, 1 row", xlToRight, 1, 0, 0},
		{"unknown direction", xlUp, 5, 0, 0},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			gotRow, gotCol := expandCornerOffset(c.direction, c.crossSpan)
			if gotRow != c.wantRow || gotCol != c.wantCol {
				t.Errorf("expandCornerOffset(%d, %d) = (%d, %d), want (%d, %d)",
					c.direction, c.crossSpan, gotRow, gotCol, c.wantRow, c.wantCol)
			}
		})
	}
}

// TestCrossSpan reads the anchor extent perpendicular to the growth direction
// off the faked COM chain: column count for "down", row count for "right".
func TestCrossSpan(t *testing.T) {
	g := newFakeGrid()
	anchor := g.rangeAt(2, 3, 4, 5) // C2:G5 -> 4 rows x 5 cols

	if got, err := crossSpan(anchor, xlDown); err != nil || got != 5 {
		t.Errorf("crossSpan(down) = (%d, %v), want (5, nil)", got, err)
	}
	if got, err := crossSpan(anchor, xlToRight); err != nil || got != 4 {
		t.Errorf("crossSpan(right) = (%d, %v), want (4, nil)", got, err)
	}
	if got, err := crossSpan(anchor, xlUp); err != nil || got != 1 {
		t.Errorf("crossSpan(up) = (%d, %v), want (1, nil) for an unused direction", got, err)
	}
}
