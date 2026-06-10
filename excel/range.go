//go:build windows

package excel

import (
	"fmt"
	"strings"

	"github.com/xll-gen/sugar"
)

// XlInsertShiftDirection values for Range.Insert.
const (
	xlShiftDown    int32 = -4121
	xlShiftToRight int32 = -4161
)

// Range is a cell, row, column, or selection of cells — the Go equivalent of
// xlwings' `Range`.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/range.html
type Range interface {
	sugar.Chain
	// Value reads the range's value. For a single-cell range it returns the
	// underlying scalar (bool, float64, string, time.Time, …). For a
	// multi-cell range it returns `[][]interface{}` shaped `[row][col]`,
	// mirroring xlwings' `range.value` semantics.
	Value() (interface{}, error)
	// SetValue sets the value for the entire range. Pass a scalar to fill
	// every cell; pass a `[][]interface{}` to write a 2-D block (the slice
	// must match the range's shape).
	SetValue(value interface{}) Range
	// Address returns the cell address as a string (e.g. "$A$1:$B$2").
	Address() (string, error)
	// Formula returns the A1-style formula text.
	Formula() (string, error)
	// SetFormula sets an A1-style formula.
	SetFormula(formula string) Range
	// Formula2 returns the modern dynamic-array formula text (Excel 365+).
	Formula2() (string, error)
	// SetFormula2 sets a dynamic-array formula. Prefer this over SetFormula
	// for code that targets Excel 365's array-spill behavior.
	SetFormula2(formula string) Range
	// NumberFormat returns the Excel number format string (e.g. "0.00").
	NumberFormat() (string, error)
	// SetNumberFormat applies an Excel number format string.
	SetNumberFormat(fmt string) Range

	// Cells returns a single-cell Range relative to this range, 1-based.
	Cells(row, col interface{}) Range
	// Offset shifts the anchor by (rowOffset, colOffset).
	Offset(rowOffset, colOffset int) Range
	// Resize returns a new range of the given shape anchored at this range's
	// top-left cell.
	Resize(rows, cols int) Range
	// Rows returns the rows collection of this range.
	Rows() Range
	// Columns returns the columns collection of this range.
	Columns() Range

	// End returns the cell at the end of the contiguous data region in the
	// given direction ("up", "down", "left", "right") — Ctrl+Arrow in the
	// UI. Equivalent to xlwings' `range.end(direction)`.
	End(direction string) Range

	// Row returns the 1-based index of the range's first row.
	Row() (int32, error)
	// Column returns the 1-based index of the range's first column.
	Column() (int32, error)
	// Count returns the number of cells in the range (rows × cols).
	Count() (int32, error)
	// Width returns the range's total width in points.
	Width() (float64, error)
	// Height returns the range's total height in points.
	Height() (float64, error)
	// ColumnWidth returns the column width (Excel character units).
	ColumnWidth() (float64, error)
	// SetColumnWidth sets the column width for all columns in the range.
	SetColumnWidth(w float64) Range
	// RowHeight returns the row height in points.
	RowHeight() (float64, error)
	// SetRowHeight sets the row height for all rows in the range.
	SetRowHeight(h float64) Range

	// Color returns the cell background color (Interior.Color) as an OLE
	// color integer (see excel.RGB). Equivalent to xlwings' `range.color`.
	Color() (int32, error)
	// SetColor fills the range background. Build values with excel.RGB.
	SetColor(color int32) Range

	// Clear clears values, formulas, and formatting.
	Clear() error
	// ClearContents clears values and formulas but preserves formatting.
	ClearContents() error
	// Delete removes the cells; remaining cells shift up (Excel default).
	Delete() error
	// Copy copies the range to the clipboard (no destination argument).
	Copy() error
	// Insert inserts cells, shifting existing ones away. shift is "down",
	// "right", or "" (Excel picks based on the range shape). Equivalent to
	// xlwings' `range.insert(shift=...)`.
	Insert(shift string) error
	// Find searches the range for a value (Excel's Ctrl+F semantics, match
	// on any part of the cell). found is false when nothing matches —
	// Excel's COM Find returns Nothing in that case, not an error.
	Find(what string) (cell Range, found bool, err error)

	// Merge merges all cells in the range into one.
	Merge() error
	// UnMerge undoes a Merge.
	UnMerge() error
	// MergeCells reports whether all cells in the range are merged.
	MergeCells() (bool, error)

	// AutoFit auto-fits the column width (or row height) for this range.
	AutoFit() error

	// Font returns the character-formatting object for this range.
	// Equivalent to xlwings' `range.font`.
	Font() Font

	// Options is the Go equivalent of xlwings' `Range.options(...)`. It
	// returns an OptionedRange that decodes the range on .Value() / .Get()
	// with the supplied conversion knobs (Scalar/Vector/Grid, Expand,
	// Header, Empty, DateFormat, Convert). See options.go for the full
	// option catalogue.
	Options(opts ...RangeOption) OptionedRange
}

type excelRange struct {
	sugar.Chain
}

// Value reads `Range.Value`. Single-cell ranges return the underlying scalar
// (bool, float64, string, time.Time, …). Multi-cell ranges return
// `[][]interface{}` shaped `[row][col]` — sugar.Chain.Value() handles the
// SAFEARRAY decode transparently.
func (r *excelRange) Value() (interface{}, error) {
	return r.Get("Value").Value()
}

func (r *excelRange) SetValue(value interface{}) Range {
	return &excelRange{r.Put("Value", value)}
}

func (r *excelRange) Address() (string, error) {
	v, err := r.Get("Address").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (r *excelRange) Formula() (string, error) {
	v, err := r.Get("Formula").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (r *excelRange) SetFormula(formula string) Range {
	return &excelRange{r.Put("Formula", formula)}
}

func (r *excelRange) Formula2() (string, error) {
	v, err := r.Get("Formula2").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (r *excelRange) SetFormula2(formula string) Range {
	return &excelRange{r.Put("Formula2", formula)}
}

func (r *excelRange) NumberFormat() (string, error) {
	v, err := r.Get("NumberFormat").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (r *excelRange) SetNumberFormat(fmt string) Range {
	return &excelRange{r.Put("NumberFormat", fmt)}
}

func (r *excelRange) Cells(row, col interface{}) Range {
	return &excelRange{r.Get("Cells", row, col)}
}

func (r *excelRange) Offset(rowOffset, colOffset int) Range {
	return &excelRange{r.Get("Offset", int32(rowOffset), int32(colOffset))}
}

func (r *excelRange) Resize(rows, cols int) Range {
	return &excelRange{r.Get("Resize", int32(rows), int32(cols))}
}

func (r *excelRange) Rows() Range {
	return &excelRange{r.Get("Rows")}
}

func (r *excelRange) Columns() Range {
	return &excelRange{r.Get("Columns")}
}

func (r *excelRange) End(direction string) Range {
	var dir int32
	switch strings.ToLower(direction) {
	case "down":
		dir = xlDown
	case "up":
		dir = xlUp
	case "left":
		dir = xlToLeft
	case "right":
		dir = xlToRight
	default:
		return &excelRange{sugar.Error(fmt.Errorf(
			"End: unsupported direction %q (use \"up\", \"down\", \"left\", or \"right\")", direction))}
	}
	return &excelRange{r.Get("End", dir)}
}

func (r *excelRange) Width() (float64, error)  { return shapeFloat(r, "Width") }
func (r *excelRange) Height() (float64, error) { return shapeFloat(r, "Height") }

func (r *excelRange) ColumnWidth() (float64, error) {
	return shapeFloat(r, "ColumnWidth")
}

func (r *excelRange) SetColumnWidth(w float64) Range {
	return &excelRange{r.Put("ColumnWidth", w)}
}

func (r *excelRange) RowHeight() (float64, error) {
	return shapeFloat(r, "RowHeight")
}

func (r *excelRange) SetRowHeight(h float64) Range {
	return &excelRange{r.Put("RowHeight", h)}
}

func (r *excelRange) Color() (int32, error) {
	v, err := r.Get("Interior").Get("Color").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

func (r *excelRange) SetColor(color int32) Range {
	inner := r.Get("Interior").Put("Color", color)
	if inner.Err() != nil {
		return &excelRange{inner}
	}
	return r
}

func (r *excelRange) Insert(shift string) error {
	switch strings.ToLower(shift) {
	case "":
		return r.Call("Insert").Err()
	case "down":
		return r.Call("Insert", xlShiftDown).Err()
	case "right":
		return r.Call("Insert", xlShiftToRight).Err()
	default:
		return fmt.Errorf("Insert: unsupported shift %q (use \"down\", \"right\", or \"\")", shift)
	}
}

func (r *excelRange) Find(what string) (Range, bool, error) {
	ch := r.Call("Find", what)
	if err := ch.Err(); err != nil {
		return nil, false, err
	}
	if !ch.IsDispatch() {
		// Excel's Find returns Nothing when there is no match.
		return nil, false, nil
	}
	return &excelRange{ch}, true, nil
}

func (r *excelRange) Row() (int32, error) {
	v, err := r.Get("Row").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

func (r *excelRange) Column() (int32, error) {
	v, err := r.Get("Column").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

func (r *excelRange) Count() (int32, error) {
	v, err := r.Get("Count").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

func (r *excelRange) Clear() error          { return r.Call("Clear").Err() }
func (r *excelRange) ClearContents() error  { return r.Call("ClearContents").Err() }
func (r *excelRange) Delete() error         { return r.Call("Delete").Err() }
func (r *excelRange) Copy() error           { return r.Call("Copy").Err() }
func (r *excelRange) Merge() error          { return r.Call("Merge").Err() }
func (r *excelRange) UnMerge() error        { return r.Call("UnMerge").Err() }

func (r *excelRange) MergeCells() (bool, error) {
	v, err := r.Get("MergeCells").Value()
	if err != nil {
		return false, err
	}
	b, _ := v.(bool)
	return b, nil
}

func (r *excelRange) AutoFit() error {
	return r.Call("EntireColumn").Call("AutoFit").Err()
}

func (r *excelRange) Font() Font {
	return &font{r.Get("Font")}
}
