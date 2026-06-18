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
	// SetFormula2Array sets a *block* of dynamic-array formulas in a single
	// COM round-trip — the batch counterpart of SetFormula2. It mirrors
	// xlwings' array assignment `range.formula2 = [["=A1", "=B1"], ...]`
	// (Go's static typing forces a separate method from the scalar setter
	// rather than an overloaded property). The argument is shaped exactly
	// like SetValue's block form: a `[][]interface{}` (or any 2-D slice such
	// as `[][]string`) matching the range's shape writes one formula string
	// per cell, and a 1-D `[]interface{}`/`[]string` writes a single row or
	// column. Each cell goes through the spill-correct Formula2 property in
	// one Put, so DA Excel never rewrites the per-cell UDF calls into the
	// implicit-intersection `=@Fn(...)` form. Use it to collapse a contiguous
	// formula column/row that would otherwise be N separate SetFormulaSpill
	// calls into one. Empty/nil cells in the block are written as blanks.
	SetFormula2Array(formulas interface{}) Range
	// SetFormulaSpill sets a formula using the dynamic-array-native COM
	// property (Formula2) when available, falling back to the legacy Formula
	// property on Excel versions that predate dynamic arrays (2016 and
	// earlier, where the Formula2 property does not exist on the COM Range).
	//
	// This is a sugar-specific convenience with no direct xlwings analogue
	// (xlwings exposes .formula and .formula2 separately). It exists because
	// writing a UDF call through the legacy Formula property on a
	// dynamic-array-aware Excel applies implicit intersection — the formula is
	// stored as `=@MyFunc(...)`, which suppresses the array spill. Formula2 is
	// the spill-correct property; SetFormulaSpill picks it automatically and
	// degrades gracefully on old Excel. Use it for any formula expected to
	// spill (a UDF returning an array, or a native dynamic-array function).
	SetFormulaSpill(formula string) Range
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
	// Unmerge undoes a Merge. xlwings `range.unmerge()`.
	Unmerge() error
	// MergeCells reports whether all cells in the range are merged.
	MergeCells() (bool, error)

	// AutoFit auto-fits both the column width and the row height of the entire
	// columns and rows intersecting this range, matching xlwings'
	// `range.autofit()`.
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

// wrapRange wraps a chain in the Range typed wrapper. It is the single
// construction point for the chain -> Range convention.
func wrapRange(c sugar.Chain) Range { return &excelRange{c} }

// Value reads `Range.Value`. Single-cell ranges return the underlying scalar
// (bool, float64, string, time.Time, …). Multi-cell ranges return
// `[][]interface{}` shaped `[row][col]` — sugar.Chain.Value() handles the
// SAFEARRAY decode transparently.
func (r *excelRange) Value() (interface{}, error) {
	return r.Get("Value").Value()
}

func (r *excelRange) SetValue(value interface{}) Range {
	return wrapRange(r.Put("Value", value))
}

func (r *excelRange) Address() (string, error) {
	return getString(r, "Address")
}

func (r *excelRange) Formula() (string, error) {
	return getString(r, "Formula")
}

func (r *excelRange) SetFormula(formula string) Range {
	return wrapRange(r.Put("Formula", formula))
}

func (r *excelRange) Formula2() (string, error) {
	return getString(r, "Formula2")
}

func (r *excelRange) SetFormula2(formula string) Range {
	return wrapRange(r.Put("Formula2", formula))
}

// SetFormula2Array writes a block of formulas through the Formula2 property in
// one COM Put. normalizeParams encodes the 2-D (or 1-D) slice into a
// VT_ARRAY|VT_VARIANT SAFEARRAY of formula strings, exactly as it does for a
// SetValue block — so the whole column/row of UDF calls lands array-native
// (no implicit-intersection `=@Fn(...)` rewrite) in a single round-trip.
func (r *excelRange) SetFormula2Array(formulas interface{}) Range {
	return wrapRange(r.Put("Formula2", formulas))
}

// SetFormulaSpill writes via the Formula2 property (dynamic-array native) and,
// only if that COM Put fails — which is how the missing Formula2 property
// surfaces on pre-dynamic-array Excel (2016 and earlier) — retries via the
// legacy Formula property. On modern Excel the first Put succeeds and Formula
// is never touched; on old Excel the fallback keeps the formula working
// (without spill, which old Excel cannot do anyway). The returned Range
// carries the error of whichever path was taken last, so callers can chain
// .Err() exactly as with SetFormula/SetFormula2.
func (r *excelRange) SetFormulaSpill(formula string) Range {
	c := r.Put("Formula2", formula)
	if c.Err() != nil {
		// Formula2 unavailable (pre-DA Excel) or rejected — fall back to the
		// legacy property so the cell still gets the formula.
		return wrapRange(r.Put("Formula", formula))
	}
	return wrapRange(c)
}

func (r *excelRange) NumberFormat() (string, error) {
	return getString(r, "NumberFormat")
}

func (r *excelRange) SetNumberFormat(fmt string) Range {
	return wrapRange(r.Put("NumberFormat", fmt))
}

func (r *excelRange) Cells(row, col interface{}) Range {
	return wrapRange(r.Get("Cells", row, col))
}

func (r *excelRange) Offset(rowOffset, colOffset int) Range {
	return wrapRange(r.Get("Offset", int32(rowOffset), int32(colOffset)))
}

func (r *excelRange) Resize(rows, cols int) Range {
	return wrapRange(r.Get("Resize", int32(rows), int32(cols)))
}

func (r *excelRange) Rows() Range {
	return wrapRange(r.Get("Rows"))
}

func (r *excelRange) Columns() Range {
	return wrapRange(r.Get("Columns"))
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
		return wrapRange(sugar.Error(fmt.Errorf(
			"End: unsupported direction %q (use \"up\", \"down\", \"left\", or \"right\")", direction)))
	}
	return wrapRange(r.Get("End", dir))
}

func (r *excelRange) Width() (float64, error)  { return getFloat64(r, "Width") }
func (r *excelRange) Height() (float64, error) { return getFloat64(r, "Height") }

func (r *excelRange) ColumnWidth() (float64, error) {
	return getFloat64(r, "ColumnWidth")
}

func (r *excelRange) SetColumnWidth(w float64) Range {
	return wrapRange(r.Put("ColumnWidth", w))
}

func (r *excelRange) RowHeight() (float64, error) {
	return getFloat64(r, "RowHeight")
}

func (r *excelRange) SetRowHeight(h float64) Range {
	return wrapRange(r.Put("RowHeight", h))
}

func (r *excelRange) Color() (int32, error) {
	return getInt32(r.Get("Interior"), "Color")
}

func (r *excelRange) SetColor(color int32) Range {
	inner := r.Get("Interior").Put("Color", color)
	if inner.Err() != nil {
		return wrapRange(inner)
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
	return wrapRange(ch), true, nil
}

func (r *excelRange) Row() (int32, error) {
	return getInt32(r, "Row")
}

func (r *excelRange) Column() (int32, error) {
	return getInt32(r, "Column")
}

func (r *excelRange) Count() (int32, error) {
	return getInt32(r, "Count")
}

func (r *excelRange) Clear() error         { return r.Call("Clear").Err() }
func (r *excelRange) ClearContents() error { return r.Call("ClearContents").Err() }
func (r *excelRange) Delete() error        { return r.Call("Delete").Err() }
func (r *excelRange) Copy() error          { return r.Call("Copy").Err() }
func (r *excelRange) Merge() error         { return r.Call("Merge").Err() }
func (r *excelRange) Unmerge() error       { return r.Call("UnMerge").Err() }

func (r *excelRange) MergeCells() (bool, error) {
	return getBool(r, "MergeCells")
}

// AutoFit fits both the column width and the row height of the cells
// intersecting this range, matching xlwings' `range.autofit()`. EntireColumn
// and EntireRow are COM *properties* (read via Get); AutoFit on the resulting
// Range objects is the method that performs the fit.
func (r *excelRange) AutoFit() error {
	if err := r.Get("EntireColumn").Call("AutoFit").Err(); err != nil {
		return err
	}
	return r.Get("EntireRow").Call("AutoFit").Err()
}

func (r *excelRange) Font() Font {
	return wrapFont(r.Get("Font"))
}
