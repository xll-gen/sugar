//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// SheetVisibility mirrors Excel's `XlSheetVisibility` enum values used by
// Worksheet.Visible. xlwings exposes these as the strings "visible",
// "hidden", "very_hidden"; we expose the COM ints directly.
type SheetVisibility int32

const (
	SheetVisible    SheetVisibility = -1 // xlSheetVisible
	SheetHidden     SheetVisibility = 0  // xlSheetHidden
	SheetVeryHidden SheetVisibility = 2  // xlSheetVeryHidden
)

// Worksheet is a single worksheet — the Go equivalent of xlwings' `Sheet`.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/sheet.html
type Worksheet interface {
	sugar.Chain
	// Range returns a Range object. Arguments can be "A1", ("A1", "B2"), or
	// any Excel-accepted address form.
	Range(cell1 interface{}, cell2 ...interface{}) Range
	// Cells returns a Range object representing a single cell at
	// (row, col), both 1-based.
	Cells(row, col interface{}) Range
	// UsedRange returns the rectangular range that bounds all cells with
	// content or formatting in the sheet.
	UsedRange() Range
	// Names returns the sheet-scoped defined-names collection. Equivalent to
	// xlwings' `sheet.names`; name strings come back sheet-qualified (e.g.
	// "Sheet1!local_name").
	Names() Names
	// Charts returns the worksheet's embedded-chart collection. Equivalent
	// to xlwings' `sheet.charts`.
	Charts() Charts
	// Shapes returns the worksheet's drawing-layer collection. Equivalent
	// to xlwings' `sheet.shapes`.
	Shapes() Shapes
	// Pictures returns the worksheet's picture collection. Equivalent to
	// xlwings' `sheet.pictures`.
	Pictures() Pictures
	// Name returns the worksheet's tab name.
	Name() (string, error)
	// SetName renames the worksheet.
	SetName(name string) Worksheet
	// Index returns the 1-based tab index of the worksheet.
	Index() (int32, error)
	// Visible returns the current visibility state.
	Visible() (SheetVisibility, error)
	// SetVisible sets the visibility state. Use SheetVisible / SheetHidden /
	// SheetVeryHidden.
	SetVisible(v SheetVisibility) Worksheet
	// Activate makes this the active worksheet.
	Activate() error
	// Delete removes this worksheet from its workbook. Excel typically
	// prompts unless DisplayAlerts is false on the parent Application.
	Delete() error
	// Clear clears both values and formatting from every cell.
	Clear() error
	// ClearContents clears values and formulas but keeps formatting.
	ClearContents() error
	// AutoFit auto-fits the width of all columns and the height of all rows
	// in UsedRange.
	AutoFit() error
}

type worksheet struct {
	sugar.Chain
}

// wrapWorksheet wraps a chain in the Worksheet typed wrapper. It is the single
// construction point for the chain -> Worksheet convention.
func wrapWorksheet(c sugar.Chain) Worksheet { return &worksheet{c} }

func (w *worksheet) Range(cell1 interface{}, cell2 ...interface{}) Range {
	if len(cell2) > 0 {
		return wrapRange(w.Get("Range", cell1, cell2[0]))
	}
	return wrapRange(w.Get("Range", cell1))
}

func (w *worksheet) Cells(row, col interface{}) Range {
	return wrapRange(w.Get("Cells", row, col))
}

func (w *worksheet) UsedRange() Range {
	return wrapRange(w.Get("UsedRange"))
}

func (w *worksheet) Names() Names {
	return wrapNames(w.Get("Names"))
}

func (w *worksheet) Charts() Charts {
	// Worksheet.ChartObjects is a method (no-arg call returns the whole
	// collection), not a property.
	return wrapCharts(w.Call("ChartObjects"))
}

func (w *worksheet) Shapes() Shapes {
	return wrapShapes(w.Get("Shapes"))
}

func (w *worksheet) Pictures() Pictures {
	// The legacy Worksheet.Pictures collection (a method, like
	// ChartObjects) provides Item/Count; Add goes through Shapes.AddPicture
	// — both wrapped together, mirroring xlwings.
	return wrapPictures(w.Call("Pictures"), w.Chain)
}

func (w *worksheet) Name() (string, error) {
	return getString(w, "Name")
}

func (w *worksheet) SetName(name string) Worksheet {
	return wrapWorksheet(w.Put("Name", name))
}

func (w *worksheet) Index() (int32, error) {
	return getInt32(w, "Index")
}

func (w *worksheet) Visible() (SheetVisibility, error) {
	v, err := getInt32(w, "Visible")
	if err != nil {
		return 0, err
	}
	return SheetVisibility(v), nil
}

func (w *worksheet) SetVisible(v SheetVisibility) Worksheet {
	return wrapWorksheet(w.Put("Visible", int32(v)))
}

func (w *worksheet) Activate() error {
	return w.Call("Activate").Err()
}

func (w *worksheet) Delete() error {
	return w.Call("Delete").Err()
}

func (w *worksheet) Clear() error {
	return w.Call("Cells").Call("Clear").Err()
}

func (w *worksheet) ClearContents() error {
	return w.Call("Cells").Call("ClearContents").Err()
}

func (w *worksheet) AutoFit() error {
	used := w.Get("UsedRange")
	if err := used.Get("Columns").Call("AutoFit").Err(); err != nil {
		return err
	}
	return used.Get("Rows").Call("AutoFit").Err()
}
