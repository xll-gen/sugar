//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Workbooks represents the Workbooks collection.
type Workbooks interface {
	sugar.Chain
	// Add creates a new empty workbook.
	Add() Workbook
	// Item returns a specific workbook by index or name.
	Item(index interface{}) Workbook
}

type workbooks struct {
	sugar.Chain
}

func (w *workbooks) Add() Workbook {
	return &workbook{w.Call("Add")}
}

func (w *workbooks) Item(index interface{}) Workbook {
	return &workbook{w.Get("Item", index)}
}

// Workbook represents a Workbook object.
type Workbook interface {
	sugar.Chain
	// Worksheets returns the collection of all worksheets in the workbook.
	Worksheets() Worksheets
	// ActiveSheet returns the worksheet that is currently active.
	ActiveSheet() Worksheet
	// Save saves the workbook.
	Save() error
	// Close closes the workbook.
	Close() error
}

type workbook struct {
	sugar.Chain
}

func (w *workbook) Worksheets() Worksheets {
	return &worksheets{w.Get("Worksheets")}
}

func (w *workbook) ActiveSheet() Worksheet {
	return &worksheet{w.Get("ActiveSheet")}
}

func (w *workbook) Save() error {
	return w.Call("Save").Err()
}

func (w *workbook) Close() error {
	return w.Call("Close").Err()
}

// Worksheets represents the Worksheets collection.
type Worksheets interface {
	sugar.Chain
	// Item returns a specific worksheet by index or name.
	Item(index interface{}) Worksheet
}

type worksheets struct {
	sugar.Chain
}

func (w *worksheets) Item(index interface{}) Worksheet {
	return &worksheet{w.Get("Item", index)}
}

// Worksheet represents a Worksheet object.
type Worksheet interface {
	sugar.Chain
	// Range returns a Range object that represents a cell or a range of cells.
	// Arguments can be "A1" or ("A1", "B2").
	Range(cell1 interface{}, cell2 ...interface{}) Range
	// Cells returns a Range object representing a single cell at (row, col).
	Cells(row, col interface{}) Range
}

type worksheet struct {
	sugar.Chain
}

func (w *worksheet) Range(cell1 interface{}, cell2 ...interface{}) Range {
	if len(cell2) > 0 {
		return &excelRange{w.Get("Range", cell1, cell2[0])}
	}
	return &excelRange{w.Get("Range", cell1)}
}

func (w *worksheet) Cells(row, col interface{}) Range {
	return &excelRange{w.Get("Cells", row, col)}
}

// Range represents a cell, a row, a column, or a selection of cells.
type Range interface {
	sugar.Chain
	// SetValue sets the value for the entire range.
	SetValue(value interface{}) Range
	// Cells returns a Range object representing a single cell relative to this range.
	Cells(row, col interface{}) Range
}

type excelRange struct {
	sugar.Chain
}

func (r *excelRange) SetValue(value interface{}) Range {
	return &excelRange{r.Put("Value", value)}
}

func (r *excelRange) Cells(row, col interface{}) Range {
	return &excelRange{r.Get("Cells", row, col)}
}
