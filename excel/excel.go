//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Application represents the Excel.Application object.
type Application interface {
	sugar.Chain
	Workbooks() Workbooks
	ActiveWorkbook() Workbook
	Quit() error
}

type application struct {
	sugar.Chain
}

func (a *application) Workbooks() Workbooks {
	return &workbooks{a.Get("Workbooks")}
}

func (a *application) ActiveWorkbook() Workbook {
	return &workbook{a.Get("ActiveWorkbook")}
}

func (a *application) Quit() error {
	return a.Call("Quit").Err()
}

// NewApplication creates a new Excel instance.
func NewApplication(ctx *sugar.Context) Application {
	return &application{ctx.Create("Excel.Application")}
}

// GetApplication attaches to a running Excel instance.
func GetApplication(ctx *sugar.Context) Application {
	return &application{ctx.GetActive("Excel.Application")}
}

// Workbooks represents the Workbooks collection.
type Workbooks interface {
	sugar.Chain
	Add() Workbook
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
	Worksheets() Worksheets
	ActiveSheet() Worksheet
	Save() error
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
	Range(cell1 interface{}, cell2 ...interface{}) Range
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

// Range represents a Range object.
type Range interface {
	sugar.Chain
	SetValue(value interface{}) Range
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
