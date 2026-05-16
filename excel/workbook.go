//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Workbook is the Go equivalent of xlwings' `Book` — a single Excel workbook.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/book.html
type Workbook interface {
	sugar.Chain
	// Worksheets returns the collection of all worksheets in the workbook.
	Worksheets() Worksheets
	// Sheets is the xlwings-named alias for Worksheets.
	Sheets() Worksheets
	// ActiveSheet returns the worksheet that is currently active in this
	// workbook.
	ActiveSheet() Worksheet
	// App returns the parent Application that owns this workbook.
	// Equivalent to xlwings' `book.app`.
	App() Application
	// Name returns the workbook's file name (e.g. "Book1.xlsx").
	Name() (string, error)
	// FullName returns the workbook's full path including the file name.
	FullName() (string, error)
	// Path returns the directory of the workbook (without the file name).
	Path() (string, error)
	// Saved reports whether the workbook has unsaved changes (true = clean).
	Saved() (bool, error)
	// SetSaved marks the workbook's modified flag. Set to true to suppress
	// the "Save changes?" prompt on Close.
	SetSaved(v bool) Workbook
	// Activate makes this workbook the active workbook in its application.
	Activate() error
	// Save saves the workbook to its existing path. The workbook must have a
	// file name (use SaveAs for a new file).
	Save() error
	// SaveAs saves the workbook to the given path. Excel infers the file
	// format from the extension.
	SaveAs(path string) error
	// Close closes the workbook. Any unsaved changes are discarded if Saved
	// is true; otherwise Excel may prompt unless DisplayAlerts is false on
	// the parent Application.
	Close() error
}

type workbook struct {
	sugar.Chain
}

func (w *workbook) Worksheets() Worksheets {
	return &worksheets{w.Get("Worksheets")}
}

func (w *workbook) Sheets() Worksheets {
	return &worksheets{w.Get("Sheets")}
}

func (w *workbook) ActiveSheet() Worksheet {
	return &worksheet{w.Get("ActiveSheet")}
}

func (w *workbook) App() Application {
	return &application{w.Get("Application")}
}

func (w *workbook) Name() (string, error) {
	v, err := w.Get("Name").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (w *workbook) FullName() (string, error) {
	v, err := w.Get("FullName").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (w *workbook) Path() (string, error) {
	v, err := w.Get("Path").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (w *workbook) Saved() (bool, error) {
	v, err := w.Get("Saved").Value()
	if err != nil {
		return false, err
	}
	b, _ := v.(bool)
	return b, nil
}

func (w *workbook) SetSaved(v bool) Workbook {
	return &workbook{w.Put("Saved", v)}
}

func (w *workbook) Activate() error {
	return w.Call("Activate").Err()
}

func (w *workbook) Save() error {
	return w.Call("Save").Err()
}

func (w *workbook) SaveAs(path string) error {
	return w.Call("SaveAs", path).Err()
}

func (w *workbook) Close() error {
	return w.Call("Close").Err()
}
