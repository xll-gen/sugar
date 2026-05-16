//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Workbooks is the Go equivalent of xlwings' `Books` collection — the
// collection of all open workbooks. xlwings exposes both `app.books` and
// `xw.books`; we keep the Workbooks name to match the Excel COM type and
// supply `Application.Books()` as the xlwings-style alias.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/books.html
type Workbooks interface {
	sugar.Chain
	// Add creates a new empty workbook.
	Add() Workbook
	// Open opens an existing workbook by file path. The path must be an
	// absolute path. xlwings reference: `books.open(fullname=...)`.
	Open(path string) Workbook
	// Item returns a specific workbook by index (1-based) or name.
	Item(index interface{}) Workbook
	// Count returns the number of open workbooks. Equivalent to len(books)
	// in xlwings.
	Count() (int32, error)
	// Active returns the currently active workbook in the parent application.
	// Equivalent to xlwings' `books.active`.
	Active() Workbook
}

type workbooks struct {
	sugar.Chain
}

func (w *workbooks) Add() Workbook {
	return &workbook{w.Call("Add")}
}

func (w *workbooks) Open(path string) Workbook {
	return &workbook{w.Call("Open", path)}
}

func (w *workbooks) Item(index interface{}) Workbook {
	return &workbook{w.Get("Item", index)}
}

func (w *workbooks) Count() (int32, error) {
	v, err := w.Get("Count").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

func (w *workbooks) Active() Workbook {
	// Workbooks has no direct `Active` property; use the parent Application's
	// ActiveWorkbook. xlwings' `books.active` resolves the same way.
	return &workbook{w.Get("Parent").Get("ActiveWorkbook")}
}
