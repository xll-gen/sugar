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
	// absolute path. Options mirror xlwings' `books.open(fullname,
	// update_links=..., read_only=..., password=...)` keywords.
	Open(path string, opts ...OpenOption) Workbook
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
	collection[Workbook]
}

// wrapWorkbooks wraps a chain in the Workbooks typed wrapper. It is the single
// construction point for the chain -> Workbooks convention.
func wrapWorkbooks(c sugar.Chain) Workbooks { return &workbooks{newCollection(c, wrapWorkbook)} }

func (w *workbooks) Add() Workbook {
	return w.add(w.Call("Add"))
}

// OpenOption configures Workbooks.Open. Build with OpenReadOnly,
// OpenPassword, OpenUpdateLinks.
type OpenOption func(*openOptions)

type openOptions struct {
	readOnly    bool
	password    string
	updateLinks *int32
}

// OpenReadOnly opens the workbook read-only (xlwings `read_only=True`).
func OpenReadOnly() OpenOption {
	return func(o *openOptions) { o.readOnly = true }
}

// OpenPassword supplies the password for a protected workbook (xlwings
// `password=...`).
func OpenPassword(pw string) OpenOption {
	return func(o *openOptions) { o.password = pw }
}

// OpenUpdateLinks controls external-link refresh on open: 0 = don't update,
// 3 = update all (Excel's XlUpdateLinks values; xlwings `update_links=`).
// Omitted, Excel applies its default prompt/settings behavior.
func OpenUpdateLinks(mode int32) OpenOption {
	return func(o *openOptions) { o.updateLinks = &mode }
}

func (w *workbooks) Open(path string, opts ...OpenOption) Workbook {
	o := openOptions{}
	for _, opt := range opts {
		opt(&o)
	}
	// Workbooks.Open's COM signature is positional:
	//   Open(FileName, UpdateLinks, ReadOnly, Format, Password, ...)
	// callOptional seeds unset optionals with Missing() and trims the trailing
	// ones so the simple Open(path) stays a 1-arg call.
	updateLinks := interface{}(sugar.Missing())
	if o.updateLinks != nil {
		updateLinks = *o.updateLinks
	}
	readOnly := interface{}(sugar.Missing())
	if o.readOnly {
		readOnly = true
	}
	format := interface{}(sugar.Missing()) // text-file column delimiter; not exposed
	password := interface{}(sugar.Missing())
	if o.password != "" {
		password = o.password
	}
	return w.add(callOptional(w, "Open", []interface{}{path}, updateLinks, readOnly, format, password))
}

func (w *workbooks) Item(index interface{}) Workbook {
	// Workbooks.Item is a parameterized property — DISPATCH_PROPERTYGET.
	return w.itemByGet(index)
}

func (w *workbooks) Count() (int32, error) {
	return w.count()
}

func (w *workbooks) Active() Workbook {
	// Workbooks has no direct `Active` property; use the parent Application's
	// ActiveWorkbook. xlwings' `books.active` resolves the same way.
	return wrapWorkbook(w.Get("Parent").Get("ActiveWorkbook"))
}
