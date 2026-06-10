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
	sugar.Chain
}

func (w *workbooks) Add() Workbook {
	return &workbook{w.Call("Add")}
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
	// sugar.Missing() skips the optionals we don't set; trailing missing
	// arguments are trimmed so the simple Open(path) stays a 1-arg call.
	args := []interface{}{
		path,
		sugar.Missing(), // UpdateLinks
		sugar.Missing(), // ReadOnly
		sugar.Missing(), // Format (text-file column delimiter; not exposed)
		sugar.Missing(), // Password
	}
	if o.updateLinks != nil {
		args[1] = *o.updateLinks
	}
	if o.readOnly {
		args[2] = true
	}
	if o.password != "" {
		args[4] = o.password
	}
	// Trim trailing Missing() placeholders.
	args = trimTrailingMissing(args)
	return &workbook{w.Call("Open", args...)}
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
