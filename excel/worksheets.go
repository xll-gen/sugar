//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Worksheets is the collection of worksheets in a workbook — the Go
// equivalent of xlwings' `sheets`.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/sheets.html
type Worksheets interface {
	sugar.Chain
	// Item returns a worksheet by 1-based index or by name.
	Item(index interface{}) Worksheet
	// Add inserts a new worksheet. With no arguments the new sheet is added
	// before the currently active sheet; pass options for finer placement.
	// xlwings reference: `sheets.add(name=None, before=None, after=None)`.
	Add(opts ...AddOption) Worksheet
	// Count returns the number of worksheets.
	Count() (int32, error)
	// Active returns the currently active worksheet in the parent workbook.
	Active() Worksheet
}

// AddOption configures Worksheets.Add. Build with AddBefore/AddAfter/AddName.
type AddOption func(*addOptions)

type addOptions struct {
	before sugar.Chain
	after  sugar.Chain
	name   string
}

// AddBefore inserts the new sheet before the given anchor sheet.
func AddBefore(anchor Worksheet) AddOption {
	return func(o *addOptions) { o.before = anchor }
}

// AddAfter inserts the new sheet after the given anchor sheet.
func AddAfter(anchor Worksheet) AddOption {
	return func(o *addOptions) { o.after = anchor }
}

// AddName names the new sheet after creation. Excel's `Worksheets.Add` does
// not accept a name argument, so we set `Name` on the result.
func AddName(name string) AddOption {
	return func(o *addOptions) { o.name = name }
}

type worksheets struct {
	collection[Worksheet]
}

// wrapWorksheets wraps a chain in the Worksheets typed wrapper. It is the
// single construction point for the chain -> Worksheets convention.
func wrapWorksheets(c sugar.Chain) Worksheets { return &worksheets{newCollection(c, wrapWorksheet)} }

func (w *worksheets) Item(index interface{}) Worksheet {
	// Worksheets.Item is a parameterized property — DISPATCH_PROPERTYGET.
	return w.itemByGet(index)
}

func (w *worksheets) Add(opts ...AddOption) Worksheet {
	o := addOptions{}
	for _, opt := range opts {
		opt(&o)
	}
	var newSheet sugar.Chain
	switch {
	case o.before != nil:
		newSheet = w.Call("Add", o.before)
	case o.after != nil:
		// `Add` signature is (Before, After, Count, Type); pass nil for Before
		// when only After is set. go-ole's nil is fine here.
		newSheet = w.Call("Add", nil, o.after)
	default:
		newSheet = w.Call("Add")
	}
	if o.name != "" {
		newSheet = newSheet.Put("Name", o.name)
	}
	return w.add(newSheet)
}

func (w *worksheets) Count() (int32, error) {
	return w.count()
}

func (w *worksheets) Active() Worksheet {
	return wrapWorksheet(w.Get("Parent").Get("ActiveSheet"))
}
