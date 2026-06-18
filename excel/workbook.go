//go:build windows

package excel

import (
	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// XlFileFormat values for Workbook.SaveAs. These mirror Excel's
// `XlFileFormat` enum; the most common formats are exposed as typed constants.
// xlwings infers the format from the file extension, but exposing the enum lets
// callers force a specific container (e.g. write .xlsx data into a macro-enabled
// .xlsm). Pass any other XlFileFormat int directly to SaveFileFormat.
type FileFormat int32

const (
	// FileFormatOpenXMLWorkbook is the default .xlsx format (xlOpenXMLWorkbook).
	FileFormatOpenXMLWorkbook FileFormat = 51
	// FileFormatOpenXMLWorkbookMacroEnabled is the .xlsm macro-enabled format
	// (xlOpenXMLWorkbookMacroEnabled).
	FileFormatOpenXMLWorkbookMacroEnabled FileFormat = 52
	// FileFormatExcel8 is the legacy .xls format (xlExcel8 / Excel 97-2003).
	FileFormatExcel8 FileFormat = 56
	// FileFormatCSV is comma-separated values (xlCSV).
	FileFormatCSV FileFormat = 6
	// FileFormatOpenXMLWorkbookBinary is the .xlsb binary format
	// (xlOpenXMLWorkbookBinary).
	FileFormatOpenXMLWorkbookBinary FileFormat = 50
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
	// Names returns the workbook-scoped defined-names collection.
	// Equivalent to xlwings' `book.names`.
	Names() Names
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
	// SaveAs saves the workbook to the given path. With no options Excel infers
	// the file format from the extension. Options mirror xlwings'
	// `book.save(path)` plus Excel COM's SaveAs keywords:
	// SaveFileFormat / SavePassword.
	SaveAs(path string, opts ...SaveAsOption) error
	// Close closes the workbook. With no options Excel discards unsaved changes
	// if Saved is true, otherwise prompts (unless DisplayAlerts is false on the
	// parent Application). Pass CloseSaveChanges(true/false) to force the
	// save-on-close decision without a prompt. xlwings analogue: `book.close()`.
	Close(opts ...CloseOption) error
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

func (w *workbook) Names() Names {
	return &names{w.Get("Names")}
}

func (w *workbook) Name() (string, error) {
	return getString(w, "Name")
}

func (w *workbook) FullName() (string, error) {
	return getString(w, "FullName")
}

func (w *workbook) Path() (string, error) {
	return getString(w, "Path")
}

func (w *workbook) Saved() (bool, error) {
	return getBool(w, "Saved")
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

// SaveAsOption configures Workbook.SaveAs. Build with SaveFileFormat and
// SavePassword, following the functional-option style of Books.Open.
type SaveAsOption func(*saveAsOptions)

type saveAsOptions struct {
	fileFormat *FileFormat
	password   string
}

// SaveFileFormat forces the on-disk format instead of letting Excel infer it
// from the file extension. Maps to SaveAs's FileFormat argument.
func SaveFileFormat(f FileFormat) SaveAsOption {
	return func(o *saveAsOptions) { o.fileFormat = &f }
}

// SavePassword sets a password to protect the saved workbook. Maps to SaveAs's
// Password argument.
func SavePassword(pw string) SaveAsOption {
	return func(o *saveAsOptions) { o.password = pw }
}

func (w *workbook) SaveAs(path string, opts ...SaveAsOption) error {
	o := saveAsOptions{}
	for _, opt := range opts {
		opt(&o)
	}
	// Workbook.SaveAs's COM signature is positional:
	//   SaveAs(Filename, FileFormat, Password, ...)
	// sugar.Missing() skips the optionals we don't set; trailing missing
	// arguments are trimmed so the simple SaveAs(path) stays a 1-arg call.
	args := []interface{}{
		path,
		sugar.Missing(), // FileFormat
		sugar.Missing(), // Password
	}
	if o.fileFormat != nil {
		args[1] = int32(*o.fileFormat)
	}
	if o.password != "" {
		args[2] = o.password
	}
	args = trimTrailingMissing(args)
	return w.Call("SaveAs", args...).Err()
}

// CloseOption configures Workbook.Close. Build with CloseSaveChanges.
type CloseOption func(*closeOptions)

type closeOptions struct {
	saveChanges *bool
}

// CloseSaveChanges forces the save-on-close decision without an interactive
// prompt: true saves pending changes, false discards them. Maps to Close's
// SaveChanges argument. Omitted, Excel applies its default (prompt / discard
// based on the Saved flag and DisplayAlerts).
func CloseSaveChanges(save bool) CloseOption {
	return func(o *closeOptions) { o.saveChanges = &save }
}

func (w *workbook) Close(opts ...CloseOption) error {
	o := closeOptions{}
	for _, opt := range opts {
		opt(&o)
	}
	if o.saveChanges == nil {
		return w.Call("Close").Err()
	}
	// Workbook.Close(SaveChanges, Filename, RouteWorkbook).
	return w.Call("Close", *o.saveChanges).Err()
}

// trimTrailingMissing drops trailing sugar.Missing() placeholders from a
// positional COM argument list so that simple calls collapse to their shortest
// form. Mirrors the inline logic in Workbooks.Open.
func trimTrailingMissing(args []interface{}) []interface{} {
	last := len(args) - 1
	for last > 0 {
		if _, isMissing := args[last].(*ole.VARIANT); !isMissing {
			break
		}
		last--
	}
	return args[:last+1]
}
