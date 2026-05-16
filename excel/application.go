//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Application is the Go equivalent of xlwings' `App` object. It wraps the
// `Excel.Application` COM object and is the root of the Excel object model.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/app.html
type Application interface {
	sugar.Chain
	// Workbooks returns the collection of all open workbooks.
	Workbooks() Workbooks
	// ActiveWorkbook returns the workbook that is currently active.
	ActiveWorkbook() Workbook
	// Quit quits the Excel application.
	Quit() error

	// Visible is the Go equivalent of xlwings' `App.visible` property getter.
	// It returns a sugar.Chain representing the COM `Visible` property value;
	// call .Value() to materialize as a bool.
	Visible() sugar.Chain
	// SetVisible is the Go equivalent of xlwings' `App.visible` property setter.
	// It sets Excel's visibility and returns the Application for fluent chaining.
	SetVisible(v bool) Application

	// DisplayAlerts is the Go equivalent of xlwings' `App.display_alerts`
	// property getter. It returns a sugar.Chain representing the COM
	// `DisplayAlerts` property value; call .Value() to materialize as a bool.
	DisplayAlerts() sugar.Chain
	// SetDisplayAlerts is the Go equivalent of xlwings' `App.display_alerts`
	// property setter. Set to false to suppress prompts and alert messages;
	// Excel will choose the default response. Returns the Application for
	// fluent chaining.
	SetDisplayAlerts(v bool) Application

	// ScreenUpdating is the Go equivalent of xlwings' `App.screen_updating`
	// property getter. It returns a sugar.Chain representing the COM
	// `ScreenUpdating` property value; call .Value() to materialize as a bool.
	ScreenUpdating() sugar.Chain
	// SetScreenUpdating is the Go equivalent of xlwings' `App.screen_updating`
	// property setter. Turn off to speed up scripts; remember to turn back on
	// when the script ends. Returns the Application for fluent chaining.
	SetScreenUpdating(v bool) Application
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

// Visible returns the current value of Excel's `Application.Visible` property
// wrapped in a sugar.Chain. Call .Value() on the returned chain to obtain the
// Go bool, or .Err() to surface any deferred COM error.
func (a *application) Visible() sugar.Chain {
	return a.Get("Visible")
}

// SetVisible sets Excel's `Application.Visible` property. The Application is
// returned so callers can fluent-chain further property writes; any COM error
// is deferred and observable via the next call's .Err().
func (a *application) SetVisible(v bool) Application {
	return &application{a.Put("Visible", v)}
}

// DisplayAlerts returns the current value of Excel's
// `Application.DisplayAlerts` property wrapped in a sugar.Chain. Call .Value()
// on the returned chain to obtain the Go bool.
func (a *application) DisplayAlerts() sugar.Chain {
	return a.Get("DisplayAlerts")
}

// SetDisplayAlerts sets Excel's `Application.DisplayAlerts` property. Returns
// the Application for fluent chaining; COM errors are deferred onto the chain.
func (a *application) SetDisplayAlerts(v bool) Application {
	return &application{a.Put("DisplayAlerts", v)}
}

// ScreenUpdating returns the current value of Excel's
// `Application.ScreenUpdating` property wrapped in a sugar.Chain. Call .Value()
// on the returned chain to obtain the Go bool.
func (a *application) ScreenUpdating() sugar.Chain {
	return a.Get("ScreenUpdating")
}

// SetScreenUpdating sets Excel's `Application.ScreenUpdating` property.
// Returns the Application for fluent chaining; COM errors are deferred onto
// the chain.
func (a *application) SetScreenUpdating(v bool) Application {
	return &application{a.Put("ScreenUpdating", v)}
}

// NewApplication creates a new Excel instance and tracks it on the given
// context's arena.
func NewApplication(ctx sugar.Context) Application {
	return &application{ctx.Create("Excel.Application")}
}

// GetApplication attaches to a running Excel instance and tracks it on the
// given context's arena.
func GetApplication(ctx sugar.Context) Application {
	return &application{ctx.GetActive("Excel.Application")}
}
