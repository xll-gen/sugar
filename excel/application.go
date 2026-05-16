//go:build windows

package excel

import (
	"os/exec"
	"strconv"

	"github.com/xll-gen/sugar"
)

// Calculation mirrors Excel's `XlCalculation` enum used by
// Application.Calculation. xlwings exposes the same idea as the strings
// "automatic" / "manual" / "semiautomatic"; we expose the COM ints directly
// because xlwings' string form is just a thin wrapper over these.
type Calculation int32

const (
	CalculationAutomatic     Calculation = -4105 // xlCalculationAutomatic
	CalculationManual        Calculation = -4135 // xlCalculationManual
	CalculationSemiautomatic Calculation = 2     // xlCalculationSemiautomatic
)

// Application is the Go equivalent of xlwings' `App` object. It wraps the
// `Excel.Application` COM object and is the root of the Excel object model.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/app.html
type Application interface {
	sugar.Chain
	// Workbooks returns the collection of all open workbooks.
	Workbooks() Workbooks
	// Books is the xlwings-style alias for Workbooks.
	Books() Workbooks
	// ActiveWorkbook returns the workbook that is currently active.
	ActiveWorkbook() Workbook
	// Quit quits the Excel application. The instance still exists in memory
	// until released; use Kill if Quit hangs.
	Quit() error
	// Kill terminates the underlying Excel process via `taskkill /PID`.
	// Equivalent to xlwings' `app.kill()`. Use only as a last resort —
	// pending writes are lost.
	Kill() error

	// Version returns the Excel application version string (e.g. "16.0").
	Version() (string, error)
	// PID returns the underlying Excel process ID.
	PID() (int32, error)
	// Hwnd returns the top-level window handle of the Excel instance.
	Hwnd() (int32, error)

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

	// Calculation returns the current calculation mode.
	Calculation() (Calculation, error)
	// SetCalculation sets the calculation mode. Set to CalculationManual to
	// stop Excel from recalculating during bulk writes.
	SetCalculation(c Calculation) Application
}

type application struct {
	sugar.Chain
}

func (a *application) Workbooks() Workbooks {
	return &workbooks{a.Get("Workbooks")}
}

func (a *application) Books() Workbooks {
	return a.Workbooks()
}

func (a *application) ActiveWorkbook() Workbook {
	return &workbook{a.Get("ActiveWorkbook")}
}

func (a *application) Quit() error {
	return a.Call("Quit").Err()
}

// Kill terminates the Excel process associated with this Application by
// looking up the PID and invoking `taskkill /F /PID`. Used only when Quit
// hangs — bypasses Excel's save/close logic, so unsaved work is lost.
func (a *application) Kill() error {
	pid, err := a.PID()
	if err != nil {
		return err
	}
	return exec.Command("taskkill", "/F", "/PID", strconv.Itoa(int(pid))).Run()
}

func (a *application) Version() (string, error) {
	v, err := a.Get("Version").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

// PID returns Excel's process ID via the `Hwnd` -> Win32 lookup path Excel
// exposes through its OM property `Application.Hwnd` plus a process-snapshot
// scan. We use the simpler approach: Excel surfaces the same PID indirectly
// via the `Application.Hinstance`-style behaviour, but the cleanest portable
// answer is the `Application` itself doesn't expose PID — we ask Excel's
// `Application.OperatingSystem`? No, the proper COM path is `Application.Hwnd`
// (a window handle) and we resolve PID with `GetWindowThreadProcessId`. We
// avoid pulling in user32 by reading `Excel`'s built-in `ProcessID` shim —
// modern Excel COM provides it indirectly through `Application.Caller`'s
// process. Since none of those are portable across Excel versions, we
// fall through to GetWindowThreadProcessId via the WinAPI.
func (a *application) PID() (int32, error) {
	hwnd, err := a.Hwnd()
	if err != nil {
		return 0, err
	}
	return pidFromHwnd(hwnd)
}

func (a *application) Hwnd() (int32, error) {
	v, err := a.Get("Hwnd").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
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

func (a *application) Calculation() (Calculation, error) {
	v, err := a.Get("Calculation").Value()
	if err != nil {
		return 0, err
	}
	return Calculation(toInt32(v)), nil
}

func (a *application) SetCalculation(c Calculation) Application {
	return &application{a.Put("Calculation", int32(c))}
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
