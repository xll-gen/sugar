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
	// PID returns the underlying Excel process ID. Windows process IDs are
	// unsigned DWORDs, so PID is reported as uint32.
	PID() (uint32, error)
	// Hwnd returns the top-level window handle of the Excel instance. Window
	// handles are pointer-sized on Win64, so Hwnd is reported as uintptr.
	Hwnd() (uintptr, error)

	// Visible is the Go equivalent of xlwings' `App.visible` property getter.
	// It returns whether the Excel application window is visible.
	Visible() (bool, error)
	// SetVisible is the Go equivalent of xlwings' `App.visible` property setter.
	// It sets Excel's visibility and returns the Application for fluent chaining.
	SetVisible(v bool) Application

	// DisplayAlerts is the Go equivalent of xlwings' `App.display_alerts`
	// property getter. It returns whether Excel shows prompts and alert
	// messages.
	DisplayAlerts() (bool, error)
	// SetDisplayAlerts is the Go equivalent of xlwings' `App.display_alerts`
	// property setter. Set to false to suppress prompts and alert messages;
	// Excel will choose the default response. Returns the Application for
	// fluent chaining.
	SetDisplayAlerts(v bool) Application

	// ScreenUpdating is the Go equivalent of xlwings' `App.screen_updating`
	// property getter. It returns whether Excel repaints the screen during
	// automation.
	ScreenUpdating() (bool, error)
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
	return exec.Command("taskkill", "/F", "/PID", strconv.FormatUint(uint64(pid), 10)).Run()
}

func (a *application) Version() (string, error) {
	v, err := a.Get("Version").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

// PID returns Excel's OS process ID. Excel's COM surface exposes the top-level
// window handle (`Application.Hwnd`) but not the process ID, so we resolve it
// the same way xlwings does via psutil: `GetWindowThreadProcessId` against the
// Hwnd. Windows process IDs are unsigned DWORDs, hence uint32.
func (a *application) PID() (uint32, error) {
	hwnd, err := a.Hwnd()
	if err != nil {
		return 0, err
	}
	return pidFromHwnd(hwnd)
}

func (a *application) Hwnd() (uintptr, error) {
	v, err := a.Get("Hwnd").Value()
	if err != nil {
		return 0, err
	}
	// Excel emits Hwnd as VT_I4; widen to uintptr (no truncation since the
	// underlying COM value is a 32-bit handle, but uintptr is the correct
	// handle-sized Go type for window handles).
	return uintptr(uint32(toInt32(v))), nil
}

// Visible returns the current value of Excel's `Application.Visible` property.
func (a *application) Visible() (bool, error) {
	v, err := a.Get("Visible").Value()
	if err != nil {
		return false, err
	}
	return toBool(v), nil
}

// SetVisible sets Excel's `Application.Visible` property. The Application is
// returned so callers can fluent-chain further property writes; any COM error
// is deferred and observable via the next call's .Err().
func (a *application) SetVisible(v bool) Application {
	return &application{a.Put("Visible", v)}
}

// DisplayAlerts returns the current value of Excel's
// `Application.DisplayAlerts` property.
func (a *application) DisplayAlerts() (bool, error) {
	v, err := a.Get("DisplayAlerts").Value()
	if err != nil {
		return false, err
	}
	return toBool(v), nil
}

// SetDisplayAlerts sets Excel's `Application.DisplayAlerts` property. Returns
// the Application for fluent chaining; COM errors are deferred onto the chain.
func (a *application) SetDisplayAlerts(v bool) Application {
	return &application{a.Put("DisplayAlerts", v)}
}

// ScreenUpdating returns the current value of Excel's
// `Application.ScreenUpdating` property.
func (a *application) ScreenUpdating() (bool, error) {
	v, err := a.Get("ScreenUpdating").Value()
	if err != nil {
		return false, err
	}
	return toBool(v), nil
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
