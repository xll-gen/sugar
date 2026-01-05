//go:build windows

package excel

import (
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/xll-gen/sugar"
)

// New creates a new Excel application instance and returns a cleanup function.
// The cleanup function ensures that the Excel application is quit and resources are released.
func New() (*ole.IDispatch, func(), error) {
	runtime.LockOSThread()
	if err := ole.CoInitialize(0); err != nil {
		runtime.UnlockOSThread()
		return nil, nil, err
	}

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, nil, err
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, nil, err
	}

	cleanup := func() {
		// Ensure we quit Excel
		// We create a new chain from the dispatch to call Quit.
		// We use Err() to handle potential errors implicitly (just returning them if needed, but here we ignore).
		sugar.From(disp).Put("DisplayAlerts", false).Call("Quit").Release()
		disp.Release()
		ole.CoUninitialize()
		runtime.UnlockOSThread()
	}

	return disp, cleanup, nil
}