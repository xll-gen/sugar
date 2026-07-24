//go:build windows

package excel

import (
	"fmt"
	"syscall"
	"unsafe"

	ole "github.com/go-ole/go-ole"
)

// Win32 helpers used by Application.Kill / Application.PID and by
// GetApplicationByPID (multi-instance attach).
//
// Excel's COM surface exposes `Application.Hwnd` (the top-level window handle)
// but not the OS process ID. xlwings resolves PID via psutil; we use the
// equivalent Win32 path: `GetWindowThreadProcessId` against the Hwnd.
//
// GetApplicationByPID needs the inverse: given an Excel PID, reach that
// instance's `Application` IDispatch. The Running Object Table (what
// `GetActiveObject("Excel.Application")` reads) is process-ambiguous — it
// returns whichever Excel registered first, and in a server process whose
// apartment never saw a ROT registration it fails outright with
// `MK_E_UNAVAILABLE` ("작업을 사용할 수 없습니다"). The reliable, instance-precise
// route is the same one the xll-gen C++ host uses (AGENTS §18.11): walk the
// window chain `XLMAIN -> XLDESK -> EXCEL7`, then pull the native object model
// off the EXCEL7 child via `AccessibleObjectFromWindow(OBJID_NATIVEOM)` and
// read its `.Application`.

const (
	objidNativeOM = 0xFFFFFFF0 // OBJID_NATIVEOM (-16 as DWORD): Excel's native object model
)

var (
	user32                       = syscall.NewLazyDLL("user32.dll")
	oleacc                       = syscall.NewLazyDLL("oleacc.dll")
	procGetWindowThreadProcessId = user32.NewProc("GetWindowThreadProcessId")
	procEnumWindows              = user32.NewProc("EnumWindows")
	procEnumChildWindows         = user32.NewProc("EnumChildWindows")
	procFindWindowEx             = user32.NewProc("FindWindowExW")
	procGetClassName             = user32.NewProc("GetClassNameW")
	procAccessibleObjectFromWnd  = oleacc.NewProc("AccessibleObjectFromWindow")
)

// pidFromHwnd returns the PID owning the given top-level window. Returns an
// error if the Win32 call fails or the window has gone away. The PID is an
// unsigned DWORD (uint32).
func pidFromHwnd(hwnd uintptr) (uint32, error) {
	var pid uint32
	ret, _, callErr := procGetWindowThreadProcessId.Call(
		hwnd,
		uintptr(unsafe.Pointer(&pid)),
	)
	if ret == 0 {
		return 0, fmt.Errorf("GetWindowThreadProcessId failed for hwnd 0x%x: %v", hwnd, callErr)
	}
	return pid, nil
}

// classNameOf reads the window class name of hwnd.
func classNameOf(hwnd uintptr) string {
	var buf [64]uint16
	n, _, _ := procGetClassName.Call(hwnd, uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf)))
	if n == 0 {
		return ""
	}
	return syscall.UTF16ToString(buf[:n])
}

// Enumeration callbacks MUST be created exactly once, at package scope.
// syscall.NewCallback allocates a C-callable thunk that the Go runtime NEVER
// frees for the life of the process. The old per-process cap (~2000) that once
// turned this leak into a hard "too many callback functions" throw is stale —
// modern Go runtimes (verified on 1.26.3) absorb 200k+ callbacks without a
// symptom, so the crash is no longer the concern. The unbounded thunk *leak*
// itself is still a real defect: an earlier version created a fresh closure
// per call to capture (pid, found), so every GetApplicationByPID — run on each
// ribbon command in a long-lived Go server — leaked one thunk permanently.
// Hoisting the callback to a package var (state threaded through lParam, below)
// remains the correct fix regardless of the cap.
//
// To keep a single reusable callback while still carrying per-call search
// state, the state travels through the enumeration's lParam as a pointer to a
// stack-local struct. EnumWindows / EnumChildWindows only dereference lParam
// for the duration of the call (they never retain it past the final callback
// invocation), so passing the address of a local is safe and the calls remain
// reentrant / goroutine-safe. The lParam parameter is typed as unsafe.Pointer
// (pointer-sized, so NewCallback accepts it) so the callback body dereferences
// it without a vet-flagged uintptr->unsafe.Pointer conversion.

// xlMainSearch is the per-call state for findXlMainForPID, threaded through the
// EnumWindows lParam.
type xlMainSearch struct {
	pid   uint32
	found uintptr
}

// enumXlMainProc is the single, reusable EnumWindows callback for
// findXlMainForPID. See the note above on why it is a package-level var.
var enumXlMainProc = syscall.NewCallback(func(hwnd uintptr, lparam unsafe.Pointer) uintptr {
	s := (*xlMainSearch)(lparam)
	if classNameOf(hwnd) != "XLMAIN" {
		return 1 // continue
	}
	if owner, err := pidFromHwnd(hwnd); err != nil || owner != s.pid {
		return 1 // continue
	}
	s.found = hwnd
	return 0 // stop
})

// enumExcel7Proc is the single, reusable EnumChildWindows callback for
// findExcel7Child (used only on the fallback path).
var enumExcel7Proc = syscall.NewCallback(func(hwnd uintptr, lparam unsafe.Pointer) uintptr {
	found := (*uintptr)(lparam)
	if classNameOf(hwnd) == "EXCEL7" {
		*found = hwnd
		return 0 // stop
	}
	return 1 // continue
})

// findXlMainForPID enumerates top-level windows and returns the first XLMAIN
// frame window owned by the target PID, or 0 if none is found. Unlike the
// in-process XLL host (which uses EnumThreadWindows on its own STA thread), the
// server runs in a separate process, so it must scan all top-level windows and
// match on PID.
func findXlMainForPID(pid uint32) uintptr {
	s := xlMainSearch{pid: pid}
	procEnumWindows.Call(enumXlMainProc, uintptr(unsafe.Pointer(&s)))
	return s.found
}

// findExcel7Child locates the EXCEL7 child window under an XLMAIN frame. It
// tries the fast direct path (XLMAIN -> XLDESK -> EXCEL7) first, then falls
// back to a recursive child enumeration — mirroring the C++ host's
// GetExcelApplication walk.
func findExcel7Child(frame uintptr) uintptr {
	xldesk, _, _ := procFindWindowEx.Call(frame, 0, classPtr("XLDESK"), 0)
	if xldesk != 0 {
		if excel7, _, _ := procFindWindowEx.Call(xldesk, 0, classPtr("EXCEL7"), 0); excel7 != 0 {
			return excel7
		}
	}

	var found uintptr
	procEnumChildWindows.Call(frame, enumExcel7Proc, uintptr(unsafe.Pointer(&found)))
	return found
}

func classPtr(s string) uintptr {
	p, _ := syscall.UTF16PtrFromString(s)
	return uintptr(unsafe.Pointer(p))
}

// applicationDispatchForPID returns the Excel `Application` IDispatch for the
// instance whose process id is pid. The returned dispatch carries one
// reference owned by the caller (callers wrap it with sugar.From, which
// AddRefs again and tracks the AddRef on the arena, then Release the raw ref
// here). Returns nil + error if the window chain or native object model is not
// reachable (e.g. Excel has no workbook open yet, so no EXCEL7 child exists).
func applicationDispatchForPID(pid uint32) (*ole.IDispatch, error) {
	frame := findXlMainForPID(pid)
	if frame == 0 {
		return nil, fmt.Errorf("no XLMAIN window found for Excel PID %d", pid)
	}
	excel7 := findExcel7Child(frame)
	if excel7 == 0 {
		return nil, fmt.Errorf("no EXCEL7 child window under Excel PID %d (no workbook open yet?)", pid)
	}

	var window *ole.IDispatch
	hr, _, _ := procAccessibleObjectFromWnd.Call(
		excel7,
		uintptr(objidNativeOM),
		uintptr(unsafe.Pointer(ole.IID_IDispatch)),
		uintptr(unsafe.Pointer(&window)),
	)
	if int32(hr) < 0 || window == nil {
		return nil, fmt.Errorf("AccessibleObjectFromWindow(OBJID_NATIVEOM) failed for Excel PID %d: hr=0x%x", pid, uint32(hr))
	}
	defer window.Release()

	// The native object model on EXCEL7 is the Excel `Window`; its
	// `.Application` is the instance we want.
	appVar, err := window.GetProperty("Application")
	if err != nil {
		return nil, fmt.Errorf("Window.Application failed for Excel PID %d: %w", pid, err)
	}
	if appVar.VT != ole.VT_DISPATCH {
		appVar.Clear()
		return nil, fmt.Errorf("Window.Application returned non-dispatch (VT=%d) for Excel PID %d", appVar.VT, pid)
	}
	app := appVar.ToIDispatch()
	if app == nil {
		appVar.Clear()
		return nil, fmt.Errorf("Window.Application returned nil dispatch for Excel PID %d", pid)
	}
	app.AddRef()   // own a ref independent of the VARIANT
	appVar.Clear() // releases the VARIANT's own ref
	return app, nil
}
