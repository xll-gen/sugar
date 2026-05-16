//go:build windows

package excel

import (
	"fmt"
	"syscall"
	"unsafe"
)

// Win32 helpers used by Application.Kill / Application.PID.
//
// Excel's COM surface exposes `Application.Hwnd` (the top-level window handle)
// but not the OS process ID. xlwings resolves PID via psutil; we use the
// equivalent Win32 path: `GetWindowThreadProcessId` against the Hwnd.

var (
	user32                       = syscall.NewLazyDLL("user32.dll")
	procGetWindowThreadProcessId = user32.NewProc("GetWindowThreadProcessId")
)

// pidFromHwnd returns the PID owning the given top-level window. Returns an
// error if the Win32 call fails or the window has gone away.
func pidFromHwnd(hwnd int32) (int32, error) {
	var pid uint32
	ret, _, callErr := procGetWindowThreadProcessId.Call(
		uintptr(hwnd),
		uintptr(unsafe.Pointer(&pid)),
	)
	if ret == 0 {
		return 0, fmt.Errorf("GetWindowThreadProcessId failed for hwnd 0x%x: %v", hwnd, callErr)
	}
	return int32(pid), nil
}
