//go:build windows

// Package testutil holds shared helpers for sugar's Windows test suites.
//
// Its flagship helper is EnsureProcessExited — the force-kill tier of the
// two-tier cleanup contract every Excel-spawning test must follow:
//
//  1. Graceful tier: the test defers DisplayAlerts(false) + Quit *inside*
//     the sugar.Do block, while COM is still initialized on the thread and
//     the Application dispatch is still alive.
//  2. Force-kill tier: the test registers EnsureProcessExited via t.Cleanup.
//     It runs after the Do block has returned (no COM needed — pure Win32),
//     waits for the process to exit on its own, and terminates it if the
//     graceful Quit hung (modal dialog, zombie EXCEL.EXE, ...).
//
// A bare `defer Quit()` alone is never enough: a hung Quit silently leaks an
// invisible Excel process that outlives the test binary.
package testutil

import (
	"fmt"
	"os"
	"syscall"
	"testing"
	"time"
	"unsafe"
)

var (
	user32                       = syscall.NewLazyDLL("user32.dll")
	procGetWindowThreadProcessId = user32.NewProc("GetWindowThreadProcessId")
)

// PIDFromHwnd resolves the process ID owning a top-level window handle.
// Excel's COM surface exposes Application.Hwnd but not the PID; this is the
// same Win32 lookup excel.Application.PID performs internally, duplicated
// here so core-layer tests that drive raw chains need not import the excel
// package.
func PIDFromHwnd(hwnd int32) (uint32, error) {
	var pid uint32
	ret, _, callErr := procGetWindowThreadProcessId.Call(
		uintptr(uint32(hwnd)),
		uintptr(unsafe.Pointer(&pid)),
	)
	if ret == 0 {
		return 0, fmt.Errorf("GetWindowThreadProcessId failed for hwnd 0x%x: %v", hwnd, callErr)
	}
	return pid, nil
}

// EnsureProcessExited waits up to grace for pid to exit on its own (the
// graceful Quit must already have been issued by the test) and force-kills
// the process if it is still alive afterwards. Intended for t.Cleanup: it
// runs after the sugar.Do block and needs no COM, only Win32 process APIs.
func EnsureProcessExited(t testing.TB, pid uint32, grace time.Duration) {
	t.Helper()
	p, err := os.FindProcess(int(pid))
	if err != nil {
		return // already exited — OpenProcess fails for dead PIDs
	}
	defer p.Release()

	exited := make(chan struct{})
	go func() {
		_, _ = p.Wait() // on Windows this works for non-child processes too
		close(exited)
	}()

	select {
	case <-exited:
		return // graceful Quit worked
	case <-time.After(grace):
	}

	t.Logf("process %d still alive %v after graceful quit; force-killing", pid, grace)
	if err := p.Kill(); err != nil {
		t.Logf("force-kill of process %d failed: %v", pid, err)
		return
	}
	select {
	case <-exited:
	case <-time.After(5 * time.Second):
		t.Logf("process %d did not exit even after TerminateProcess", pid)
	}
}
