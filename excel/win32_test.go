//go:build windows

// Regression tests for the window-enumeration callbacks used by
// GetApplicationByPID. These run Excel-free: the search helpers just enumerate
// the desktop's windows and return 0 when no Excel instance matches.

package excel

import "testing"

// TestEnumCallbacksArePackageLevel pins the fix for the callback-accumulation
// fatal bug. syscall.NewCallback allocates a C-callable thunk the Go runtime
// NEVER frees. The pre-fix code minted a fresh closure inside findXlMainForPID
// / findExcel7Child on every call to capture (pid, found); because
// GetApplicationByPID runs on the ribbon-command attach path, a long-lived Go
// server leaked one thunk per command. Historically the process capped at
// runtime maxCallback (2000) and died with an unrecoverable "too many callback
// functions" throw; newer runtimes raise that ceiling, but the leak is still
// unbounded growth, so the fix — one thunk created once at package scope —
// stands regardless of the exact cap.
//
// This test guards the fix structurally, which is robust across Go versions:
// the thunk addresses must be non-zero and must NEVER change across many
// repeated searches. A per-call NewCallback (the regression) would allocate a
// new thunk each call, so a captured address would no longer match. The loop
// also confirms the search helpers stay functional (return 0 when no Excel
// matches) across repeated invocations.
func TestEnumCallbacksArePackageLevel(t *testing.T) {
	if enumXlMainProc == 0 || enumExcel7Proc == 0 {
		t.Fatal("enumeration callbacks must be package-level syscall.NewCallback thunks (non-zero)")
	}
	xlMain, excel7 := enumXlMainProc, enumExcel7Proc

	const iterations = 500
	for i := 0; i < iterations; i++ {
		// PID 0xFFFFFFFE owns no window, so the XLMAIN callback runs for every
		// top-level window, matches none, and returns 0. Exercises the full
		// callback body without depending on Excel being present.
		if got := findXlMainForPID(0xFFFFFFFE); got != 0 {
			t.Fatalf("iteration %d: findXlMainForPID(bogus) = 0x%x, want 0", i, got)
		}
		// frame 0 makes EnumChildWindows enumerate top-level windows (its
		// documented NULL behavior), so this takes findExcel7Child's fallback
		// callback path; no EXCEL7 top-level window exists, so it returns 0.
		if got := findExcel7Child(0); got != 0 {
			t.Fatalf("iteration %d: findExcel7Child(0) = 0x%x, want 0", i, got)
		}
	}

	if enumXlMainProc != xlMain || enumExcel7Proc != excel7 {
		t.Fatal("enumeration callback thunks changed during repeated calls; they must be created exactly once")
	}
}
