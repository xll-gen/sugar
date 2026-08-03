//go:build windows && excel_integration

// Executable copy of README.md's "Expression-Based Automation" example.
//
// The example drives live Excel, so it sits behind the excel_integration tag
// with the rest of the Office-bound suite:
//
//	go test -tags=excel_integration -run TestREADME ./expression/ -count=1
//
// Why this file exists: the README block is the first code a reader copies, and
// a prose-only fix to it is indistinguishable from no fix at all — nothing in
// the build would say whether it runs. Keeping the same text here means a drift
// between README and reality is a compile or test failure rather than a review
// miss. If you edit one, edit the other.

package expression_test

import (
	"testing"
	"time"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/expression"
	"github.com/xll-gen/sugar/internal/testutil"
)

// TestREADMEExpressionExample runs README.md's expression example verbatim.
//
// It is a Test rather than a Go Example on purpose: an Example receives no
// *testing.T, so it cannot register the force-kill tier of the two-tier Excel
// cleanup contract (AGENTS §6, v0.9.1) — and a README example that leaks an
// invisible EXCEL.EXE on every test run is not an improvement.
func TestREADMEExpressionExample(t *testing.T) {
	err := sugar.Do(func(ctx sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		// --- test infrastructure, not part of the README block ---------------
		// Force-kill tier, registered before any work so a hung Quit cannot
		// outlive the test. The graceful tier (DisplayAlerts + Quit) is inside
		// the README block itself, because a reader needs that too.
		if h, err := excel.Get("Hwnd").Value(); err == nil {
			if hwnd, ok := h.(int32); ok {
				if pid, err := testutil.PIDFromHwnd(hwnd); err == nil && pid != 0 {
					t.Cleanup(func() { testutil.EnsureProcessExited(t, pid, 5*time.Second) })
				}
			}
		}
		// --- README block begins ---------------------------------------------
		if err := excel.Put("DisplayAlerts", false).Err(); err != nil {
			return err
		}
		defer excel.Call("Quit")

		// Set a property at the end of a path: Workbooks.Add() is a method
		// call, ActiveSheet and Name are properties.
		if err := expression.Put(excel, "Workbooks.Add().ActiveSheet.Name", "Sugar"); err != nil {
			return err
		}

		// Read it back.
		name, err := expression.Get(excel, "ActiveSheet.Name")
		if err != nil {
			return err
		}
		// --- README block ends -----------------------------------------------
		if name != "Sugar" {
			t.Errorf("ActiveSheet.Name = %v (%T), want %q", name, name, "Sugar")
		}

		// The documented boundary, pinned: Excel's ARGUMENTED PROPERTIES
		// (Range(...), Cells(...), Offset(...)) are not reachable through call
		// syntax. The expression engine issues a CallNode as DISPATCH_METHOD and
		// Excel answers DISP_E_MEMBERNOTFOUND for those members — the same
		// property-vs-method trap excel/dispatch_kind_test.go exists for. This
		// is what the README example used to promise and never did. If the
		// assertion starts failing because the engine gained a property
		// fallback, update expression.go's package doc and the README note too.
		if err := expression.Put(excel, "ActiveSheet.Range('A1').Value", "Hello Sugar!"); err == nil {
			t.Error("expression.Put through Range('A1') unexpectedly succeeded; " +
				"the argumented-property limitation documented in expression.go and README.md is stale")
		}
		// The raw chain reaches it, because Range is read with Get (propget) —
		// this is the form the README now shows for cell I/O.
		if err := excel.Get("ActiveSheet").Get("Range", "A1").Put("Value", "Hello Sugar!").Err(); err != nil {
			t.Errorf("chain write to A1: %v", err)
		}
		v, err := excel.Get("ActiveSheet").Get("Range", "A1").Get("Value").Value()
		if err != nil || v != "Hello Sugar!" {
			t.Errorf("A1 = %v (%T) err=%v, want %q", v, v, err, "Hello Sugar!")
		}
		return nil
	})
	if err != nil {
		t.Fatalf("sugar.Do: %v", err)
	}
}
