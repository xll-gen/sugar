//go:build windows && excel_integration

// Excel-bound goroutine-isolation regression test for sugar.Go. Gated behind
// the excel_integration build tag (it spawns two real Excel instances):
//
//	go test -tags=excel_integration ./...

package sugar_test

import (
	"testing"
	"time"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/internal/testutil"
)

// TestGo_TwoExcelInstancesIsolated is the goroutine-safety regression test
// from the AGENTS.md backlog: two sugar.Go goroutines each drive their own
// Excel instance on their own OS thread, concurrently, and must not
// interfere. Skips when Excel is not installed.
func TestGo_TwoExcelInstancesIsolated(t *testing.T) {
	type result struct {
		hwnd  interface{}
		pid   uint32
		value interface{}
		err   error
	}

	run := func(marker string, out chan<- result) <-chan error {
		return sugar.Go(func(ctx sugar.Context) error {
			excel := ctx.Create("Excel.Application")
			if err := excel.Err(); err != nil {
				out <- result{err: err}
				return nil // reported via channel; Skip decision is the test's
			}
			// Graceful tier of the two-tier cleanup; the force-kill tier is
			// registered by the receiving side from the reported PID.
			defer excel.Put("DisplayAlerts", false).Call("Quit")
			excel.Put("Visible", false)

			var pid uint32
			if hwnd, err := excel.Get("Hwnd").Value(); err == nil {
				pid, _ = testutil.PIDFromHwnd(toInt32(hwnd))
			}

			sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
			if err := sheet.Get("Range", "A1").Put("Value", marker).Err(); err != nil {
				out <- result{pid: pid, err: err}
				return nil
			}
			v, err := sheet.Get("Range", "A1").Get("Value").Value()
			if err != nil {
				out <- result{pid: pid, err: err}
				return nil
			}
			hwnd, err := excel.Get("Hwnd").Value()
			out <- result{hwnd: hwnd, pid: pid, value: v, err: err}
			return nil
		})
	}

	ch1 := make(chan result, 1)
	ch2 := make(chan result, 1)
	done1 := run("from-goroutine-1", ch1)
	done2 := run("from-goroutine-2", ch2)

	r1, r2 := <-ch1, <-ch2
	for _, r := range []result{r1, r2} {
		if r.pid != 0 {
			pid := r.pid
			t.Cleanup(func() { testutil.EnsureProcessExited(t, pid, 5*time.Second) })
		}
	}
	if err := <-done1; err != nil {
		t.Fatalf("goroutine 1 terminal error: %v", err)
	}
	if err := <-done2; err != nil {
		t.Fatalf("goroutine 2 terminal error: %v", err)
	}
	if r1.err != nil || r2.err != nil {
		t.Skipf("Excel not usable in this environment: %v / %v", r1.err, r2.err)
	}

	if r1.value != "from-goroutine-1" {
		t.Errorf("goroutine 1 read back %v", r1.value)
	}
	if r2.value != "from-goroutine-2" {
		t.Errorf("goroutine 2 read back %v", r2.value)
	}
	if r1.hwnd == r2.hwnd {
		t.Errorf("expected two distinct Excel instances, both have Hwnd %v", r1.hwnd)
	}
}
