//go:build windows

// Regression tests for the Runner's COM-initialization handling and for
// cross-goroutine isolation of sugar.Go. These use lightweight scripting COM
// objects (Scripting.Dictionary) where possible so they run fast on any
// Windows host; the goroutine-isolation test needs Excel and skips without
// it.

package sugar_test

import (
	"runtime"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// TestDo_OnPreInitializedSTAThread runs sugar.Do on a thread the host has
// already CoInitialize'd — the situation inside an XLL or a GUI app.
// CoInitialize then returns S_FALSE, which go-ole surfaces as an error;
// before v0.8.0 sugar.Do failed outright on such threads.
func TestDo_OnPreInitializedSTAThread(t *testing.T) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := ole.CoInitialize(0); err != nil {
		t.Fatalf("host CoInitialize: %v", err)
	}
	defer ole.CoUninitialize()

	err := sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			return err
		}
		if err := dict.Call("Add", "k", "v").Err(); err != nil {
			return err
		}
		v, err := dict.Get("Item", "k").Value()
		if err != nil {
			return err
		}
		if v != "v" {
			t.Errorf("dictionary round trip: got %v", v)
		}
		return nil
	})
	if err != nil {
		t.Fatalf("sugar.Do on pre-initialized thread: %v", err)
	}

	// The thread must still be usable by the host afterwards (the arena
	// must not have over-released the host's COM init count).
	d := sugar.Create("Scripting.Dictionary")
	if err := d.Err(); err != nil {
		t.Fatalf("thread unusable after sugar.Do: %v", err)
	}
	_ = d.Release()
}

// TestValueChainsAreTracked pins the BSTR-leak fix: chains carrying plain
// VARIANT results (not IDispatch) must be registered with the arena so
// Release() VariantClears them at scope end. We observe tracking through the
// chain's post-release behavior: a released chain reports a nil value.
func TestValueChainsAreTracked(t *testing.T) {
	var valueChain sugar.Chain

	err := sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			return err
		}
		if err := dict.Call("Add", "k", "hello").Err(); err != nil {
			return err
		}

		valueChain = dict.Get("Item", "k")
		v, err := valueChain.Value()
		if err != nil || v != "hello" {
			t.Fatalf("inside scope: got %v err=%v; want hello", v, err)
		}
		return nil
	})
	if err != nil {
		t.Fatalf("sugar.Do: %v", err)
	}

	// After the Do block the arena has released everything it tracked. A
	// tracked value chain has had its VARIANT cleared, so Value() now
	// returns nil. If the chain had leaked (untracked), the BSTR would
	// still be live here.
	v, err := valueChain.Value()
	if err != nil {
		t.Fatalf("after scope: %v", err)
	}
	if v != nil {
		t.Errorf("value chain was not released by the arena: got %v", v)
	}
}

// TestGo_TwoExcelInstancesIsolated is the goroutine-safety regression test
// from the AGENTS.md backlog: two sugar.Go goroutines each drive their own
// Excel instance on their own OS thread, concurrently, and must not
// interfere. Skips when Excel is not installed.
func TestGo_TwoExcelInstancesIsolated(t *testing.T) {
	type result struct {
		hwnd  interface{}
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
			defer excel.Put("DisplayAlerts", false).Call("Quit")
			excel.Put("Visible", false)

			sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
			if err := sheet.Get("Range", "A1").Put("Value", marker).Err(); err != nil {
				out <- result{err: err}
				return nil
			}
			v, err := sheet.Get("Range", "A1").Get("Value").Value()
			if err != nil {
				out <- result{err: err}
				return nil
			}
			hwnd, err := excel.Get("Hwnd").Value()
			out <- result{hwnd: hwnd, value: v, err: err}
			return nil
		})
	}

	ch1 := make(chan result, 1)
	ch2 := make(chan result, 1)
	done1 := run("from-goroutine-1", ch1)
	done2 := run("from-goroutine-2", ch2)

	r1, r2 := <-ch1, <-ch2
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
