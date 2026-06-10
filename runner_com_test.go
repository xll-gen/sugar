//go:build windows

// Regression tests for the Runner's COM-initialization handling. These use
// lightweight scripting COM objects (Scripting.Dictionary) so they run fast
// on any Windows host with no Excel required. The Excel-bound goroutine-
// isolation test lives in runner_excel_com_test.go behind the
// excel_integration build tag.

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
