//go:build windows

// Regression tests for chain result classification (handleResult / IsDispatch /
// Value). These drive Scripting.Dictionary — present on every Windows host — so
// `go test ./...` runs them without spawning Office processes.

package sugar_test

import (
	"testing"

	"github.com/xll-gen/sugar"
)

// (The VT_UNKNOWN handleResult / Value regression lives in
// chain_unknown_test.go — an internal package test — because a bare IUnknown
// result is not reachable through the public API against Scripting.Dictionary,
// whose _NewEnum is not name-resolvable.)

// TestIsDispatch_ObjectChains pins the item-3 fix: IsDispatch must report true
// for a chain that holds a live IDispatch directly (Create / Fork), not only
// when the last Get/Call produced a VT_DISPATCH result. Before the fix these
// reported false because lastResult was nil.
func TestIsDispatch_ObjectChains(t *testing.T) {
	err := sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			t.Skipf("Scripting.Dictionary unavailable: %v", err)
			return nil
		}

		// Create: disp != nil, lastResult == nil.
		if !dict.IsDispatch() {
			t.Error("Create chain should report IsDispatch() == true")
		}

		// Fork: a fresh reference to the same object, also lastResult == nil.
		fork := dict.Fork()
		if err := fork.Err(); err != nil {
			t.Fatalf("Fork: %v", err)
		}
		if !fork.IsDispatch() {
			t.Error("Fork chain should report IsDispatch() == true")
		}

		// Invariant preserved: a scalar value chain is not a dispatch.
		if dict.Get("Count").IsDispatch() {
			t.Error("scalar value chain (Count) should report IsDispatch() == false")
		}
		return nil
	})
	if err != nil {
		t.Fatalf("sugar.Do: %v", err)
	}
}

// TestIsDispatch_NilChain confirms a nil-dispatch chain (the shape a COM
// Nothing result takes) still reports false after the item-3 widening.
func TestIsDispatch_NilChain(t *testing.T) {
	if sugar.From(nil).IsDispatch() {
		t.Error("nil-dispatch chain should report IsDispatch() == false")
	}
}
