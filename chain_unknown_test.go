//go:build windows

// Internal (package sugar) regression tests for the item-1 VT_UNKNOWN handling.
// They live inside the package because a bare IUnknown result is not reachable
// through the public API against the Excel-free Scripting.Dictionary server
// (its _NewEnum enumerator is not name-resolvable), so the tests build the
// VT_UNKNOWN VARIANT directly and drive handleResult / Value. A Scripting
// object doubles as a live IUnknown that also supports IDispatch.

package sugar

import (
	"testing"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// TestHandleResult_VTUnknownPromotedToDispatch pins the promotion path: a
// VT_UNKNOWN result whose IUnknown supports IDispatch must become a dispatch
// chain (Store-able, IsDispatch true), not a value chain that would hand back a
// dangling raw pointer.
func TestHandleResult_VTUnknownPromotedToDispatch(t *testing.T) {
	err := Do(func(ctx Context) error {
		recv, ok := ctx.Create("Scripting.Dictionary").(*chain)
		if !ok || recv.err != nil {
			t.Skipf("Scripting.Dictionary unavailable: ok=%v err=%v", ok, recv.err)
			return nil
		}

		// A second live object, handed to handleResult as a VT_UNKNOWN VARIANT
		// that owns the CreateObject reference.
		unk, e := oleutil.CreateObject("Scripting.Dictionary")
		if e != nil {
			t.Skipf("CreateObject: %v", e)
			return nil
		}
		v := ole.NewVariant(ole.VT_UNKNOWN, int64(uintptr(unsafe.Pointer(unk))))

		out := recv.handleResult(&v, nil)
		if err := out.Err(); err != nil {
			t.Fatalf("handleResult(VT_UNKNOWN): %v", err)
		}
		if !out.IsDispatch() {
			t.Error("a VT_UNKNOWN result that QIs to IDispatch must be a dispatch chain")
		}
		// The promoted chain holds disp directly with no lastResult (the original
		// VARIANT was Cleared), so Value() reports (nil, nil) like a From/Fork
		// chain — crucially NOT a live raw interface pointer.
		if v, err := out.Value(); err != nil || v != nil {
			t.Errorf("promoted chain Value() = (%v, %v); want (nil, nil)", v, err)
		}
		d, err := out.Store()
		if err != nil {
			t.Fatalf("Store on promoted chain: %v", err)
		}
		d.Release()
		return nil
	})
	if err != nil {
		t.Fatalf("Do: %v", err)
	}
}

// TestValue_VTUnknownDemotesToNil pins the defensive Value() branch: a chain
// whose lastResult is a bare VT_UNKNOWN must return nil rather than the raw
// interface pointer (which go-ole's Value() would surface without an AddRef).
func TestValue_VTUnknownDemotesToNil(t *testing.T) {
	err := Do(func(ctx Context) error {
		unk, e := oleutil.CreateObject("Scripting.Dictionary")
		if e != nil {
			t.Skipf("CreateObject: %v", e)
			return nil
		}
		v := ole.NewVariant(ole.VT_UNKNOWN, int64(uintptr(unsafe.Pointer(unk))))
		c := &chain{lastResult: &v}

		got, err := c.Value()
		if err != nil {
			t.Fatalf("Value(): %v", err)
		}
		if got != nil {
			t.Errorf("VT_UNKNOWN Value() should demote to nil, got %T %v", got, got)
		}

		v.Clear() // release the object reference the VARIANT held
		return nil
	})
	if err != nil {
		t.Fatalf("Do: %v", err)
	}
}

// TestVTUnknownChain_IsIndistinguishableFromNothing pins the fact the
// expression package's comparison contract rests on: through the exported Chain
// interface, a chain whose lastResult is a bare VT_UNKNOWN presents EXACTLY the
// same three answers as COM `Nothing`.
//
//	IsDispatch() == false   Store() == error   Value() == (nil, nil)
//
// That is why expression's object-operand guard (which tests IsDispatch) cannot
// refuse such an operand, and why `x == nil` answers TRUE for it — see
// expression/operator_contract_test.go
// TestComparison_NonDispatchObjectChainComparesEqualToNil, which models this
// shape from outside the package. Neither test proves the contract alone: this
// one shows a real chain has the shape, that one shows what the engine does with
// it.
//
// The state is not reachable through a COM call — handleResult promotes a
// QI-able IUnknown to a dispatch chain and degrades the rest to an empty chain
// — so this builds it directly, the same way TestValue_VTUnknownDemotesToNil
// does. Widening IsDispatch to cover VT_UNKNOWN fails this test — and THIS TEST
// ONLY.
//
// (Corrected 2026-08-03. The wording here used to add "AND breaks the
// IsDispatch/Store pairing that TestHandleResult_VTUnknownPromotedToDispatch
// asserts, since Store reads disp". That is false, and it was checked by
// running it: adding `|| c.lastResult.VT == ole.VT_UNKNOWN` to IsDispatch
// (sugar.go:635) leaves TestHandleResult_VTUnknownPromotedToDispatch PASSING.
// The promoted chain that test builds holds `disp` directly with lastResult nil
// — handleResult Clear()s the source VARIANT and returns &chain{disp: newDisp}
// (sugar.go:214, 221-224) — so neither IsDispatch nor Store ever reaches the
// widened arm there. If you make that change, THIS is the only test that stops
// you, which is exactly why it exists.)
//
// The objection to widening is therefore semantic, not test-coverage: Store()
// reads `disp` (sugar.go:602) and would keep answering "nil dispatch", so a
// widened IsDispatch would produce a chain that claims to be an object and then
// refuses to hand one over.
func TestVTUnknownChain_IsIndistinguishableFromNothing(t *testing.T) {
	err := Do(func(ctx Context) error {
		unk, e := oleutil.CreateObject("Scripting.Dictionary")
		if e != nil {
			t.Skipf("CreateObject: %v", e)
			return nil
		}
		v := ole.NewVariant(ole.VT_UNKNOWN, int64(uintptr(unsafe.Pointer(unk))))
		defer v.Clear()

		unknownChain := &chain{lastResult: &v}
		// The Nothing chain handleResult produces for a NULL VT_DISPATCH result
		// and for a non-IDispatch-capable IUnknown: no disp, no lastResult.
		nothingChain := &chain{}

		for _, tc := range []struct {
			name string
			c    *chain
		}{
			{"VT_UNKNOWN", unknownChain},
			{"Nothing", nothingChain},
		} {
			if tc.c.IsDispatch() {
				t.Errorf("%s: IsDispatch() = true, want false", tc.name)
			}
			if d, err := tc.c.Store(); err == nil {
				d.Release()
				t.Errorf("%s: Store() succeeded, want an error", tc.name)
			}
			if got, err := tc.c.Value(); err != nil || got != nil {
				t.Errorf("%s: Value() = (%v, %v), want (nil, nil)", tc.name, got, err)
			}
		}
		return nil
	})
	if err != nil {
		t.Fatalf("Do: %v", err)
	}
}
