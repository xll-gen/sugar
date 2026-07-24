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
