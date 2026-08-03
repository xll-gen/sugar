//go:build windows

// Excel-free regression test for ForEach's handling of a COM `Nothing` item.
//
// A collection is free to enumerate a VT_DISPATCH item whose pointer is NULL
// (VB's `Nothing`). go-ole's VARIANT.ToIDispatch returns a nil *IDispatch for
// that shape, so calling AddRef on it dereferences a nil RawVTable and panics —
// taking down the caller's whole goroutine. ForEach must skip the item instead,
// which is the convention handleResult already implements for a Nothing result.
//
// The test drives the real ForEach code path with a hand-built COM server: a
// fake IDispatch exposing `_NewEnum` and a fake IEnumVARIANT handing out a
// scripted item list. No Excel, and no CoInitialize — every call here is a
// direct vtable dispatch plus oleaut32's VariantClear, neither of which needs an
// initialized apartment (the same reasoning safearray_test.go relies on).

package sugar

import (
	"runtime"
	"strings"
	"syscall"
	"testing"
	"unsafe"

	"github.com/go-ole/go-ole"
)

// COM HRESULTs and DISPIDs used by the fake server.
const (
	fakeSOK                 = 0x00000000
	fakeSFalse              = 0x00000001
	fakeENoInterface        = 0x80004002
	fakeENotImpl            = 0x80004001
	fakeEFail               = 0x80004005
	fakeDispEUnknownName    = 0x80020006
	fakeDispEMemberNotFound = 0x80020003
	fakeDispIDNewEnum       = -4 // DISPID_NEWENUM
)

// iidEnumVARIANT is the IID ForEach queries for ({00020404-...}).
var iidEnumVARIANT = func() *ole.GUID {
	g, err := ole.IIDFromString("{00020404-0000-0000-C000-000000000046}")
	if err != nil {
		panic(err)
	}
	return g
}()

// fakeCollection is a COM object whose first word is a vtable pointer, so it can
// be handed to go-ole as an *ole.IDispatch. It answers `_NewEnum` with its
// embedded enumerator.
type fakeCollection struct {
	vtbl *ole.IDispatchVtbl
	refs int32
	enum *fakeEnum
}

// fakeEnum is the matching IEnumVARIANT: it hands out `items` one at a time.
type fakeEnum struct {
	vtbl  *ole.IEnumVARIANTVtbl
	refs  int32
	items []ole.VARIANT
	pos   int

	// failAtPos > 0 makes Next return a FAILURE HRESULT once pos reaches it,
	// modelling an enumeration that breaks part-way through (a collection mutated
	// under the enumerator, a server that dies mid-iteration). Without this the
	// fake could only ever produce the clean-exhaustion path, so the difference
	// between "ended" and "broke" was untestable -- which is exactly the
	// distinction ForEach used to collapse.
	failAtPos int
}

var (
	fakeCollectionVtbl = &ole.IDispatchVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(fakeCollectionQueryInterface),
			AddRef:         syscall.NewCallback(fakeCollectionAddRef),
			Release:        syscall.NewCallback(fakeCollectionRelease),
		},
		GetTypeInfoCount: syscall.NewCallback(fakeNotImpl2),
		GetTypeInfo:      syscall.NewCallback(fakeNotImpl3),
		GetIDsOfNames:    syscall.NewCallback(fakeCollectionGetIDsOfNames),
		Invoke:           syscall.NewCallback(fakeCollectionInvoke),
	}

	fakeEnumVtbl = &ole.IEnumVARIANTVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(fakeEnumQueryInterface),
			AddRef:         syscall.NewCallback(fakeEnumAddRef),
			Release:        syscall.NewCallback(fakeEnumRelease),
		},
		Next:  syscall.NewCallback(fakeEnumNext),
		Skip:  syscall.NewCallback(fakeNotImpl2),
		Reset: syscall.NewCallback(fakeNotImpl1),
		Clone: syscall.NewCallback(fakeNotImpl2),
	}
)

// newFakeCollection wires a collection to a fresh enumerator. Items are seeded
// afterwards so a test can embed the collection's own pointer as an item.
func newFakeCollection() *fakeCollection {
	return &fakeCollection{
		vtbl: fakeCollectionVtbl,
		enum: &fakeEnum{vtbl: fakeEnumVtbl},
	}
}

// dispatch reinterprets the fake as an *ole.IDispatch (its first field is the
// vtable pointer, matching ole.IUnknown's layout).
func (f *fakeCollection) dispatch() *ole.IDispatch {
	return (*ole.IDispatch)(unsafe.Pointer(f))
}

// variant returns a VT_DISPATCH VARIANT pointing at the fake collection, for use
// as an enumerated item.
func (f *fakeCollection) variant() ole.VARIANT {
	return ole.NewVariant(ole.VT_DISPATCH, int64(uintptr(unsafe.Pointer(f))))
}

// fakePtr reinterprets a COM callback's uintptr argument as a typed Go pointer.
//
// Windows hands `this` and every out-parameter address to a syscall.NewCallback
// thunk as a uintptr, so a COM server written in Go has to convert back. The
// direct `unsafe.Pointer(x)` form trips `go vet`'s unsafeptr check on every such
// line, even though the value really is a live Go pointer here: the fake objects
// stay reachable from the test for its whole duration (see the runtime.KeepAlive
// calls) and Go's GC does not move heap objects. Funnelling the conversion
// through this one helper keeps `go vet ./...` clean and puts the justification
// in exactly one place.
func fakePtr[T any](p uintptr) *T {
	return (*T)(*(*unsafe.Pointer)(unsafe.Pointer(&p)))
}

func fakeNotImpl1(this uintptr) uintptr       { return fakeENotImpl }
func fakeNotImpl2(this, a uintptr) uintptr    { return fakeENotImpl }
func fakeNotImpl3(this, a, b uintptr) uintptr { return fakeENotImpl }

func fakeCollectionQueryInterface(this, iid, ppv uintptr) uintptr {
	obj := fakePtr[fakeCollection](this)
	g := fakePtr[ole.GUID](iid)
	switch {
	case ole.IsEqualGUID(g, ole.IID_IUnknown), ole.IsEqualGUID(g, ole.IID_IDispatch):
		obj.refs++
		*fakePtr[uintptr](ppv) = this
		return fakeSOK
	case ole.IsEqualGUID(g, iidEnumVARIANT):
		obj.enum.refs++
		*fakePtr[uintptr](ppv) = uintptr(unsafe.Pointer(obj.enum))
		return fakeSOK
	}
	*fakePtr[uintptr](ppv) = 0
	return fakeENoInterface
}

func fakeCollectionAddRef(this uintptr) uintptr {
	obj := fakePtr[fakeCollection](this)
	obj.refs++
	return uintptr(obj.refs)
}

func fakeCollectionRelease(this uintptr) uintptr {
	obj := fakePtr[fakeCollection](this)
	obj.refs--
	return uintptr(obj.refs)
}

// fakeCollectionGetIDsOfNames resolves only `_NewEnum`.
func fakeCollectionGetIDsOfNames(this, riid, rgszNames, cNames, lcid, rgDispID uintptr) uintptr {
	name := ole.LpOleStrToString(*fakePtr[*uint16](rgszNames))
	if name != "_NewEnum" {
		return fakeDispEUnknownName
	}
	*fakePtr[int32](rgDispID) = fakeDispIDNewEnum
	return fakeSOK
}

// fakeCollectionInvoke returns the enumerator as a VT_UNKNOWN, AddRef'd, exactly
// as a real collection's `_NewEnum` propget does.
func fakeCollectionInvoke(this, dispID, riid, lcid, flags, dispParams, varResult, excepInfo, argErr uintptr) uintptr {
	obj := fakePtr[fakeCollection](this)
	if int32(dispID) != fakeDispIDNewEnum || varResult == 0 {
		return fakeDispEMemberNotFound
	}
	obj.enum.refs++
	*fakePtr[ole.VARIANT](varResult) = ole.NewVariant(
		ole.VT_UNKNOWN, int64(uintptr(unsafe.Pointer(obj.enum))))
	return fakeSOK
}

func fakeEnumQueryInterface(this, iid, ppv uintptr) uintptr {
	obj := fakePtr[fakeEnum](this)
	g := fakePtr[ole.GUID](iid)
	if ole.IsEqualGUID(g, ole.IID_IUnknown) || ole.IsEqualGUID(g, iidEnumVARIANT) {
		obj.refs++
		*fakePtr[uintptr](ppv) = this
		return fakeSOK
	}
	*fakePtr[uintptr](ppv) = 0
	return fakeENoInterface
}

func fakeEnumAddRef(this uintptr) uintptr {
	obj := fakePtr[fakeEnum](this)
	obj.refs++
	return uintptr(obj.refs)
}

func fakeEnumRelease(this uintptr) uintptr {
	obj := fakePtr[fakeEnum](this)
	obj.refs--
	return uintptr(obj.refs)
}

// fakeEnumNext hands out the next scripted item, AddRef'ing object items per the
// IEnumVARIANT contract (the caller owns the returned reference).
func fakeEnumNext(this, celt, rgVar, pceltFetched uintptr) uintptr {
	obj := fakePtr[fakeEnum](this)
	if pceltFetched != 0 {
		*fakePtr[uint32](pceltFetched) = 0
	}
	if obj.failAtPos > 0 && obj.pos >= obj.failAtPos {
		return fakeEFail
	}
	if obj.pos >= len(obj.items) || celt == 0 || rgVar == 0 {
		return fakeSFalse
	}
	item := obj.items[obj.pos]
	obj.pos++
	if item.VT == ole.VT_DISPATCH && item.Val != 0 {
		fakePtr[fakeCollection](uintptr(item.Val)).refs++
	}
	*fakePtr[ole.VARIANT](rgVar) = item
	if pceltFetched != 0 {
		*fakePtr[uint32](pceltFetched) = 1
	}
	return fakeSOK
}

// TestForEach_SkipsNullDispatchItem is the regression: the first enumerated item
// is a VT_DISPATCH carrying a NULL pointer (COM `Nothing`). Before the fix
// ForEach called AddRef on the nil *IDispatch and panicked; it must instead skip
// the item and keep iterating, delivering only the real object.
func TestForEach_SkipsNullDispatchItem(t *testing.T) {
	coll := newFakeCollection()
	coll.enum.items = []ole.VARIANT{
		ole.NewVariant(ole.VT_DISPATCH, 0), // COM Nothing -> skipped, no panic
		coll.variant(),                     // a real object -> delivered
		ole.NewVariant(ole.VT_I4, 42),      // a scalar     -> skipped (pre-existing)
	}

	c := &chain{disp: coll.dispatch()}

	var delivered int
	res := c.ForEach(func(item Chain) error {
		delivered++
		if !item.IsDispatch() {
			t.Errorf("item %d: expected a dispatch chain", delivered)
		}
		return nil
	})
	if err := res.Err(); err != nil {
		t.Fatalf("ForEach returned an error: %v", err)
	}
	if delivered != 1 {
		t.Errorf("callback ran %d times, want 1 (the null item and the scalar must be skipped)", delivered)
	}
	if coll.enum.pos != len(coll.enum.items) {
		t.Errorf("enumerator stopped after %d of %d items", coll.enum.pos, len(coll.enum.items))
	}
	// Every reference the fake handed out must have come back: the null item
	// takes no reference, the delivered object is Released by the item chain, and
	// the enumerator is Released by ForEach's defers.
	if coll.refs != 0 {
		t.Errorf("collection refcount leaked/underflowed: %d, want 0", coll.refs)
	}
	if coll.enum.refs != 0 {
		t.Errorf("enumerator refcount leaked/underflowed: %d, want 0", coll.enum.refs)
	}
	runtime.KeepAlive(coll)
}

// TestForEach_AllNullDispatchItems is the degenerate case: a collection whose
// every item is Nothing iterates to completion with no callback and no error.
func TestForEach_AllNullDispatchItems(t *testing.T) {
	coll := newFakeCollection()
	coll.enum.items = []ole.VARIANT{
		ole.NewVariant(ole.VT_DISPATCH, 0),
		ole.NewVariant(ole.VT_DISPATCH, 0),
	}

	c := &chain{disp: coll.dispatch()}
	called := false
	res := c.ForEach(func(item Chain) error {
		called = true
		return nil
	})
	if err := res.Err(); err != nil {
		t.Fatalf("ForEach returned an error: %v", err)
	}
	if called {
		t.Error("callback ran for a COM Nothing item")
	}
	if coll.enum.refs != 0 {
		t.Errorf("enumerator refcount leaked/underflowed: %d, want 0", coll.enum.refs)
	}
	runtime.KeepAlive(coll)
}

// TestForEach_PropagatesEnumNextError is the regression for the silent truncation
// (2026-08-03).
//
// enum.Next() reports two different things through one return path: "no more items"
// (S_FALSE, fetched == 0) and "the enumeration FAILED". ForEach used to `break` on
// both, then `return c` -- the receiver -- so a collection that broke after two of
// four items produced two callbacks, a nil Err(), and a caller convinced it had seen
// everything. There is no way to detect the loss from the outside: the callback count
// is whatever it is, and a short collection is a legitimate outcome.
//
// The scripted enumerator here hands out two items and then fails, so the assertion
// is exactly the distinction: the callback saw the first two, AND Err() is non-nil.
func TestForEach_PropagatesEnumNextError(t *testing.T) {
	coll := newFakeCollection()
	coll.enum.items = []ole.VARIANT{
		coll.variant(),
		coll.variant(),
		coll.variant(), // never reached: Next fails before handing this one out
	}
	coll.enum.failAtPos = 2

	c := &chain{disp: coll.dispatch()}

	var delivered int
	res := c.ForEach(func(item Chain) error {
		delivered++
		return nil
	})

	if err := res.Err(); err == nil {
		t.Fatalf("ForEach reported success after IEnumVARIANT.Next FAILED at item %d; "+
			"the caller processed %d of %d items believing the enumeration was complete",
			coll.enum.failAtPos, delivered, len(coll.enum.items))
	} else if !strings.Contains(err.Error(), "ForEach") {
		t.Errorf("error should name the operation that failed, got %q", err.Error())
	}
	if delivered != 2 {
		t.Errorf("callback ran %d times, want 2 (the items delivered before the failure)", delivered)
	}
	// The failure path must not leak: both items handed out are Released by their
	// item chains, and the enumerator by ForEach's defer, error or not.
	if coll.refs != 0 {
		t.Errorf("collection refcount leaked/underflowed on the error path: %d, want 0", coll.refs)
	}
	if coll.enum.refs != 0 {
		t.Errorf("enumerator refcount leaked/underflowed on the error path: %d, want 0", coll.enum.refs)
	}
	runtime.KeepAlive(coll)
}

// TestForEach_CleanExhaustionStillSucceeds pins the other side of the split: normal
// exhaustion (S_FALSE) must remain a SUCCESS. Without this, "propagate the error"
// could be satisfied by failing every enumeration.
func TestForEach_CleanExhaustionStillSucceeds(t *testing.T) {
	coll := newFakeCollection()
	coll.enum.items = []ole.VARIANT{coll.variant(), coll.variant()}
	// failAtPos left 0: the enumerator exhausts cleanly.

	c := &chain{disp: coll.dispatch()}
	var delivered int
	res := c.ForEach(func(item Chain) error {
		delivered++
		return nil
	})
	if err := res.Err(); err != nil {
		t.Fatalf("clean exhaustion must not be an error, got %v", err)
	}
	if delivered != 2 {
		t.Errorf("callback ran %d times, want 2", delivered)
	}
	if coll.refs != 0 || coll.enum.refs != 0 {
		t.Errorf("refcounts: collection=%d enum=%d, want 0/0", coll.refs, coll.enum.refs)
	}
	runtime.KeepAlive(coll)
}
