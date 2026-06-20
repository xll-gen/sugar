//go:build windows

// Excel-free regression tests for the shared collection base (R38).
//
// These pin the COM call each collection's Count/Item/Add delegates to, using
// a fake sugar.Chain that records Get/Call/Put invocations instead of touching
// real Excel. They run under plain `go test ./...` (no excel_integration tag),
// so a refactor that changed *which* COM verb (Get vs Call) or *which* member
// ("Count"/"Item"/"Add") a collection uses would fail here.

package excel

import (
	"reflect"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// recordedCall captures one COM verb dispatched against a fakeChain.
type recordedCall struct {
	verb   string // "Get", "Call", or "Put"
	member string
	args   []interface{}
}

// fakeChain is a minimal sugar.Chain that records Get/Call/Put and returns a
// child chain sharing the same log. Value() returns a fixed scalar so Count's
// getInt32 path resolves without Excel. Children fan out from a root so a whole
// chain of operations records into one ordered log.
type fakeChain struct {
	root *fakeRoot
}

type fakeRoot struct {
	calls   []recordedCall
	value   interface{} // returned by Value() (Count reads this)
	valErr  error
	callErr error // seeded into child chains' Err()
}

func newFakeChain() *fakeChain {
	return &fakeChain{root: &fakeRoot{value: int32(7)}}
}

func (f *fakeChain) record(verb, member string, args []interface{}) sugar.Chain {
	f.root.calls = append(f.root.calls, recordedCall{verb: verb, member: member, args: args})
	return &fakeChain{root: f.root}
}

func (f *fakeChain) Get(prop string, params ...interface{}) sugar.Chain {
	return f.record("Get", prop, params)
}
func (f *fakeChain) Call(method string, params ...interface{}) sugar.Chain {
	return f.record("Call", method, params)
}
func (f *fakeChain) Put(prop string, params ...interface{}) sugar.Chain {
	return f.record("Put", prop, params)
}
func (f *fakeChain) ForEach(cb func(item sugar.Chain) error) sugar.Chain { return f }
func (f *fakeChain) Fork() sugar.Chain                                   { return &fakeChain{root: f.root} }
func (f *fakeChain) Store() (*ole.IDispatch, error)                      { return nil, nil }
func (f *fakeChain) Release() error                                      { return nil }
func (f *fakeChain) IsDispatch() bool                                    { return true }
func (f *fakeChain) Value() (interface{}, error)                         { return f.root.value, f.root.valErr }
func (f *fakeChain) Err() error                                          { return f.root.callErr }

// lastCall returns the most recently recorded verb on the root log.
func (f *fakeChain) lastCall(t *testing.T) recordedCall {
	t.Helper()
	if len(f.root.calls) == 0 {
		t.Fatal("no COM calls recorded")
	}
	return f.root.calls[len(f.root.calls)-1]
}

func (f *fakeChain) findCall(t *testing.T, verb, member string) recordedCall {
	t.Helper()
	for _, c := range f.root.calls {
		if c.verb == verb && c.member == member {
			return c
		}
	}
	t.Fatalf("expected a %s(%q) call; recorded: %+v", verb, member, f.root.calls)
	return recordedCall{}
}

// wantCall asserts the most recent recorded call matches verb+member+args.
func wantCall(t *testing.T, got recordedCall, verb, member string, args ...interface{}) {
	t.Helper()
	if got.verb != verb || got.member != member {
		t.Errorf("verb/member: got %s(%q), want %s(%q)", got.verb, got.member, verb, member)
	}
	if len(args) == 0 {
		if len(got.args) != 0 {
			t.Errorf("args: got %+v, want none", got.args)
		}
		return
	}
	if !reflect.DeepEqual(got.args, args) {
		t.Errorf("args: got %+v, want %+v", got.args, args)
	}
}

// --- Count delegates to getInt32(source, "Count") for all six ---

func TestCollections_CountDelegatesToGetCount(t *testing.T) {
	cases := []struct {
		name  string
		build func(*fakeChain) interface{ Count() (int32, error) }
	}{
		{"charts", func(f *fakeChain) interface{ Count() (int32, error) } { return wrapCharts(f) }},
		{"workbooks", func(f *fakeChain) interface{ Count() (int32, error) } { return wrapWorkbooks(f) }},
		{"worksheets", func(f *fakeChain) interface{ Count() (int32, error) } { return wrapWorksheets(f) }},
		{"names", func(f *fakeChain) interface{ Count() (int32, error) } { return wrapNames(f) }},
		{"shapes", func(f *fakeChain) interface{ Count() (int32, error) } { return wrapShapes(f) }},
		// pictures re-fetches the snapshot collection first (sheet.Call("Pictures")),
		// then reads Count off it; covered separately below for the extra hop.
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			f := newFakeChain()
			f.root.value = int32(42)
			coll := tc.build(f)
			n, err := coll.Count()
			if err != nil {
				t.Fatalf("Count error: %v", err)
			}
			if n != 42 {
				t.Errorf("Count value: got %d, want 42 (from faked VARIANT)", n)
			}
			wantCall(t, f.findCall(t, "Get", "Count"), "Get", "Count")
		})
	}
}

// --- Item: Charts/Names/Shapes use Call("Item"); Workbooks/Worksheets use Get("Item") ---

func TestCollections_ItemVerb(t *testing.T) {
	t.Run("charts-call", func(t *testing.T) {
		f := newFakeChain()
		wrapCharts(f).Item(int32(2))
		wantCall(t, f.lastCall(t), "Call", "Item", int32(2))
	})
	t.Run("names-call", func(t *testing.T) {
		f := newFakeChain()
		wrapNames(f).Item("Region")
		wantCall(t, f.lastCall(t), "Call", "Item", "Region")
	})
	t.Run("shapes-call", func(t *testing.T) {
		f := newFakeChain()
		wrapShapes(f).Item(int32(3))
		wantCall(t, f.lastCall(t), "Call", "Item", int32(3))
	})
	t.Run("workbooks-get", func(t *testing.T) {
		f := newFakeChain()
		wrapWorkbooks(f).Item("Book1.xlsx")
		wantCall(t, f.lastCall(t), "Get", "Item", "Book1.xlsx")
	})
	t.Run("worksheets-get", func(t *testing.T) {
		f := newFakeChain()
		wrapWorksheets(f).Item(int32(1))
		wantCall(t, f.lastCall(t), "Get", "Item", int32(1))
	})
}

// --- Pictures re-fetches the snapshot collection (sheet.Call("Pictures")) before each Item/Count ---

func TestPictures_ItemAndCountReFetchSnapshot(t *testing.T) {
	t.Run("item", func(t *testing.T) {
		sheet := newFakeChain()
		snap := newFakeChain() // the chain passed as the snapshot anchor
		wrapPictures(snap, sheet).Item(int32(1))
		// Item must go sheet.Call("Pictures") then .Call("Item", 1) on the result.
		wantCall(t, sheet.findCall(t, "Call", "Pictures"), "Call", "Pictures")
		wantCall(t, sheet.findCall(t, "Call", "Item"), "Call", "Item", int32(1))
	})
	t.Run("count", func(t *testing.T) {
		sheet := newFakeChain()
		sheet.root.value = int32(5)
		snap := newFakeChain()
		n, err := wrapPictures(snap, sheet).Count()
		if err != nil || n != 5 {
			t.Fatalf("Count: got %d err=%v, want 5", n, err)
		}
		wantCall(t, sheet.findCall(t, "Call", "Pictures"), "Call", "Pictures")
		wantCall(t, sheet.findCall(t, "Get", "Count"), "Get", "Count")
	})
}

// --- Add wraps a Call("Add", ...) result (the verb/member is preserved) ---

func TestCollections_AddDelegatesToCallAdd(t *testing.T) {
	t.Run("charts", func(t *testing.T) {
		f := newFakeChain()
		wrapCharts(f).Add(ChartAt(10, 20), ChartSize(300, 200))
		wantCall(t, f.findCall(t, "Call", "Add"), "Call", "Add", 10.0, 20.0, 300.0, 200.0)
	})
	t.Run("workbooks", func(t *testing.T) {
		f := newFakeChain()
		wrapWorkbooks(f).Add()
		wantCall(t, f.findCall(t, "Call", "Add"), "Call", "Add")
	})
	t.Run("worksheets", func(t *testing.T) {
		f := newFakeChain()
		wrapWorksheets(f).Add()
		wantCall(t, f.findCall(t, "Call", "Add"), "Call", "Add")
	})
	t.Run("names", func(t *testing.T) {
		f := newFakeChain()
		wrapNames(f).Add("Region", "=Sheet1!$A$1")
		wantCall(t, f.findCall(t, "Call", "Add"), "Call", "Add", "Region", "=Sheet1!$A$1")
	})
}
