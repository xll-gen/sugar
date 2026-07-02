//go:build windows

// Unit tests for the SAFEARRAY encode path. SafeArray* APIs live in
// oleaut32 and do not require CoInitialize, so encode→decode round trips
// run without any COM server — these are fast, Excel-free tests.

package sugar

import (
	"math"
	"reflect"
	"testing"
	"time"

	// Embed the IANA zone database so LoadLocation works deterministically on
	// any host (Windows ships no /usr/share/zoneinfo, and CI images vary). The
	// VT_DATE zone-drift tests below need real zones; this keeps them from
	// flaking to t.Skip. Test-only import — it does not bloat the library.
	_ "time/tzdata"

	ole "github.com/go-ole/go-ole"
)

// roundTrip encodes a Go slice and decodes it back through the same
// machinery Range.Value uses.
func roundTrip(t *testing.T, in interface{}) interface{} {
	t.Helper()
	v, err := encodeVariantArray(in)
	if err != nil {
		t.Fatalf("encodeVariantArray(%T): %v", in, err)
	}
	defer v.Clear()
	out, err := decodeVariantArray(v)
	if err != nil {
		t.Fatalf("decodeVariantArray: %v", err)
	}
	return out
}

func TestEncodeDecode_Grid(t *testing.T) {
	in := [][]interface{}{
		{"name", 1.5, true},
		{nil, -2.0, "x"},
	}
	got := roundTrip(t, in)
	if !reflect.DeepEqual(got, in) {
		t.Errorf("got %v, want %v", got, in)
	}
}

func TestEncodeDecode_Vector(t *testing.T) {
	in := []interface{}{1.0, "two", false}
	got := roundTrip(t, in)
	if !reflect.DeepEqual(got, in) {
		t.Errorf("got %v, want %v", got, in)
	}
}

// TestEncodeDecode_TypedSlices covers the reflect widening path: typed Go
// slices encode like their []interface{} counterparts. Integers come back
// as float64 (VT_R8) — Excel's native number representation.
func TestEncodeDecode_TypedSlices(t *testing.T) {
	got := roundTrip(t, [][]float64{{1.5, 2.5}, {3.5, 4.5}})
	want := [][]interface{}{{1.5, 2.5}, {3.5, 4.5}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("[][]float64: got %v, want %v", got, want)
	}

	got = roundTrip(t, []int{1, 2, 3})
	want1 := []interface{}{1.0, 2.0, 3.0}
	if !reflect.DeepEqual(got, want1) {
		t.Errorf("[]int: got %v, want %v", got, want1)
	}

	got = roundTrip(t, [][]string{{"a", "b"}, {"c", "d"}})
	want = [][]interface{}{{"a", "b"}, {"c", "d"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("[][]string: got %v, want %v", got, want)
	}
}

// TestEncodeDecode_FormulaGrid documents the marshaling path used by
// Range.SetFormula2Array: a column of formula strings (the shape the showcase
// build writes to B5:B15) survives the SAFEARRAY encode→decode round trip
// unchanged. This is Excel-free because SafeArray* APIs need no CoInitialize.
func TestEncodeDecode_FormulaGrid(t *testing.T) {
	in := [][]interface{}{
		{"=Add(2,3)"},
		{"=Multiply(1.5,4)"},
		{`=Greet("Excel")`},
	}
	got := roundTrip(t, in)
	if !reflect.DeepEqual(got, in) {
		t.Errorf("formula grid round-trip: got %v, want %v", got, in)
	}
}

func TestEncodeDecode_Date(t *testing.T) {
	// Use an explicit DST zone rather than time.Local so the test is
	// deterministic on any host and actually exercises the zone-drift path
	// (America/New_York's UTC offset differs between the 1899 epoch and 2026).
	loc, err := time.LoadLocation("America/New_York")
	if err != nil {
		t.Skipf("America/New_York unavailable: %v", err)
	}
	in := time.Date(2026, 6, 10, 15, 30, 0, 0, loc)
	out := roundTrip(t, []interface{}{in})
	cells, ok := out.([]interface{})
	if !ok || len(cells) != 1 {
		t.Fatalf("unexpected shape %T %v", out, out)
	}
	ts, ok := cells[0].(time.Time)
	if !ok {
		t.Fatalf("expected time.Time, got %T", cells[0])
	}
	if ts.Year() != 2026 || ts.Month() != 6 || ts.Day() != 10 ||
		ts.Hour() != 15 || ts.Minute() != 30 {
		t.Errorf("date mismatch: got %v, want %v", ts, in)
	}
}

func TestEncode_RaggedRowsError(t *testing.T) {
	_, err := encodeVariantArray([][]interface{}{{1.0, 2.0}, {3.0}})
	if err == nil {
		t.Error("expected error for ragged rows")
	}
}

func TestEncode_UnsupportedCellError(t *testing.T) {
	_, err := encodeVariantArray([]interface{}{struct{}{}})
	if err == nil {
		t.Error("expected error for unsupported cell type")
	}
}

// dateSerial encodes t through scalarToVariant and returns the raw OLE VT_DATE
// serial (day count since 1899-12-30). Excel stores dates as this double.
func dateSerial(t *testing.T, ts time.Time) float64 {
	t.Helper()
	var out ole.VARIANT
	ole.VariantInit(&out)
	if err := scalarToVariant(ts, &out); err != nil {
		t.Fatalf("scalarToVariant(%v): %v", ts, err)
	}
	if out.VT != ole.VT_DATE {
		t.Fatalf("expected VT_DATE, got VT=%d", out.VT)
	}
	return math.Float64frombits(uint64(out.Val))
}

// TestScalarToVariant_DateZoneDrift is the regression test for the VT_DATE
// wall-clock encoding bug. scalarToVariant must encode a date by its wall-clock
// fields, zone-independent: midnight on a given day must map to that day's
// integer Excel serial in every zone, never to x.xx that Excel rounds back to
// the previous day (23:xx). The pre-fix code subtracted two absolute instants
// in x.Location(), folding in the offset difference between x's date and the
// 1899 epoch — 60 min in DST zones, +08:27:52 LMT (≈32 min) for IANA
// Asia/Seoul — which pushed midnight into the previous day.
func TestScalarToVariant_DateZoneDrift(t *testing.T) {
	utcMidnight := dateSerial(t, time.Date(2026, 6, 10, 0, 0, 0, 0, time.UTC))
	if utcMidnight != math.Trunc(utcMidnight) {
		t.Fatalf("UTC midnight is not an integer serial: %v", utcMidnight)
	}

	for _, name := range []string{"America/New_York", "Europe/Berlin", "Asia/Seoul"} {
		loc, err := time.LoadLocation(name)
		if err != nil {
			t.Skipf("zone %s unavailable: %v", name, err)
		}
		got := dateSerial(t, time.Date(2026, 6, 10, 0, 0, 0, 0, loc))
		if got != math.Trunc(got) {
			t.Errorf("%s: midnight encoded to non-integer serial %v (zone drift)", name, got)
		}
		if got != utcMidnight {
			t.Errorf("%s: serial %v != UTC-midnight serial %v (zone drift)", name, got, utcMidnight)
		}
	}
}

// TestScalarToVariant_DateWallClock checks that a non-midnight wall-clock time
// survives with its hour/minute intact regardless of zone — the property the
// 313-314 comment promises ("what the user sees is what Excel shows").
func TestScalarToVariant_DateWallClock(t *testing.T) {
	ref := dateSerial(t, time.Date(2026, 6, 10, 15, 30, 0, 0, time.UTC))
	for _, name := range []string{"America/New_York", "Asia/Seoul"} {
		loc, err := time.LoadLocation(name)
		if err != nil {
			t.Skipf("zone %s unavailable: %v", name, err)
		}
		got := dateSerial(t, time.Date(2026, 6, 10, 15, 30, 0, 0, loc))
		if math.Abs(got-ref) > 1e-9 {
			t.Errorf("%s: 15:30 wall-clock serial %v != UTC 15:30 serial %v", name, got, ref)
		}
	}
}

func TestNeedsArrayEncoding(t *testing.T) {
	cases := []struct {
		in   interface{}
		want bool
	}{
		{[][]interface{}{}, true},
		{[]interface{}{}, true},
		{[][]float64{}, true},
		{[]int{1}, true},
		{[]byte("x"), false},   // go-ole native VT_UI1 array
		{[]string{"x"}, false}, // go-ole native VT_BSTR array
		{"scalar", false},
		{42, false},
		{nil, false},
	}
	for _, c := range cases {
		if got := needsArrayEncoding(c.in); got != c.want {
			t.Errorf("needsArrayEncoding(%T) = %v, want %v", c.in, got, c.want)
		}
	}
}
