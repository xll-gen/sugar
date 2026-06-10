//go:build windows

// Unit tests for the SAFEARRAY encode path. SafeArray* APIs live in
// oleaut32 and do not require CoInitialize, so encode→decode round trips
// run without any COM server — these are fast, Excel-free tests.

package sugar

import (
	"reflect"
	"testing"
	"time"
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

func TestEncodeDecode_Date(t *testing.T) {
	in := time.Date(2026, 6, 10, 15, 30, 0, 0, time.Local)
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
