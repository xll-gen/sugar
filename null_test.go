//go:build windows

// Excel-free unit tests for the VT_NULL sentinel. Building VARIANTs and
// SAFEARRAYs needs no COM server and no CoInitialize, so these run under plain
// `go test ./...`.

package sugar

import (
	"testing"

	"github.com/go-ole/go-ole"
)

// TestDecodeVariantScalar_Null pins the gap this type exists to close: go-ole's
// (*VARIANT).Value() has no VT_NULL case, so a "no single value" result — what
// Excel returns from a scalar property read on a range whose cells disagree —
// decoded to a bare nil, byte-for-byte identical to VT_EMPTY ("no value"). The
// typed getters downstream then coerced that nil to "" / 0 / false and returned
// it with a nil error.
func TestDecodeVariantScalar_Null(t *testing.T) {
	v := ole.NewVariant(ole.VT_NULL, 0)
	if raw := v.Value(); raw != nil {
		t.Fatalf("precondition: go-ole VT_NULL Value() now returns %v (%T); test premise stale", raw, raw)
	}
	got := decodeVariantScalar(&v)
	if got == nil {
		t.Fatalf("VT_NULL decoded to a bare nil — indistinguishable from VT_EMPTY")
	}
	if _, ok := got.(Null); !ok {
		t.Fatalf("VT_NULL decode: got %v (%T), want sugar.Null", got, got)
	}
	if got != interface{}(Null{}) {
		t.Errorf("Null must be comparable by value: %v != Null{}", got)
	}
}

// TestDecodeVariantScalar_EmptyStaysNil is the mandatory other half: the guard
// must not be satisfiable by turning every absent value into Null. VT_EMPTY is
// a genuinely empty cell and stays nil, which is what every existing consumer
// (and the Options Empty() substitution) keys on.
func TestDecodeVariantScalar_EmptyStaysNil(t *testing.T) {
	var v ole.VARIANT
	ole.VariantInit(&v)
	if got := decodeVariantScalar(&v); got != nil {
		t.Errorf("VT_EMPTY decode: got %v (%T), want nil", got, got)
	}
	// The omitted-optional marker is also not Null.
	missing := ole.NewVariant(ole.VT_ERROR, dispEParamNotFound)
	if got := decodeVariantScalar(&missing); got != nil {
		t.Errorf("DISP_E_PARAMNOTFOUND decode: got %v (%T), want nil", got, got)
	}
}

// TestIsNull pins the exported recognizer. The rows that matter are the
// negatives: nil (VT_EMPTY) and CellError (VT_ERROR) are the two values a
// caller is most likely to conflate with Null, and all three used to be
// reachable only as "nil or not nil".
func TestIsNull(t *testing.T) {
	cases := []struct {
		name string
		v    interface{}
		want bool
	}{
		{"Null value", Null{}, true},
		{"Null pointer is not Null", &Null{}, false},
		{"nil (VT_EMPTY)", nil, false},
		{"empty string", "", false},
		{"zero float", 0.0, false},
		{"false", false, false},
		{"cell error", CellError{SCode: 0x800A07D7}, false},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := IsNull(tc.v); got != tc.want {
				t.Errorf("IsNull(%#v) = %v, want %v", tc.v, got, tc.want)
			}
		})
	}
}

// TestNull_String keeps the rendering distinct from CellError's "#NULL!", which
// is the unrelated worksheet error value xlErrNull (cvErr 2000).
func TestNull_String(t *testing.T) {
	if got := (Null{}).String(); got != "Null" {
		t.Errorf("Null.String() = %q, want %q", got, "Null")
	}
	if got := (CellError{SCode: 0x800A07D0}).String(); got != "#NULL!" {
		t.Errorf("precondition: CellError 2000 renders %q; the two Null-ish values must stay distinct", got)
	}
}

// TestSafeArray_NullCellRoundTrips covers the array half of the decoder. A
// VT_NULL cell inside a SAFEARRAY is not an Excel `Range.Value` shape (blank
// cells arrive as VT_EMPTY there), but it is the ordinary representation of a
// SQL NULL in the 2-D array an ADO-style COM server hands back — and sugar's
// core is a general COM layer. Decode must keep it distinct from a blank, and
// encode must be able to write it back, otherwise reading such a grid and
// writing it out again would fail with "unsupported cell type".
func TestSafeArray_NullCellRoundTrips(t *testing.T) {
	src := []interface{}{Null{}, nil, "x"}
	v, err := encodeVariantArray(src)
	if err != nil {
		t.Fatalf("encodeVariantArray with a Null cell: %v", err)
	}
	defer v.Clear()

	got, err := decodeVariantArray(v)
	if err != nil {
		t.Fatalf("decodeVariantArray: %v", err)
	}
	row, ok := got.([]interface{})
	if !ok {
		t.Fatalf("decode shape: got %T, want []interface{}", got)
	}
	if len(row) != 3 {
		t.Fatalf("decode length: got %d, want 3", len(row))
	}
	if !IsNull(row[0]) {
		t.Errorf("cell 0: got %v (%T), want Null", row[0], row[0])
	}
	if row[1] != nil {
		t.Errorf("cell 1 (VT_EMPTY): got %v (%T), want nil", row[1], row[1])
	}
	if row[2] != "x" {
		t.Errorf("cell 2: got %v (%T), want \"x\"", row[2], row[2])
	}
}
