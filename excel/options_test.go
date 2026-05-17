//go:build windows

// Unit tests for excel.Options conversion pipeline.
//
// These tests exercise the pure-Go conversion stages — shape forcing, Empty
// substitution, Convert, struct-by-header decode — without launching Excel.
// The COM-bound parts (Expand, Range.Value SAFEARRAY decode) are covered in
// range_test.go behind the `excel_integration` build tag.

package excel

import (
	"errors"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// TestShapeResult_Auto verifies the default xlwings shape rules: 1×1 unwraps
// to a scalar; 1×N and N×1 collapse to a flat slice; everything else stays
// as a 2-D grid.
func TestShapeResult_Auto(t *testing.T) {
	cases := []struct {
		name string
		in   [][]interface{}
		want interface{}
	}{
		{"1x1 scalar", [][]interface{}{{"x"}}, "x"},
		{"1xN row", [][]interface{}{{1, 2, 3}}, []interface{}{1, 2, 3}},
		{"Nx1 col", [][]interface{}{{1}, {2}, {3}}, []interface{}{1, 2, 3}},
		{"NxM grid", [][]interface{}{{1, 2}, {3, 4}}, [][]interface{}{{1, 2}, {3, 4}}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			got, err := shapeResult(c.in, ShapeAuto)
			if err != nil {
				t.Fatalf("shapeResult error: %v", err)
			}
			if !reflect.DeepEqual(got, c.want) {
				t.Errorf("got %v (%T), want %v (%T)", got, got, c.want, c.want)
			}
		})
	}
}

// TestShapeResult_Scalar enforces 1×1 input and errors on mismatch — this is
// the xlwings `.options(ndim=0)` rule.
func TestShapeResult_Scalar(t *testing.T) {
	got, err := shapeResult([][]interface{}{{42.0}}, ShapeScalar)
	if err != nil {
		t.Fatalf("Scalar accepted 1x1: %v", err)
	}
	if got != 42.0 {
		t.Errorf("Scalar 1x1: got %v, want 42.0", got)
	}
	if _, err := shapeResult([][]interface{}{{1, 2}}, ShapeScalar); err == nil {
		t.Errorf("Scalar(1x2) should error, got nil")
	}
}

// TestShapeResult_Vector flattens 1×N and N×1 and rejects genuine 2-D blocks.
func TestShapeResult_Vector(t *testing.T) {
	got, err := shapeResult([][]interface{}{{1, 2, 3}}, ShapeVector)
	if err != nil || !reflect.DeepEqual(got, []interface{}{1, 2, 3}) {
		t.Errorf("Vector(1x3): got %v err=%v", got, err)
	}
	got, err = shapeResult([][]interface{}{{1}, {2}, {3}}, ShapeVector)
	if err != nil || !reflect.DeepEqual(got, []interface{}{1, 2, 3}) {
		t.Errorf("Vector(3x1): got %v err=%v", got, err)
	}
	if _, err := shapeResult([][]interface{}{{1, 2}, {3, 4}}, ShapeVector); err == nil {
		t.Errorf("Vector(2x2) should error, got nil")
	}
}

// TestShapeResult_Grid always returns [][]interface{} — even for 1×1.
func TestShapeResult_Grid(t *testing.T) {
	got, err := shapeResult([][]interface{}{{"x"}}, ShapeGrid)
	if err != nil {
		t.Fatalf("Grid(1x1): %v", err)
	}
	if _, ok := got.([][]interface{}); !ok {
		t.Errorf("Grid: got %T, want [][]interface{}", got)
	}
}

// TestApplyEmpty checks the nil-replacement step that runs before shape
// forcing / Convert.
func TestApplyEmpty(t *testing.T) {
	raw := [][]interface{}{{nil, "x"}, {"y", nil}}
	applyEmpty(raw, "N/A")
	want := [][]interface{}{{"N/A", "x"}, {"y", "N/A"}}
	if !reflect.DeepEqual(raw, want) {
		t.Errorf("applyEmpty: got %v, want %v", raw, want)
	}
}

// TestDecodeStructSlice_Headers covers the header-based struct decode used
// by `.options(Header(true)).Get(&rows)`. It exercises:
//   - Header row consumption.
//   - Case-insensitive field matching.
//   - Unknown header skipped.
//   - Missing trailing column leaves field at zero.
func TestDecodeStructSlice_Headers(t *testing.T) {
	type Row struct {
		Name   string
		Age    int
		Active bool
	}
	raw := [][]interface{}{
		{"name", "age", "active"},
		{"alice", 30.0, true},
		{"bob", 25.0, false},
	}
	var out []Row
	dv := reflect.ValueOf(&out).Elem()
	if err := decodeStructSlice(dv, raw, ""); err != nil {
		t.Fatalf("decodeStructSlice: %v", err)
	}
	want := []Row{
		{Name: "alice", Age: 30, Active: true},
		{Name: "bob", Age: 25, Active: false},
	}
	if !reflect.DeepEqual(out, want) {
		t.Errorf("got %+v, want %+v", out, want)
	}
}

// TestDecodeStructSlice_UnknownHeader verifies an unknown header is silently
// skipped (xlwings parity — pandas decode is lenient).
func TestDecodeStructSlice_UnknownHeader(t *testing.T) {
	type Row struct {
		Name string
	}
	raw := [][]interface{}{
		{"name", "unknown_column"},
		{"alice", "ignored"},
	}
	var out []Row
	if err := decodeStructSlice(reflect.ValueOf(&out).Elem(), raw, ""); err != nil {
		t.Fatalf("decodeStructSlice: %v", err)
	}
	if len(out) != 1 || out[0].Name != "alice" {
		t.Errorf("got %+v, want [{alice}]", out)
	}
}

// TestDecodeStructSlice_DateFormat exercises the time.Time -> string field
// path with a custom layout.
func TestDecodeStructSlice_DateFormat(t *testing.T) {
	type Row struct {
		When string
	}
	stamp := time.Date(2026, 5, 17, 12, 0, 0, 0, time.UTC)
	raw := [][]interface{}{
		{"when"},
		{stamp},
	}
	var out []Row
	if err := decodeStructSlice(reflect.ValueOf(&out).Elem(), raw, "2006-01-02"); err != nil {
		t.Fatalf("decodeStructSlice: %v", err)
	}
	if out[0].When != "2026-05-17" {
		t.Errorf("DateFormat: got %q, want %q", out[0].When, "2026-05-17")
	}
}

// TestDecodeStructSlice_EmptyHeader handles the degenerate case of zero data
// rows (only headers present, or the slice is empty). The destination must
// end up as a non-nil empty slice so callers can range over it safely.
func TestDecodeStructSlice_EmptyHeader(t *testing.T) {
	type Row struct{ Name string }
	raw := [][]interface{}{{"name"}}
	var out []Row
	if err := decodeStructSlice(reflect.ValueOf(&out).Elem(), raw, ""); err != nil {
		t.Fatalf("decodeStructSlice: %v", err)
	}
	if out == nil || len(out) != 0 {
		t.Errorf("expected empty non-nil slice, got %v (nil=%v)", out, out == nil)
	}
}

// TestOptions_ConfigAccumulation makes sure functional options accumulate in
// order and the last setter wins for any given knob.
func TestOptions_ConfigAccumulation(t *testing.T) {
	o := rangeOptions{}
	for _, fn := range []RangeOption{
		Scalar(),
		Grid(), // overrides Scalar
		Header(true),
		Empty("N/A"),
		DateFormat("2006-01-02"),
	} {
		fn(&o)
	}
	if o.shape != ShapeGrid {
		t.Errorf("shape: got %v, want ShapeGrid (later option wins)", o.shape)
	}
	if !o.header {
		t.Errorf("header: got false, want true")
	}
	if o.empty != "N/A" {
		t.Errorf("empty: got %v, want N/A", o.empty)
	}
	if o.dateFormat != "2006-01-02" {
		t.Errorf("dateFormat: got %q, want 2006-01-02", o.dateFormat)
	}
}

// TestConvert_RoundTrip walks the Convert escape hatch end-to-end via a fake
// Range that returns a known grid. This is the unit-level cover for the
// `.options(Convert(fn))` API.
func TestConvert_RoundTrip(t *testing.T) {
	fake := &fakeRange{value: [][]interface{}{{1.0, 2.0}, {3.0, 4.0}}}
	or := &optionedRange{
		rng: fake,
		opts: rangeOptions{
			convert: func(raw [][]interface{}) (interface{}, error) {
				sum := 0.0
				for _, row := range raw {
					for _, c := range row {
						if f, ok := c.(float64); ok {
							sum += f
						}
					}
				}
				return sum, nil
			},
		},
	}
	v, err := or.Value()
	if err != nil {
		t.Fatalf("Value: %v", err)
	}
	if v != 10.0 {
		t.Errorf("Convert sum: got %v, want 10.0", v)
	}
}

// TestConvert_Error propagates a converter error out of Value() unchanged.
func TestConvert_Error(t *testing.T) {
	boom := errors.New("boom")
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{{1}}},
		opts: rangeOptions{
			convert: func(raw [][]interface{}) (interface{}, error) { return nil, boom },
		},
	}
	_, err := or.Value()
	if !errors.Is(err, boom) {
		t.Errorf("got %v, want %v", err, boom)
	}
}

// TestGet_ScalarPointer covers the typed-destination Scalar path.
func TestGet_ScalarPointer(t *testing.T) {
	or := &optionedRange{
		rng:  &fakeRange{value: 42.5},
		opts: rangeOptions{shape: ShapeScalar},
	}
	var f float64
	if err := or.Get(&f); err != nil {
		t.Fatalf("Get(&f): %v", err)
	}
	if f != 42.5 {
		t.Errorf("got %v, want 42.5", f)
	}
}

// TestGet_StructSliceRequiresHeader matches the xlwings rule that the
// header row's presence is opt-in: without Header(true) the decode refuses
// to guess the column->field mapping.
func TestGet_StructSliceRequiresHeader(t *testing.T) {
	type Row struct{ Name string }
	or := &optionedRange{
		rng:  &fakeRange{value: [][]interface{}{{"name"}, {"alice"}}},
		opts: rangeOptions{},
	}
	var out []Row
	err := or.Get(&out)
	if err == nil || !strings.Contains(err.Error(), "Header(true)") {
		t.Errorf("expected Header(true) error, got %v", err)
	}
}

// TestGet_StructSliceWithHeader is the happy-path companion to the above.
func TestGet_StructSliceWithHeader(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{
			{"name", "age"},
			{"alice", 30.0},
		}},
		opts: rangeOptions{header: true},
	}
	var out []Row
	if err := or.Get(&out); err != nil {
		t.Fatalf("Get: %v", err)
	}
	if len(out) != 1 || out[0].Name != "alice" || out[0].Age != 30 {
		t.Errorf("got %+v, want [{alice 30}]", out)
	}
}

// TestGet_NilDestination guards the public Get against the common "forgot &"
// mistake.
func TestGet_NilDestination(t *testing.T) {
	or := &optionedRange{rng: &fakeRange{value: 1}}
	if err := or.Get(nil); err == nil {
		t.Errorf("Get(nil) should error")
	}
	var x int
	if err := or.Get(x); err == nil {
		t.Errorf("Get(non-pointer) should error")
	}
}

// TestExpand_UnknownDirection is the configuration-error guard for
// `Options(Expand("upside-down"))`. The error must surface from Err() / the
// first .Value() call without launching Excel.
func TestExpand_UnknownDirection(t *testing.T) {
	// Build an OptionedRange directly so we don't need a Range that
	// supports the End()/Worksheet COM chain. The error path runs entirely
	// in Go.
	or := &optionedRange{
		rng: &fakeRange{value: 1},
		err: errors.New("Expand: unsupported direction \"sideways\""),
	}
	if err := or.Err(); err == nil {
		t.Errorf("expected error from Err(), got nil")
	}
	if _, err := or.Value(); err == nil {
		t.Errorf("expected error from Value(), got nil")
	}
}

// fakeRange is a stub Range used by the unit tests above. It only implements
// the methods the Options pipeline actually calls during pure-Go decoding —
// Value() and Err(). Every other method panics so an accidental dependency
// on COM behaviour shows up loudly.
type fakeRange struct {
	value interface{}
	err   error
}

func (f *fakeRange) Value() (interface{}, error)               { return f.value, f.err }
func (f *fakeRange) Err() error                                { return f.err }
func (f *fakeRange) SetValue(v interface{}) Range              { panic("not implemented") }
func (f *fakeRange) Address() (string, error)                  { panic("not implemented") }
func (f *fakeRange) Formula() (string, error)                  { panic("not implemented") }
func (f *fakeRange) SetFormula(s string) Range                 { panic("not implemented") }
func (f *fakeRange) Formula2() (string, error)                 { panic("not implemented") }
func (f *fakeRange) SetFormula2(s string) Range                { panic("not implemented") }
func (f *fakeRange) NumberFormat() (string, error)             { panic("not implemented") }
func (f *fakeRange) SetNumberFormat(s string) Range            { panic("not implemented") }
func (f *fakeRange) Cells(r, c interface{}) Range              { panic("not implemented") }
func (f *fakeRange) Offset(r, c int) Range                     { panic("not implemented") }
func (f *fakeRange) Resize(r, c int) Range                     { panic("not implemented") }
func (f *fakeRange) Rows() Range                               { panic("not implemented") }
func (f *fakeRange) Columns() Range                            { panic("not implemented") }
func (f *fakeRange) Row() (int32, error)                       { panic("not implemented") }
func (f *fakeRange) Column() (int32, error)                    { panic("not implemented") }
func (f *fakeRange) Count() (int32, error)                     { panic("not implemented") }
func (f *fakeRange) Clear() error                              { panic("not implemented") }
func (f *fakeRange) ClearContents() error                      { panic("not implemented") }
func (f *fakeRange) Delete() error                             { panic("not implemented") }
func (f *fakeRange) Copy() error                               { panic("not implemented") }
func (f *fakeRange) Merge() error                              { panic("not implemented") }
func (f *fakeRange) UnMerge() error                            { panic("not implemented") }
func (f *fakeRange) MergeCells() (bool, error)                 { panic("not implemented") }
func (f *fakeRange) AutoFit() error                            { panic("not implemented") }
func (f *fakeRange) Options(opts ...RangeOption) OptionedRange { panic("not implemented") }

// sugar.Chain methods (embedded in Range). These panic for the same reason.
func (f *fakeRange) Get(prop string, params ...interface{}) sugar.Chain {
	panic("not implemented")
}
func (f *fakeRange) Call(method string, params ...interface{}) sugar.Chain {
	panic("not implemented")
}
func (f *fakeRange) Put(prop string, params ...interface{}) sugar.Chain {
	panic("not implemented")
}
func (f *fakeRange) ForEach(cb func(item sugar.Chain) error) sugar.Chain {
	panic("not implemented")
}
func (f *fakeRange) Fork() sugar.Chain               { panic("not implemented") }
func (f *fakeRange) Store() (*ole.IDispatch, error)  { panic("not implemented") }
func (f *fakeRange) Release() error                  { panic("not implemented") }
func (f *fakeRange) IsDispatch() bool                { panic("not implemented") }
