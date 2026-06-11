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
	"testing"
	"time"
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
			got, err := shapeResult(c.in, NDimAuto)
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
	got, err := shapeResult([][]interface{}{{42.0}}, NDimScalar)
	if err != nil {
		t.Fatalf("Scalar accepted 1x1: %v", err)
	}
	if got != 42.0 {
		t.Errorf("Scalar 1x1: got %v, want 42.0", got)
	}
	if _, err := shapeResult([][]interface{}{{1, 2}}, NDimScalar); err == nil {
		t.Errorf("Scalar(1x2) should error, got nil")
	}
}

// TestShapeResult_Vector flattens 1×N and N×1 and rejects genuine 2-D blocks.
func TestShapeResult_Vector(t *testing.T) {
	got, err := shapeResult([][]interface{}{{1, 2, 3}}, NDimVector)
	if err != nil || !reflect.DeepEqual(got, []interface{}{1, 2, 3}) {
		t.Errorf("Vector(1x3): got %v err=%v", got, err)
	}
	got, err = shapeResult([][]interface{}{{1}, {2}, {3}}, NDimVector)
	if err != nil || !reflect.DeepEqual(got, []interface{}{1, 2, 3}) {
		t.Errorf("Vector(3x1): got %v err=%v", got, err)
	}
	if _, err := shapeResult([][]interface{}{{1, 2}, {3, 4}}, NDimVector); err == nil {
		t.Errorf("Vector(2x2) should error, got nil")
	}
}

// TestShapeResult_Grid always returns [][]interface{} — even for 1×1.
func TestShapeResult_Grid(t *testing.T) {
	got, err := shapeResult([][]interface{}{{"x"}}, NDimGrid)
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

// TestDecodeStructSlice_EmbeddedStruct verifies that headers matching fields
// promoted from an embedded struct decode into the promoted field, not the
// embedded struct itself. FieldByName returns a multi-level Index path for
// promoted fields; storing only Index[0] (the embedded field's index) and
// using Field() would target the embedded struct and fail to assign a scalar.
func TestDecodeStructSlice_EmbeddedStruct(t *testing.T) {
	type Base struct {
		Name string
		Age  int
	}
	type Row struct {
		Base
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
		{Base: Base{Name: "alice", Age: 30}, Active: true},
		{Base: Base{Name: "bob", Age: 25}, Active: false},
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
	if o.shape != NDimGrid {
		t.Errorf("shape: got %v, want NDimGrid (later option wins)", o.shape)
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
		opts: rangeOptions{shape: NDimScalar},
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
func TestGet_StructSlicePositional(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	// Without Header(true) every row is data; columns map to exported
	// fields in declaration order.
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{
			{"alice", 30.0},
			{"bob", 25.0},
		}},
		opts: rangeOptions{},
	}
	var out []Row
	if err := or.Get(&out); err != nil {
		t.Fatalf("Get: %v", err)
	}
	want := []Row{{"alice", 30}, {"bob", 25}}
	if !reflect.DeepEqual(out, want) {
		t.Errorf("got %+v, want %+v", out, want)
	}
}

// TestGet_StructSlicePositional_ShortRow checks lenient handling: rows with
// fewer columns than fields leave the remaining fields at zero values.
func TestGet_StructSlicePositional_ShortRow(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	or := &optionedRange{
		rng:  &fakeRange{value: [][]interface{}{{"alice"}}},
		opts: rangeOptions{},
	}
	var out []Row
	if err := or.Get(&out); err != nil {
		t.Fatalf("Get: %v", err)
	}
	if len(out) != 1 || out[0].Name != "alice" || out[0].Age != 0 {
		t.Errorf("got %+v, want [{alice 0}]", out)
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

// TestApplyIndex covers the xlwings `index=n` analogue: leading columns are
// dropped, n<=0 is a no-op, and rows shorter than n become empty.
func TestApplyIndex(t *testing.T) {
	raw := [][]interface{}{
		{"idx1", "a", 1.0},
		{"idx2", "b", 2.0},
	}
	got := applyIndex(raw, 1)
	want := [][]interface{}{{"a", 1.0}, {"b", 2.0}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("Index(1): got %v, want %v", got, want)
	}

	if !reflect.DeepEqual(applyIndex(raw, 0), raw) {
		t.Errorf("Index(0) should be a no-op")
	}
	if !reflect.DeepEqual(applyIndex(raw, -1), raw) {
		t.Errorf("Index(-1) should be a no-op")
	}

	short := applyIndex([][]interface{}{{"only"}}, 2)
	if len(short) != 1 || len(short[0]) != 0 {
		t.Errorf("Index past row length should give empty rows, got %v", short)
	}
}

// TestGet_IndexWithHeaderDecode combines Index with the header struct
// decode: the index column disappears before headers are interpreted.
func TestGet_IndexWithHeaderDecode(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{
			{"id", "name", "age"},
			{1.0, "alice", 30.0},
		}},
		opts: rangeOptions{header: true, index: 1},
	}
	var out []Row
	if err := or.Get(&out); err != nil {
		t.Fatalf("Get: %v", err)
	}
	if len(out) != 1 || out[0].Name != "alice" || out[0].Age != 30 {
		t.Errorf("got %+v, want [{alice 30}]", out)
	}
}

// TestConvertTo_TypedDestination verifies the generic, compile-time-checked
// Convert flavor feeding a *T destination.
func TestConvertTo_TypedDestination(t *testing.T) {
	type Stats struct{ Sum float64 }
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{{1.0, 2.0}, {3.0, 4.0}}},
	}
	opt := ConvertTo(func(raw [][]interface{}) (Stats, error) {
		s := Stats{}
		for _, row := range raw {
			for _, c := range row {
				s.Sum += c.(float64)
			}
		}
		return s, nil
	})
	opt(&or.opts)

	var stats Stats
	if err := or.Get(&stats); err != nil {
		t.Fatalf("Get: %v", err)
	}
	if stats.Sum != 10.0 {
		t.Errorf("Sum: got %v, want 10.0", stats.Sum)
	}
}

// TestSet_StructRowsWithHeader is the write-direction mirror of the header
// decode: field names become the first row.
func TestSet_StructRowsWithHeader(t *testing.T) {
	type Row struct {
		Name string
		Age  float64
	}
	fake := &fakeRange{}
	or := &optionedRange{rng: fake, opts: rangeOptions{header: true}}

	err := or.Set([]Row{{"alice", 30}, {"bob", 25}})
	if err != nil {
		t.Fatalf("Set: %v", err)
	}
	want := [][]interface{}{
		{"Name", "Age"},
		{"alice", 30.0},
		{"bob", 25.0},
	}
	if !reflect.DeepEqual(fake.setValue, want) {
		t.Errorf("written grid: got %v, want %v", fake.setValue, want)
	}
	if fake.resizedRows != 3 || fake.resizedCols != 2 {
		t.Errorf("resize: got %dx%d, want 3x2", fake.resizedRows, fake.resizedCols)
	}
}

// TestSet_StructRowsPositional omits the header row without Header(true).
func TestSet_StructRowsPositional(t *testing.T) {
	type Row struct {
		Name string
		Age  float64
	}
	fake := &fakeRange{}
	or := &optionedRange{rng: fake}

	if err := or.Set([]Row{{"alice", 30}}); err != nil {
		t.Fatalf("Set: %v", err)
	}
	want := [][]interface{}{{"alice", 30.0}}
	if !reflect.DeepEqual(fake.setValue, want) {
		t.Errorf("written grid: got %v, want %v", fake.setValue, want)
	}
	if fake.resizedRows != 1 || fake.resizedCols != 2 {
		t.Errorf("resize: got %dx%d, want 1x2", fake.resizedRows, fake.resizedCols)
	}
}

// TestSet_SliceShapes checks the resize arithmetic for plain 1-D and 2-D
// slices and the zero-size no-op.
func TestSet_SliceShapes(t *testing.T) {
	fake := &fakeRange{}
	or := &optionedRange{rng: fake}

	if err := or.Set([][]float64{{1, 2, 3}, {4, 5, 6}}); err != nil {
		t.Fatalf("Set 2-D: %v", err)
	}
	if fake.resizedRows != 2 || fake.resizedCols != 3 {
		t.Errorf("2-D resize: got %dx%d, want 2x3", fake.resizedRows, fake.resizedCols)
	}

	if err := or.Set([]string{"a", "b"}); err != nil {
		t.Fatalf("Set 1-D: %v", err)
	}
	if fake.resizedRows != 1 || fake.resizedCols != 2 {
		t.Errorf("1-D resize: got %dx%d, want 1x2", fake.resizedRows, fake.resizedCols)
	}

	fake2 := &fakeRange{}
	or2 := &optionedRange{rng: fake2}
	if err := or2.Set([]string{}); err != nil {
		t.Fatalf("Set empty: %v", err)
	}
	if fake2.setValue != nil {
		t.Errorf("empty source should be a no-op, wrote %v", fake2.setValue)
	}
}

// TestSet_NoExportedFields rejects structs the encoder cannot project.
func TestSet_NoExportedFields(t *testing.T) {
	type hidden struct{ x int } //nolint:unused
	or := &optionedRange{rng: &fakeRange{}}
	if err := or.Set([]hidden{{1}}); err == nil {
		t.Error("expected error for struct with no exported fields")
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
// the methods the Options pipeline actually calls during pure-Go decoding
// and encoding — Value/Err/SetValue/Resize. The embedded nil Range interface
// supplies the rest of the method set: calling any of them panics with a
// nil-pointer dereference, so an accidental dependency on COM behaviour
// shows up loudly, and new Range methods don't require new stubs here.
type fakeRange struct {
	Range // nil — panics on any method not overridden below

	value interface{}
	err   error

	// write recording for Options.Set tests
	setValue                 interface{}
	resizedRows, resizedCols int
}

func (f *fakeRange) Value() (interface{}, error)  { return f.value, f.err }
func (f *fakeRange) Err() error                   { return f.err }
func (f *fakeRange) SetValue(v interface{}) Range { f.setValue = v; return f }
func (f *fakeRange) Resize(r, c int) Range        { f.resizedRows, f.resizedCols = r, c; return f }
