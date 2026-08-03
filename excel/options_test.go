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

// TestValidExpandDirection covers the direction validator used for the early
// Options() error (the expansion itself is deferred to read time).
func TestValidExpandDirection(t *testing.T) {
	for _, ok := range []string{"table", "down", "right"} {
		if !validExpandDirection(ok) {
			t.Errorf("validExpandDirection(%q) = false, want true", ok)
		}
	}
	for _, bad := range []string{"", "up", "left", "sideways", "TABLE"} {
		if validExpandDirection(bad) {
			t.Errorf("validExpandDirection(%q) = true, want false", bad)
		}
	}
}

// TestOptions_ExpandBadDirection confirms an invalid Expand direction is
// reported eagerly through Err() (Excel-free: Options() validates the string
// but never touches COM — the actual expansion is deferred to Value()/Get()).
func TestOptions_ExpandBadDirection(t *testing.T) {
	r := &excelRange{sugar.Error(nil)}
	or := r.Options(Expand("sideways"))
	if or.Err() == nil {
		t.Error("Options(Expand(\"sideways\")).Err() = nil, want a direction error")
	}
	// A valid direction must NOT set an eager error.
	if err := r.Options(Expand("down")).Err(); err != nil {
		t.Errorf("Options(Expand(\"down\")).Err() = %v, want nil (expansion deferred)", err)
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

// TestDecodeStructSlicePositional_EmbeddedStruct is the positional twin of the
// test above. The two read paths used to disagree: the header path flattens
// (FieldByName traverses embedded structs and yields a multi-level Index),
// while the positional path scanned only st.NumField() and handed assignField a
// whole struct destination — a hard error, "cannot assign float64 to Base".
func TestDecodeStructSlicePositional_EmbeddedStruct(t *testing.T) {
	type Base struct {
		ID int
	}
	type Row struct {
		Base
		Name string
	}
	raw := [][]interface{}{
		{1.0, "a"},
		{2.0, "b"},
	}
	var out []Row
	if err := decodeStructSlicePositional(reflect.ValueOf(&out).Elem(), raw, ""); err != nil {
		t.Fatalf("decodeStructSlicePositional: %v", err)
	}
	want := []Row{
		{Base: Base{ID: 1}, Name: "a"},
		{Base: Base{ID: 2}, Name: "b"},
	}
	if !reflect.DeepEqual(out, want) {
		t.Errorf("got %+v, want %+v", out, want)
	}
}

// TestSet_StructRowsEmbeddedStruct is the write-direction mirror: the grid must
// carry the promoted LEAF fields, one column each. Before the unification the
// embedded struct went into a single cell and scalarToVariant rejected it with
// "unsupported cell type".
func TestSet_StructRowsEmbeddedStruct(t *testing.T) {
	type Base struct {
		ID int
	}
	type Row struct {
		Base
		Name string
	}
	fake := &fakeRange{}
	or := &optionedRange{rng: fake, opts: rangeOptions{header: true}}
	if err := or.Set([]Row{{Base{1}, "a"}}); err != nil {
		t.Fatalf("Set: %v", err)
	}
	want := [][]interface{}{
		{"ID", "Name"},
		{1, "a"},
	}
	if !reflect.DeepEqual(fake.setValue, want) {
		t.Errorf("written grid: got %v, want %v", fake.setValue, want)
	}
	if fake.resizedRows != 2 || fake.resizedCols != 2 {
		t.Errorf("resize: got %dx%d, want 2x2", fake.resizedRows, fake.resizedCols)
	}
}

// TestSet_GetRoundTripEmbedded is the test that pins all THREE field-collection
// rules agreeing: Set(Header(true)) writes the header row using the promoted
// leaf names, and the header decode — which resolves those names through
// FieldByName — must read the same struct back. It goes red if a later change
// "fixes" only one path.
func TestSet_GetRoundTripEmbedded(t *testing.T) {
	type Base struct {
		ID   int
		Tag  string
		Kept bool
	}
	type Row struct {
		Base
		Name string
	}
	src := []Row{
		{Base{1, "x", true}, "a"},
		{Base{2, "y", false}, "b"},
	}
	fake := &fakeRange{}
	or := &optionedRange{rng: fake, opts: rangeOptions{header: true}}
	if err := or.Set(src); err != nil {
		t.Fatalf("Set: %v", err)
	}
	grid, ok := fake.setValue.([][]interface{})
	if !ok {
		t.Fatalf("Set wrote %T, want [][]interface{}", fake.setValue)
	}
	// The fake Range accepts any Go value, so a cell holding a whole struct
	// would round-trip through decodeStructSlice and this test would pass
	// vacuously — real Excel rejects it at scalarToVariant ("unsupported cell
	// type"). Assert the cells are the scalars the SAFEARRAY encoder accepts.
	assertScalarCells(t, grid)
	var out []Row
	if err := decodeStructSlice(reflect.ValueOf(&out).Elem(), grid, ""); err != nil {
		t.Fatalf("decodeStructSlice: %v", err)
	}
	if !reflect.DeepEqual(out, src) {
		t.Errorf("round trip: got %+v, want %+v", out, src)
	}
}

// assertScalarCells fails the test if any cell of grid is a composite value the
// SAFEARRAY encoder (sugar.scalarToVariant) refuses — the fake Range does not
// model that rejection, so the assertion has to.
func assertScalarCells(t *testing.T, grid [][]interface{}) {
	t.Helper()
	for r, row := range grid {
		for c, cell := range row {
			if cell == nil {
				continue
			}
			if _, isTime := cell.(time.Time); isTime {
				continue
			}
			switch reflect.ValueOf(cell).Kind() {
			case reflect.Struct, reflect.Map, reflect.Slice, reflect.Array, reflect.Ptr:
				t.Errorf("cell (%d,%d) is %T (%v) — not a value Excel can store",
					r, c, cell, cell)
			}
		}
	}
}

// TestStructFields_EmbeddedTimeIsALeaf is the carve-out pin that stops the
// flattening from regressing behaviour that WORKS today. time.Time is a struct
// with no exported fields, so expanding it would yield zero columns and silently
// drop the field — while an embedded time.Time already decodes positionally
// today, because AssignableTo(time.Time) matches at field level.
func TestStructFields_EmbeddedTimeIsALeaf(t *testing.T) {
	type R struct {
		time.Time
		Name string
	}
	fields, err := structFields(reflect.TypeOf(R{}))
	if err != nil {
		t.Fatalf("structFields: %v", err)
	}
	if len(fields) != 2 {
		names := make([]string, len(fields))
		for i, f := range fields {
			names[i] = f.Name
		}
		t.Fatalf("structFields(R) = %v (%d columns), want 2 (Time, Name)", names, len(fields))
	}
	if fields[0].Name != "Time" || fields[1].Name != "Name" {
		t.Errorf("structFields(R) names = %q,%q; want Time,Name", fields[0].Name, fields[1].Name)
	}
	// And the positional decode must still populate it.
	when := time.Date(2026, 8, 3, 10, 0, 0, 0, time.UTC)
	var out []R
	if err := decodeStructSlicePositional(reflect.ValueOf(&out).Elem(), [][]interface{}{{when, "x"}}, ""); err != nil {
		t.Fatalf("decodeStructSlicePositional: %v", err)
	}
	if len(out) != 1 || !out[0].Time.Equal(when) || out[0].Name != "x" {
		t.Errorf("got %+v, want [{%v x}]", out, when)
	}
}

// TestStructFields_NamedStructIsNotExpanded pins the other carve-out: only
// ANONYMOUS embedded structs expand. A named struct field stays one column,
// matching Go's own promotion rule and the header path, which cannot reach
// inside it either.
func TestStructFields_NamedStructIsNotExpanded(t *testing.T) {
	type Base struct {
		ID int
	}
	type R struct {
		Inner Base
		Name  string
	}
	fields, err := structFields(reflect.TypeOf(R{}))
	if err != nil {
		t.Fatalf("structFields: %v", err)
	}
	if len(fields) != 2 || fields[0].Name != "Inner" || fields[1].Name != "Name" {
		t.Fatalf("structFields(R) = %+v, want [Inner Name]", fields)
	}
	if fields[0].Type != reflect.TypeOf(Base{}) {
		t.Errorf("Inner column type = %v, want %v (unexpanded)", fields[0].Type, reflect.TypeOf(Base{}))
	}
}

// TestStructFields_NestedEmbedAndUnexported covers the remaining shapes: an
// embed inside an embed flattens depth-first, unexported leaves are skipped, and
// an embedded struct with no exported leaves stays a single column so the
// surrounding column indices do not shift.
func TestStructFields_NestedEmbedAndUnexported(t *testing.T) {
	type Inner struct {
		X       int
		skipped string //nolint:unused
	}
	type Mid struct {
		Inner
		Y int
	}
	type Empty struct{ hidden int } //nolint:unused
	type R struct {
		Mid
		Empty
		Z int
	}
	fields, err := structFields(reflect.TypeOf(R{}))
	if err != nil {
		t.Fatalf("structFields: %v", err)
	}
	var got []string
	for _, f := range fields {
		got = append(got, f.Name)
	}
	want := []string{"X", "Y", "Empty", "Z"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("structFields(R) = %v, want %v", got, want)
	}
	if !reflect.DeepEqual(fields[0].Index, []int{0, 0, 0}) {
		t.Errorf("X Index = %v, want [0 0 0]", fields[0].Index)
	}
}

// TestStructFields_AmbiguousPromotedNameRejected is the safety valve Go's
// promotion rules force: when a flattened leaf name does not resolve back
// through FieldByName to the same field, Set would write a column that Get
// cannot read back. Rather than invent a resolution the header path does not
// share, structFields refuses the struct.
func TestStructFields_AmbiguousPromotedNameRejected(t *testing.T) {
	type A struct{ Name string }
	type B struct{ Name string }
	type Ambiguous struct {
		A
		B
	}
	if _, err := structFields(reflect.TypeOf(Ambiguous{})); err == nil {
		t.Error("structFields on two embeds sharing a leaf name should error")
	}
	// Shadowing is the same hazard: Go resolves Name to the outer field, so the
	// inner column would be written and then read back into the wrong field.
	type Shadow struct {
		A
		Name string
	}
	if _, err := structFields(reflect.TypeOf(Shadow{})); err == nil {
		t.Error("structFields on a shadowed embedded leaf name should error")
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

// TestAssign_IntegerIntoStringIsDigitsNotRune pins the integer-source /
// string-destination rule: Go's reflect Convert turns an integer into the
// string holding that CODE POINT (int32(65) -> "A", int(0) -> "\x00"), which is
// never what a spreadsheet caller means. The digits must win.
func TestAssign_IntegerIntoStringIsDigitsNotRune(t *testing.T) {
	cases := []struct {
		val  interface{}
		want string
	}{
		{int32(65), "65"},
		{int(0), "0"},
		{int64(1), "1"},
		{uint8(65), "65"},
		{int(-45), "-45"},
	}
	for _, c := range cases {
		var s string
		if err := assign(reflect.ValueOf(&s).Elem(), c.val); err != nil {
			t.Fatalf("assign(%T(%v)): %v", c.val, c.val, err)
		}
		if s != c.want {
			t.Errorf("assign(%T(%v)) = %q, want %q", c.val, c.val, s, c.want)
		}
		// assignField must agree — the two functions carried the same
		// unguarded ConvertibleTo branch, and patching only one would leave
		// the header and positional struct paths inconsistent.
		var fs string
		if err := assignField(reflect.ValueOf(&fs).Elem(), c.val, ""); err != nil {
			t.Fatalf("assignField(%T(%v)): %v", c.val, c.val, err)
		}
		if fs != c.want {
			t.Errorf("assignField(%T(%v)) = %q, want %q", c.val, c.val, fs, c.want)
		}
	}
	// A float source is NOT ConvertibleTo string, so it never had the rune
	// problem: it falls through to assignField's Sprint fallback (and to a
	// clean error in assign, which has no fallback). Pinned so the new
	// integer branch is not mistaken for what makes floats work.
	var f string
	if err := assignField(reflect.ValueOf(&f).Elem(), float64(65), ""); err != nil {
		t.Fatalf("assignField(float64): %v", err)
	}
	if f != "65" {
		t.Errorf("assignField(float64(65)) = %q, want \"65\"", f)
	}
	if err := assign(reflect.ValueOf(&f).Elem(), float64(65)); err == nil {
		t.Errorf("assign(float64 -> string) = nil error, want the no-conversion error")
	}
}

// TestAssignField_EmptyIntIntoStringField is the reachable trigger for the rune
// trap: Options(Empty(0)) substitutes an int into the grid, and a string field
// then received "\x00". Driven through the real Options().Get pipeline.
func TestAssignField_EmptyIntIntoStringField(t *testing.T) {
	type Row struct {
		Name string
	}
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{
			{"Name"},
			{nil},
		}},
		opts: rangeOptions{header: true, empty: 0},
	}
	var out []Row
	if err := or.Get(&out); err != nil {
		t.Fatalf("Get: %v", err)
	}
	if len(out) != 1 {
		t.Fatalf("got %d rows, want 1", len(out))
	}
	if out[0].Name != "0" {
		t.Errorf("Empty(0) into a string field = %q, want %q", out[0].Name, "0")
	}
}

// TestAssign_ByteSliceStillConvertsToString is a NEGATIVE pin: the integer
// guard is keyed on the source KIND, so []byte (reflect.Slice, not Uint8) keeps
// its existing ConvertibleTo behaviour.
func TestAssign_ByteSliceStillConvertsToString(t *testing.T) {
	var s string
	if err := assign(reflect.ValueOf(&s).Elem(), []byte("hi")); err != nil {
		t.Fatalf("assign([]byte): %v", err)
	}
	if s != "hi" {
		t.Errorf("assign([]byte(\"hi\")) = %q, want \"hi\"", s)
	}
	var s2 string
	if err := assign(reflect.ValueOf(&s2).Elem(), []rune("hi")); err != nil {
		t.Fatalf("assign([]rune): %v", err)
	}
	if s2 != "hi" {
		t.Errorf("assign([]rune(\"hi\")) = %q, want \"hi\"", s2)
	}
}

// TestAssignField_NamedStringDestination is the other NEGATIVE pin: a named
// string type must still take the string fast path in both assign and
// assignField, and a named INTEGER type with a String() method must render
// through that method (fmt.Sprint), not as a rune.
func TestAssignField_NamedStringDestination(t *testing.T) {
	type Code string
	var c Code
	if err := assignField(reflect.ValueOf(&c).Elem(), "AB", ""); err != nil {
		t.Fatalf("assignField(string -> Code): %v", err)
	}
	if c != Code("AB") {
		t.Errorf("assignField(\"AB\") = %q, want \"AB\"", string(c))
	}
	var c2 Code
	if err := assign(reflect.ValueOf(&c2).Elem(), "AB"); err != nil {
		t.Fatalf("assign(string -> Code): %v", err)
	}
	if c2 != Code("AB") {
		t.Errorf("assign(\"AB\") = %q, want \"AB\"", string(c2))
	}
	// time.Duration is an int64 kind whose Convert-to-string yields U+FFFD
	// garbage; fmt.Sprint renders "1s".
	var s string
	if err := assign(reflect.ValueOf(&s).Elem(), time.Second); err != nil {
		t.Fatalf("assign(time.Duration): %v", err)
	}
	if s != "1s" {
		t.Errorf("assign(time.Second) = %q, want \"1s\"", s)
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

// TestGet_NilValueEmptyCell is the regression test for the assign() nil panic.
// An empty Excel cell (VT_EMPTY) surfaces as a nil value: Range.Value() → nil,
// which readGrid wraps as [][]interface{}{{nil}}, shapeResult collapses back to
// a scalar nil, and Get(dst) calls assign(elem, nil). The pre-fix assign ran
// the interface{} fast path (dst.Set(reflect.ValueOf(nil))) before the nil
// check, and reflect.ValueOf(nil) is the zero Value, so Set panicked with
// "reflect: call of reflect.Value.Set on zero Value". Both *interface{} and a
// concrete destination must accept nil and yield the zero value without
// panicking.
func TestGet_NilValueEmptyCell(t *testing.T) {
	// *interface{} destination — the documented, common case.
	var iface interface{} = "sentinel" // start non-nil to prove it is cleared
	or := &optionedRange{rng: &fakeRange{value: nil}, opts: rangeOptions{shape: NDimScalar}}
	if err := or.Get(&iface); err != nil {
		t.Fatalf("Get(&interface{}) on empty cell: %v", err)
	}
	if iface != nil {
		t.Errorf("empty cell should decode to nil interface, got %v (%T)", iface, iface)
	}

	// Concrete destination — nil must leave it at the zero value.
	var s string = "sentinel"
	or2 := &optionedRange{rng: &fakeRange{value: nil}, opts: rangeOptions{shape: NDimScalar}}
	if err := or2.Get(&s); err != nil {
		t.Fatalf("Get(&string) on empty cell: %v", err)
	}
	if s != "" {
		t.Errorf("empty cell should decode to zero string, got %q", s)
	}
}

// TestGet_ConvertNilResult covers the sibling path the same fix guards: a
// Convert callback that returns nil (options.go convert path) must land in
// assign() as nil without panicking.
func TestGet_ConvertNilResult(t *testing.T) {
	or := &optionedRange{
		rng: &fakeRange{value: [][]interface{}{{1.0}}},
		opts: rangeOptions{
			convert: func(raw [][]interface{}) (interface{}, error) { return nil, nil },
		},
	}
	var iface interface{} = "sentinel"
	if err := or.Get(&iface); err != nil {
		t.Fatalf("Get with nil-returning Convert: %v", err)
	}
	if iface != nil {
		t.Errorf("nil Convert result should decode to nil, got %v (%T)", iface, iface)
	}
}

// TestNeighborOffset maps each End() direction to the adjacent-cell delta.
func TestNeighborOffset(t *testing.T) {
	if dr, dc := neighborOffset(xlDown); dr != 1 || dc != 0 {
		t.Errorf("xlDown: got (%d,%d), want (1,0)", dr, dc)
	}
	if dr, dc := neighborOffset(xlToRight); dr != 0 || dc != 1 {
		t.Errorf("xlToRight: got (%d,%d), want (0,1)", dr, dc)
	}
}

// TestCellBlank covers the xlwings `raw_value in (None, "")` blank test.
func TestCellBlank(t *testing.T) {
	cases := []struct {
		v    interface{}
		want bool
	}{
		{nil, true},
		{"", true},
		{"x", false},
		{0.0, false}, // a zero number is data, not blank
		{false, false},
	}
	for _, c := range cases {
		got, err := cellBlank(&fakeRange{value: c.v})
		if err != nil {
			t.Fatalf("cellBlank(%v): %v", c.v, err)
		}
		if got != c.want {
			t.Errorf("cellBlank(%#v) = %v, want %v", c.v, got, c.want)
		}
	}
}

// TestEndpointAddr_BlankNeighborGuard is the Excel-free cover for the defect-4
// guard: when the cell adjacent to the origin in the expansion direction is
// blank, End() must NOT be called (fakeRange's embedded nil Range panics on
// Get), and the origin is its own endpoint. Pre-fix, End() was called
// unconditionally, which here would panic and against Excel would overshoot to
// a distant data island / the sheet edge.
func TestEndpointAddr_BlankNeighborGuard(t *testing.T) {
	// nil neighbor below the anchor.
	origin := &fakeRange{address: "$A$1", offsetValue: nil}
	got, err := endpointAddr(origin, xlDown)
	if err != nil {
		t.Fatalf("endpointAddr(down, blank): %v", err)
	}
	if got != "$A$1" {
		t.Errorf("blank down-neighbor: got %q, want $A$1 (guard should skip End)", got)
	}

	// empty-string neighbor to the right.
	origin2 := &fakeRange{address: "$A$1", offsetValue: ""}
	got, err = endpointAddr(origin2, xlToRight)
	if err != nil {
		t.Fatalf("endpointAddr(right, blank): %v", err)
	}
	if got != "$A$1" {
		t.Errorf("blank right-neighbor: got %q, want $A$1", got)
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

	// expand blank-neighbor guard tests: Address() returns address, and
	// Offset(...) returns a cell whose Value() is offsetValue.
	address     string
	offsetValue interface{}
}

func (f *fakeRange) Value() (interface{}, error)  { return f.value, f.err }
func (f *fakeRange) Err() error                   { return f.err }
func (f *fakeRange) SetValue(v interface{}) Range { f.setValue = v; return f }
func (f *fakeRange) Resize(r, c int) Range        { f.resizedRows, f.resizedCols = r, c; return f }
func (f *fakeRange) Address() (string, error)     { return f.address, f.err }
func (f *fakeRange) Offset(r, c int) Range        { return &fakeRange{value: f.offsetValue} }

// TestAssignField_NullIsNotSprintedIntoAString closes the last forgery hole the
// sugar.Null sentinel opened. assignField's last-resort branch is
// `fmt.Sprint(val)` for a string destination, and Null HAS a String() method —
// so a Null cell decoding into a string struct field would silently write the
// literal text "Null" and look like real cell data. (Before the sentinel it
// silently wrote "" instead, which was wrong more quietly.)
//
// The sibling `assign` needs no explicit arm: a struct source is neither
// Assignable nor Convertible to any scalar destination, so it already reaches
// the "cannot assign" error — asserted here so that stays true.
func TestAssignField_NullIsNotSprintedIntoAString(t *testing.T) {
	var s string
	err := assignField(reflect.ValueOf(&s).Elem(), sugar.Null{}, "")
	if err == nil {
		t.Fatalf("assignField(Null -> string) = %q with no error; want a refusal", s)
	}
	if s == "Null" {
		t.Errorf("the sentinel's String() was forged into the field as cell data")
	}
	if !strings.Contains(err.Error(), "Null") {
		t.Errorf("error %v should name Null", err)
	}

	// An interface{} field still receives the sentinel — a caller asking for raw
	// values must be able to see it, which is the whole point of the type.
	var any interface{}
	if err := assignField(reflect.ValueOf(&any).Elem(), sugar.Null{}, ""); err != nil {
		t.Fatalf("assignField(Null -> interface{}): %v", err)
	}
	if !sugar.IsNull(any) {
		t.Errorf("interface{} field = %v (%T); want the Null sentinel", any, any)
	}

	// assign's numeric path: already an error, pinned so it stays one.
	var f float64
	if err := assign(reflect.ValueOf(&f).Elem(), sugar.Null{}); err == nil {
		t.Errorf("assign(Null -> float64) = %v with no error; want a refusal", f)
	}
}
