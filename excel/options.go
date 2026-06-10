//go:build windows

package excel

import (
	"errors"
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/xll-gen/sugar"
)

// XlDirection enumerates the COM `XlDirection` values used by `Range.End`.
// We only need these four for the xlwings-parity `Expand` option.
const (
	xlUp      int32 = -4162
	xlToLeft  int32 = -4159
	xlDown    int32 = -4121
	xlToRight int32 = -4161
)

// NDim is the dimension-forcing knob exposed by xlwings' `.options(ndim=...)`.
// Scalar / Vector / Grid mirror xlwings' implicit ndim=0/1/2. (Named NDim
// rather than Shape to leave the Shape identifier for the drawing-layer
// object, matching xlwings.)
type NDim int

const (
	// NDimAuto leaves the result in its natural xlwings shape: a single
	// scalar for 1×1, a flat `[]interface{}` for 1×N or N×1, and
	// `[][]interface{}` for everything else. This is the default.
	NDimAuto NDim = iota
	// NDimScalar forces a 1×1 read. The Value()/Get() call errors out if
	// the underlying range has more than one cell.
	NDimScalar
	// NDimVector forces a 1-D slice. `1×N` is read row-wise; `N×1` is read
	// column-wise. Anything else returns an error.
	NDimVector
	// NDimGrid forces a `[][]interface{}` result, even for `1×1`, `1×N`,
	// or `N×1` ranges. Equivalent to xlwings' `ndim=2`.
	NDimGrid
)

// rangeOptions accumulates the knobs passed to Range.Options(...). It is an
// internal record; callers use the exported option helpers (Scalar, Vector,
// Grid, Header, Expand, Empty, DateFormat, Convert).
type rangeOptions struct {
	shape      NDim
	header     bool
	index      int
	empty      interface{}
	dateFormat string
	expand     string // "", "table", "down", "right"
	convert    func(raw [][]interface{}) (interface{}, error)
}

// RangeOption is the functional-options type accepted by Range.Options. Each
// option mutates a private rangeOptions struct; the rangeOptions value is
// stored on the returned OptionedRange and consumed by Value/Get.
type RangeOption func(*rangeOptions)

// Scalar is the xlwings `.options(ndim=0)` analogue: force a single-cell read
// and return the underlying scalar (or error on a multi-cell range).
func Scalar() RangeOption {
	return func(o *rangeOptions) { o.shape = NDimScalar }
}

// Vector is the xlwings `.options(ndim=1)` analogue: force a 1-D slice result.
func Vector() RangeOption {
	return func(o *rangeOptions) { o.shape = NDimVector }
}

// Grid is the xlwings `.options(ndim=2)` analogue: always return [][]interface{}
// (one row per Excel row).
func Grid() RangeOption {
	return func(o *rangeOptions) { o.shape = NDimGrid }
}

// Header treats the first row of the read result as struct field names when
// decoding into a `*[]SomeStruct` destination. Matches xlwings'
// `.options(pd.DataFrame, header=1, ...)` flag for the struct case.
//
// When true, the first row is consumed as headers; remaining rows become
// elements of the destination slice. The default false treats every row as
// data and decodes positionally: column 0 fills the struct's first exported
// field, column 1 the second, and so on.
func Header(on bool) RangeOption {
	return func(o *rangeOptions) { o.header = on }
}

// Index drops the first n columns of the read before any further conversion.
// This is the Go analogue of xlwings' `.options(index=n)` for DataFrames,
// where the leading columns form the index rather than data. n <= 0 is a
// no-op.
func Index(n int) RangeOption {
	return func(o *rangeOptions) { o.index = n }
}

// Empty replaces nil cells with the provided value during decoding. Mirrors
// xlwings' `.options(empty=...)`. The replacement is applied after the raw
// 2-D read but before shape forcing / struct decode.
func Empty(value interface{}) RangeOption {
	return func(o *rangeOptions) { o.empty = value }
}

// DateFormat sets a `time.Format` layout used when assigning time.Time values
// into string struct fields. xlwings hands the layout to Python's
// `datetime.strftime`; in Go we use the equivalent `time.Format` layout.
//
// If unset, time.Time values are written into string fields via time.Time.String().
func DateFormat(layout string) RangeOption {
	return func(o *rangeOptions) { o.dateFormat = layout }
}

// Expand auto-grows the range from its anchor (top-left cell) before reading.
// Supported directions mirror xlwings:
//
//   - "table" — grow down then right until the first empty cell in each
//     direction (xlwings' default for `expand=...`).
//   - "down"  — grow down only.
//   - "right" — grow right only.
//
// An unknown direction defers an error onto the OptionedRange's chain so the
// caller observes it from .Value() / .Get() / .Err().
func Expand(direction string) RangeOption {
	return func(o *rangeOptions) { o.expand = strings.ToLower(direction) }
}

// Convert installs a caller-supplied function that receives the raw 2-D read
// (after Empty replacement, before shape forcing or struct decode) and
// returns the value handed back to Value()/Get(). Mirrors xlwings'
// `.options(MyConverter, ...)` escape hatch.
//
// When Convert is set, Shape, Header, and struct decode are bypassed —
// the function owns the projection end-to-end.
func Convert(fn func(raw [][]interface{}) (interface{}, error)) RangeOption {
	return func(o *rangeOptions) { o.convert = fn }
}

// ConvertTo is the type-safe flavor of Convert: the converter returns a
// concrete T instead of interface{}, so its signature is checked at compile
// time. Pair it with a *T destination in Get:
//
//	rng.Options(excel.ConvertTo(func(raw [][]interface{}) (Stats, error) {
//	    ...
//	})).Get(&stats)
func ConvertTo[T any](fn func(raw [][]interface{}) (T, error)) RangeOption {
	return Convert(func(raw [][]interface{}) (interface{}, error) {
		return fn(raw)
	})
}

// OptionedRange is the deferred-read view returned by Range.Options(...). It
// captures the original Range plus a set of conversion options and runs the
// conversion lazily on Value()/Get().
//
// OptionedRange follows the sugar Chain contract for arena-tracking: the
// underlying Range is reused as-is, so any IDispatch references stay owned
// by the parent context.
type OptionedRange interface {
	// Value returns the decoded value applying every configured option.
	// xlwings parity: `rng.options(...).value`.
	Value() (interface{}, error)
	// Get decodes the read into the supplied destination pointer. Supported
	// destinations:
	//
	//   - *interface{}    — same as Value().
	//   - *string, *float64, *int, *int64, *bool, *time.Time — scalar reads.
	//   - *[]any          — flat 1-D copy of the read.
	//   - *[][]any        — full 2-D grid.
	//   - *[]MyStruct     — struct decode: by header row with Header(true),
	//     positionally (column order = exported field order) without.
	//
	// Returns an error if the read shape and destination cannot be reconciled.
	Get(dst interface{}) error
	// Set writes src into the sheet anchored at this range's top-left cell,
	// auto-resizing to fit — the write-direction mirror of Get. Supported
	// sources:
	//
	//   - []T (T struct)  — one row per element, exported fields in
	//     declaration order. With Header(true) a header row of field names
	//     is written first.
	//   - [][]… 2-D slices — written as a rows×cols block.
	//   - []… 1-D slices   — written as a single row.
	//   - scalars          — plain single-cell write (no resize).
	//
	// xlwings analogue: `rng.options(...).value = data`.
	Set(src interface{}) error
	// Err returns the first deferred error captured while building this
	// OptionedRange (e.g. an invalid Expand direction).
	Err() error
}

// optionedRange is the concrete OptionedRange. We capture the *Range* (typed)
// so Expand can call Range/Range-returning helpers on it without re-reading
// any state.
type optionedRange struct {
	rng  Range
	opts rangeOptions
	err  error
}

// Options is the Go equivalent of xlwings' `Range.options(...)`. Pass any
// combination of Shape (Scalar/Vector/Grid), Expand, Header, Empty,
// DateFormat, or Convert; call .Value() or .Get() on the returned
// OptionedRange to materialize the result.
//
// xlwings reference: https://docs.xlwings.org/en/stable/converters.html
func (r *excelRange) Options(opts ...RangeOption) OptionedRange {
	o := rangeOptions{}
	for _, fn := range opts {
		fn(&o)
	}
	or := &optionedRange{rng: r, opts: o}
	// Resolve Expand eagerly so the user sees a configuration error before
	// the eventual Value() call. The expanded range replaces r in or.rng.
	if o.expand != "" {
		expanded, err := applyExpand(r, o.expand)
		if err != nil {
			or.err = err
		} else {
			or.rng = expanded
		}
	}
	return or
}

// applyExpand walks the COM `Range.End(direction)` chain to grow `anchor`
// into the contiguous block in the requested direction(s). xlwings calls
// these grown ranges "current_region-like" — when the anchor is already
// blank the grown range is just the anchor itself.
//
// Implementation note: Excel COM accepts string addresses for the
// `Worksheet.Range(cell1, cell2)` form, and go-ole's Invoke dispatcher
// likewise marshals strings as `VT_BSTR`. Marshalling chain-wrapped
// IDispatch results would require unwrapping the chain — passing the
// resolved address string is both simpler and matches xlwings' own
// internal "$A$1:$C$2" formulation.
func applyExpand(anchor Range, direction string) (Range, error) {
	switch direction {
	case "down":
		return expandFromEnd(anchor, xlDown)
	case "right":
		return expandFromEnd(anchor, xlToRight)
	case "table":
		startAddr, err := anchor.Get("Cells", 1, 1).Get("Address").Value()
		if err != nil {
			return nil, fmt.Errorf("expand(table): anchor address: %w", err)
		}
		// End(xlDown) from the anchor's top-left, then End(xlToRight) from
		// the resulting row's anchor column gives the bottom-right corner
		// of the contiguous block — matching xlwings' table expansion.
		bottomEnd := anchor.Get("End", xlDown)
		rightEnd := bottomEnd.Get("End", xlToRight)
		endAddr, err := rightEnd.Get("Address").Value()
		if err != nil {
			return nil, fmt.Errorf("expand(table): bottom-right address: %w", err)
		}
		parent := anchor.Get("Worksheet")
		joined := parent.Get("Range", toString(startAddr), toString(endAddr))
		return &excelRange{joined}, nil
	default:
		return nil, fmt.Errorf("Expand: unsupported direction %q (use \"table\", \"down\", or \"right\")", direction)
	}
}

// expandFromEnd creates a new Range spanning anchor through anchor.End(dir).
// Used for the "down" and "right" Expand variants. We resolve both endpoint
// addresses to strings before re-asking Worksheet.Range — see applyExpand's
// note on why we don't pass chains as COM parameters.
func expandFromEnd(anchor Range, direction int32) (Range, error) {
	startAddr, err := anchor.Get("Address").Value()
	if err != nil {
		return nil, fmt.Errorf("expand: anchor address: %w", err)
	}
	endAddr, err := anchor.Get("End", direction).Get("Address").Value()
	if err != nil {
		return nil, fmt.Errorf("expand: end address: %w", err)
	}
	parent := anchor.Get("Worksheet")
	joined := parent.Get("Range", toString(startAddr), toString(endAddr))
	return &excelRange{joined}, nil
}

// Err exposes any deferred construction error (e.g. invalid Expand direction).
func (o *optionedRange) Err() error {
	if o.err != nil {
		return o.err
	}
	return o.rng.Err()
}

// Value reads the underlying range and applies every configured option.
func (o *optionedRange) Value() (interface{}, error) {
	if o.err != nil {
		return nil, o.err
	}
	raw, err := readGrid(o.rng)
	if err != nil {
		return nil, err
	}
	raw = applyIndex(raw, o.opts.index)
	if o.opts.empty != nil {
		applyEmpty(raw, o.opts.empty)
	}
	if o.opts.convert != nil {
		return o.opts.convert(raw)
	}
	return shapeResult(raw, o.opts.shape)
}

// Get decodes into a typed destination. See OptionedRange.Get for supported
// destinations.
func (o *optionedRange) Get(dst interface{}) error {
	if o.err != nil {
		return o.err
	}
	if dst == nil {
		return errors.New("Options.Get: destination is nil")
	}
	dv := reflect.ValueOf(dst)
	if dv.Kind() != reflect.Ptr || dv.IsNil() {
		return fmt.Errorf("Options.Get: destination must be a non-nil pointer, got %T", dst)
	}
	raw, err := readGrid(o.rng)
	if err != nil {
		return err
	}
	raw = applyIndex(raw, o.opts.index)
	if o.opts.empty != nil {
		applyEmpty(raw, o.opts.empty)
	}
	if o.opts.convert != nil {
		out, err := o.opts.convert(raw)
		if err != nil {
			return err
		}
		return assign(dv.Elem(), out)
	}
	// Struct slice decode: detect *[]Struct destinations even when shape is
	// unset. With Header(true) the first row names the target fields; without
	// it every row is data and columns map to exported fields positionally.
	elem := dv.Elem()
	if elem.Kind() == reflect.Slice {
		et := elem.Type().Elem()
		if et.Kind() == reflect.Struct && et != reflect.TypeOf(time.Time{}) {
			if o.opts.header {
				return decodeStructSlice(elem, raw, o.opts.dateFormat)
			}
			return decodeStructSlicePositional(elem, raw, o.opts.dateFormat)
		}
	}
	val, err := shapeResult(raw, o.opts.shape)
	if err != nil {
		return err
	}
	return assign(elem, val)
}

// Set writes src anchored at the range's top-left cell. See
// OptionedRange.Set for the supported source shapes.
func (o *optionedRange) Set(src interface{}) error {
	if o.err != nil {
		return o.err
	}
	if src == nil {
		return errors.New("Options.Set: source is nil")
	}
	rv := reflect.ValueOf(src)
	if rv.Kind() != reflect.Slice {
		return o.rng.SetValue(src).Err()
	}
	et := rv.Type().Elem()
	switch {
	case et.Kind() == reflect.Struct && et != reflect.TypeOf(time.Time{}):
		grid, err := structRowsToGrid(rv, o.opts.header)
		if err != nil {
			return err
		}
		return o.resizeAndSet(len(grid), gridCols(grid), grid)
	case et.Kind() == reflect.Slice:
		rows := rv.Len()
		cols := 0
		if rows > 0 {
			cols = rv.Index(0).Len()
		}
		return o.resizeAndSet(rows, cols, src)
	default:
		return o.resizeAndSet(1, rv.Len(), src)
	}
}

func gridCols(grid [][]interface{}) int {
	if len(grid) == 0 {
		return 0
	}
	return len(grid[0])
}

// resizeAndSet grows the anchor range to rows×cols and writes v into it.
// Zero-sized sources are a no-op.
func (o *optionedRange) resizeAndSet(rows, cols int, v interface{}) error {
	if rows == 0 || cols == 0 {
		return nil
	}
	return o.rng.Resize(rows, cols).SetValue(v).Err()
}

// structRowsToGrid flattens a slice of structs into a [][]interface{} —
// exported fields in declaration order, optionally preceded by a header row
// of field names. The write-direction mirror of decodeStructSlice /
// decodeStructSlicePositional.
func structRowsToGrid(rv reflect.Value, includeHeader bool) ([][]interface{}, error) {
	st := rv.Type().Elem()
	var fields []reflect.StructField
	for fi := 0; fi < st.NumField(); fi++ {
		if st.Field(fi).PkgPath == "" { // exported only
			fields = append(fields, st.Field(fi))
		}
	}
	if len(fields) == 0 {
		return nil, fmt.Errorf("Options.Set: struct %s has no exported fields", st)
	}
	grid := make([][]interface{}, 0, rv.Len()+1)
	if includeHeader {
		hdr := make([]interface{}, len(fields))
		for i, f := range fields {
			hdr[i] = f.Name
		}
		grid = append(grid, hdr)
	}
	for r := 0; r < rv.Len(); r++ {
		item := rv.Index(r)
		row := make([]interface{}, len(fields))
		for i, f := range fields {
			row[i] = item.FieldByIndex(f.Index).Interface()
		}
		grid = append(grid, row)
	}
	return grid, nil
}

// readGrid normalises any Range.Value() result into [][]interface{} so the
// downstream option pipeline sees a single shape. Single-cell results become
// a 1x1 grid; 1-D slice results (rare — Excel returns 2-D for multi-cell
// ranges) become a 1xN row.
func readGrid(r Range) ([][]interface{}, error) {
	v, err := r.Value()
	if err != nil {
		return nil, err
	}
	switch t := v.(type) {
	case nil:
		return [][]interface{}{{nil}}, nil
	case [][]interface{}:
		return t, nil
	case []interface{}:
		return [][]interface{}{t}, nil
	default:
		return [][]interface{}{{t}}, nil
	}
}

// applyIndex drops the first n columns of every row — the xlwings
// `index=n` analogue. Rows shorter than n become empty rows rather than
// erroring, matching the lenient decode style of the rest of the pipeline.
func applyIndex(raw [][]interface{}, n int) [][]interface{} {
	if n <= 0 {
		return raw
	}
	out := make([][]interface{}, len(raw))
	for r := range raw {
		if n >= len(raw[r]) {
			out[r] = []interface{}{}
			continue
		}
		out[r] = raw[r][n:]
	}
	return out
}

// applyEmpty walks every cell in raw and replaces nil with the configured
// substitute. xlwings parity: `.options(empty="N/A")`.
func applyEmpty(raw [][]interface{}, fill interface{}) {
	for r := range raw {
		for c := range raw[r] {
			if raw[r][c] == nil {
				raw[r][c] = fill
			}
		}
	}
}

// shapeResult coerces a [][]interface{} into the shape the caller asked for.
// NDimAuto reproduces xlwings' implicit rules: 1×1 → scalar, 1×N or N×1 →
// flat slice, everything else → 2-D grid.
func shapeResult(raw [][]interface{}, shape NDim) (interface{}, error) {
	rows := len(raw)
	cols := 0
	if rows > 0 {
		cols = len(raw[0])
	}
	switch shape {
	case NDimScalar:
		if rows != 1 || cols != 1 {
			return nil, fmt.Errorf("Options.Scalar: range is %dx%d, expected 1x1", rows, cols)
		}
		return raw[0][0], nil
	case NDimVector:
		return flatten(raw)
	case NDimGrid:
		return raw, nil
	default: // NDimAuto
		if rows == 1 && cols == 1 {
			return raw[0][0], nil
		}
		if rows == 1 || cols == 1 {
			out, _ := flatten(raw)
			return out, nil
		}
		return raw, nil
	}
}

// flatten turns a 2-D read into a 1-D slice. Used by NDimVector. Returns an
// error if the input is genuinely 2-D (multiple rows AND multiple cols).
func flatten(raw [][]interface{}) ([]interface{}, error) {
	rows := len(raw)
	if rows == 0 {
		return []interface{}{}, nil
	}
	cols := len(raw[0])
	if rows == 1 {
		out := make([]interface{}, cols)
		copy(out, raw[0])
		return out, nil
	}
	if cols == 1 {
		out := make([]interface{}, rows)
		for i := range raw {
			out[i] = raw[i][0]
		}
		return out, nil
	}
	return nil, fmt.Errorf("Options.Vector: range is %dx%d; one dimension must be 1", rows, cols)
}

// decodeStructSlice walks a [][]interface{} where row 0 is field headers and
// rows 1..N are records, populating dst (a reflect.Value pointing to a slice
// of structs) with one struct per data row. xlwings parity: the struct-slice
// equivalent of `.options(pd.DataFrame, header=1, ...).value`.
//
// Field matching is case-insensitive; an exact match wins over a fold match.
// Unknown headers are silently skipped (consistent with xlwings' lenient
// pandas decode). Missing headers leave the struct field at its zero value.
func decodeStructSlice(dst reflect.Value, raw [][]interface{}, dateFormat string) error {
	if len(raw) == 0 {
		dst.Set(reflect.MakeSlice(dst.Type(), 0, 0))
		return nil
	}
	headers := raw[0]
	st := dst.Type().Elem()
	// Build header-name -> field-index map.
	colToField := make([]int, len(headers))
	for i := range colToField {
		colToField[i] = -1
	}
	for ci, h := range headers {
		name := strings.TrimSpace(fmt.Sprint(h))
		if name == "" {
			continue
		}
		// exact match first
		if f, ok := st.FieldByName(name); ok {
			colToField[ci] = f.Index[0]
			continue
		}
		// case-insensitive fallback (exported fields only — an unexported
		// fold match would fail CanSet later)
		for fi := 0; fi < st.NumField(); fi++ {
			if st.Field(fi).PkgPath == "" && strings.EqualFold(st.Field(fi).Name, name) {
				colToField[ci] = fi
				break
			}
		}
	}
	out := reflect.MakeSlice(dst.Type(), 0, len(raw)-1)
	for r := 1; r < len(raw); r++ {
		row := raw[r]
		item := reflect.New(st).Elem()
		for ci, fi := range colToField {
			if fi < 0 || ci >= len(row) {
				continue
			}
			if err := assignField(item.Field(fi), row[ci], dateFormat); err != nil {
				return fmt.Errorf("row %d, column %q: %w", r, headers[ci], err)
			}
		}
		out = reflect.Append(out, item)
	}
	dst.Set(out)
	return nil
}

// decodeStructSlicePositional decodes every row of raw into one struct,
// mapping column 0 to the struct's first exported field, column 1 to the
// second, and so on — the Header(false) counterpart of decodeStructSlice.
// Extra columns are ignored; missing columns leave fields at their zero
// value.
func decodeStructSlicePositional(dst reflect.Value, raw [][]interface{}, dateFormat string) error {
	st := dst.Type().Elem()
	var fields []int
	for fi := 0; fi < st.NumField(); fi++ {
		if st.Field(fi).PkgPath == "" { // exported only
			fields = append(fields, fi)
		}
	}
	out := reflect.MakeSlice(dst.Type(), 0, len(raw))
	for r, row := range raw {
		item := reflect.New(st).Elem()
		for ci, fi := range fields {
			if ci >= len(row) {
				break
			}
			if err := assignField(item.Field(fi), row[ci], dateFormat); err != nil {
				return fmt.Errorf("row %d, column %d: %w", r, ci, err)
			}
		}
		out = reflect.Append(out, item)
	}
	dst.Set(out)
	return nil
}

// assign writes `val` into the addressable reflect.Value `dst`. Used for the
// non-struct Options.Get path; performs the same coercions assignField does
// for scalar destinations.
func assign(dst reflect.Value, val interface{}) error {
	if !dst.CanSet() {
		return errors.New("Options.Get: destination is not settable")
	}
	// Fast path: interface{} destination accepts anything.
	if dst.Kind() == reflect.Interface && dst.NumMethod() == 0 {
		dst.Set(reflect.ValueOf(val))
		return nil
	}
	if val == nil {
		dst.Set(reflect.Zero(dst.Type()))
		return nil
	}
	vv := reflect.ValueOf(val)
	if vv.Type().AssignableTo(dst.Type()) {
		dst.Set(vv)
		return nil
	}
	if vv.Type().ConvertibleTo(dst.Type()) {
		dst.Set(vv.Convert(dst.Type()))
		return nil
	}
	return fmt.Errorf("Options.Get: cannot assign %T to %s", val, dst.Type())
}

// assignField mirrors `assign` but knows the dateFormat hint for time.Time
// values landing in string fields.
func assignField(dst reflect.Value, val interface{}, dateFormat string) error {
	if !dst.CanSet() {
		return errors.New("field is not settable")
	}
	if val == nil {
		dst.Set(reflect.Zero(dst.Type()))
		return nil
	}
	// Special case: time.Time -> string (formatted) when DateFormat is set.
	if t, ok := val.(time.Time); ok && dst.Kind() == reflect.String {
		if dateFormat != "" {
			dst.SetString(t.Format(dateFormat))
		} else {
			dst.SetString(t.String())
		}
		return nil
	}
	vv := reflect.ValueOf(val)
	if vv.Type().AssignableTo(dst.Type()) {
		dst.Set(vv)
		return nil
	}
	if vv.Type().ConvertibleTo(dst.Type()) {
		dst.Set(vv.Convert(dst.Type()))
		return nil
	}
	// Best-effort string conversion: Excel often hands strings to numeric
	// destinations via TEXT-typed cells. Use Sprint as a last resort.
	if dst.Kind() == reflect.String {
		dst.SetString(fmt.Sprint(val))
		return nil
	}
	return fmt.Errorf("cannot assign %T to %s", val, dst.Type())
}

// Compile-time check that sugar.Chain stays satisfied by excelRange via the
// embedded Chain (we did not change that — this is a guard against future
// refactors).
var _ sugar.Chain = (*excelRange)(nil)
