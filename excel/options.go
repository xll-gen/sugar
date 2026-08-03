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
	//   - []T (T struct)  — one row per element, one column per flattened
	//     exported leaf field in declaration order: anonymous embedded structs
	//     are expanded in place, while a named struct field and an embedded
	//     time.Time each stay a single value cell. With Header(true) a header
	//     row of those leaf names is written first — and because a leaf's name
	//     is also its Go promoted name, the grid reads straight back through
	//     Options(Header(true)).Get.
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
	// Validate the Expand direction eagerly (a config typo surfaces before the
	// read), but DEFER the actual expansion to Value()/Get(). xlwings evaluates
	// options only on value access, so a stored OptionedRange re-discovers the
	// current data block on every read — a range that grows after Options() is
	// captured must include the new rows/columns. See effectiveRange.
	if o.expand != "" && !validExpandDirection(o.expand) {
		or.err = fmt.Errorf("Expand: unsupported direction %q (use \"table\", \"down\", or \"right\")", o.expand)
	}
	return or
}

// validExpandDirection reports whether s is one of the supported Expand
// directions. Used by Options() for early validation without running the
// (deferred) expansion.
func validExpandDirection(s string) bool {
	switch s {
	case "table", "down", "right":
		return true
	}
	return false
}

// effectiveRange resolves the range to read from. When an Expand direction is
// configured, applyExpand is re-run here — at read time — so each Value()/Get()
// sees the current extent of the data block (xlwings' "options are only
// evaluated when accessing the values" semantics). Without Expand it returns
// the original anchor range unchanged.
func (o *optionedRange) effectiveRange() (Range, error) {
	if o.opts.expand == "" {
		return o.rng, nil
	}
	return applyExpand(o.rng, o.opts.expand)
}

// applyExpand walks the COM `Range.End(direction)` chain to grow `anchor`
// into the contiguous block in the requested direction(s), from the anchor's
// top-left cell (its "origin").
//
// Blank-neighbor guard (xlwings parity): before calling End() in a direction
// we check the cell immediately adjacent to the origin in that direction. If
// that neighbor is empty, End() would jump to the sheet boundary (row
// 1,048,576 or column XFD) and drag in every blank cell up to a distant data
// island, so we do NOT expand that dimension — the origin is its own endpoint.
// This mirrors xlwings' expansion.py, which only calls end() when the adjacent
// raw_value is non-empty. For "table" both the down and right dimensions are
// guarded independently. The guard is a three-rung ladder, not a single probe
// — see endpointCell for what the second rung buys.
//
// # Multi-area anchors
//
// A multi-area anchor (`Range("A1:C1,E1:F1")`, or any comma-joined address) is
// reduced to its FIRST AREA and the result is always ONE rectangle. That is not
// a choice made here: Excel itself reports `Cells(1,1)`, `Rows.Count` and
// `Columns.Count` for the first area only, and those are the three values the
// expansion consumes. No error is returned.
//
// Returning an error was considered and REJECTED. xlwings has no Areas concept
// at all, and `_xlwindows.Range` rebuilds any range from its first-area coords
// (`coords` = sheet/row/column/Rows.Count/Columns.Count, then
// `Range(Cells(row, col), Cells(row+nrows-1, col+ncols-1))`), so upstream
// rectangularizes a multi-area range even more aggressively than sugar does.
// Erroring would therefore both break every caller that passes a multi-area
// address today and move sugar away from the behavior it is modelled on. sugar
// exposes no `Areas` API, so there is nothing else to reconcile.
// Pinned by TestExpand_MultiAreaAnchorUsesFirstAreaOnly.
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
		origin := anchor.Cells(1, 1)
		// Bottom-left corner (grow down) and top-right corner (grow right),
		// each guarded against a blank neighbor. Worksheet.Range(cornerA,
		// cornerB) is the bounding rectangle of the two, i.e. the full block
		// origin-row..bottom-row × origin-col..right-col.
		bottomAddr, err := endpointAddr(origin, xlDown)
		if err != nil {
			return nil, fmt.Errorf("expand(table): bottom address: %w", err)
		}
		rightAddr, err := endpointAddr(origin, xlToRight)
		if err != nil {
			return nil, fmt.Errorf("expand(table): right address: %w", err)
		}
		parent := anchor.Get("Worksheet")
		joined := parent.Get("Range", bottomAddr, rightAddr)
		return wrapRange(joined), nil
	default:
		return nil, fmt.Errorf("Expand: unsupported direction %q (use \"table\", \"down\", or \"right\")", direction)
	}
}

// expandFromEnd creates a new Range spanning the anchor's origin through the
// far corner of the block grown in `direction`. Used for the "down" and "right"
// Expand variants. We resolve both corner addresses to strings before
// re-asking Worksheet.Range — see applyExpand's note on why we don't pass
// chains as COM parameters.
//
// The far corner is NOT the End(direction) cell itself: a multi-cell anchor
// keeps its extent on the axis *perpendicular* to the growth direction, exactly
// as the "table" branch keeps both axes. `Range("A1:C1")` grown down is
// A1:C<end>, never A1:A<end> — building the rectangle from two addresses that
// both sit in the anchor's first column collapses it to a single column and
// silently truncates the read (columns B and C would vanish with err == nil).
// xlwings' VerticalExpander does the same, ending its range at
// `(end_row, rng.column + rng.shape[1] - 1)`.
//
// On a multi-area anchor every input this reads — Cells(1,1) and the crossSpan
// counts — is first-area-only in Excel, so the result is the first area's block
// grown in `direction`, as one rectangle. See applyExpand for why that is
// deliberate and why it is not an error.
func expandFromEnd(anchor Range, direction int32) (Range, error) {
	origin := anchor.Cells(1, 1)
	startAddr, err := origin.Address()
	if err != nil {
		return nil, fmt.Errorf("expand: anchor address: %w", err)
	}
	cross, err := crossSpan(anchor, direction)
	if err != nil {
		return nil, fmt.Errorf("expand: anchor span: %w", err)
	}
	end, err := endpointCell(origin, direction)
	if err != nil {
		return nil, fmt.Errorf("expand: end cell: %w", err)
	}
	// Widen (down) / deepen (right) the endpoint cell back to the anchor's own
	// span so the two addresses bound the whole rectangle.
	if dr, dc := expandCornerOffset(direction, cross); dr != 0 || dc != 0 {
		end = end.Offset(dr, dc)
	}
	endAddr, err := end.Address()
	if err != nil {
		return nil, fmt.Errorf("expand: end address: %w", err)
	}
	parent := anchor.Get("Worksheet")
	joined := parent.Get("Range", startAddr, endAddr)
	return wrapRange(joined), nil
}

// crossSpan returns the anchor's size along the axis perpendicular to the
// expansion direction — the column count when growing down, the row count when
// growing right. That is the span the expanded rectangle must preserve. Other
// directions report 1 (no cross-axis widening); "table" derives both extents
// from the data block itself and never calls this.
//
// On a multi-area anchor Excel's Rows.Count / Columns.Count describe the FIRST
// AREA, so that is the span preserved here — see applyExpand.
func crossSpan(anchor Range, direction int32) (int, error) {
	var (
		span int32
		err  error
	)
	switch direction {
	case xlDown:
		span, err = anchor.Columns().Count()
	case xlToRight:
		span, err = anchor.Rows().Count()
	default:
		return 1, nil
	}
	if err != nil {
		return 0, err
	}
	if span < 1 {
		// A COM Count of 0 (or an unexpected VARIANT shape narrowing to 0)
		// would otherwise shift the corner backwards and invert the rectangle.
		span = 1
	}
	return int(span), nil
}

// expandCornerOffset returns the (row, col) delta from the End(direction)
// endpoint cell to the far corner of the expanded rectangle, given the anchor's
// perpendicular span (see crossSpan). Growing down moves the corner right by
// the anchor's column span; growing right moves it down by the row span. A
// single-cell-wide cross axis needs no shift, which keeps the common 1x1 anchor
// on exactly the same COM traffic as before.
//
// Pure arithmetic, so it is verifiable without Excel.
// (The parameter is named `cross` rather than `crossSpan` so it does not shadow
// the crossSpan function inside this scope.)
func expandCornerOffset(direction int32, cross int) (rowOff, colOff int) {
	if cross <= 1 {
		return 0, 0
	}
	switch direction {
	case xlDown:
		return 0, cross - 1
	case xlToRight:
		return cross - 1, 0
	}
	return 0, 0
}

// endpointCell returns the far cell of the contiguous block starting at the
// single-cell origin in the given direction, applying the blank-neighbor guard
// (see applyExpand): when the adjacent cell is empty the origin is its own
// endpoint, so End() is never called into empty space.
//
// The three rungs mirror xlwings' expansion.py exactly (TableExpander /
// VerticalExpander / HorizontalExpander all use the same ladder):
//
//	neighbor blank        -> origin is its own endpoint
//	second neighbor blank -> the neighbor is the endpoint
//	otherwise             -> neighbor.End(direction)
//
// Two things hang on the last rung starting from the *neighbor* rather than
// from the origin, which is what sugar used to do:
//
//   - A blank origin no longer truncates the block. Excel's End() from an
//     empty cell jumps to the *first* non-empty cell instead of the last cell
//     of a run, so `A1:E10` with an empty top-left corner (a table whose
//     header row starts at B1 and whose row labels start at A2 — the most
//     ordinary spreadsheet layout there is) expanded down collapsed to A1:A2
//     and read two cells instead of ten. End() is now only ever called from a
//     cell the ladder has already proven non-empty.
//   - The middle rung is what makes that safe for a two-cell block: with the
//     second neighbor blank, End() from the neighbor would jump *past* the
//     block to the next data island (or the sheet edge), so the neighbor is
//     returned directly without calling End() at all.
//
// For a non-empty origin every rung agrees with the old single probe, so this
// widens the accepted layouts without moving any case that already worked.
func endpointCell(origin Range, direction int32) (Range, error) {
	dr, dc := neighborOffset(direction)
	if dr == 0 && dc == 0 {
		return origin, nil
	}

	neighbor := origin.Offset(dr, dc)
	blank, err := cellBlank(neighbor)
	if err != nil {
		return nil, err
	}
	if blank {
		return origin, nil
	}

	blank, err = cellBlank(origin.Offset(2*dr, 2*dc))
	if err != nil {
		return nil, err
	}
	if blank {
		return neighbor, nil
	}

	return wrapRange(neighbor.Get("End", direction)), nil
}

// endpointAddr is endpointCell's address — the form the "table" branch needs to
// build its bounding box from two opposite corners.
func endpointAddr(origin Range, direction int32) (string, error) {
	end, err := endpointCell(origin, direction)
	if err != nil {
		return "", err
	}
	return end.Address()
}

// neighborOffset returns the (row, col) delta from an anchor to the cell
// immediately adjacent in the given End() direction.
func neighborOffset(direction int32) (int, int) {
	switch direction {
	case xlDown:
		return 1, 0
	case xlToRight:
		return 0, 1
	}
	return 0, 0
}

// cellBlank reports whether a single-cell range holds no data — nil (an empty
// cell) or an empty string. Matches xlwings' `raw_value in (None, "")` test.
func cellBlank(cell Range) (bool, error) {
	v, err := cell.Value()
	if err != nil {
		return false, err
	}
	switch t := v.(type) {
	case nil:
		return true, nil
	case string:
		return t == "", nil
	}
	return false, nil
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
	rng, err := o.effectiveRange()
	if err != nil {
		return nil, err
	}
	raw, err := readGrid(rng)
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
	rng, err := o.effectiveRange()
	if err != nil {
		return err
	}
	raw, err := readGrid(rng)
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
		if et.Kind() == reflect.Struct && et != timeType {
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
	case et.Kind() == reflect.Struct && et != timeType:
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

// timeType is the reflect.Type of time.Time — the one struct the conversion
// layer treats as a scalar cell value everywhere (see Get's struct-slice
// detection, Set's element-type switch and structFields' leaf rule).
var timeType = reflect.TypeOf(time.Time{})

// structFields returns the flattened, ordered leaf fields of st: exported
// top-level fields in declaration order, with ANONYMOUS embedded structs
// expanded depth-first in place. Each returned StructField carries the full
// multi-level .Index path, so item.FieldByIndex(f.Index) reaches the leaf.
//
// It is the single field-collection rule shared by the two write/positional-read
// sites, and it is deliberately congruent with what the header path
// (decodeStructSlice) gets from FieldByName / FieldByNameFunc — which traverse
// embedded structs and already return a multi-level Index. Before this existed
// there were three rules: the header path flattened, while the positional decode
// and the Set grid builder both scanned only st.NumField(), so an anonymous
// embedded struct became ONE column and then failed hard ("cannot assign
// float64 to Base" on read, "unsupported cell type" on write). Nothing worked on
// either of those paths, which is why unifying them cannot regress a success.
//
// Three carve-outs, each of which would otherwise break behaviour that works:
//
//  1. time.Time is a LEAF. It is a struct with no exported fields, so expanding
//     it would contribute zero columns and silently drop the field — yet an
//     embedded time.Time decodes correctly today, because AssignableTo matches
//     at field level. (Carve-out 3 would catch time.Time as well, since it has
//     no exported leaves, so removing THIS test alone changes no output —
//     verified by mutation. It is kept as the named, intentional rule: the
//     conversion layer treats time.Time as a scalar everywhere else too, and a
//     future time.Time with an exported field must not silently re-shape every
//     grid.)
//  2. Only ANONYMOUS embedded fields expand. A named struct field stays one
//     cell, matching Go's promotion rule and the header path, which cannot see
//     inside it either. Embedded *pointers* to structs are also leaves: the
//     header path reaches through them via FieldByIndexErr, but a write has no
//     sane column count for a nil one.
//  3. An embedded struct with no exported leaves stays a single column rather
//     than vanishing, so the surrounding column indices never shift silently.
//
// Ambiguity is refused, not resolved. When a flattened leaf's promoted name does
// not resolve back through st.FieldByName to the same field — two embeds sharing
// a leaf name (FieldByName reports not-ok), or an outer field shadowing an
// embedded one — Set would write a column Get could not read back into the same
// place. Inventing a resolution the header path does not share is exactly how
// the three rules diverged in the first place, so such a struct is an error.
func structFields(st reflect.Type) ([]reflect.StructField, error) {
	var out []reflect.StructField
	collectStructFields(st, nil, &out)
	for _, f := range out {
		g, ok := st.FieldByName(f.Name)
		if !ok || !sameIndex(g.Index, f.Index) {
			return nil, fmt.Errorf(
				"struct %s: field %q is ambiguous or shadowed under Go's field-promotion "+
					"rules, so a header column written for it cannot be read back — "+
					"rename it or drop the embedding", st, f.Name)
		}
	}
	return out, nil
}

// collectStructFields is structFields' depth-first walker. prefix is the Index
// path of the embedded chain currently being expanded.
func collectStructFields(st reflect.Type, prefix []int, out *[]reflect.StructField) {
	for i := 0; i < st.NumField(); i++ {
		f := st.Field(i)
		if f.Anonymous && f.Type.Kind() == reflect.Struct && f.Type != timeType {
			before := len(*out)
			collectStructFields(f.Type, append(prefix, i), out)
			if len(*out) > before {
				continue
			}
			// Carve-out 3: no exported leaves inside — fall through and keep
			// the embedded struct itself as one column.
		}
		if f.PkgPath != "" {
			// Unexported and not expandable. (An unexported *embedded* type is
			// still expanded above: its exported leaves are settable.)
			continue
		}
		leaf := f
		leaf.Index = append(append([]int{}, prefix...), i)
		*out = append(*out, leaf)
	}
}

func sameIndex(a, b []int) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

// structRowsToGrid flattens a slice of structs into a [][]interface{} — the
// flattened exported leaf fields (see structFields) in declaration order,
// optionally preceded by a header row of field names. The write-direction mirror
// of decodeStructSlice / decodeStructSlicePositional.
//
// The header row carries each leaf's own name, which is also its Go *promoted*
// name — so decodeStructSlice's FieldByName lookup resolves the column straight
// back to the field it came from and the Set -> Get round trip closes.
func structRowsToGrid(rv reflect.Value, includeHeader bool) ([][]interface{}, error) {
	st := rv.Type().Elem()
	fields, err := structFields(st)
	if err != nil {
		return nil, fmt.Errorf("Options.Set: %w", err)
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
			// FieldByIndexErr (not FieldByIndex) so a leaf reached through a nil
			// embedded pointer writes an empty cell instead of panicking — the
			// same tolerance decodeStructSlice already has on the read side.
			// (structFields does not expand pointer embeds today, so this is a
			// guard against a future widening, not a live path.)
			fv, err := item.FieldByIndexErr(f.Index)
			if err != nil {
				row[i] = nil
				continue
			}
			row[i] = fv.Interface()
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
	// Build header-name -> field-index-path map. Paths (not single indices)
	// because FieldByName resolves promoted fields from embedded structs to a
	// multi-level Index; a nil path means the column has no matching field.
	colToField := make([][]int, len(headers))
	for ci, h := range headers {
		name := strings.TrimSpace(fmt.Sprint(h))
		if name == "" {
			continue
		}
		// Exact match first (case-sensitive); both lookups traverse embedded
		// structs and yield a full Index path. We require an exported leaf
		// (PkgPath == "") so an unexported fold match can't slip through and
		// fail CanSet later.
		if f, ok := st.FieldByName(name); ok && f.PkgPath == "" {
			colToField[ci] = f.Index
			continue
		}
		// Case-insensitive fallback. FieldByNameFunc (not a top-level NumField
		// scan) so headers also match fields promoted from embedded structs.
		if f, ok := st.FieldByNameFunc(func(s string) bool { return strings.EqualFold(s, name) }); ok && f.PkgPath == "" {
			colToField[ci] = f.Index
		}
	}
	out := reflect.MakeSlice(dst.Type(), 0, len(raw)-1)
	for r := 1; r < len(raw); r++ {
		row := raw[r]
		item := reflect.New(st).Elem()
		for ci, path := range colToField {
			if path == nil || ci >= len(row) {
				continue
			}
			// FieldByIndexErr (not FieldByIndex) so a field promoted through a
			// nil embedded *pointer* is skipped gracefully instead of panicking.
			fv, err := item.FieldByIndexErr(path)
			if err != nil {
				continue
			}
			if err := assignField(fv, row[ci], dateFormat); err != nil {
				return fmt.Errorf("row %d, column %q: %w", r, headers[ci], err)
			}
		}
		out = reflect.Append(out, item)
	}
	dst.Set(out)
	return nil
}

// decodeStructSlicePositional decodes every row of raw into one struct,
// mapping column 0 to the struct's first flattened exported leaf field, column 1
// to the second, and so on — the Header(false) counterpart of decodeStructSlice.
// Extra columns are ignored; missing columns leave fields at their zero
// value.
//
// The column list comes from structFields, the same collector structRowsToGrid
// writes with, so a positional Set and a positional Get agree column for column
// (and anonymous embedded structs are flattened here exactly as the header path
// flattens them via FieldByName).
func decodeStructSlicePositional(dst reflect.Value, raw [][]interface{}, dateFormat string) error {
	st := dst.Type().Elem()
	fields, err := structFields(st)
	if err != nil {
		return err
	}
	out := reflect.MakeSlice(dst.Type(), 0, len(raw))
	for r, row := range raw {
		item := reflect.New(st).Elem()
		for ci, f := range fields {
			if ci >= len(row) {
				break
			}
			// FieldByIndexErr so a leaf behind a nil embedded pointer is skipped
			// rather than panicking, matching decodeStructSlice.
			fv, err := item.FieldByIndexErr(f.Index)
			if err != nil {
				continue
			}
			if err := assignField(fv, row[ci], dateFormat); err != nil {
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
	// nil must be handled before the interface{} fast path: reflect.ValueOf(nil)
	// is the zero Value, and Set(zero Value) panics. An empty cell (VT_EMPTY)
	// flows here as nil — via readGrid's [][]interface{}{{nil}} → shapeResult →
	// Get(&v) — and *interface{} is a documented destination, so this is a
	// common input, not an edge case. Leaving dst at its zero value is correct
	// for every kind (nil interface, "", 0, ...). Mirrors assignField's order.
	if val == nil {
		dst.Set(reflect.Zero(dst.Type()))
		return nil
	}
	// Fast path: interface{} destination accepts anything.
	if dst.Kind() == reflect.Interface && dst.NumMethod() == 0 {
		dst.Set(reflect.ValueOf(val))
		return nil
	}
	vv := reflect.ValueOf(val)
	if vv.Type().AssignableTo(dst.Type()) {
		dst.Set(vv)
		return nil
	}
	if setStringFromInteger(dst, vv, val) {
		return nil
	}
	if vv.Type().ConvertibleTo(dst.Type()) {
		dst.Set(vv.Convert(dst.Type()))
		return nil
	}
	return fmt.Errorf("Options.Get: cannot assign %T to %s", val, dst.Type())
}

// setStringFromInteger writes an INTEGER-kinded source into a string-kinded
// destination as its decimal digits and reports true; it reports false (writing
// nothing) for every other type pairing.
//
// This guard must run BEFORE the ConvertibleTo branch in assign/assignField,
// because Go's reflect Convert reads an integer -> string conversion as
// "the string containing that CODE POINT", not "the digits": int32(65) becomes
// "A", int(0) becomes "\x00", int64(1) becomes "\x01", and time.Duration(1e9)
// becomes U+FFFD garbage. Nothing in a spreadsheet pipeline wants that, and it
// used to happen silently — the reachable trigger is Options(Empty(0)), whose
// substituted int 0 landed in string struct fields as a NUL byte.
//
// It is keyed on the SOURCE kind, which is what keeps it narrow: []byte and
// []rune are reflect.Slice (not Uint8/Int32), so `[]byte("hi") -> "hi"` still
// goes through ConvertibleTo unchanged; a named string source keeps its
// conversion; float sources were never ConvertibleTo string and still reach
// assignField's Sprint fallback. dst.Kind() (not dst.Type()) is tested so a
// NAMED string destination is covered too — SetString works on those.
//
// fmt.Sprint rather than strconv so a named integer type with a String()
// method renders sensibly (time.Duration -> "1s").
//
// This is a guard, not redundancy (AGENTS §5 rule 10): deleting it restores the
// rune conversion, and `TestAssign_IntegerIntoStringIsDigitsNotRune` /
// `TestAssignField_EmptyIntIntoStringField` are what say so.
func setStringFromInteger(dst reflect.Value, vv reflect.Value, val interface{}) bool {
	if dst.Kind() != reflect.String {
		return false
	}
	switch vv.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		dst.SetString(fmt.Sprint(val))
		return true
	}
	return false
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
	// Integer source into a string destination: digits, never a rune. Must sit
	// ahead of ConvertibleTo — see setStringFromInteger. (The Sprint fallback
	// below would produce the same text, but it is unreachable for integer
	// sources because ConvertibleTo matches them first.)
	if setStringFromInteger(dst, vv, val) {
		return nil
	}
	if vv.Type().ConvertibleTo(dst.Type()) {
		dst.Set(vv.Convert(dst.Type()))
		return nil
	}
	// The VT_NULL sentinel must never reach the Sprint fallback below.
	// sugar.Null HAS a String() method, so Sprint would write the literal text
	// "Null" into the field and it would read as genuine cell data — the same
	// forgery class as the "[[=1+1 =2+2]]" string stringFromVariant refuses.
	// This sits AFTER the AssignableTo branch on purpose: an interface{} field
	// still receives the sentinel, which is what a caller asking for raw values
	// needs. (assign needs no such arm — a struct source is neither assignable
	// nor convertible to a scalar destination, so it already errors.)
	if sugar.IsNull(val) {
		return fmt.Errorf(
			"cannot assign %T to %s: the cell is Null (no single value); "+
				"decode into an interface{} field and test it with sugar.IsNull",
			val, dst.Type())
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
