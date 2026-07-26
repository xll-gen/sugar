//go:build windows

// Unit tests for the SAFEARRAY encode path. SafeArray* APIs live in
// oleaut32 and do not require CoInitialize, so encode→decode round trips
// run without any COM server — these are fast, Excel-free tests.

package sugar

import (
	"fmt"
	"math"
	"reflect"
	"testing"
	"time"
	"unsafe"

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

// TestDecodeVariantArray_TypedArrayRejected builds a *typed* SAFEARRAY
// (VT_ARRAY|VT_BSTR — the shape a general COM server can return, not just
// Excel's VT_VARIANT) and confirms decodeVariantArray rejects it instead of
// feeding VT_VARIANT-assuming getElement a poisoned buffer (which silently
// misdecoded and leaked a BSTR per string cell). SafeArray* APIs need no
// CoInitialize, so this is Excel-free.
func TestDecodeVariantArray_TypedArrayRejected(t *testing.T) {
	bounds := []safeArrayBound{{cElements: 2, lLbound: 0}}
	sa, _, _ := procSafeArrayCreate.Call(
		uintptr(ole.VT_BSTR),
		uintptr(uint32(1)),
		uintptr(unsafe.Pointer(&bounds[0])),
	)
	if sa == 0 {
		t.Fatal("SafeArrayCreate(VT_BSTR) failed")
	}
	defer procSafeArrayDestroy.Call(sa)

	v := ole.NewVariant(ole.VT_ARRAY|ole.VT_BSTR, int64(sa))
	if _, err := decodeVariantArray(&v); err == nil {
		t.Error("expected error decoding a typed (VT_BSTR) SAFEARRAY, got nil")
	}
}

// TestDecodeVariantArray_ByrefRejected confirms a VT_BYREF array VARIANT is
// rejected before its Val (a SAFEARRAY**, not SAFEARRAY*) is dereferenced.
func TestDecodeVariantArray_ByrefRejected(t *testing.T) {
	v := ole.NewVariant(ole.VT_ARRAY|ole.VT_VARIANT|ole.VT_BYREF, 0)
	if _, err := decodeVariantArray(&v); err == nil {
		t.Error("expected error decoding a VT_BYREF array, got nil")
	}
}

// TestNormalizeParams_TimeScalar is the Excel-free unit cover for defect 3:
// a scalar time.Time argument must be normalized to a VT_DATE VARIANT (the
// same wall-clock encoding a [][]any block uses) rather than handed to go-ole,
// which would marshal it as a locale-dependent VT_BSTR string.
func TestNormalizeParams_TimeScalar(t *testing.T) {
	ts := time.Date(2026, 7, 2, 9, 30, 0, 0, time.UTC)
	out, cleanup, err := normalizeParams([]interface{}{ts})
	if err != nil {
		t.Fatalf("normalizeParams: %v", err)
	}
	defer cleanup()
	if len(out) != 1 {
		t.Fatalf("expected 1 arg, got %d", len(out))
	}
	v, ok := out[0].(*ole.VARIANT)
	if !ok {
		t.Fatalf("expected *ole.VARIANT, got %T", out[0])
	}
	if v.VT != ole.VT_DATE {
		t.Errorf("expected VT_DATE, got VT=0x%x", v.VT)
	}
	// The serial must match the standalone scalarToVariant encoding.
	want := dateSerial(t, ts)
	if got := math.Float64frombits(uint64(v.Val)); got != want {
		t.Errorf("VT_DATE serial: got %v, want %v", got, want)
	}
}

// TestNormalizeParams_TimePointer covers the *time.Time branch (nil stays nil;
// non-nil encodes VT_DATE).
func TestNormalizeParams_TimePointer(t *testing.T) {
	ts := time.Date(2026, 7, 2, 9, 30, 0, 0, time.UTC)
	out, cleanup, err := normalizeParams([]interface{}{&ts})
	if err != nil {
		t.Fatalf("normalizeParams(*time.Time): %v", err)
	}
	defer cleanup()
	v, ok := out[0].(*ole.VARIANT)
	if !ok || v.VT != ole.VT_DATE {
		t.Fatalf("expected VT_DATE VARIANT, got %T (VT=%v)", out[0], out[0])
	}

	var nilPtr *time.Time
	out2, cleanup2, err := normalizeParams([]interface{}{nilPtr})
	if err != nil {
		t.Fatalf("normalizeParams(nil *time.Time): %v", err)
	}
	defer cleanup2()
	if out2[0] != nil {
		t.Errorf("nil *time.Time should pass through as nil, got %T %v", out2[0], out2[0])
	}
}

// TestDecodeVariantScalar_Currency is the Excel-free regression for the
// VT_CY gap in go-ole's Value(): a currency VARIANT (an int64 scaled by 1e-4)
// must decode to the plain float64 amount, not the bare nil go-ole returns.
func TestDecodeVariantScalar_Currency(t *testing.T) {
	// 12.34 currency == 123400 in CY units (1e-4 scale).
	v := ole.NewVariant(ole.VT_CY, 123400)
	if raw := v.Value(); raw != nil {
		t.Fatalf("precondition: go-ole VT_CY Value() now returns %v (%T); test premise stale", raw, raw)
	}
	got := decodeVariantScalar(&v)
	if got != 12.34 {
		t.Errorf("VT_CY decode: got %v (%T), want 12.34", got, got)
	}
}

// TestDecodeVariantScalar_Decimal covers the VT_DECIMAL gap: the DECIMAL
// overlays the VARIANT, and VarR8FromDec must recover the float64 value.
func TestDecodeVariantScalar_Decimal(t *testing.T) {
	var v ole.VARIANT
	ole.VariantInit(&v)
	v.VT = ole.VT_DECIMAL
	// Lay 12.34 into the overlaid DECIMAL: coefficient 1234, scale 2.
	dec := (*oleDecimal)(unsafe.Pointer(&v))
	dec.scale = 2
	dec.sign = 0
	dec.hi32 = 0
	dec.lo64 = 1234
	got := decodeVariantScalar(&v)
	gf, ok := got.(float64)
	if !ok || math.Abs(gf-12.34) > 1e-9 {
		t.Errorf("VT_DECIMAL decode: got %v (%T), want 12.34", got, got)
	}
}

// TestDecodeVariantScalar_Error covers the VT_ERROR gap: a worksheet error
// cell must become a typed CellError (so it is distinguishable from a blank
// cell), while the DISP_E_PARAMNOTFOUND marker stays nil.
func TestDecodeVariantScalar_Error(t *testing.T) {
	// #DIV/0! == cvErr 2007, SCODE 0x800A0000 | 2007.
	div0 := ole.NewVariant(ole.VT_ERROR, 0x800A07D7)
	got := decodeVariantScalar(&div0)
	ce, ok := got.(CellError)
	if !ok {
		t.Fatalf("VT_ERROR decode: got %T, want CellError", got)
	}
	if ce.String() != "#DIV/0!" {
		t.Errorf("CellError.String(): got %q, want #DIV/0!", ce.String())
	}
	if ce.SCode != 0x800A07D7 {
		t.Errorf("CellError.SCode: got 0x%08X, want 0x800A07D7", ce.SCode)
	}

	// The omitted-optional-parameter marker is not a worksheet error.
	missing := ole.NewVariant(ole.VT_ERROR, dispEParamNotFound)
	if got := decodeVariantScalar(&missing); got != nil {
		t.Errorf("DISP_E_PARAMNOTFOUND should decode to nil, got %v (%T)", got, got)
	}
}

// TestCellError_String spot-checks the CVErr → Excel text mapping.
func TestCellError_String(t *testing.T) {
	cases := map[uint32]string{
		0x800A07D0: "#NULL!",  // 2000
		0x800A07D7: "#DIV/0!", // 2007
		0x800A07DF: "#VALUE!", // 2015
		0x800A07E7: "#REF!",   // 2023
		0x800A07ED: "#NAME?",  // 2029
		0x800A07F4: "#NUM!",   // 2036
		0x800A07FA: "#N/A",    // 2042
	}
	for scode, want := range cases {
		if got := (CellError{SCode: scode}).String(); got != want {
			t.Errorf("CellError{0x%08X}.String() = %q, want %q", scode, got, want)
		}
	}
	// Unknown code falls back to the hex form.
	if got := (CellError{SCode: 0xDEADBEEF}).String(); got != "#ERR(0xDEADBEEF)" {
		t.Errorf("unknown CellError.String() = %q, want #ERR(0xDEADBEEF)", got)
	}
}

// TestSafeArrayDataLayout pins the assumption the bulk encode/decode paths are
// built on: a SAFEARRAY's element buffer is **column-major** — the dimension-1
// index (rows, for the `[row][col]` arrays Excel returns) varies fastest, so
// cell (r, c) of a rows x cols array sits at linear offset `c*rows + r`.
//
// Getting this backwards transposes every multi-cell Range read and write, and
// the transposition is invisible on square grids — hence an asymmetric 2x3.
// The elemsize check guards the second assumption: that go-ole's VARIANT struct
// is ABI-identical to the native one, so the COM buffer can be aliased as a
// []ole.VARIANT.
func TestSafeArrayDataLayout(t *testing.T) {
	const rows, cols = 2, 3
	sa, err := createVariantSafeArray([]safeArrayBound{
		{cElements: rows, lLbound: 0},
		{cElements: cols, lLbound: 0},
	})
	if err != nil {
		t.Fatal(err)
	}
	defer procSafeArrayDestroy.Call(sa)

	// Write through the OS's own index arithmetic.
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			if err := putElement(sa, []int32{int32(r), int32(c)}, float64(10*r+c)); err != nil {
				t.Fatalf("putElement(%d,%d): %v", r, c, err)
			}
		}
	}

	if size, _, _ := procSafeArrayGetElemsize.Call(sa); size != uintptr(variantSize) {
		t.Fatalf("SafeArrayGetElemsize = %d, but sizeof(ole.VARIANT) = %d — the COM buffer cannot be aliased as []ole.VARIANT", size, variantSize)
	}

	cells, unlock, err := accessVariantData(sa, rows*cols)
	if err != nil {
		t.Fatalf("accessVariantData: %v", err)
	}
	defer unlock()

	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			got := cells[c*rows+r].Value()
			if want := float64(10*r + c); got != want {
				t.Errorf("linear[%d] (c*rows+r for r=%d,c=%d) = %v, want %v — SAFEARRAY layout is not column-major",
					c*rows+r, r, c, got, want)
			}
		}
	}
}

// TestAccessVariantData_RejectsTypedArray covers the element-size guard: a
// typed (non-VT_VARIANT) SAFEARRAY has a different element width, so aliasing
// its buffer as []ole.VARIANT would read garbage. accessVariantData must refuse
// so the caller falls back to the per-element API instead.
func TestAccessVariantData_RejectsTypedArray(t *testing.T) {
	bounds := []safeArrayBound{{cElements: 4, lLbound: 0}}
	sa, _, _ := procSafeArrayCreate.Call(
		uintptr(ole.VT_R8), // 8-byte elements, not sizeof(VARIANT)
		uintptr(uint32(1)),
		uintptr(unsafe.Pointer(&bounds[0])),
	)
	if sa == 0 {
		t.Fatal("SafeArrayCreate(VT_R8) failed")
	}
	defer procSafeArrayDestroy.Call(sa)

	cells, unlock, err := accessVariantData(sa, 4)
	if err == nil {
		unlock()
		t.Fatalf("expected an element-size error for a VT_R8 array, got %d cells", len(cells))
	}
}

// bulkGrid is the shared fixture for the bulk-vs-per-element cross-checks: an
// asymmetric grid mixing every cell type the encoder supports.
func bulkGrid(rows, cols int) [][]interface{} {
	g := make([][]interface{}, rows)
	for r := range g {
		row := make([]interface{}, cols)
		for c := range row {
			switch (r + c) % 4 {
			case 0:
				row[c] = float64(r)*1000 + float64(c)
			case 1:
				row[c] = fmt.Sprintf("s(%d,%d)", r, c)
			case 2:
				row[c] = (r+c)%8 == 2
			case 3:
				row[c] = nil
			}
		}
		g[r] = row
	}
	return g
}

// mustCreate allocates a rows x cols VT_VARIANT SAFEARRAY for a test.
func mustCreate(t *testing.T, rows, cols int) uintptr {
	t.Helper()
	sa, err := createVariantSafeArray([]safeArrayBound{
		{cElements: uint32(rows)},
		{cElements: uint32(cols)},
	})
	if err != nil {
		t.Fatalf("createVariantSafeArray(%dx%d): %v", rows, cols, err)
	}
	return sa
}

// readPerElement reads a grid back through oleaut32's SafeArrayGetElement.
func readPerElement(t *testing.T, sa uintptr, rows, cols int) [][]interface{} {
	t.Helper()
	out := make([][]interface{}, rows)
	for r := 0; r < rows; r++ {
		out[r] = make([]interface{}, cols)
		for c := 0; c < cols; c++ {
			v, err := getElement(sa, []int32{int32(r), int32(c)})
			if err != nil {
				t.Fatalf("getElement(%d,%d): %v", r, c, err)
			}
			out[r][c] = v
		}
	}
	return out
}

// viaPerElement round-trips src entirely through the per-element API, giving
// the reference grid the bulk paths must reproduce.
func viaPerElement(t *testing.T, src [][]interface{}, rows, cols int) [][]interface{} {
	t.Helper()
	sa := mustCreate(t, rows, cols)
	defer procSafeArrayDestroy.Call(sa)
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			if err := putElement(sa, []int32{int32(r), int32(c)}, src[r][c]); err != nil {
				t.Fatalf("putElement(%d,%d): %v", r, c, err)
			}
		}
	}
	return readPerElement(t, sa, rows, cols)
}

// TestBulkMatchesPerElement is the correctness anchor for the
// SafeArrayAccessData fast paths: for every shape, the bulk encode+decode must
// produce byte-for-byte the same grid as the SafeArrayGetElement /
// SafeArrayPutElement paths, which use oleaut32's own index arithmetic rather
// than sugar's `c*rows + r`.
//
// Both mixed directions are checked (bulk-write/slow-read and
// slow-write/bulk-read), so a transposition in either path is caught even
// though a matched pair of transposes would cancel out.
func TestBulkMatchesPerElement(t *testing.T) {
	shapes := [][2]int{{1, 1}, {1, 7}, {7, 1}, {2, 3}, {3, 2}, {5, 9}, {64, 33}}
	for _, s := range shapes {
		rows, cols := s[0], s[1]
		name := fmt.Sprintf("%dx%d", rows, cols)
		src := bulkGrid(rows, cols)

		// Reference: write and read entirely through the per-element API.
		ref := viaPerElement(t, src, rows, cols)

		// 1. bulk write -> per-element read.
		bulkSA := mustCreate(t, rows, cols)
		if err := fill2D(bulkSA, src, rows, cols); err != nil {
			t.Fatalf("%s: fill2D: %v", name, err)
		}
		gotSlowRead := readPerElement(t, bulkSA, rows, cols)
		procSafeArrayDestroy.Call(bulkSA)
		if !reflect.DeepEqual(gotSlowRead, ref) {
			t.Errorf("%s: bulk write + per-element read = %v, want %v", name, gotSlowRead, ref)
		}

		// 2. per-element write -> bulk read.
		slowSA := mustCreate(t, rows, cols)
		for r := 0; r < rows; r++ {
			for c := 0; c < cols; c++ {
				if err := putElement(slowSA, []int32{int32(r), int32(c)}, src[r][c]); err != nil {
					t.Fatalf("%s: putElement: %v", name, err)
				}
			}
		}
		gotBulkRead, err := decode2D(slowSA)
		if err != nil {
			t.Fatalf("%s: decode2D: %v", name, err)
		}
		if !reflect.DeepEqual(gotBulkRead, ref) {
			t.Errorf("%s: per-element write + bulk read = %v, want %v", name, gotBulkRead, ref)
		}
		procSafeArrayDestroy.Call(slowSA)

		// 3. the production path end to end.
		v, err := encodeVariantArray(src)
		if err != nil {
			t.Fatalf("%s: encodeVariantArray: %v", name, err)
		}
		gotFull, err := decodeVariantArray(v)
		if err != nil {
			t.Fatalf("%s: decodeVariantArray: %v", name, err)
		}
		v.Clear()
		if !reflect.DeepEqual(gotFull, ref) {
			t.Errorf("%s: bulk round trip = %v, want %v", name, gotFull, ref)
		}
	}
}

// TestBulkMatchesPerElement1D is the 1-D twin of TestBulkMatchesPerElement.
func TestBulkMatchesPerElement1D(t *testing.T) {
	for _, n := range []int{1, 2, 17, 256} {
		src := bulkGrid(1, n)[0]

		slowSA, err := createVariantSafeArray([]safeArrayBound{{cElements: uint32(n)}})
		if err != nil {
			t.Fatal(err)
		}
		for i, val := range src {
			if err := putElement(slowSA, []int32{int32(i)}, val); err != nil {
				t.Fatalf("n=%d: putElement: %v", n, err)
			}
		}
		ref := make([]interface{}, n)
		for i := range ref {
			v, err := getElement(slowSA, []int32{int32(i)})
			if err != nil {
				t.Fatalf("n=%d: getElement: %v", n, err)
			}
			ref[i] = v
		}
		gotBulkRead, err := decode1D(slowSA)
		if err != nil {
			t.Fatalf("n=%d: decode1D: %v", n, err)
		}
		procSafeArrayDestroy.Call(slowSA)
		if !reflect.DeepEqual(gotBulkRead, ref) {
			t.Errorf("n=%d: bulk read = %v, want %v", n, gotBulkRead, ref)
		}

		bulkSA, err := createVariantSafeArray([]safeArrayBound{{cElements: uint32(n)}})
		if err != nil {
			t.Fatal(err)
		}
		if err := fill1D(bulkSA, src); err != nil {
			t.Fatalf("n=%d: fill1D: %v", n, err)
		}
		for i := range ref {
			got, err := getElement(bulkSA, []int32{int32(i)})
			if err != nil {
				t.Fatalf("n=%d: getElement: %v", n, err)
			}
			if !reflect.DeepEqual(got, ref[i]) {
				t.Errorf("n=%d: bulk write cell %d = %v, want %v", n, i, got, ref[i])
			}
		}
		procSafeArrayDestroy.Call(bulkSA)
	}
}

// TestEncodeDecode_LargeGrid round-trips a grid big enough to cross every
// interesting boundary of the bulk paths (multi-page data buffer, the shared
// flat backing allocation decode2D slices per row) and asserts exact value
// identity, not just shape.
func TestEncodeDecode_LargeGrid(t *testing.T) {
	const rows, cols = 257, 129 // asymmetric and coprime-ish, so a transpose shows
	src := bulkGrid(rows, cols)

	v, err := encodeVariantArray(src)
	if err != nil {
		t.Fatalf("encodeVariantArray: %v", err)
	}
	defer v.Clear()

	out, err := decodeVariantArray(v)
	if err != nil {
		t.Fatalf("decodeVariantArray: %v", err)
	}
	grid, ok := out.([][]interface{})
	if !ok {
		t.Fatalf("decoded to %T, want [][]interface{}", out)
	}
	if len(grid) != rows {
		t.Fatalf("got %d rows, want %d", len(grid), rows)
	}
	for r := range grid {
		if len(grid[r]) != cols {
			t.Fatalf("row %d has %d cols, want %d", r, len(grid[r]), cols)
		}
		for c := range grid[r] {
			if !reflect.DeepEqual(grid[r][c], src[r][c]) {
				t.Fatalf("cell (%d,%d) = %v (%T), want %v (%T)", r, c, grid[r][c], grid[r][c], src[r][c], src[r][c])
			}
		}
	}
}

// TestEncodeDecode_Empty pins the zero-element edge of the bulk paths:
// SafeArrayAccessData may hand back a nil data pointer for an empty array, and
// accessVariantData must turn that into the (no-op) fallback rather than
// aliasing nil.
func TestEncodeDecode_Empty(t *testing.T) {
	for _, in := range []interface{}{[]interface{}{}, [][]interface{}{}} {
		v, err := encodeVariantArray(in)
		if err != nil {
			t.Fatalf("encodeVariantArray(%T{}): %v", in, err)
		}
		out, err := decodeVariantArray(v)
		if err != nil {
			t.Fatalf("decodeVariantArray(%T{}): %v", in, err)
		}
		v.Clear()
		if rv := reflect.ValueOf(out); !rv.IsValid() || rv.Kind() != reflect.Slice || rv.Len() != 0 {
			t.Errorf("%T{} round-tripped to %#v, want an empty slice", in, out)
		}
	}
}

// TestEncode_UnsupportedCellMidGridDestroys covers the bulk encode failure
// path: an unsupported cell in the *middle* of a grid (after string cells have
// already had their BSTRs handed to the array) must unlock the array before
// SafeArrayDestroy runs, or the destroy fails with DISP_E_ARRAYISLOCKED and the
// whole array leaks. A leaked lock is not directly observable, so this asserts
// the reachable symptom: a clean error and no panic/crash, repeated enough to
// make a leak visible under -race / heap growth.
func TestEncode_UnsupportedCellMidGridDestroys(t *testing.T) {
	for i := 0; i < 100; i++ {
		src := bulkGrid(8, 8)
		src[5][5] = struct{ nope int }{}
		if _, err := encodeVariantArray(src); err == nil {
			t.Fatal("expected an error for an unsupported cell type")
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
