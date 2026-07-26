//go:build windows

package sugar

import (
	"fmt"
	"math"
	"reflect"
	"syscall"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
)

// 2-D SAFEARRAY decoding for VARIANT-of-array values.
//
// go-ole v1.3.0's `SafeArrayConversion.ToValueArray` is hard-wired to 1-D and
// silently corrupts multi-dim arrays. Excel's `Range.Value` always returns a
// 2-D `SAFEARRAY` of `VARIANT`, so we drive `oleaut32.dll` directly.
//
// Bulk transfer is the default: `SafeArrayAccessData` locks the array once and
// the whole grid is walked as a `[]ole.VARIANT` view over the element buffer.
// The per-element `SafeArrayGetElement` / `SafeArrayPutElement` entry points
// are kept as a fallback for the (unreachable in practice) case where the
// array's element width is not `sizeof(VARIANT)`, and as the independent
// reference the bulk paths are tested against. See the "Bulk element access"
// block below for the layout and locking rules.

var oleaut32 = syscall.NewLazyDLL("oleaut32.dll")

var (
	procSafeArrayGetDim       = oleaut32.NewProc("SafeArrayGetDim")
	procSafeArrayGetLBound    = oleaut32.NewProc("SafeArrayGetLBound")
	procSafeArrayGetUBound    = oleaut32.NewProc("SafeArrayGetUBound")
	procSafeArrayGetElement   = oleaut32.NewProc("SafeArrayGetElement")
	procSafeArrayGetVartype   = oleaut32.NewProc("SafeArrayGetVartype")
	procSafeArrayGetElemsize  = oleaut32.NewProc("SafeArrayGetElemsize")
	procSafeArrayCreate       = oleaut32.NewProc("SafeArrayCreate")
	procSafeArrayPutElement   = oleaut32.NewProc("SafeArrayPutElement")
	procSafeArrayAccessData   = oleaut32.NewProc("SafeArrayAccessData")
	procSafeArrayUnaccessData = oleaut32.NewProc("SafeArrayUnaccessData")
	procSafeArrayDestroy      = oleaut32.NewProc("SafeArrayDestroy")
	procVarR8FromDec          = oleaut32.NewProc("VarR8FromDec")
)

// variantSize is the byte width of a VARIANT as Go sees it (24 on amd64, 16 on
// 386). accessVariantData cross-checks it against the SAFEARRAY's own element
// size before aliasing the COM buffer as a []ole.VARIANT.
const variantSize = unsafe.Sizeof(ole.VARIANT{})

// Bulk element access.
//
// The per-element `SafeArrayGetElement` / `SafeArrayPutElement` APIs cost a
// `syscall.LazyProc.Call` plus a VARIANT deep copy *per cell*, which dominates
// the cost of reading or writing a large `Range.Value` block (measured: ~350 ns
// per cell, i.e. ~90 ms for a 500x500 grid — about half of the end-to-end
// `Range.Value()` time even across the process boundary). `SafeArrayAccessData`
// locks the array once and hands back the raw element buffer, so the whole grid
// is walked with plain memory access.
//
// # Data layout
//
// A SAFEARRAY's element buffer is column-major: the **first** index (the one
// `SafeArrayGetLBound(psa, 1)` describes) varies fastest. For the 2-D
// `[row][col]` arrays Excel uses — dimension 1 = rows, dimension 2 = columns —
// the linear offset of cell (r, c) is therefore `c*rows + r`, **not**
// `r*cols + c`. This is verified two ways: `TestSafeArrayDataLayout` asserts the
// ordering directly against oleaut32, and `TestBulkMatchesPerElement`
// cross-checks the bulk paths against the `SafeArrayGetElement`/`PutElement`
// results (the OS's own index arithmetic) on asymmetric grids.
//
// # Locking contract
//
// Every successful `accessVariantData` returns an `unlock` closure that must
// run before the array is destroyed — `SafeArrayDestroy` fails with
// `DISP_E_ARRAYISLOCKED` on a locked array. Callers `defer unlock()` inside a
// helper whose scope closes before any destroy in the caller.

// accessVariantData locks a VT_VARIANT SAFEARRAY and returns its `n` elements
// as a Go slice aliasing the COM-owned buffer, together with the matching
// unlock function.
//
// The returned slice is a *view*: its VARIANTs are owned by the SAFEARRAY. Do
// not Clear them on the read path (that would free the array's own BSTRs), and
// on the write path storing a VARIANT transfers ownership of any BSTR it holds
// to the array (destroyed later by SafeArrayDestroy / VariantClear).
//
// An element-size mismatch is reported as an error so the caller can fall back
// to the per-element API rather than reinterpret a foreign layout.
func accessVariantData(sa uintptr, n int) ([]ole.VARIANT, func(), error) {
	size, _, _ := procSafeArrayGetElemsize.Call(sa)
	if size != uintptr(variantSize) {
		return nil, nil, fmt.Errorf("SafeArrayGetElemsize returned %d, want %d (VARIANT)", size, variantSize)
	}
	// Let oleaut32 write the element pointer straight into an unsafe.Pointer
	// variable; a uintptr round-trip would be both a `go vet` violation and,
	// in principle, a GC-invisible pointer.
	var data unsafe.Pointer
	hr, _, _ := procSafeArrayAccessData.Call(sa, uintptr(unsafe.Pointer(&data)))
	if hr != 0 {
		return nil, nil, fmt.Errorf("SafeArrayAccessData failed: 0x%x", hr)
	}
	if data == nil {
		procSafeArrayUnaccessData.Call(sa)
		return nil, nil, fmt.Errorf("SafeArrayAccessData returned a nil data pointer")
	}
	cells := unsafe.Slice((*ole.VARIANT)(data), n)
	return cells, func() { procSafeArrayUnaccessData.Call(sa) }, nil
}

// decodeVariantArray turns a VARIANT carrying `VT_ARRAY | VT_VARIANT` — the
// shape Excel returns from `Range.Value` — into a Go value:
//
//   - A 2-D array becomes `[][]interface{}` indexed as `[row][col]`.
//   - A 1-D array becomes `[]interface{}`.
//   - Other shapes return an error rather than silent corruption.
//
// Inner VARIANT cells are converted via `(*VARIANT).Value()`. Dates land as
// `time.Time` (go-ole already does this) — matching xlwings.
func decodeVariantArray(v *ole.VARIANT) (interface{}, error) {
	if v == nil {
		return nil, nil
	}
	if v.VT&ole.VT_ARRAY == 0 {
		return nil, fmt.Errorf("decodeVariantArray: VARIANT type 0x%x is not an array", v.VT)
	}
	// A VT_BYREF array stores a SAFEARRAY** (a pointer to the array pointer) in
	// Val, not a SAFEARRAY* — dereferencing Val as the array directly would
	// read the wrong pointer and misdecode or access-violate. Some COM servers
	// hand back byref arrays; reject them explicitly rather than corrupting.
	if v.VT&ole.VT_BYREF != 0 {
		return nil, fmt.Errorf("decodeVariantArray: VT_BYREF arrays (VT 0x%x) are not supported", v.VT)
	}
	// The VARIANT's `Val` field is `int64` but actually stores the SAFEARRAY
	// pointer. Reinterpret the bytes via &Val — not a uintptr round-trip —
	// to keep go vet happy with `unsafe.Pointer` rules. The handle is then
	// carried as a uintptr (the same convention the encode path uses): the
	// memory is COM-allocated, so Go's GC never moves or frees it.
	sa := uintptr(*(*unsafe.Pointer)(unsafe.Pointer(&v.Val)))
	if sa == 0 {
		return nil, nil
	}
	// getElement hands SafeArrayGetElement a VARIANT output buffer, which is
	// only correct for a VT_VARIANT element type. A typed SAFEARRAY
	// (VT_ARRAY|VT_BSTR / R8 / …) — common from non-Excel COM servers, since
	// the core is a general COM layer — would have its element bytes copied
	// into the head of the VARIANT, poisoning its VT tag: the read would then
	// silently return nil/wrong scalars (and leak a BSTR per string cell).
	// Query the array's real element type directly; the SAFEARRAY is more
	// authoritative than the enclosing VARIANT's VT, which some servers fill
	// loosely. Reject anything but VT_VARIANT rather than misdecode.
	vt, err := safeArrayVartype(sa)
	if err != nil {
		return nil, err
	}
	if vt != ole.VT_VARIANT {
		return nil, fmt.Errorf("decodeVariantArray: SAFEARRAY element type VT 0x%x is not VT_VARIANT; typed arrays are not supported", vt)
	}

	dims, _, _ := procSafeArrayGetDim.Call(sa)
	switch uint32(dims) {
	case 1:
		return decode1D(sa)
	case 2:
		return decode2D(sa)
	default:
		return nil, fmt.Errorf("decodeVariantArray: unsupported dimension count %d", dims)
	}
}

// dimLen returns the element count of one SAFEARRAY dimension. The subtraction
// is done in int64 because SafeArrayGetLBound/GetUBound are signed 32-bit and a
// hostile or corrupt descriptor (lo = MinInt32, hi = MaxInt32) would wrap an
// int32 count to 0 or a negative. The count feeds unsafe.Slice on the raw
// element buffer, so it must never be a wrapped value.
func dimLen(lo, hi int32) (int, error) {
	n := int64(hi) - int64(lo) + 1
	if n <= 0 {
		return 0, nil
	}
	if n > int64(maxInt) {
		return 0, fmt.Errorf("SAFEARRAY dimension of %d elements exceeds the addressable range", n)
	}
	return int(n), nil
}

// maxInt is the largest value of Go's int on this build (32- or 64-bit).
const maxInt = int(^uint(0) >> 1)

func decode1D(sa uintptr) ([]interface{}, error) {
	lo, hi, err := bounds(sa, 1)
	if err != nil {
		return nil, err
	}
	n, err := dimLen(lo, hi)
	if err != nil {
		return nil, err
	}
	if n <= 0 {
		return []interface{}{}, nil
	}
	out := make([]interface{}, n)

	cells, unlock, err := accessVariantData(sa, n)
	if err != nil {
		// Fall back to the per-element API — correct, just slower.
		for i := int32(0); i < int32(n); i++ {
			val, err := getElement(sa, []int32{lo + i})
			if err != nil {
				return nil, err
			}
			out[i] = val
		}
		return out, nil
	}
	defer unlock()

	for i := range out {
		out[i] = decodeArrayCell(&cells[i])
	}
	return out, nil
}

func decode2D(sa uintptr) ([][]interface{}, error) {
	rLo, rHi, err := bounds(sa, 1)
	if err != nil {
		return nil, err
	}
	cLo, cHi, err := bounds(sa, 2)
	if err != nil {
		return nil, err
	}
	rows, err := dimLen(rLo, rHi)
	if err != nil {
		return nil, err
	}
	cols, err := dimLen(cLo, cHi)
	if err != nil {
		return nil, err
	}
	if rows <= 0 || cols <= 0 {
		return [][]interface{}{}, nil
	}
	// rows*cols is the length handed to unsafe.Slice; a wrapped product would
	// alias memory past the element buffer. (Excel's own limits make this
	// unreachable, but the decoder accepts arrays from any COM server.)
	if cols > maxInt/rows {
		return nil, fmt.Errorf("decodeVariantArray: %dx%d SAFEARRAY exceeds the addressable range", rows, cols)
	}
	out := make([][]interface{}, rows)
	// One backing allocation for every cell, re-sliced per row: a 500x500 read
	// drops 500 separate row allocations to one.
	flat := make([]interface{}, rows*cols)
	for r := range out {
		out[r] = flat[r*cols : (r+1)*cols : (r+1)*cols]
	}

	cells, unlock, err := accessVariantData(sa, rows*cols)
	if err != nil {
		// Fall back to the per-element API — correct, just slower.
		for r := 0; r < rows; r++ {
			for c := 0; c < cols; c++ {
				val, err := getElement(sa, []int32{rLo + int32(r), cLo + int32(c)})
				if err != nil {
					return nil, err
				}
				out[r][c] = val
			}
		}
		return out, nil
	}
	defer unlock()

	// Column-major: dimension 1 (the row index) varies fastest, so walking a
	// column at a time is also the cache-friendly order.
	for c := 0; c < cols; c++ {
		col := cells[c*rows : (c+1)*rows]
		for r := 0; r < rows; r++ {
			out[r][c] = decodeArrayCell(&col[r])
		}
	}
	return out, nil
}

// decodeArrayCell converts one SAFEARRAY cell — a VARIANT still *owned by the
// array* — to a Go value. It is the bulk-access twin of getElement and must
// keep the same conversion semantics, minus the Clear (the array owns the
// element; clearing it here would free the array's own BSTRs).
//
// Object cells degrade to nil for the same reason getElement degrades them:
// (*VARIANT).Value() hands back the raw interface pointer with no AddRef, and a
// value grid cannot own a live COM reference. Returning it would outlive the
// array (the caller's VARIANT is VariantClear'd by the arena) and dangle.
func decodeArrayCell(v *ole.VARIANT) interface{} {
	if v.VT == ole.VT_DISPATCH || v.VT == ole.VT_UNKNOWN {
		return nil
	}
	return decodeVariantScalar(v)
}

// safeArrayVartype returns the element VARTYPE of a SAFEARRAY via
// oleaut32!SafeArrayGetVartype. Used to reject typed arrays before the
// VT_VARIANT-assuming element decode runs.
func safeArrayVartype(sa uintptr) (ole.VT, error) {
	var vt ole.VT
	hr, _, _ := procSafeArrayGetVartype.Call(sa, uintptr(unsafe.Pointer(&vt)))
	if hr != 0 {
		return 0, fmt.Errorf("SafeArrayGetVartype failed: 0x%x", hr)
	}
	return vt, nil
}

func bounds(sa uintptr, dim uint32) (int32, int32, error) {
	var lo, hi int32
	hr, _, _ := procSafeArrayGetLBound.Call(sa, uintptr(dim), uintptr(unsafe.Pointer(&lo)))
	if hr != 0 {
		return 0, 0, fmt.Errorf("SafeArrayGetLBound(dim=%d) failed: 0x%x", dim, hr)
	}
	hr, _, _ = procSafeArrayGetUBound.Call(sa, uintptr(dim), uintptr(unsafe.Pointer(&hi)))
	if hr != 0 {
		return 0, 0, fmt.Errorf("SafeArrayGetUBound(dim=%d) failed: 0x%x", dim, hr)
	}
	return lo, hi, nil
}

// SAFEARRAY encoding — the write-direction mirror of decodeVariantArray.
//
// go-ole's Invoke marshals only flat scalar types ([]byte / []string aside)
// and panics on anything else, so writing a block to `Range.Value` needs a
// hand-built `VT_ARRAY | VT_VARIANT` SAFEARRAY.

type safeArrayBound struct {
	cElements uint32
	lLbound   int32
}

// encodeVariantArray builds a VARIANT carrying a VT_ARRAY|VT_VARIANT
// SAFEARRAY from a Go slice:
//
//   - `[]interface{}` becomes a 1-D array (Excel reads it as a row).
//   - `[][]interface{}` becomes a 2-D array indexed `[row][col]`. Rows must
//     be equal length.
//   - Any other slice ([]float64, [][]float64, []string rows, …) is widened
//     to the shapes above via reflection: an element kind of Slice means
//     2-D, anything else 1-D.
//
// The returned VARIANT owns the SAFEARRAY; the caller must Clear() it after
// the COM call (VariantClear destroys the array).
func encodeVariantArray(value interface{}) (*ole.VARIANT, error) {
	switch v := value.(type) {
	case []interface{}:
		return encode1D(v)
	case [][]interface{}:
		return encode2D(v)
	}
	rv := reflect.ValueOf(value)
	if rv.Kind() != reflect.Slice {
		return nil, fmt.Errorf("encodeVariantArray: unsupported type %T", value)
	}
	if rv.Type().Elem().Kind() == reflect.Slice {
		grid := make([][]interface{}, rv.Len())
		for r := 0; r < rv.Len(); r++ {
			row := rv.Index(r)
			cells := make([]interface{}, row.Len())
			for c := 0; c < row.Len(); c++ {
				cells[c] = row.Index(c).Interface()
			}
			grid[r] = cells
		}
		return encode2D(grid)
	}
	vec := make([]interface{}, rv.Len())
	for i := 0; i < rv.Len(); i++ {
		vec[i] = rv.Index(i).Interface()
	}
	return encode1D(vec)
}

func encode1D(src []interface{}) (*ole.VARIANT, error) {
	bounds := []safeArrayBound{{cElements: uint32(len(src)), lLbound: 0}}
	sa, err := createVariantSafeArray(bounds)
	if err != nil {
		return nil, err
	}
	if err := fill1D(sa, src); err != nil {
		// fill1D has already released its data lock, so the destroy (which
		// fails with DISP_E_ARRAYISLOCKED on a locked array) can proceed. It
		// VariantClears every cell, freeing whatever was written before the
		// failure.
		procSafeArrayDestroy.Call(sa)
		return nil, err
	}
	return wrapSafeArray(sa), nil
}

// fill1D writes src into a freshly created (zero-initialized) 1-D VT_VARIANT
// SAFEARRAY. On the bulk path each VARIANT is stored by value, which transfers
// BSTR ownership to the array — unlike putElement, which deep-copies and then
// clears the temporary.
func fill1D(sa uintptr, src []interface{}) error {
	cells, unlock, err := accessVariantData(sa, len(src))
	if err != nil {
		// Fall back to the per-element API — correct, just slower.
		for i, val := range src {
			if err := putElement(sa, []int32{int32(i)}, val); err != nil {
				return err
			}
		}
		return nil
	}
	defer unlock()

	for i, val := range src {
		if err := scalarToVariant(val, &cells[i]); err != nil {
			return err
		}
	}
	return nil
}

func encode2D(src [][]interface{}) (*ole.VARIANT, error) {
	rows := len(src)
	cols := 0
	if rows > 0 {
		cols = len(src[0])
	}
	for r, row := range src {
		if len(row) != cols {
			return nil, fmt.Errorf("encodeVariantArray: ragged rows — row 0 has %d columns, row %d has %d", cols, r, len(row))
		}
	}
	bounds := []safeArrayBound{
		{cElements: uint32(rows), lLbound: 0},
		{cElements: uint32(cols), lLbound: 0},
	}
	sa, err := createVariantSafeArray(bounds)
	if err != nil {
		return nil, err
	}
	if err := fill2D(sa, src, rows, cols); err != nil {
		// fill2D has already released its data lock (see fill1D).
		procSafeArrayDestroy.Call(sa)
		return nil, err
	}
	return wrapSafeArray(sa), nil
}

// fill2D writes src into a freshly created (zero-initialized) 2-D VT_VARIANT
// SAFEARRAY. See fill1D for the BSTR-ownership note and accessVariantData for
// the column-major offset rule.
func fill2D(sa uintptr, src [][]interface{}, rows, cols int) error {
	cells, unlock, err := accessVariantData(sa, rows*cols)
	if err != nil {
		// Fall back to the per-element API — correct, just slower.
		for r, row := range src {
			for c, val := range row {
				if err := putElement(sa, []int32{int32(r), int32(c)}, val); err != nil {
					return err
				}
			}
		}
		return nil
	}
	defer unlock()

	for r, row := range src {
		for c, val := range row {
			if err := scalarToVariant(val, &cells[c*rows+r]); err != nil {
				return err
			}
		}
	}
	return nil
}

// createVariantSafeArray allocates a VT_VARIANT SAFEARRAY. The handle stays
// a uintptr through the encode path: the memory is COM-allocated, so Go's GC
// never moves or frees it, and uintptr avoids vet's unsafe.Pointer rules.
func createVariantSafeArray(bounds []safeArrayBound) (uintptr, error) {
	sa, _, _ := procSafeArrayCreate.Call(
		uintptr(ole.VT_VARIANT),
		uintptr(uint32(len(bounds))),
		uintptr(unsafe.Pointer(&bounds[0])),
	)
	if sa == 0 {
		return 0, fmt.Errorf("SafeArrayCreate failed (dims=%d)", len(bounds))
	}
	return sa, nil
}

// wrapSafeArray packages a SAFEARRAY handle into a VARIANT that owns it.
func wrapSafeArray(sa uintptr) *ole.VARIANT {
	v := ole.NewVariant(ole.VT_ARRAY|ole.VT_VARIANT, int64(sa))
	return &v
}

// putElement writes one Go scalar into a VT_ARRAY|VT_VARIANT SAFEARRAY cell.
// SafeArrayPutElement deep-copies the VARIANT (BSTRs included), so the
// temporary cell is cleared afterwards.
func putElement(sa uintptr, indices []int32, val interface{}) error {
	var cell ole.VARIANT
	if err := scalarToVariant(val, &cell); err != nil {
		return err
	}
	defer cell.Clear()
	hr, _, _ := procSafeArrayPutElement.Call(
		sa,
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(&cell)),
	)
	if hr != 0 {
		return fmt.Errorf("SafeArrayPutElement failed: 0x%x", hr)
	}
	return nil
}

// scalarToVariant converts a Go scalar to a VARIANT cell value. Supported:
// nil (VT_EMPTY), bool, string, all int/uint widths, float32/64, time.Time
// (VT_DATE). Excel stores all numbers as doubles, so numeric width loss is
// not a concern on the COM side.
//
// `out` must be an empty/uninitialized VARIANT the caller owns — this function
// overwrites it without clearing, so passing a VARIANT that already holds a
// BSTR or an interface pointer leaks it. Zeroing is done in Go rather than via
// ole.VariantInit (an oleaut32 call) because the bulk encode path runs this
// once per cell and the syscall dominated a large block write; a zeroed VARIANT
// *is* VT_EMPTY, which is exactly what VariantInit produces.
func scalarToVariant(val interface{}, out *ole.VARIANT) error {
	*out = ole.VARIANT{}
	switch x := val.(type) {
	case nil:
		// VT_EMPTY — an empty cell.
	case bool:
		if x {
			*out = ole.NewVariant(ole.VT_BOOL, 0xffff)
		} else {
			*out = ole.NewVariant(ole.VT_BOOL, 0)
		}
	case string:
		*out = ole.NewVariant(ole.VT_BSTR, int64(uintptr(unsafe.Pointer(ole.SysAllocStringLen(x)))))
	case int:
		setNumericVariant(out, float64(x))
	case int8:
		setNumericVariant(out, float64(x))
	case int16:
		setNumericVariant(out, float64(x))
	case int32:
		setNumericVariant(out, float64(x))
	case int64:
		setNumericVariant(out, float64(x))
	case uint:
		setNumericVariant(out, float64(x))
	case uint8:
		setNumericVariant(out, float64(x))
	case uint16:
		setNumericVariant(out, float64(x))
	case uint32:
		setNumericVariant(out, float64(x))
	case uint64:
		setNumericVariant(out, float64(x))
	case float32:
		setNumericVariant(out, float64(x))
	case float64:
		setNumericVariant(out, x)
	case time.Time:
		// OLE dates are timezone-naive day counts since 1899-12-30; use the
		// wall-clock reading of x so what the user sees is what Excel shows.
		//
		// Subtracting two absolute instants (x and the 1899 epoch) would fold
		// in the difference between x's UTC offset and the epoch's offset in
		// x.Location(). For zones whose historical offset differs from the
		// modern one — e.g. IANA Asia/Seoul carries an LMT of +08:27:52 at the
		// 1899 epoch, and any DST zone shifts by 60 min — that drift pushes a
		// midnight date back to the previous day (23:xx). Decompose x into its
		// wall-clock fields and rebuild them in UTC so the elapsed-time
		// arithmetic is zone-neutral. go-ole's GetVariantDate decodes VT_DATE
		// back into a naive-UTC time.Time, so this makes the round trip
		// symmetric to the minute.
		n := time.Date(x.Year(), x.Month(), x.Day(), x.Hour(), x.Minute(), x.Second(), x.Nanosecond(), time.UTC)
		days := n.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24.0
		*out = ole.NewVariant(ole.VT_DATE, int64(math.Float64bits(days)))
	default:
		return fmt.Errorf("scalarToVariant: unsupported cell type %T", val)
	}
	return nil
}

// setNumericVariant stores a float64 as VT_R8 — Excel's native cell number
// representation.
func setNumericVariant(out *ole.VARIANT, f float64) {
	*out = ole.NewVariant(ole.VT_R8, int64(math.Float64bits(f)))
}

// vtTypeMask isolates the base VARTYPE from a VARIANT's VT field, stripping
// the VT_ARRAY / VT_BYREF / VT_VECTOR flag bits (Win32 VT_TYPEMASK = 0x0fff).
const vtTypeMask ole.VT = 0x0fff

// CellError is the typed Go representation of an Excel error VARIANT (VT_ERROR)
// — a cell holding #DIV/0!, #N/A, #VALUE!, and the like. go-ole v1.3.0's
// (*VARIANT).Value() has no VT_ERROR case and returns a bare nil, which is
// indistinguishable from an empty cell; decodeVariantScalar returns a CellError
// instead so callers can tell "error cell" from "blank cell" apart.
//
// SCode is the raw COM SCODE from the VARIANT. Excel encodes its worksheet
// error values as 0x800A0000 | cvErr, so SCode&0xffff recovers the CVErr code
// (2007 for #DIV/0!, 2042 for #N/A, …). String / Error render the familiar
// Excel error text, matching how xlwings surfaces error cells.
type CellError struct {
	SCode uint32
}

// String renders the Excel error text for the CellError's SCODE (e.g.
// "#DIV/0!"). Unknown codes fall back to a hex form.
func (e CellError) String() string {
	switch e.SCode & 0xffff {
	case 2000:
		return "#NULL!"
	case 2007:
		return "#DIV/0!"
	case 2015:
		return "#VALUE!"
	case 2023:
		return "#REF!"
	case 2029:
		return "#NAME?"
	case 2036:
		return "#NUM!"
	case 2042:
		return "#N/A"
	case 2043:
		return "#GETTING_DATA"
	case 2045:
		return "#SPILL!"
	case 2047:
		return "#CONNECT!"
	case 2048:
		return "#BLOCKED!"
	case 2049:
		return "#UNKNOWN!"
	case 2050:
		return "#FIELD!"
	case 2051:
		return "#CALC!"
	}
	return fmt.Sprintf("#ERR(0x%08X)", e.SCode)
}

// Error lets CellError satisfy the error interface — an error cell is a
// legitimate error condition when a caller wants to treat it as one.
func (e CellError) Error() string { return e.String() }

// oleDecimal mirrors the Win32 DECIMAL struct (16 bytes). When a VARIANT holds
// VT_DECIMAL the DECIMAL overlays the whole VARIANT starting at offset 0 — its
// wReserved field aligns with the VARIANT's VT field — so a *VARIANT can be
// reinterpreted as a *oleDecimal.
type oleDecimal struct {
	wReserved uint16
	scale     uint8
	sign      uint8
	hi32      uint32
	lo64      uint64
}

// decodeVariantScalar converts a scalar VARIANT to a Go value. It covers the
// VT_CY (currency), VT_DECIMAL, and VT_ERROR cases that go-ole v1.3.0's
// (*VARIANT).Value() switch omits — those fall through to a bare nil, making a
// currency-formatted or #DIV/0! cell indistinguishable from an empty one.
// Everything else delegates to go-ole's Value().
//
//   - VT_CY:      the OLE CY is an int64 scaled by 1e-4; return Val/10000 as
//     float64 (xlwings returns currency cells as plain numbers).
//   - VT_DECIMAL: converted to float64 via oleaut32!VarR8FromDec.
//   - VT_ERROR:   returned as a typed CellError, except DISP_E_PARAMNOTFOUND
//     (the "omitted optional parameter" marker, never a real cell value),
//     which stays nil.
//
// VT_BYREF variants of CY/DECIMAL (Val holds a pointer to the value) are
// dereferenced. The unsafe.Pointer(&Val) reinterpret pattern mirrors
// decodeVariantArray's SAFEARRAY handling and keeps `go vet` happy.
func decodeVariantScalar(v *ole.VARIANT) interface{} {
	byref := v.VT&ole.VT_BYREF != 0
	switch v.VT & vtTypeMask {
	case ole.VT_CY:
		cy := v.Val
		if byref {
			p := *(*unsafe.Pointer)(unsafe.Pointer(&v.Val))
			if p == nil {
				return nil
			}
			cy = *(*int64)(p)
		}
		return float64(cy) / 10000.0
	case ole.VT_DECIMAL:
		dec := (*oleDecimal)(unsafe.Pointer(v))
		if byref {
			p := *(*unsafe.Pointer)(unsafe.Pointer(&v.Val))
			if p == nil {
				return nil
			}
			dec = (*oleDecimal)(p)
		}
		var out float64
		hr, _, _ := procVarR8FromDec.Call(
			uintptr(unsafe.Pointer(dec)),
			uintptr(unsafe.Pointer(&out)),
		)
		if hr != 0 {
			return nil
		}
		return out
	case ole.VT_ERROR:
		scode := uint32(v.Val)
		if scode == dispEParamNotFound {
			// The omitted-optional-parameter marker, not a worksheet error.
			return nil
		}
		return CellError{SCode: scode}
	}
	return v.Value()
}

// getElement reads one VARIANT cell from a VT_ARRAY|VT_VARIANT SAFEARRAY at
// `indices` (one entry per dimension), converts it to a Go value, and clears
// the temporary VARIANT.
func getElement(sa uintptr, indices []int32) (interface{}, error) {
	var v ole.VARIANT
	ole.VariantInit(&v)
	hr, _, _ := procSafeArrayGetElement.Call(
		sa,
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(&v)),
	)
	if hr != 0 {
		return nil, fmt.Errorf("SafeArrayGetElement failed: 0x%x", hr)
	}
	val := decodeVariantScalar(&v)
	vt := v.VT
	v.Clear()
	// For object-typed cells, v.Value() returned the raw COM interface pointer
	// (ToIDispatch/ToIUnknown — no AddRef), which v.Clear() has just Released.
	// Returning val here would hand the caller a dangling pointer (use-after-
	// free / refcount underflow). A value grid can't represent a live object
	// anyway, so degrade such cells to nil. Scalars (VT_R8/BSTR/BOOL/DATE) are
	// copied by Value(), so returning them after Clear is safe.
	if vt == ole.VT_DISPATCH || vt == ole.VT_UNKNOWN {
		return nil, nil
	}
	return val, nil
}
