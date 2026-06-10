//go:build windows

package sugar

import (
	"fmt"
	"math"
	"syscall"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
)

// 2-D SAFEARRAY decoding for VARIANT-of-array values.
//
// go-ole v1.3.0's `SafeArrayConversion.ToValueArray` is hard-wired to 1-D and
// silently corrupts multi-dim arrays. Excel's `Range.Value` always returns a
// 2-D `SAFEARRAY` of `VARIANT`, so we reach for `oleaut32.dll!SafeArrayGetElement`
// directly with a properly shaped index array.

var oleaut32 = syscall.NewLazyDLL("oleaut32.dll")

var (
	procSafeArrayGetDim     = oleaut32.NewProc("SafeArrayGetDim")
	procSafeArrayGetLBound  = oleaut32.NewProc("SafeArrayGetLBound")
	procSafeArrayGetUBound  = oleaut32.NewProc("SafeArrayGetUBound")
	procSafeArrayGetElement = oleaut32.NewProc("SafeArrayGetElement")
	procSafeArrayCreate     = oleaut32.NewProc("SafeArrayCreate")
	procSafeArrayPutElement = oleaut32.NewProc("SafeArrayPutElement")
	procSafeArrayDestroy    = oleaut32.NewProc("SafeArrayDestroy")
)

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
	// The VARIANT's `Val` field is `int64` but actually stores the SAFEARRAY
	// pointer. Reinterpret the bytes via &Val — not a uintptr round-trip —
	// to keep go vet happy with `unsafe.Pointer` rules.
	sa := *(*unsafe.Pointer)(unsafe.Pointer(&v.Val))
	if sa == nil {
		return nil, nil
	}

	dims, _, _ := procSafeArrayGetDim.Call(uintptr(sa))
	switch uint32(dims) {
	case 1:
		return decode1D(sa)
	case 2:
		return decode2D(sa)
	default:
		return nil, fmt.Errorf("decodeVariantArray: unsupported dimension count %d", dims)
	}
}

func decode1D(sa unsafe.Pointer) ([]interface{}, error) {
	lo, hi, err := bounds(sa, 1)
	if err != nil {
		return nil, err
	}
	n := int(hi - lo + 1)
	if n <= 0 {
		return []interface{}{}, nil
	}
	out := make([]interface{}, n)
	for i := int32(0); i < int32(n); i++ {
		val, err := getElement(sa, []int32{lo + i})
		if err != nil {
			return nil, err
		}
		out[i] = val
	}
	return out, nil
}

func decode2D(sa unsafe.Pointer) ([][]interface{}, error) {
	rLo, rHi, err := bounds(sa, 1)
	if err != nil {
		return nil, err
	}
	cLo, cHi, err := bounds(sa, 2)
	if err != nil {
		return nil, err
	}
	rows := int(rHi - rLo + 1)
	cols := int(cHi - cLo + 1)
	if rows <= 0 || cols <= 0 {
		return [][]interface{}{}, nil
	}
	out := make([][]interface{}, rows)
	for r := 0; r < rows; r++ {
		row := make([]interface{}, cols)
		for c := 0; c < cols; c++ {
			val, err := getElement(sa, []int32{rLo + int32(r), cLo + int32(c)})
			if err != nil {
				return nil, err
			}
			row[c] = val
		}
		out[r] = row
	}
	return out, nil
}

func bounds(sa unsafe.Pointer, dim uint32) (int32, int32, error) {
	var lo, hi int32
	hr, _, _ := procSafeArrayGetLBound.Call(uintptr(sa), uintptr(dim), uintptr(unsafe.Pointer(&lo)))
	if hr != 0 {
		return 0, 0, fmt.Errorf("SafeArrayGetLBound(dim=%d) failed: 0x%x", dim, hr)
	}
	hr, _, _ = procSafeArrayGetUBound.Call(uintptr(sa), uintptr(dim), uintptr(unsafe.Pointer(&hi)))
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
	return nil, fmt.Errorf("encodeVariantArray: unsupported type %T", value)
}

func encode1D(src []interface{}) (*ole.VARIANT, error) {
	bounds := []safeArrayBound{{cElements: uint32(len(src)), lLbound: 0}}
	sa, err := createVariantSafeArray(bounds)
	if err != nil {
		return nil, err
	}
	for i, val := range src {
		if err := putElement(sa, []int32{int32(i)}, val); err != nil {
			procSafeArrayDestroy.Call(sa)
			return nil, err
		}
	}
	return wrapSafeArray(sa), nil
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
	for r, row := range src {
		for c, val := range row {
			if err := putElement(sa, []int32{int32(r), int32(c)}, val); err != nil {
				procSafeArrayDestroy.Call(sa)
				return nil, err
			}
		}
	}
	return wrapSafeArray(sa), nil
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
func scalarToVariant(val interface{}, out *ole.VARIANT) error {
	ole.VariantInit(out)
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
		days := x.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, x.Location())).Hours() / 24.0
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

// getElement reads one VARIANT cell from a VT_ARRAY|VT_VARIANT SAFEARRAY at
// `indices` (one entry per dimension), converts it to a Go value, and clears
// the temporary VARIANT.
func getElement(sa unsafe.Pointer, indices []int32) (interface{}, error) {
	var v ole.VARIANT
	ole.VariantInit(&v)
	hr, _, _ := procSafeArrayGetElement.Call(
		uintptr(sa),
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(&v)),
	)
	if hr != 0 {
		return nil, fmt.Errorf("SafeArrayGetElement failed: 0x%x", hr)
	}
	val := v.Value()
	v.Clear()
	return val, nil
}
