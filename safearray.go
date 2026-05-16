//go:build windows

package sugar

import (
	"fmt"
	"syscall"
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
