//go:build windows

// Excel-free unit tests for the Missing()/IsMissing() omitted-optional marker.
// Building a VT_ERROR VARIANT and reading its fields needs no COM server and no
// CoInitialize, so these run under plain `go test ./...`.

package sugar

import (
	"testing"

	"github.com/go-ole/go-ole"
)

// TestIsMissing pins the VALUE-based recognition of the omitted-optional
// placeholder. The distinction that matters is the second case: a worksheet
// error VARIANT (#DIV/0!) is also VT_ERROR, and a real VARIANT of any other
// type is also a *ole.VARIANT — a type-only check calls all of them "Missing"
// and silently drops them from a positional COM argument list.
func TestIsMissing(t *testing.T) {
	divZero := ole.NewVariant(ole.VT_ERROR, 0x800A07D7) // Excel #DIV/0! cell error
	r8 := ole.NewVariant(ole.VT_R8, 0)
	bstr := ole.NewVariant(ole.VT_BSTR, 0)
	// A BYREF/flagged VT_ERROR carrying the same SCODE must still be
	// recognised — IsMissing masks the VT flag bits.
	byrefMissing := ole.NewVariant(ole.VT_ERROR|ole.VT_BYREF, dispEParamNotFound)

	cases := []struct {
		name string
		v    interface{}
		want bool
	}{
		{"Missing()", Missing(), true},
		{"byref VT_ERROR marker", &byrefMissing, true},
		{"#DIV/0! cell error", &divZero, false},
		{"VT_R8 variant", &r8, false},
		{"VT_BSTR variant", &bstr, false},
		{"nil interface", nil, false},
		{"typed nil *VARIANT", (*ole.VARIANT)(nil), false},
		{"int", int(1), false},
		{"string", "x", false},
		{"VARIANT by value", ole.NewVariant(ole.VT_ERROR, dispEParamNotFound), false},
	}
	for _, c := range cases {
		if got := IsMissing(c.v); got != c.want {
			t.Errorf("IsMissing(%s) = %v, want %v", c.name, got, c.want)
		}
	}
}
