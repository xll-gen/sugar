//go:build windows

// Excel-free unit tests for the Names error-classification helper. These run
// under plain `go test ./...` (no excel_integration tag) because they exercise
// pure Go logic over synthetic COM errors.

package excel

import (
	"errors"
	"testing"

	ole "github.com/go-ole/go-ole"
)

// TestIsNameNotFound pins the item-2 fix: only the not-found HRESULT classes
// may be folded into a clean "absent" result; every other COM failure must be
// reported as an error (isNameNotFound == false), so Contains can propagate it.
func TestIsNameNotFound(t *testing.T) {
	cases := []struct {
		name string
		err  error
		want bool
	}{
		{"DISP_E_BADINDEX is not-found", ole.NewError(hrDispEBadIndex), true},
		{"Excel item-not-found is not-found", ole.NewError(hrXlItemNotFound), true},
		{"access denied is a real error", ole.NewError(0x80070005), false},
		{"RPC disconnected is a real error", ole.NewError(0x80010108), false},
		{"non-OLE error is not not-found", errors.New("boom"), false},
		{"nil is not not-found", nil, false},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := isNameNotFound(tc.err); got != tc.want {
				t.Errorf("isNameNotFound(%v) = %v; want %v", tc.err, got, tc.want)
			}
		})
	}
}
