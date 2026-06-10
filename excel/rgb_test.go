//go:build windows

// Unit test for the pure excel.RGB helper — no COM server involved, so it
// runs under plain `go test ./...` (unlike the Font integration tests in
// font_test.go, which need real Excel).

package excel_test

import (
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestRGB pins the OLE color packing: blue occupies the high byte.
func TestRGB(t *testing.T) {
	if got := excel.RGB(255, 0, 0); got != 255 {
		t.Errorf("RGB(red): got %d, want 255", got)
	}
	if got := excel.RGB(0, 0, 255); got != 0xFF0000 {
		t.Errorf("RGB(blue): got %d, want %d", got, 0xFF0000)
	}
	if got := excel.RGB(0x12, 0x34, 0x56); got != 0x563412 {
		t.Errorf("RGB mixed: got %#x, want 0x563412", got)
	}
}
