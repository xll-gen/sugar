//go:build windows && excel_integration

// Integration tests for excel.Font (reached via Range.Font()).
// Build with `-tags=excel_integration`.

package excel_test

import (
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestFont_RoundTrip sets every Font property and reads it back.
func TestFont_RoundTrip(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("styled")
		f := sheet.Range("A1").Font()

		red := excel.RGB(255, 0, 0)
		res := f.SetName("Arial").SetSize(14).SetBold(true).SetItalic(true).SetColor(red)
		if err := res.Err(); err != nil {
			t.Fatalf("font setters: %v", err)
		}

		name, err := f.Name()
		if err != nil || name != "Arial" {
			t.Errorf("Name: got %q err=%v; want Arial", name, err)
		}
		size, err := f.Size()
		if err != nil || size != 14 {
			t.Errorf("Size: got %v err=%v; want 14", size, err)
		}
		bold, err := f.Bold()
		if err != nil || !bold {
			t.Errorf("Bold: got %v err=%v; want true", bold, err)
		}
		italic, err := f.Italic()
		if err != nil || !italic {
			t.Errorf("Italic: got %v err=%v; want true", italic, err)
		}
		color, err := f.Color()
		if err != nil || color != red {
			t.Errorf("Color: got %d err=%v; want %d", color, err, red)
		}
	})
}

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
