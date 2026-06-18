//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// RGB packs red/green/blue components into the OLE color integer Excel's
// Color properties use (&HBBGGRR — blue in the high byte). xlwings accepts
// (r, g, b) tuples; this is the Go spelling.
func RGB(r, g, b uint8) int32 {
	return int32(r) | int32(g)<<8 | int32(b)<<16
}

// Font is the character formatting of a range — the Go equivalent of
// xlwings' `Font`, reached via Range.Font().
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/font.html
type Font interface {
	sugar.Chain
	// Name returns the font name (e.g. "Calibri").
	Name() (string, error)
	// SetName sets the font name.
	SetName(name string) Font
	// Size returns the font size in points.
	Size() (float64, error)
	// SetSize sets the font size in points.
	SetSize(size float64) Font
	// Bold reports whether the font is bold.
	Bold() (bool, error)
	// SetBold toggles bold.
	SetBold(on bool) Font
	// Italic reports whether the font is italic.
	Italic() (bool, error)
	// SetItalic toggles italic.
	SetItalic(on bool) Font
	// Color returns the font color as an OLE color integer (see RGB).
	Color() (int32, error)
	// SetColor sets the font color. Build the value with excel.RGB.
	SetColor(color int32) Font
}

type font struct {
	sugar.Chain
}

func (f *font) Name() (string, error) {
	return getString(f, "Name")
}

func (f *font) SetName(name string) Font {
	return &font{f.Put("Name", name)}
}

func (f *font) Size() (float64, error) {
	return getFloat64(f, "Size")
}

func (f *font) SetSize(size float64) Font {
	return &font{f.Put("Size", size)}
}

func (f *font) Bold() (bool, error) {
	return getBool(f, "Bold")
}

func (f *font) SetBold(on bool) Font {
	return &font{f.Put("Bold", on)}
}

func (f *font) Italic() (bool, error) {
	return getBool(f, "Italic")
}

func (f *font) SetItalic(on bool) Font {
	return &font{f.Put("Italic", on)}
}

func (f *font) Color() (int32, error) {
	return getInt32(f, "Color")
}

func (f *font) SetColor(color int32) Font {
	return &font{f.Put("Color", color)}
}
