//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Picture is an image placed on a worksheet — the Go equivalent of xlwings'
// `Picture`. The chain points at the underlying Shape/Picture dispatch;
// both expose the same Name/geometry/Delete surface.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/picture.html
type Picture interface {
	sugar.Chain
	// Name returns the picture's name.
	Name() (string, error)
	// SetName renames the picture.
	SetName(name string) Picture
	// Left / Top / Width / Height are the position and size in points.
	Left() (float64, error)
	Top() (float64, error)
	Width() (float64, error)
	Height() (float64, error)
	// SetPosition moves and resizes the picture (points).
	SetPosition(left, top, width, height float64) Picture
	// Delete removes the picture from its worksheet.
	Delete() error
}

type picture struct {
	sugar.Chain
}

// wrapPicture wraps a chain in the Picture typed wrapper. It is the single
// construction point for the chain -> Picture convention.
func wrapPicture(c sugar.Chain) Picture { return &picture{c} }

func (p *picture) Name() (string, error) {
	return getString(p, "Name")
}

func (p *picture) SetName(name string) Picture {
	return wrapPicture(p.Put("Name", name))
}

func (p *picture) Left() (float64, error)   { return getFloat64(p, "Left") }
func (p *picture) Top() (float64, error)    { return getFloat64(p, "Top") }
func (p *picture) Width() (float64, error)  { return getFloat64(p, "Width") }
func (p *picture) Height() (float64, error) { return getFloat64(p, "Height") }

func (p *picture) SetPosition(left, top, width, height float64) Picture {
	next := p.Put("Left", left).Put("Top", top).Put("Width", width).Put("Height", height)
	return wrapPicture(next)
}

func (p *picture) Delete() error {
	return p.Call("Delete").Err()
}
