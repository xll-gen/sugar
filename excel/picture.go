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

func (p *picture) Name() (string, error) {
	v, err := p.Get("Name").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (p *picture) SetName(name string) Picture {
	return &picture{p.Put("Name", name)}
}

func (p *picture) Left() (float64, error)   { return shapeFloat(p, "Left") }
func (p *picture) Top() (float64, error)    { return shapeFloat(p, "Top") }
func (p *picture) Width() (float64, error)  { return shapeFloat(p, "Width") }
func (p *picture) Height() (float64, error) { return shapeFloat(p, "Height") }

func (p *picture) SetPosition(left, top, width, height float64) Picture {
	next := p.Put("Left", left).Put("Top", top).Put("Width", width).Put("Height", height)
	return &picture{next}
}

func (p *picture) Delete() error {
	return p.Call("Delete").Err()
}
