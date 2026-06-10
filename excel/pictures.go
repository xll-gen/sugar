//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// MsoTriState values used by Shapes.AddPicture.
const (
	msoFalse int32 = 0
	msoTrue  int32 = -1
)

// Pictures is the picture collection of a worksheet — the Go equivalent of
// xlwings' `pictures`. Item/Count come from the legacy Worksheet.Pictures
// COM collection; Add inserts through Shapes.AddPicture (the modern entry
// point), mirroring xlwings' implementation.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/pictures.html
type Pictures interface {
	sugar.Chain
	// Add inserts an image file (absolute path) onto the worksheet. With no
	// options the picture lands at (0, 0) in its native size; use PictureAt
	// / PictureSize / PictureName to adjust.
	Add(filename string, opts ...PictureOption) Picture
	// Item returns a picture by 1-based index or by name.
	Item(index interface{}) Picture
	// Count returns the number of pictures on the worksheet.
	Count() (int32, error)
}

// PictureOption configures Pictures.Add. Build with PictureAt, PictureSize,
// PictureName.
type PictureOption func(*pictureOptions)

type pictureOptions struct {
	left, top     float64
	width, height float64
	name          string
}

// PictureAt places the new picture's top-left corner at (left, top) points.
func PictureAt(left, top float64) PictureOption {
	return func(o *pictureOptions) { o.left, o.top = left, top }
}

// PictureSize scales the new picture to width × height points. Without it
// the image keeps its native size.
func PictureSize(width, height float64) PictureOption {
	return func(o *pictureOptions) { o.width, o.height = width, height }
}

// PictureName names the picture after insertion (xlwings' `name=` kwarg).
func PictureName(name string) PictureOption {
	return func(o *pictureOptions) { o.name = name }
}

type pictures struct {
	sugar.Chain             // a Worksheet.Pictures snapshot (Err/ForEach anchor)
	sheet       sugar.Chain // the parent worksheet
}

// collection re-fetches the legacy Pictures collection. The COM object
// returned by Worksheet.Pictures() is a snapshot of the pictures that
// existed at call time (its Count never grows), so every lookup must go
// through a fresh call — exactly what xlwings' `api` property does.
func (p *pictures) collection() sugar.Chain {
	return p.sheet.Call("Pictures")
}

func (p *pictures) Add(filename string, opts ...PictureOption) Picture {
	o := pictureOptions{width: -1, height: -1} // -1 keeps the native size
	for _, opt := range opts {
		opt(&o)
	}
	shp := p.sheet.Get("Shapes").Call("AddPicture",
		filename, msoFalse, msoTrue, o.left, o.top, o.width, o.height)
	pic := &picture{shp}
	if o.name != "" && shp.Err() == nil {
		pic.Chain = shp.Put("Name", o.name)
	}
	return pic
}

func (p *pictures) Item(index interface{}) Picture {
	// Pictures.Item is a method (like Names.Item), not a parameterized
	// property.
	return &picture{p.collection().Call("Item", index)}
}

func (p *pictures) Count() (int32, error) {
	v, err := p.collection().Get("Count").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}
