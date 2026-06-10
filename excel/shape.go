//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// ShapeType mirrors Office's `MsoShapeType` enum. Only the values the
// wrappers branch on are named; cast any other MsoShapeType integer
// directly: `excel.ShapeType(17)`.
type ShapeType int32

const (
	ShapeTypePicture       ShapeType = 13 // msoPicture
	ShapeTypeChart         ShapeType = 3  // msoChart
	ShapeTypeAutoShape     ShapeType = 1  // msoAutoShape
	ShapeTypeTextBox       ShapeType = 17 // msoTextBox
	ShapeTypeLinkedPicture ShapeType = 11 // msoLinkedPicture
)

// Shape is a drawing-layer object on a worksheet — the Go equivalent of
// xlwings' `Shape`. Charts and pictures are shapes too; the typed Chart and
// Picture wrappers are preferred when the kind is known.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/shape.html
type Shape interface {
	sugar.Chain
	// Name returns the shape's name.
	Name() (string, error)
	// SetName renames the shape.
	SetName(name string) Shape
	// Type returns the MsoShapeType of the shape.
	Type() (ShapeType, error)
	// Left / Top / Width / Height are the position and size in points.
	Left() (float64, error)
	Top() (float64, error)
	Width() (float64, error)
	Height() (float64, error)
	// SetPosition moves and resizes the shape (points).
	SetPosition(left, top, width, height float64) Shape
	// Delete removes the shape from its worksheet.
	Delete() error
}

type shape struct {
	sugar.Chain
}

func (s *shape) Name() (string, error) {
	v, err := s.Get("Name").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (s *shape) SetName(name string) Shape {
	return &shape{s.Put("Name", name)}
}

func (s *shape) Type() (ShapeType, error) {
	v, err := s.Get("Type").Value()
	if err != nil {
		return 0, err
	}
	return ShapeType(toInt32(v)), nil
}

func (s *shape) Left() (float64, error)   { return shapeFloat(s, "Left") }
func (s *shape) Top() (float64, error)    { return shapeFloat(s, "Top") }
func (s *shape) Width() (float64, error)  { return shapeFloat(s, "Width") }
func (s *shape) Height() (float64, error) { return shapeFloat(s, "Height") }

func shapeFloat(c sugar.Chain, prop string) (float64, error) {
	v, err := c.Get(prop).Value()
	if err != nil {
		return 0, err
	}
	return toFloat64(v), nil
}

func (s *shape) SetPosition(left, top, width, height float64) Shape {
	next := s.Put("Left", left).Put("Top", top).Put("Width", width).Put("Height", height)
	return &shape{next}
}

func (s *shape) Delete() error {
	return s.Call("Delete").Err()
}
