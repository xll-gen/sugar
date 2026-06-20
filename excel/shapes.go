//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Shapes is the drawing-layer collection of a worksheet — the Go equivalent
// of xlwings' `shapes`.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/shapes.html
type Shapes interface {
	sugar.Chain
	// Item returns a shape by 1-based index or by name.
	Item(index interface{}) Shape
	// Count returns the number of shapes on the worksheet.
	Count() (int32, error)
	// ForEachShape iterates the collection with the typed wrapper. Iteration
	// stops when fn returns a non-nil error (sugar.ErrForEachBreak to stop
	// cleanly); the error surfaces on the returned chain like sugar.ForEach.
	ForEachShape(fn func(s Shape) error) sugar.Chain
}

type shapes struct {
	collection[Shape]
}

// wrapShapes wraps a chain in the Shapes typed wrapper. It is the single
// construction point for the chain -> Shapes convention.
func wrapShapes(c sugar.Chain) Shapes { return &shapes{newCollection(c, wrapShape)} }

func (s *shapes) Item(index interface{}) Shape {
	// Shapes.Item is a method in the type library, like Names.Item.
	return s.itemByCall(index)
}

func (s *shapes) Count() (int32, error) {
	return s.count()
}

func (s *shapes) ForEachShape(fn func(sh Shape) error) sugar.Chain {
	return s.ForEach(func(item sugar.Chain) error {
		return fn(wrapShape(item))
	})
}
