//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Charts is the collection of embedded charts on a worksheet — the Go
// equivalent of xlwings' `charts`. It wraps Excel COM's `ChartObjects`
// collection.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/charts.html
type Charts interface {
	sugar.Chain
	// Add creates a new empty chart object at (left, top) with the given
	// size, all in points. xlwings' defaults are (0, 0, 355, 211). Follow
	// with SetSourceData and SetChartType.
	Add(left, top, width, height float64) Chart
	// Item returns a chart by 1-based index or by name.
	Item(index interface{}) Chart
	// Count returns the number of chart objects on the worksheet.
	Count() (int32, error)
}

type charts struct {
	sugar.Chain
}

func (c *charts) Add(left, top, width, height float64) Chart {
	return &chart{c.Call("Add", left, top, width, height)}
}

func (c *charts) Item(index interface{}) Chart {
	// ChartObjects.Item is a method (like Names.Item), not a parameterized
	// property — DISPATCH_METHOD required.
	return &chart{c.Call("Item", index)}
}

func (c *charts) Count() (int32, error) {
	v, err := c.Get("Count").Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}
