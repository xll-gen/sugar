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
	// Add creates a new empty chart object. With no options it uses xlwings'
	// defaults: position (0, 0) and size (355, 211) points. Override either
	// with ChartAt / ChartSize. Follow with SetSourceData and SetChartType.
	//
	// xlwings analogue: `charts.add(left=0, top=0, width=355, height=211)`.
	Add(opts ...ChartOption) Chart
	// Item returns a chart by 1-based index or by name.
	Item(index interface{}) Chart
	// Count returns the number of chart objects on the worksheet.
	Count() (int32, error)
}

type charts struct {
	sugar.Chain
}

// ChartOption configures Charts.Add. Build with ChartAt and ChartSize, mirroring
// the functional-option style of Pictures.Add / Worksheets.Add / Books.Open.
type ChartOption func(*chartOptions)

type chartOptions struct {
	left, top     float64
	width, height float64
}

// ChartAt sets the top-left position of the new chart in points. xlwings'
// `left`/`top` keywords; defaults to (0, 0).
func ChartAt(left, top float64) ChartOption {
	return func(o *chartOptions) {
		o.left = left
		o.top = top
	}
}

// ChartSize sets the width and height of the new chart in points. xlwings'
// `width`/`height` keywords; defaults to (355, 211).
func ChartSize(width, height float64) ChartOption {
	return func(o *chartOptions) {
		o.width = width
		o.height = height
	}
}

func (c *charts) Add(opts ...ChartOption) Chart {
	// xlwings defaults: left=0, top=0, width=355, height=211.
	o := chartOptions{left: 0, top: 0, width: 355, height: 211}
	for _, opt := range opts {
		opt(&o)
	}
	return &chart{c.Call("Add", o.left, o.top, o.width, o.height)}
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
