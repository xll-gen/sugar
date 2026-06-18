//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// ChartType mirrors Excel's `XlChartType` enum. Only the most common values
// are named here; any other XlChartType integer can be cast directly:
// `excel.ChartType(76)`.
type ChartType int32

const (
	ChartArea             ChartType = 1     // xlArea
	ChartLine             ChartType = 4     // xlLine
	ChartPie              ChartType = 5     // xlPie
	ChartColumnClustered  ChartType = 51    // xlColumnClustered
	ChartColumnStacked    ChartType = 52    // xlColumnStacked
	ChartBarClustered     ChartType = 57    // xlBarClustered
	ChartXYScatter        ChartType = -4169 // xlXYScatter
	ChartXYScatterLines   ChartType = 74    // xlXYScatterLines
)

// xlTypePDF is the XlFixedFormatType value for PDF export.
const xlTypePDF int32 = 0

// Chart is an embedded chart on a worksheet — the Go equivalent of xlwings'
// `Chart`.
//
// Like xlwings, this type fuses Excel COM's two-level model: the chain
// points at the *ChartObject* (the positioned container), and chart-level
// members (ChartType, SetSourceData, exports) reach through its inner
// `.Chart` automatically.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/chart.html
type Chart interface {
	sugar.Chain
	// Name returns the chart object's name.
	Name() (string, error)
	// SetName renames the chart object.
	SetName(name string) Chart
	// ChartType returns the chart's XlChartType.
	ChartType() (ChartType, error)
	// SetChartType changes the chart's type (e.g. excel.ChartLine).
	SetChartType(t ChartType) Chart
	// SetSourceData points the chart at the data range.
	SetSourceData(source Range) error
	// Left / Top / Width / Height are the position and size in points.
	Left() (float64, error)
	Top() (float64, error)
	Width() (float64, error)
	Height() (float64, error)
	// SetPosition moves and resizes the chart object (points).
	SetPosition(left, top, width, height float64) Chart
	// ToPNG exports the chart as a PNG image at the given absolute path.
	ToPNG(path string) error
	// ToPDF exports the chart sheet content as a PDF at the given absolute
	// path.
	ToPDF(path string) error
	// Delete removes the chart object from its worksheet.
	Delete() error
}

type chart struct {
	sugar.Chain // the ChartObject dispatch
}

// inner returns the ChartObject's inner Chart dispatch.
func (c *chart) inner() sugar.Chain {
	return c.Get("Chart")
}

func (c *chart) Name() (string, error) {
	return getString(c, "Name")
}

func (c *chart) SetName(name string) Chart {
	return &chart{c.Put("Name", name)}
}

func (c *chart) ChartType() (ChartType, error) {
	v, err := getInt32(c.inner(), "ChartType")
	if err != nil {
		return 0, err
	}
	return ChartType(v), nil
}

func (c *chart) SetChartType(t ChartType) Chart {
	inner := c.inner().Put("ChartType", int32(t))
	if inner.Err() != nil {
		// Put returns an error-only chain on failure; surfacing it keeps
		// the fluent contract (callers check .Err() at the end).
		return &chart{inner}
	}
	return c
}

func (c *chart) SetSourceData(source Range) error {
	return c.inner().Call("SetSourceData", source).Err()
}

func (c *chart) Left() (float64, error)   { return getFloat64(c, "Left") }
func (c *chart) Top() (float64, error)    { return getFloat64(c, "Top") }
func (c *chart) Width() (float64, error)  { return getFloat64(c, "Width") }
func (c *chart) Height() (float64, error) { return getFloat64(c, "Height") }

func (c *chart) SetPosition(left, top, width, height float64) Chart {
	next := c.Put("Left", left).Put("Top", top).Put("Width", width).Put("Height", height)
	return &chart{next}
}

func (c *chart) ToPNG(path string) error {
	return c.inner().Call("Export", path, "PNG").Err()
}

func (c *chart) ToPDF(path string) error {
	return c.inner().Call("ExportAsFixedFormat", xlTypePDF, path).Err()
}

func (c *chart) Delete() error {
	return c.Call("Delete").Err()
}
