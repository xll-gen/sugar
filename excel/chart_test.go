//go:build windows && excel_integration

// Integration tests for excel.Chart / excel.Charts.
// Build with `-tags=excel_integration`. Skipped on machines without Excel.

package excel_test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// seedChartData writes a small two-series table and returns its range.
func seedChartData(t *testing.T, sheet excel.Worksheet) excel.Range {
	t.Helper()
	data := [][]interface{}{
		{"Month", "Sales"},
		{"Jan", 10.0},
		{"Feb", 20.0},
		{"Mar", 15.0},
	}
	rng := sheet.Range("A1", "B4")
	if err := rng.SetValue(data).Err(); err != nil {
		t.Fatalf("seed data: %v", err)
	}
	return rng
}

// TestCharts_AddCountDelete walks the collection lifecycle: Add grows Count,
// Item resolves by index and by name, Delete shrinks Count.
func TestCharts_AddCountDelete(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		src := seedChartData(t, sheet)
		charts := sheet.Charts()

		before, err := charts.Count()
		if err != nil {
			t.Fatalf("Count: %v", err)
		}

		ch := charts.Add(excel.ChartAt(10, 10), excel.ChartSize(300, 200))
		if err := ch.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}
		if err := ch.SetSourceData(src); err != nil {
			t.Fatalf("SetSourceData: %v", err)
		}

		after, err := charts.Count()
		if err != nil || after != before+1 {
			t.Errorf("Count after Add: got %d err=%v; want %d", after, err, before+1)
		}

		named := ch.SetName("SalesChart")
		if err := named.Err(); err != nil {
			t.Fatalf("SetName: %v", err)
		}
		got, err := charts.Item("SalesChart").Name()
		if err != nil || got != "SalesChart" {
			t.Errorf("Item by name: got %q err=%v; want SalesChart", got, err)
		}

		if err := charts.Item(1).Delete(); err != nil {
			t.Fatalf("Delete: %v", err)
		}
		final, err := charts.Count()
		if err != nil || final != before {
			t.Errorf("Count after Delete: got %d err=%v; want %d", final, err, before)
		}
	})
}

// TestChart_TypeAndGeometry round-trips ChartType and the position/size
// getters against the values passed to Add/SetPosition.
func TestChart_TypeAndGeometry(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		src := seedChartData(t, sheet)

		ch := sheet.Charts().Add()
		if err := ch.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}
		if err := ch.SetSourceData(src); err != nil {
			t.Fatalf("SetSourceData: %v", err)
		}

		if err := ch.SetChartType(excel.ChartLine).Err(); err != nil {
			t.Fatalf("SetChartType: %v", err)
		}
		ct, err := ch.ChartType()
		if err != nil || ct != excel.ChartLine {
			t.Errorf("ChartType: got %d err=%v; want %d", ct, err, excel.ChartLine)
		}

		moved := ch.SetPosition(50, 60, 400, 250)
		if err := moved.Err(); err != nil {
			t.Fatalf("SetPosition: %v", err)
		}
		left, _ := moved.Left()
		top, _ := moved.Top()
		w, _ := moved.Width()
		h, _ := moved.Height()
		if left != 50 || top != 60 || w != 400 || h != 250 {
			t.Errorf("geometry: got (%v,%v,%v,%v), want (50,60,400,250)", left, top, w, h)
		}
	})
}

// TestChart_ToPNG exports a chart image and checks a non-empty PNG file
// lands on disk — the xlwings `chart.to_png()` path.
func TestChart_ToPNG(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		src := seedChartData(t, sheet)

		ch := sheet.Charts().Add()
		if err := ch.SetSourceData(src); err != nil {
			t.Fatalf("SetSourceData: %v", err)
		}
		if err := ch.SetChartType(excel.ChartColumnClustered).Err(); err != nil {
			t.Fatalf("SetChartType: %v", err)
		}

		path := filepath.Join(t.TempDir(), "chart.png")
		if err := ch.ToPNG(path); err != nil {
			t.Fatalf("ToPNG: %v", err)
		}
		info, err := os.Stat(path)
		if err != nil {
			t.Fatalf("exported file missing: %v", err)
		}
		if info.Size() == 0 {
			t.Errorf("exported PNG is empty")
		}
	})
}
