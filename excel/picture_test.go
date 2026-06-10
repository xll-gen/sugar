//go:build windows && excel_integration

// Integration tests for excel.Picture / excel.Pictures and excel.Shape /
// excel.Shapes. Build with `-tags=excel_integration`.

package excel_test

import (
	"path/filepath"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// makePNG renders a tiny chart and exports it — a self-contained way to get
// a real image file without shipping test fixtures.
func makePNG(t *testing.T, sheet excel.Worksheet) string {
	t.Helper()
	if err := sheet.Range("A1", "B3").SetValue([][]interface{}{
		{"x", "y"},
		{1.0, 2.0},
		{2.0, 4.0},
	}).Err(); err != nil {
		t.Fatalf("seed: %v", err)
	}
	ch := sheet.Charts().Add(excel.ChartSize(200, 150))
	if err := ch.SetSourceData(sheet.Range("A1", "B3")); err != nil {
		t.Fatalf("SetSourceData: %v", err)
	}
	path := filepath.Join(t.TempDir(), "source.png")
	if err := ch.ToPNG(path); err != nil {
		t.Fatalf("ToPNG: %v", err)
	}
	if err := ch.Delete(); err != nil {
		t.Fatalf("chart cleanup: %v", err)
	}
	return path
}

// TestPictures_AddItemCountDelete walks the picture lifecycle, including the
// PictureAt / PictureSize / PictureName options.
func TestPictures_AddItemCountDelete(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		png := makePNG(t, sheet)
		pics := sheet.Pictures()

		before, err := pics.Count()
		if err != nil {
			t.Fatalf("Count: %v", err)
		}

		pic := pics.Add(png,
			excel.PictureAt(30, 40),
			excel.PictureSize(120, 90),
			excel.PictureName("Logo"))
		if err := pic.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}

		after, err := pics.Count()
		if err != nil || after != before+1 {
			t.Errorf("Count after Add: got %d err=%v; want %d", after, err, before+1)
		}

		got, err := pics.Item("Logo").Name()
		if err != nil || got != "Logo" {
			t.Errorf("Item by name: got %q err=%v; want Logo", got, err)
		}

		left, _ := pic.Left()
		top, _ := pic.Top()
		w, _ := pic.Width()
		h, _ := pic.Height()
		if left != 30 || top != 40 || w != 120 || h != 90 {
			t.Errorf("geometry: got (%v,%v,%v,%v), want (30,40,120,90)", left, top, w, h)
		}

		if err := pics.Item("Logo").Delete(); err != nil {
			t.Fatalf("Delete: %v", err)
		}
		final, err := pics.Count()
		if err != nil || final != before {
			t.Errorf("Count after Delete: got %d err=%v; want %d", final, err, before)
		}
	})
}

// TestPictures_AddNativeSize omits PictureSize: the image keeps its natural
// dimensions (the COM AddPicture -1/-1 convention).
func TestPictures_AddNativeSize(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		png := makePNG(t, sheet)

		pic := sheet.Pictures().Add(png)
		if err := pic.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}
		w, err := pic.Width()
		if err != nil || w <= 0 {
			t.Errorf("native width: got %v err=%v; want > 0", w, err)
		}
	})
}

// TestShapes_TypedIteration exercises Shapes.Count, Item, Type, and the
// typed ForEachShape iterator — pictures and charts both live on the
// drawing layer.
func TestShapes_TypedIteration(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		png := makePNG(t, sheet)
		if err := sheet.Pictures().Add(png, excel.PictureName("P1")).Err(); err != nil {
			t.Fatalf("Add picture: %v", err)
		}
		if err := sheet.Charts().Add(excel.ChartAt(200, 0), excel.ChartSize(150, 100)).Err(); err != nil {
			t.Fatalf("Add chart: %v", err)
		}

		shapes := sheet.Shapes()
		count, err := shapes.Count()
		if err != nil || count != 2 {
			t.Fatalf("Shapes.Count: got %d err=%v; want 2", count, err)
		}

		st, err := shapes.Item("P1").Type()
		if err != nil || st != excel.ShapeTypePicture {
			t.Errorf("Item(P1).Type: got %d err=%v; want %d (picture)", st, err, excel.ShapeTypePicture)
		}

		var pictureCount, chartCount int
		res := shapes.ForEachShape(func(s excel.Shape) error {
			st, err := s.Type()
			if err != nil {
				return err
			}
			switch st {
			case excel.ShapeTypePicture:
				pictureCount++
			case excel.ShapeTypeChart:
				chartCount++
			}
			return nil
		})
		if err := res.Err(); err != nil {
			t.Fatalf("ForEachShape: %v", err)
		}
		if pictureCount != 1 || chartCount != 1 {
			t.Errorf("typed iteration: pictures=%d charts=%d; want 1/1", pictureCount, chartCount)
		}
	})
}
