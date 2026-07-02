//go:build windows && excel_integration

// COM tests for chain argument normalization (normalizeParams):
//   - Go 2-D slices must marshal to VT_ARRAY|VT_VARIANT SAFEARRAYs so block
//     writes to Range.Value work in one COM call.
//   - Chain values must marshal to raw IDispatch so COM methods taking
//     object arguments (e.g. Worksheets.Add(Before:=...)) work.
//
// Before these were normalized, both paths hit go-ole's `panic("unknown
// type")` inside Invoke.
//
// Gated behind the excel_integration build tag (these spawn real Excel):
//
//	go test -tags=excel_integration ./...

package sugar_test

import (
	"testing"
	"time"

	// Deterministic IANA zones for the VT_DATE round-trip test below.
	_ "time/tzdata"

	"github.com/xll-gen/sugar"
)

// TestChain_PutGridValue writes a [][]interface{} block to a multi-cell
// range in a single Put and reads it back through the SAFEARRAY decoder.
func TestChain_PutGridValue(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil {
			return nil
		}
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
		rng := sheet.Get("Range", "A1:C2")

		grid := [][]interface{}{
			{"name", 1.5, true},
			{nil, -2.0, "x"},
		}
		if err := rng.Put("Value", grid).Err(); err != nil {
			t.Fatalf("Put grid value: %v", err)
		}

		got, err := sheet.Get("Range", "A1:C2").Get("Value").Value()
		if err != nil {
			t.Fatalf("read back: %v", err)
		}
		rows, ok := got.([][]interface{})
		if !ok || len(rows) != 2 || len(rows[0]) != 3 {
			t.Fatalf("expected 2x3 [][]interface{}, got %T %v", got, got)
		}
		if rows[0][0] != "name" || rows[0][1] != 1.5 || rows[0][2] != true {
			t.Errorf("row 0 mismatch: %v", rows[0])
		}
		if rows[1][0] != nil || rows[1][1] != -2.0 || rows[1][2] != "x" {
			t.Errorf("row 1 mismatch: %v", rows[1])
		}
		return nil
	})
}

// TestChain_PutGridValue_Date checks that time.Time cells survive the
// VT_DATE encode/decode round trip with wall-clock semantics.
func TestChain_PutGridValue_Date(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil {
			return nil
		}
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
		// Use an explicit DST zone instead of time.Local so the round trip is
		// zone-independent and actually exercises the VT_DATE wall-clock path
		// (America/New_York's UTC offset in 2026 differs from the 1899 epoch's).
		loc, err := time.LoadLocation("America/New_York")
		if err != nil {
			t.Skipf("America/New_York unavailable: %v", err)
		}
		want := time.Date(2026, 6, 10, 15, 30, 0, 0, loc)

		err = sheet.Get("Range", "A1:A1").
			Put("Value", [][]interface{}{{want}}).Err()
		if err != nil {
			t.Fatalf("Put date: %v", err)
		}

		got, err := sheet.Get("Range", "A1").Get("Value").Value()
		if err != nil {
			t.Fatalf("read back: %v", err)
		}
		ts, ok := got.(time.Time)
		if !ok {
			t.Fatalf("expected time.Time, got %T %v", got, got)
		}
		// Compare wall-clock fields: OLE dates carry no zone.
		if ts.Year() != want.Year() || ts.Month() != want.Month() ||
			ts.Day() != want.Day() || ts.Hour() != want.Hour() ||
			ts.Minute() != want.Minute() {
			t.Errorf("date mismatch: want %v, got %v", want, ts)
		}
		return nil
	})
}

// TestChain_PutVectorValue writes a 1-D []interface{} — Excel reads it as a
// row vector.
func TestChain_PutVectorValue(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil {
			return nil
		}
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
		err := sheet.Get("Range", "A1:C1").
			Put("Value", []interface{}{1.0, 2.0, 3.0}).Err()
		if err != nil {
			t.Fatalf("Put vector: %v", err)
		}

		got, err := sheet.Get("Range", "B1").Get("Value").Value()
		if err != nil || got != 2.0 {
			t.Errorf("B1: want 2.0, got %v err=%v", got, err)
		}
		return nil
	})
}

// TestChain_ChainAsArgument passes a Chain (a Worksheet dispatch) as a COM
// method argument: Worksheets.Add(Before:=sheet1).
func TestChain_ChainAsArgument(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil {
			return nil
		}
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wb := excel.Get("Workbooks").Call("Add")
		sheets := wb.Get("Worksheets")
		first := sheets.Get("Item", 1)

		added := sheets.Call("Add", first)
		if err := added.Err(); err != nil {
			t.Fatalf("Add(Before:=sheet): %v", err)
		}

		idx, err := added.Get("Index").Value()
		if err != nil {
			t.Fatalf("Index: %v", err)
		}
		if toInt(idx) != 1 {
			t.Errorf("new sheet should be inserted at index 1, got %v", idx)
		}
		return nil
	})
}

// TestChain_RaggedGridFails verifies ragged 2-D input surfaces as a chain
// error instead of a panic or silent corruption.
func TestChain_RaggedGridFails(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil {
			return nil
		}
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		sheet := excel.Get("Workbooks").Call("Add").Get("ActiveSheet")
		err := sheet.Get("Range", "A1:B2").
			Put("Value", [][]interface{}{{1.0, 2.0}, {3.0}}).Err()
		if err == nil {
			t.Error("expected error for ragged rows, got nil")
		}
		return nil
	})
}

func toInt(v interface{}) int {
	switch x := v.(type) {
	case int32:
		return int(x)
	case int64:
		return int(x)
	case float64:
		return int(x)
	case int:
		return x
	}
	return -1
}
