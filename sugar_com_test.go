//go:build windows

package sugar_test

import (
	"errors"
	"testing"

	"github.com/xll-gen/sugar"
)

func setupExcel(t *testing.T, ctx sugar.Context) sugar.Chain {
	excel := ctx.Create("Excel.Application")
	if err := excel.Err(); err != nil {
		t.Logf("CRITICAL: excel.Create failed: %v", err)
		t.Skip("Excel not installed or failed to create:", err)
		return nil
	}
	
	excel.Put("Visible", false)
	return excel
}

func TestChain_Properties(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		err := excel.Put("Visible", false).Err()
		if err != nil {
			t.Errorf("failed to set Visible: %v", err)
		}

		val, err := excel.Get("Visible").Value()
		if err != nil {
			t.Errorf("failed to get Visible: %v", err)
		}
		if visible, ok := val.(bool); !ok || visible != false {
			t.Errorf("expected Visible to be false, got %v", val)
		}
		return nil
	})
}

func TestChain_Methods(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wb := excel.Get("Workbooks").Call("Add")
		if err := wb.Err(); err != nil {
			t.Errorf("failed to add workbook: %v", err)
		}

		wbs := excel.Get("Workbooks")
		if err := wbs.Err(); err != nil {
			t.Fatalf("failed to get Workbooks: %v", err)
		}

		count, err := wbs.Get("Count").Value()
		if err != nil {
			t.Errorf("failed to get workbooks count: %v", err)
		}
		
		var countInt int
		switch v := count.(type) {
		case int32: countInt = int(v)
		case int64: countInt = int(v)
		case int:   countInt = v
		default:
			t.Fatalf("unexpected type for count: %T", count)
		}

		if countInt < 1 {
			t.Errorf("expected at least 1 workbook, got %d", countInt)
		}
		return nil
	})
}

func TestChain_Store(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wbDisp, err := excel.Get("Workbooks").Call("Add").Store()
		if err != nil {
			t.Fatalf("failed to store workbook: %v", err)
		}
		defer wbDisp.Release()

		wb := ctx.From(wbDisp)
		sheetDisp, err := wb.Get("ActiveSheet").Store()
		if err != nil {
			t.Fatalf("failed to store sheet: %v", err)
		}
		defer sheetDisp.Release()

		sheet := ctx.From(sheetDisp)
		err = sheet.Get("Cells", 1, 1).Put("Value", "Sugar").Err()
		if err != nil {
			t.Errorf("failed to set cell value: %v", err)
		}

		val, err := sheet.Get("Cells", 1, 1).Get("Value").Value()
		if err != nil {
			t.Errorf("failed to get cell value: %v", err)
		}
		if val != "Sugar" {
			t.Errorf("expected 'Sugar', got %v", val)
		}
		return nil
	})
}

func TestChain_Errors(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		err := excel.Get("NonExistentProperty").Get("Another").Err()
		if err == nil {
			t.Error("expected error for non-existent property, got nil")
		}
		return nil
	})
}

func TestChain_ValueRestrictions(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		_, err := excel.Get("Workbooks").Value()
		if err == nil {
			t.Error("expected error when calling Value() on an IDispatch result, got nil")
		}
		expectedErr := "result is IDispatch, use Store"
		if err != nil && err.Error() != expectedErr {
			t.Errorf("expected error %q, got %q", expectedErr, err.Error())
		}
		return nil
	})
}

func TestChain_IsDispatch(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		isDisp := excel.Get("Workbooks").IsDispatch()
		if !isDisp {
			t.Error("expected IsDispatch() to be true for Workbooks")
		}
		
		isDisp = excel.Get("Visible").IsDispatch()
		if isDisp {
			t.Error("expected IsDispatch() to be false for Visible (bool)")
		}
		return nil
	})
}

func TestChain_ForEach(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wbs := excel.Get("Workbooks")
		for i := 0; i < 2; i++ {
			wbs.Call("Add")
		}

		count := 0
		err := wbs.ForEach(func(item sugar.Chain) error {
			count++
			return nil
		}).Err()

		if err != nil {
			t.Fatalf("ForEach failed: %v", err)
		}
		if count < 2 {
			t.Errorf("expected at least 2 workbooks, counted %d", count)
		}

		count = 0
		err = wbs.ForEach(func(item sugar.Chain) error {
			count++
			return sugar.ErrForEachBreak
		}).Err()

		if !errors.Is(err, sugar.ErrForEachBreak) {
			t.Fatalf("expected ErrForEachBreak, got %v", err)
		}
		if count != 1 {
			t.Errorf("expected count 1 with ErrForEachBreak, got %d", count)
		}

		// Test with additional information
		err = wbs.ForEach(func(item sugar.Chain) error {
			return &sugar.ForEachBreak{Value: "captured data"}
		}).Err()

		var feBreak *sugar.ForEachBreak
		if !errors.As(err, &feBreak) {
			t.Fatalf("expected ForEachBreak, got %v", err)
		}
		if feBreak.Value != "captured data" {
			t.Errorf("expected 'captured data', got %v", feBreak.Value)
		}

		err = wbs.ForEach(func(item sugar.Chain) error {
			return errors.New("custom error")
		}).Err()

		if err == nil || err.Error() != "custom error" {
			t.Errorf("expected custom error, got %v", err)
		}
		return nil
	})
}

func TestChain_Fork(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wbChain := excel.Get("Workbooks").Call("Add").Fork()
		if err := wbChain.Err(); err != nil {
			t.Fatalf("Fork failed: %v", err)
		}
		
		err := wbChain.Put("Saved", true).Err()
		if err != nil {
			t.Errorf("failed to use forked chain: %v", err)
		}
		return nil
	})
}