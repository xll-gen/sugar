//go:build windows

package sugar_test

import (
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

func initExcel(t *testing.T) (*ole.IDispatch, func()) {
	t.Helper()
	
	disp, cleanup, err := excel.New()
	if err != nil {
		t.Skip("Excel not installed or failed to create:", err)
	}

	// Set Visible = false
	if err := sugar.From(disp).Put("Visible", false).Release(); err != nil {
		cleanup()
		t.Fatalf("failed to set Visible=false: %v", err)
	}

	return disp, cleanup
}

func TestChain_Properties(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)
	
	// Test Put and Get
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
}

func TestChain_Methods(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)

	// Test nested Get and Call
	// Workbooks returns IDispatch, so we MUST Store it or just Call on the chain without Value.
	// Call("Add") returns a Workbook object.
	wb, err := excel.Get("Workbooks").Call("Add").Store()
	if err != nil {
		t.Errorf("failed to add workbook: %v", err)
	} else {
		wb.Release()
	}

	// Re-acquire Workbooks to get Count
	// Note: We must create a new chain because the previous 'excel' chain was consumed/modified by Store().
	excel = sugar.From(disp)
	wbs, err := excel.Get("Workbooks").Store()
	if err != nil {
		t.Fatalf("failed to get Workbooks: %v", err)
	}
	defer wbs.Release()

	count, err := sugar.From(wbs).Get("Count").Value()
	if err != nil {
		t.Errorf("failed to get workbooks count: %v", err)
	}
	
	var countInt int
	switch v := count.(type) {
	case int32:
		countInt = int(v)
	case int64:
		countInt = int(v)
	case int:
		countInt = v
	default:
		t.Fatalf("unexpected type for count: %T", count)
	}

	if countInt < 1 {
		t.Errorf("expected at least 1 workbook, got %d", countInt)
	}
}

func TestChain_Store(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)

	// Test Store
	wb, err := excel.Get("Workbooks").Call("Add").Store()
	if err != nil {
		t.Fatalf("failed to store workbook: %v", err)
	}
	defer wb.Release()

	sheet, err := sugar.From(wb).Get("ActiveSheet").Store()
	if err != nil {
		t.Fatalf("failed to store sheet: %v", err)
	}
	defer sheet.Release()

	// Use stored sheet
	err = sugar.From(sheet).Get("Cells", 1, 1).Put("Value", "Sugar").Release()
	if err != nil {
		t.Errorf("failed to set cell value: %v", err)
	}

	val, err := sugar.From(sheet).Get("Cells", 1, 1).Get("Value").Value()
	if err != nil {
		t.Errorf("failed to get cell value: %v", err)
	}
	if val != "Sugar" {
		t.Errorf("expected 'Sugar', got %v", val)
	}
}

func TestChain_Errors(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)

	// Test error propagation in chain
	err := excel.Get("NonExistentProperty").Get("Another").Release()
	if err == nil {
		t.Error("expected error for non-existent property, got nil")
	}
}

func TestChain_ValueRestrictions(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)

	_, err := excel.Get("Workbooks").Value()
	if err == nil {
		t.Error("expected error when calling Value() on an IDispatch result, got nil")
	}
	expectedErr := "value cannot return IDispatch, use Store() instead"
	if err != nil && err.Error() != expectedErr {
		t.Errorf("expected error %q, got %q", expectedErr, err.Error())
	}
}

func TestChain_IsDispatch(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	excel := sugar.From(disp)

	isDisp := excel.Get("Workbooks").IsDispatch()
	if !isDisp {
		t.Error("expected IsDispatch() to be true for Workbooks")
	}
	
	excel = sugar.From(disp)
		isDisp = excel.Get("Visible").IsDispatch()
		if isDisp {
			t.Error("expected IsDispatch() to be false for Visible (bool)")
		}
	}
	
	func TestChain_ForEach(t *testing.T) {
		disp, cleanup := initExcel(t)
		defer cleanup()
	
		// Get Workbooks collection safely
		wbs, err := sugar.From(disp).Get("Workbooks").Store()
		if err != nil {
			t.Fatalf("failed to get Workbooks: %v", err)
		}
		defer wbs.Release()
	
		// Add 2 workbooks
		for i := 0; i < 2; i++ {
			// Call Add on Workbooks collection
			wb, err := sugar.From(wbs).Call("Add").Store()
			if err != nil {
				t.Fatalf("failed to add workbook %d: %v", i, err)
			}
			wb.Release()
		}
	
		// 1. Count workbooks
		count := 0
		err = sugar.From(wbs).ForEach(func(item *sugar.Chain) bool {
			count++
			return true
		}).Err()
	
		if err != nil {
			t.Fatalf("ForEach failed: %v", err)
		}
		// Note: Excel starts with 0 or 1 workbook depending on settings/version, 
		// plus we added 2. So count should be at least 2.
		if count < 2 {
			t.Errorf("expected at least 2 workbooks, counted %d", count)
		}
	
		// 2. Test early exit
		count = 0
		err = sugar.From(wbs).ForEach(func(item *sugar.Chain) bool {
			count++
			return false // Stop after first item
		}).Err()
	
		if err != nil {
			t.Fatalf("ForEach with early exit failed: %v", err)
		}
		if count != 1 {
			t.Errorf("expected count to be 1 after early exit, got %d", count)
		}
	}
	