//go:build windows

package sugar_test

import (
	"testing"

	"github.com/xll-gen/sugar"
)

// setupExcel creates an Excel instance using the context.
// It sets up a deferred Quit to ensure the process terminates.
func setupExcel(t *testing.T, ctx *sugar.Context) *sugar.Chain {
	excel := ctx.Create("Excel.Application")
	if err := excel.Err(); err != nil {
		t.Logf("CRITICAL: excel.Create failed: %v", err)
		t.Skip("Excel not installed or failed to create:", err)
		return nil
	}
	
	// Ensure Excel quits. We use defer within the Do function scope.
	// Since this helper is called inside Do, the caller should defer the Quit 
	// or we rely on the caller to do it. 
	// But defer here executes when setupExcel returns, which is TOO EARLY.
	// So we can't defer Quit here.
	// We return the chain and let the caller handle Quit, OR we attach a cleanup to t (but t.Cleanup runs after Do returns?).
	// t.Cleanup runs after the test function returns. If Do blocks, t.Cleanup runs after Do.
	// So t.Cleanup is a safe place to Quit?
	// No, because by then ctx might be released (Do returns -> ctx.Release -> t.Cleanup).
	// If ctx releases the object, we can't call Quit on it properly?
	// Actually, if we release the object, we can't call methods.
	// So Quit MUST happen before ctx.Release.
	// Therefore, Quit must be deferred inside the Do callback.
	
	// We'll require callers to defer Quit.
	
	excel.Put("Visible", false)
	return excel
}

func TestChain_Properties(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		
		// Ensure Quit
		defer excel.Put("DisplayAlerts", false).Call("Quit")

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
	})
}

func TestChain_Methods(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Test nested Get and Call
		// Immutable chain: calls return new chains, auto-tracked by ctx.
		wb := excel.Get("Workbooks").Call("Add")
		if err := wb.Err(); err != nil {
			t.Errorf("failed to add workbook: %v", err)
		}

		// Re-acquire Workbooks to get Count
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
	})
}

func TestChain_Store(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Test Store (manual management mixed with context)
		// Store removes the object from Context management (if we implemented Detach, 
		// but currently Store just returns a new reference. Wait, does Store detach?)
		// Checking sugar.go: Store returns AddRef'd disp. It doesn't detach the Chain from Context.
		// So ctx will release its copy, user must release their copy.
		
		wbDisp, err := excel.Get("Workbooks").Call("Add").Store()
		if err != nil {
			t.Fatalf("failed to store workbook: %v", err)
		}
		defer wbDisp.Release()

		// Create a chain from stored disp using context
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
	})
}

func TestChain_Errors(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Test error propagation in chain
		err := excel.Get("NonExistentProperty").Get("Another").Err()
		if err == nil {
			t.Error("expected error for non-existent property, got nil")
		}
	})
}

func TestChain_ValueRestrictions(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		_, err := excel.Get("Workbooks").Value()
		if err == nil {
			t.Error("expected error when calling Value() on an IDispatch result, got nil")
		}
		expectedErr := "result is IDispatch, use Store"
		if err != nil && err.Error() != expectedErr {
			t.Errorf("expected error %q, got %q", expectedErr, err.Error())
		}
	})
}

func TestChain_IsDispatch(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		isDisp := excel.Get("Workbooks").IsDispatch()
		if !isDisp {
			t.Error("expected IsDispatch() to be true for Workbooks")
		}
		
		// Immutable chain: excel.Get("Visible") works on App object
		isDisp = excel.Get("Visible").IsDispatch()
		if isDisp {
			t.Error("expected IsDispatch() to be false for Visible (bool)")
		}
	})
}

func TestChain_ForEach(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		wbs := excel.Get("Workbooks")
		
		// Add 2 workbooks
		for i := 0; i < 2; i++ {
			wbs.Call("Add")
		}

		// 1. Count workbooks
		count := 0
		err := wbs.ForEach(func(item *sugar.Chain) bool {
			count++
			return true
		}).Err()

		if err != nil {
			t.Fatalf("ForEach failed: %v", err)
		}
		if count < 2 {
			t.Errorf("expected at least 2 workbooks, counted %d", count)
		}

		// 2. Test early exit
		count = 0
		err = wbs.ForEach(func(item *sugar.Chain) bool {
			count++
			return false // Stop after first item
		}).Err()

		if err != nil {
			t.Fatalf("ForEach with early exit failed: %v", err)
		}
		if count != 1 {
			t.Errorf("expected count to be 1 after early exit, got %d", count)
		}
	})
}

func TestChain_Fork(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// 1. Create a workbook and Fork (which is just explicit clone now)
		wbChain := excel.Get("Workbooks").Call("Add").Fork()
		
		if err := wbChain.Err(); err != nil {
			t.Fatalf("Fork failed: %v", err)
		}
		
		// 2. Use the forked chain
		err := wbChain.Put("Saved", true).Err()
		if err != nil {
			t.Errorf("failed to use forked chain: %v", err)
		}
		
		// 3. No manual Release needed (handled by ctx)
	})
}