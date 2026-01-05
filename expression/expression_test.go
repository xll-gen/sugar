//go:build windows

package expression

import (
	"testing"

	"github.com/xll-gen/sugar"
)

// setupExcel creates an Excel instance using the context.
// Caller must defer Quit inside Do block.
func setupExcel(t *testing.T, ctx *sugar.Context) *sugar.Chain {
	excel := ctx.Create("Excel.Application")
	if err := excel.Err(); err != nil {
		t.Skip("Excel not installed or failed to create:", err)
		return nil
	}
	return excel
}

func TestGet_Property(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Test with *sugar.Chain (using context-tracked chain)
		version, err := Get(excel, "Version")
		if err != nil {
			t.Fatalf("Get failed with *sugar.Chain: %v", err)
		}
		if version == nil {
			t.Fatal("Get returned nil for Version")
		}
		if vStr, ok := version.(string); !ok || vStr == "" {
			t.Errorf("Expected a version string, got %v", version)
		}
		t.Logf("Excel version (from *sugar.Chain): %v", version)

		// Test with *ole.IDispatch
		// Since we want to test Get(disp), we need a raw disp.
		// We can use Store() to get one, or just use reflection hacks, 
		// but standard way is to just pass a Chain if we have one.
		// To test IDispatch branch, we need to extract it.
		// But Store() detaches it from context? No, Store() returns AddRef'd copy.
		// So we can use Store() to get a disp to test with.
		
		disp, err := excel.Store()
		if err != nil {
			t.Fatalf("Failed to get raw dispatch: %v", err)
		}
		defer disp.Release() // Release our manual copy

		version, err = Get(disp, "Version")
		if err != nil {
			t.Fatalf("Get failed with *ole.IDispatch: %v", err)
		}
		if version == nil {
			t.Fatal("Get returned nil for Version (IDispatch)")
		}
		if vStr, ok := version.(string); !ok || vStr == "" {
			t.Errorf("Expected a version string from IDispatch, got %v", version)
		}
		t.Logf("Excel version (from *ole.IDispatch): %v", version)
	})
}

func TestGet_MethodCall(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Add a new workbook
		// "Workbooks" returns an object, so we MUST use Store.
		workbooks, err := Store(excel, "Workbooks")
		if err != nil {
			t.Fatalf("Failed to get Workbooks object: %v", err)
		}
		defer workbooks.Release()

		// Call Add. Add returns a Workbook object.
		// Since we are discarding the result (underscore), we can use Get or Store?
		// But Get errors on objects. So we must use Store.
		wb, err := Store(workbooks, "Add()")
		if err != nil {
			t.Fatalf("Failed to call Workbooks.Add(): %v", err)
		}
		wb.Release()

		// Check that a workbook was added
		count, err := Get(excel, "Workbooks.Count")
		if err != nil {
			t.Fatalf("Failed to get Workbooks.Count: %v", err)
		}
		
		// Convert to int
		var countInt int
		switch c := count.(type) {
		case int32:
			countInt = int(c)
		case int64:
			countInt = int(c)
		case int:
			countInt = c
		default:
			t.Logf("Warning: Count type is %T", count)
		}
		
		if countInt < 1 {
			t.Errorf("Expected at least 1 workbook, got %v", count)
		}
	})
}

func TestGet_MethodCallWithArgs(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Add a workbook to ensure we have a context where evaluation works properly
		wb, err := Store(excel, "Workbooks.Add()")
		if err != nil {
			t.Fatalf("Failed to add workbook: %v", err)
		}
		wb.Release()

		// Use Application.Evaluate("A1") to test passing string arguments to a method.
		// Evaluate returns a Range object, so we must use Store.
		rng, err := Store(excel, "Evaluate('A1')")
		if err != nil {
			t.Fatalf("Failed to call Evaluate('A1'): %v", err)
		}
		rng.Release()
	})
}

func TestPut_Property(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Add a workbook and select a cell
		wb, err := Store(excel, "Workbooks.Add()")
		if err != nil {
			t.Fatalf("Failed to add workbook: %v", err)
		}
		wb.Release()

		// Set value of A1 using Put
		err = Put(excel, "ActiveCell.Value", "Hello")
		if err != nil {
			t.Fatalf("Put failed: %v", err)
		}

		// Verify the value was set
		value, err := Get(excel, "ActiveCell.Value")
		if err != nil {
			t.Fatalf("Failed to get ActiveCell.Value after Put: %v", err)
		}
		if value != "Hello" {
			t.Errorf("Expected ActiveCell.Value to be 'Hello', got '%v'", value)
		}
	})
}

func TestErrorHandling(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		excel := setupExcel(t, ctx)
		if excel == nil { return }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		// Test invalid property access
		_, err := Get(excel, "InvalidProperty")
		if err == nil {
			t.Error("Expected error for invalid property, got nil")
		}

		// Test invalid method call
		_, err = Get(excel, "InvalidMethod()")
		if err == nil {
			t.Error("Expected error for invalid method, got nil")
		}

		// Test invalid expression for Put
		err = Put(excel, "Workbooks.Add()", "some-value")
		if err == nil {
			t.Error("Expected error for using a method call in Put, got nil")
		}
	})
}