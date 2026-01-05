//go:build windows

package expression

import (
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

// setupExcel a helper function to create an Excel instance for testing.
// It returns *ole.IDispatch and a cleanup function.
func setupExcel(t *testing.T) (*ole.IDispatch, func()) {
	t.Helper()
	
	disp, cleanup, err := excel.New()
	if err != nil {
		t.Skip("Excel not installed or failed to create:", err)
	}

	return disp, cleanup
}

func TestGet_Property(t *testing.T) {
	disp, cleanup := setupExcel(t)
	defer cleanup()

	// Test with *sugar.Chain (create a new chain from disp)
	chain := sugar.From(disp)
	version, err := Get(chain, "Version")
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
	// Since Get calls Value() which releases the chain created internally from disp,
	// but From(disp) does not own disp, disp is safe.
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
}

func TestGet_MethodCall(t *testing.T) {
	disp, cleanup := setupExcel(t)
	defer cleanup()

	// Add a new workbook
	// Get returns *IDispatch now? No, Get returns value. Store returns *IDispatch.
	// "Workbooks" returns an object, so we MUST use Store.
	workbooks, err := Store(disp, "Workbooks")
	if err != nil {
		t.Fatalf("Failed to get Workbooks object: %v", err)
	}
	defer workbooks.Release()

	// Call Add. Add returns a Workbook object.
	// Since we are discarding the result (underscore), we can use Get or Store?
	// If we use Get, it will error because it returns an object.
	// We should probably use Store and release it, or just ensure no error if we don't care about result?
	// But Get errors on objects.
	// So we must use Store.
	wb, err := Store(workbooks, "Add()")
	if err != nil {
		t.Fatalf("Failed to call Workbooks.Add(): %v", err)
	}
	wb.Release()

	// Check that a workbook was added
	count, err := Get(disp, "Workbooks.Count")
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
}

func TestGet_MethodCallWithArgs(t *testing.T) {
	disp, cleanup := setupExcel(t)
	defer cleanup()

	// Add a workbook to ensure we have a context where evaluation works properly
	wb, err := Store(disp, "Workbooks.Add()")
	if err != nil {
		t.Fatalf("Failed to add workbook: %v", err)
	}
	wb.Release()

	// Use Application.Evaluate("A1") to test passing string arguments to a method.
	// Evaluate returns a Range object, so we must use Store.
	rng, err := Store(disp, "Evaluate('A1')")
	if err != nil {
		t.Fatalf("Failed to call Evaluate('A1'): %v", err)
	}
	rng.Release()
}

func TestPut_Property(t *testing.T) {
	disp, cleanup := setupExcel(t)
	defer cleanup()

	// Add a workbook and select a cell
	wb, err := Store(disp, "Workbooks.Add()")
	if err != nil {
		t.Fatalf("Failed to add workbook: %v", err)
	}
	wb.Release()

	// Set value of A1 using Put
	err = Put(disp, "ActiveCell.Value", "Hello")
	if err != nil {
		t.Fatalf("Put failed: %v", err)
	}

	// Verify the value was set
	value, err := Get(disp, "ActiveCell.Value")
	if err != nil {
		t.Fatalf("Failed to get ActiveCell.Value after Put: %v", err)
	}
	if value != "Hello" {
		t.Errorf("Expected ActiveCell.Value to be 'Hello', got '%v'", value)
	}
}

func TestErrorHandling(t *testing.T) {
	disp, cleanup := setupExcel(t)
	defer cleanup()

	// Test invalid property access
	_, err := Get(disp, "InvalidProperty")
	if err == nil {
		t.Error("Expected error for invalid property, got nil")
	}

	// Test invalid method call
	_, err = Get(disp, "InvalidMethod()")
	if err == nil {
		t.Error("Expected error for invalid method, got nil")
	}

	// Test invalid expression for Put
	err = Put(disp, "Workbooks.Add()", "some-value")
	if err == nil {
		t.Error("Expected error for using a method call in Put, got nil")
	}
}
