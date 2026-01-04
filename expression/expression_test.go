//go:build windows

package expression

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// setupExcel a helper function to create an Excel instance for testing.
// It returns a sugar.Chain and a cleanup function.
func setupExcel(t *testing.T) (*sugar.Chain, func()) {
	t.Helper()
	ole.CoInitialize(0)
	excel, err := sugar.Create("Excel.Application")
	if err != nil {
		t.Fatalf("Failed to create Excel instance: %v", err)
	}

	// Make Excel visible for debugging if needed, but hide it in final tests.
	_, err = Get(excel, "Visible")
	if err != nil {
		t.Logf("Could not set Excel visibility: %v", err)
	}
	excel.Put("Visible", false)

	cleanup := func() {
		// Suppress errors during cleanup
		excel.Call("Quit").Release()
		ole.CoUninitialize()
	}

	return excel, cleanup
}

func TestGet_Property(t *testing.T) {
	excel, cleanup := setupExcel(t)
	defer cleanup()

	// Test with *sugar.Chain
	version, err := Get(excel, "Version")
	if err != nil {
		t.Fatalf("Get failed with *sugar.Chain: %v", err)
	}
	if version == "" {
		t.Errorf("Expected a version string, got empty string")
	}
	t.Logf("Excel version (from *sugar.Chain): %v", version)

	// Test with *ole.IDispatch
	disp, err := excel.Store()
	if err != nil {
		t.Fatalf("Failed to store IDispatch: %v", err)
	}
	defer disp.Release()

	version, err = Get(disp, "Version")
	if err != nil {
		t.Fatalf("Get failed with *ole.IDispatch: %v", err)
	}
	if version == "" {
		t.Errorf("Expected a version string from IDispatch, got empty string")
	}
	t.Logf("Excel version (from *ole.IDispatch): %v", version)
}

func TestGet_MethodCall(t *testing.T) {
	excel, cleanup := setupExcel(t)
	defer cleanup()

	// Add a new workbook
	workbooks, err := Get(excel, "Workbooks")
	if err != nil {
		t.Fatalf("Failed to get Workbooks object: %v", err)
	}

	disp, ok := workbooks.(*ole.IDispatch)
	if !ok {
		t.Fatalf("Workbooks is not an IDispatch, but %T", workbooks)
	}
	defer disp.Release()

	_, err = Get(disp, "Add()")
	if err != nil {
		t.Fatalf("Failed to call Workbooks.Add(): %v", err)
	}

	// Check that a workbook was added
	count, err := Get(excel, "Workbooks.Count")
	if err != nil {
		t.Fatalf("Failed to get Workbooks.Count: %v", err)
	}
	if count.(int32) != 1 {
		t.Errorf("Expected 1 workbook, got %d", count)
	}
}

func TestGet_MethodCallWithArgs(t *testing.T) {
	excel, cleanup := setupExcel(t)
	defer cleanup()

	// Get the Workbooks collection
	workbooks, err := excel.Get("Workbooks").Store()
	if err != nil {
		t.Fatalf("Failed to get Workbooks object: %v", err)
	}
	defer workbooks.Release()

	// Create a temporary file to open
	tempDir := t.TempDir()
	tempFilePath := filepath.Join(tempDir, "test.xlsx")
	f, err := os.Create(tempFilePath)
	if err != nil {
		t.Fatalf("Failed to create temp file: %v", err)
	}
	f.Close()

	// Open the workbook
	_, err = Get(workbooks, "Open('"+tempFilePath+"')")
	if err != nil {
		// Excel paths can be tricky. Let's try to make it absolute.
		absPath, _ := filepath.Abs(tempFilePath)
		_, err = Get(workbooks, "Open('"+absPath+"')")
	}
	if err != nil {
		t.Fatalf("Failed to call Workbooks.Open() with argument: %v", err)
	}

	// Check that the workbook is open
	activeWorkbookName, err := Get(excel, "ActiveWorkbook.Name")
	if err != nil {
		t.Fatalf("Failed to get ActiveWorkbook.Name: %v", err)
	}
	if activeWorkbookName != "test.xlsx" {
		t.Errorf("Expected active workbook name to be 'test.xlsx', got '%v'", activeWorkbookName)
	}
}

func TestPut_Property(t *testing.T) {
	excel, cleanup := setupExcel(t)
	defer cleanup()

	// Add a workbook and select a cell
	_, err := Get(excel, "Workbooks.Add()")
	if err != nil {
		t.Fatalf("Failed to add workbook: %v", err)
	}

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
}

func TestErrorHandling(t *testing.T) {
	excel, cleanup := setupExcel(t)
	defer cleanup()

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
}
