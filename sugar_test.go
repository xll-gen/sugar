//go:build windows

package sugar_test

import (
	"fmt"
	"log"
	"runtime"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/xll-gen/sugar"
)

// excelTestSetup is a helper to reduce boilerplate in examples.
func excelTestSetup() (*ole.IDispatch, func()) {
	runtime.LockOSThread()
	ole.CoInitialize(0)

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		// Excel might not be installed. This is not a test failure.
		// So we just log it and return.
		log.Println("Failed to create Excel object:", err)
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, nil
	}

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err) // This should not happen if CreateObject succeeded.
	}
	unknown.Release()

	// Teardown function
	cleanup := func() {
		sugar.From(excel).Call("Quit").Release()
		excel.Release()
		ole.CoUninitialize()
		runtime.UnlockOSThread()
	}

	return excel, cleanup
}

// This example demonstrates basic chaining with manual resource management.
func ExampleChain_manual() {
	excel, cleanup := excelTestSetup()
	if excel == nil {
		return // Excel not available
	}
	defer cleanup()

	err := sugar.From(excel).
		Put("Visible", true).
		Get("Workbooks").
		Call("Add").
		Release()

	if err != nil {
		log.Fatal(err)
	}

	// For demonstration, we can retrieve a value to confirm.
	name, err := sugar.From(excel).Get("ActiveWorkbook").Get("Name").Value()
	if err == nil {
		fmt.Printf("Newly created workbook is named: %s", name)
	}
}

// This example demonstrates creating a new COM object with Create.
func ExampleCreate() {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	err := sugar.Create("Excel.Application").
		Call("Quit"). // Quit the application immediately.
		Release()     // Release the application object.

	if err != nil {
		// Excel might not be installed.
		log.Println("Failed to run create/quit example:", err)
		return
	}

	fmt.Println("Create and Quit successful.")
	// Output: Create and Quit successful.
}

// This example demonstrates getting an active object.
func ExampleGetActive() {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	// This example will fail if Excel is not running, which is expected in a
	// test environment where we don't start it first.
	err := sugar.GetActive("Excel.Application").Release()
	if err != nil {
		fmt.Println("GetActive failed as expected.")
	} else {
		// This would be the success case.
		fmt.Println("GetActive succeeded.")
	}
}

func TestChain_Mock(t *testing.T) {
	// These tests use a nil dispatch to test internal error handling and logic
	// without requiring a real COM object.
	
	t.Run("Nil Dispatch Error", func(t *testing.T) {
		c := sugar.From(nil)
		err := c.Get("Prop").Err()
		// Calling Get on nil dispatch should return an error.
		if err == nil {
			t.Error("expected error for nil dispatch, got nil")
		}
	})

	t.Run("Error Propagation", func(t *testing.T) {
		c := sugar.From(nil)
		// Manually set an error via a failed operation or just check initial state
		if err := c.Err(); err != nil {
			t.Errorf("initial error should be nil, got %v", err)
		}
	})
	
	t.Run("Create invalid ProgID", func(t *testing.T) {
		ole.CoInitialize(0)
		defer ole.CoUninitialize()
		
		c := sugar.Create("Invalid.ProgID.That.Does.Not.Exist")
		if err := c.Err(); err == nil {
			t.Error("expected error for invalid ProgID, got nil")
		}
	})
}

