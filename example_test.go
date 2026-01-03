//go:build windows

package sugar

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// excelTestSetup is a helper to reduce boilerplate in examples.
func excelTestSetup() (*ole.IDispatch, func()) {
	ole.CoInitialize(0)

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		// Excel might not be installed. This is not a test failure.
		// So we just log it and return.
		log.Println("Failed to create Excel object:", err)
		ole.CoUninitialize()
		return nil, nil
	}

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err) // This should not happen if CreateObject succeeded.
	}
	unknown.Release()

	// Teardown function
	cleanup := func() {
		From(excel).Call("Quit").Release()
		excel.Release()
		ole.CoUninitialize()
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

	err := From(excel).
		Put("Visible", true).
		Get("Workbooks").
		Call("Add").
		Release()

	if err != nil {
		log.Fatal(err)
	}

	// For demonstration, we can retrieve a value to confirm.
	name, err := From(excel).Get("ActiveWorkbook").Get("Name").Value()
	if err == nil {
		fmt.Printf("Newly created workbook is named: %s", name)
	}
	// Output: Newly created workbook is named: Book1
}

// This example demonstrates how to get a value from the chain.
func ExampleChain_Value() {
	excel, cleanup := excelTestSetup()
	if excel == nil {
		return
	}
	defer cleanup()

	From(excel).Get("Workbooks").Call("Add").Release() // Ensure there's a workbook

	val, err := From(excel).
		Get("ActiveSheet").
		Get("Name").
		Value()

	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Active sheet name: %v", val)
	// Output: Active sheet name: Sheet1
}

// This example demonstrates using the Store method to save and reuse
// an intermediate object.
func ExampleChain_Store() {
	excel, cleanup := excelTestSetup()
	if excel == nil {
		return
	}
	defer cleanup()

	From(excel).Get("Workbooks").Call("Add").Release()

	sheet, err := From(excel).
		Get("ActiveSheet").
		Store()
	if err != nil {
		log.Fatal(err)
	}
	defer sheet.Release()

	// Now 'sheet' can be used multiple times, ensuring each chain is terminated.
	From(sheet).Get("Cells", 1, 1).Put("Value", "Hello").Release()
	From(sheet).Get("Cells", 1, 2).Put("Value", "World").Release()

	val, _ := From(sheet).Get("Cells", 1, 1).Get("Value").Value()
	fmt.Printf("Cell A1 contains: %v", val)
	// Output: Cell A1 contains: Hello
}

// This example demonstrates using auto-release mode.
func ExampleChain_AutoRelease() {
	excel, cleanup := excelTestSetup()
	if excel == nil {
		return
	}
	defer cleanup()

	err := From(excel).
		AutoRelease().
		Put("Visible", true).
		Get("Workbooks").
		Call("Add").
		Err()

	if err != nil {
		log.Fatal(err)
	}

	// No need to call Release(), resources are managed by the GC.
	fmt.Println("AutoRelease chain executed without error.")
	// Output: AutoRelease chain executed without error.
}
