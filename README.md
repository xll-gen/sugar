# sugar: A fluent, chainable API for COM automation in Go

`sugar` is a Go library that provides a fluent, chainable API for Component Object Model (COM) automation on Windows. It acts as a syntactic sugar layer on top of the powerful `go-ole` library, simplifying COM interactions and making your code more readable and expressive.

This library is designed to manage the complexity of COM object lifecycles, offering both manual and automatic resource management to prevent memory leaks.

## Features

- **Fluent, Chainable Interface:** Write clean and readable automation scripts.
- **Simplified Object Lifecycle:** Clear and predictable resource management.
- **Easy Object Creation:** Functions to create new COM objects (`Create`), attach to existing ones (`GetActive`), or wrap your own (`From`).
- **Error Handling:** Errors are captured and handled at the end of the chain.

## Prerequisites

- **Windows Only:** This library depends on `go-ole` and the underlying COM technology, which is specific to the Windows operating system.
- **Go:** Version 1.18 or higher.

## Installation

```sh
go get github.com/xll-gen/sugar
```

## Quick Start

Here's a simple example of how to launch Microsoft Excel, make it visible, and add a new workbook.

```go
package main

import (
	"log"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

func main() {
	// Initialize the COM library for the current goroutine.
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	// Create a new Excel application instance.
	err := sugar.Create("Excel.Application").
		Put("Visible", true).  // Set the 'Visible' property to true.
		Get("Workbooks").       // Get the Workbooks collection.
		Call("Add").            // Call the 'Add' method to create a new workbook.
		Release()              // Release all COM objects created in the chain.

	if err != nil {
		log.Fatalf("Excel automation failed: %v", err)
	}
}
```

## Core Concepts

### 1. The Chain

`sugar` works by creating a "chain" of operations. The chain starts with an initial object and continues with a series of `Get`, `Call`, or `Put` actions.

- `Get(property, ...args)`: Retrieves a property, which can be another object.
- `Call(method, ...args)`: Executes a method.
- `Put(property, ...args)`: Sets the value of a property.

Each of these methods returns the chain, allowing you to append the next operation.

### 2. Terminal Methods

A chain must always end with a **terminal method**. This is critical for releasing COM objects and handling errors. If you forget to call a terminal method, you will leak resources.

- `Release()`: Releases all COM objects acquired during the chain and returns any error that occurred. This is the most common terminal method for chains that do not return a value.
- `Value()`: Returns the value from the last `Get` or `Call` operation as an `interface{}`. It also releases all resources.
- `Err()`: Returns the first error that occurred during the chain. It also releases all resources.

### 3. Creating a Chain

There are three ways to start a chain:

- `sugar.Create(progID string)`: Creates a new COM object (e.g., `"Excel.Application"`). The chain takes ownership of this object and is responsible for releasing it.
- `sugar.GetActive(progID string)`: Attaches to a running instance of a COM object. The chain takes ownership and will release it.
- `sugar.From(disp *ole.IDispatch)`: Wraps an existing `*ole.IDispatch` object that you manage yourself. The chain does **not** take ownership of the initial object; you are responsible for releasing it.

### 4. Resource Management

All objects created in a chain are tracked and released only when a terminal method (`Release`, `Value`, `Err`) is called. This gives you deterministic and immediate cleanup.

```go
// All intermediate objects (Workbooks, new Workbook) are released at the end.
err := sugar.Create("Excel.Application").
    Get("Workbooks").
    Call("Add").
    Release() // Releases Excel, the Workbooks object, and the new Workbook object.
```

## Advanced Usage

### Retrieving a Value

Use the `Value()` terminal method to get the result of the last operation.

```go
// Get the name of the active worksheet.
val, err := sugar.Create("Excel.Application").
    Get("Workbooks").
    Call("Add").
    Get("ActiveSheet").
    Get("Name").
    Value() // val will contain "Sheet1"

if err != nil {
    log.Fatal(err)
}
fmt.Printf("Active sheet name: %v", val)
```

### Storing and Reusing an Object

Sometimes you need to get an object from a chain and use it multiple times. The `Store()` method allows you to do this.

`Store()` is a terminal method that ends the chain and transfers ownership of the *current* COM object to you. You are then responsible for calling `Release()` on it when you are done. Because it is a terminal method, you do not need to call `Release()` or `Err()` after it.

A common pattern is to create and store the main application object first, and then use it to build new, independent chains.

```go
// Create the Excel application and store the object.
excel, err := sugar.Create("Excel.Application").Store()
if err != nil {
    log.Fatalf("Failed to create Excel object: %v", err)
}
// IMPORTANT: You are now responsible for releasing the object.
defer excel.Release()

// Now you can use 'excel' to start multiple chains.
// Make Excel visible and add a workbook.
err = sugar.From(excel).
    Put("Visible", true).
    Get("Workbooks").
    Call("Add").
    Release() // Terminate this chain.

if err != nil {
    log.Fatal(err)
}

// Get the active sheet and write some values to it.
sheet, err := sugar.From(excel).Get("ActiveSheet").Store()
if err != nil {
    log.Fatal(err)
}
defer sheet.Release() // You also own this object now.

sugar.From(sheet).Get("Cells", 1, 1).Put("Value", "Hello").Release()
sugar.From(sheet).Get("Cells", 1, 2).Put("Value", "World").Release()
```

## Expression-Based Automation

The `expression` subpackage provides a powerful way to interact with COM objects using simple string expressions, which is ideal for simplifying complex or deeply nested operations.

Here is a complete example of how to use it with Excel. This code will create a new Excel instance, add a workbook, write a value to a cell using `expression.Put`, and read it back using `expression.Get`.

```go
package main

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/expression"
)

func main() {
	// Initialize COM for the current goroutine.
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	// Create a new Excel application instance.
	excel, err := sugar.Create("Excel.Application").Store()
	if err != nil {
		log.Fatalf("Failed to create Excel object: %v", err)
	}
	// Ensure the Excel process is quit when we are done.
	defer excel.Release()
	defer sugar.From(excel).Call("Quit").Release()

	// Ensure there is a workbook.
	sugar.From(excel).Get("Workbooks").Call("Add").Release()

	// Use expression.Put to set a cell's value.
	err = expression.Put(excel, "ActiveSheet.Cells(1, 1).Value", "Hello from expression!")
	if err != nil {
		log.Fatal(err)
	}

	// Use expression.Get to retrieve the value.
	val, err := expression.Get(excel, "ActiveSheet.Cells(1, 1).Value")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Cell A1 contains: %v", val)
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
