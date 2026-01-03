# sugar: A fluent, chainable API for COM automation in Go

`sugar` is a Go library that provides a fluent, chainable API for Component Object Model (COM) automation on Windows. It acts as a syntactic sugar layer on top of the powerful `go-ole` library, simplifying COM interactions and making your code more readable and expressive.

This library is designed to manage the complexity of COM object lifecycles, offering both manual and automatic resource management to prevent memory leaks.

## Features

- **Fluent, Chainable Interface:** Write clean and readable automation scripts.
- **Simplified Object Lifecycle:** Clear and predictable resource management.
- **Flexible Resource Management:** Choose between manual control (default) or garbage-collector-based automatic cleanup.
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

`sugar` provides two modes for managing the lifecycle of COM objects created *during* a chain (i.e., from `Get` or `Call`).

#### Manual Mode (Default)

By default, all objects created in a chain are tracked and released only when a terminal method (`Release`, `Value`, `Err`) is called. This gives you deterministic and immediate cleanup.

```go
// All intermediate objects (Workbooks, new Workbook) are released at the end.
err := sugar.Create("Excel.Application").
    Get("Workbooks").
    Call("Add").
    Release() // Releases Excel, the Workbooks object, and the new Workbook object.
```

#### Automatic Mode (`AutoRelease`)

You can opt into garbage-collector-based cleanup by calling `AutoRelease()` on the chain. In this mode, you don't need to call a terminal method to release resources, but you should still use one to check for errors.

`AutoRelease` is useful for "fire-and-forget" scenarios or when object lifetimes are complex. However, it makes cleanup non-deterministic.

```go
err := sugar.Create("Excel.Application").
    AutoRelease().      // Switch to automatic mode.
    Put("Visible", true).
    Err()               // Check for errors. Resources are released by the GC.
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

Sometimes you need to get an object from the middle of a chain and use it multiple times. The `Store()` method allows you to do this.

When you use `Store()`, you take ownership of the object. You are responsible for calling `Release()` on it when you are done, regardless of whether the chain is in `AutoRelease` mode.

```go
var sheet *ole.IDispatch

err := sugar.Create("Excel.Application").
    Get("Workbooks").
    Call("Add").
    Get("ActiveSheet").
    Store(&sheet). // Store the IDispatch for the worksheet.
    Release()      // Terminate the chain.

if err != nil {
    log.Fatal(err)
}

// Now you can use 'sheet' in new chains.
defer sheet.Release() // IMPORTANT: You must release it yourself.

sugar.From(sheet).Put("Cells", 1, 1, "Hello").Release()
sugar.From(sheet).Put("Cells", 1, 2, "World").Release()
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
