# sugar: Sweeten your Windows automation.

> **Warning:** This project is currently in the **Alpha stage (v0.x.x)**. APIs are subject to change and breaking changes may occur until the v1.0.0 release.

`sugar` is a flexible and safe Go library for Component Object Model (COM) automation on Windows. Built on top of the powerful `go-ole` library, it introduces **Immutability** and the **Arena (Context) pattern** to help you write clean code without worrying about resource leaks.

## Key Features

- **Standard Execution Pattern (`Do`/`Go`):** Automatically handles thread locking (`LockOSThread`) and COM initialization (`CoInitialize`).
- **Immutable Chain:** All operations (`Get`, `Call`, etc.) return a new Chain instance, preventing side effects on original objects.
- **Automatic Resource Management (Arena):** All COM objects created within a context are automatically released in reverse order when the block completes.
- **Standard `context.Context` Integration:** Leverage Go's standard context features for cancellation, timeouts, and value passing.
- **Expression-Based Automation:** Navigate complex object hierarchies using a single string expression.

## Installation

```sh
go get github.com/xll-gen/sugar
```

## Quick Start

A simple example using `sugar.Do` to launch Excel and add a workbook. Resource cleanup is handled automatically.

```go
package main

import (
	"log"
	"github.com/xll-gen/sugar"
)

func main() {
	// sugar.Do guarantees COM initialization and automatic resource cleanup.
	err := sugar.Do(func(ctx *sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			return err
		}
		
		// Schedule Excel to quit
		defer excel.Call("Quit")

		// Method chaining (Immutable pattern)
		excel.Put("Visible", true).
			Get("Workbooks").
			Call("Add")
            
		// When the function returns, excel, workbooks, and the new workbook 
		// objects are all automatically released.
		return nil
	})

	if err != nil {
		log.Fatalf("Automation failed: %v", err)
	}
}
```

## Core Concepts

### 1. Standard Execution (`sugar.Do` & `sugar.Go`)

COM is sensitive to the execution thread. `sugar` provides safe entry points to manage this.

- **`sugar.Do`**: Locks the current goroutine to an OS thread and executes synchronously.
- **`sugar.Go`**: Starts a new goroutine (new OS thread) and independently initializes the COM environment for asynchronous work.

### 2. Immutable Chain

Methods like `Get`, `Call`, and `ForEach` always return a **NEW `Chain` instance**.

```go
workbooks := excel.Get("Workbooks") // 'excel' still points to Application
wb := workbooks.Call("Add")         // 'workbooks' still points to the Workbooks collection
```

### 3. Iteration with `ForEach`

You can iterate over COM collections (like Workbooks, Sheets, or Ranges) using the `ForEach` method. Each item is provided as a new `Chain` instance.

```go
sugar.Do(func(ctx *sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    workbooks := excel.Get("Workbooks")

    // Iterate through all open workbooks
    workbooks.ForEach(func(wb *sugar.Chain) error {
        name, _ := wb.Get("Name").Value()
        fmt.Printf("Workbook: %v\n", name)
        return nil // Return nil to continue, sugar.ErrBreak to stop
    })
    return nil
})
```

### 4. Arena Context

The `sugar.Context` acts as a resource collector (Arena). Any object created via `ctx.Create`, `ctx.From`, or derived from a chain is automatically registered with that context and cleaned up when the `Do` block ends.

**Manual `Release()` calls are no longer necessary.**

### 5. Nested Scopes

If you want to clean up resources for a specific part of a function early, use `ctx.Do` to create a nested arena.

```go
sugar.Do(func(ctx *sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    
    ctx.Do(func(innerCtx *sugar.Context) error {
        // Objects created in this block are released immediately when it ends.
        wb := excel.Get("Workbooks").Call("Add")
        return nil
    }) 
    // 'wb' is released here, while 'excel' remains valid.
    return nil
})
```

## Expression-Based Automation (Subpackage)

The `expression` package allows you to manipulate complex hierarchies with a single line of code. Intermediate objects created during evaluation are automatically managed.

```go
import "github.com/xll-gen/sugar/expression"

sugar.Do(func(ctx *sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    
    // Create a new workbook (automatically tracked by ctx)
    excel.Get("Workbooks").Call("Add")

	// Set complex paths at once
    expression.Put(excel, "ActiveSheet.Range('A1').Value", "Hello Sugar!")
    
    // Read values
    val, _ := expression.Get(excel, "ActiveSheet.Range('A1').Value")
    fmt.Println(val)
    return nil
})
```

## Considerations

- **Windows Only:** This library depends on Windows COM technology and only works on Windows OS.
- **Object Sharing Between Threads:** Sharing raw `IDispatch` pointers between threads (goroutines) without proper marshaling is dangerous. We recommend creating independent objects in each goroutine using `sugar.Go`.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
