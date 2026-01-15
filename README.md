# sugar: Sweeten your Windows automation.

> **Warning:** This project is currently in the **Alpha stage (v0.x.x)**. APIs are subject to change and breaking changes may occur until the v1.0.0 release.

`sugar` is a flexible and safe Go library for Component Object Model (COM) automation on Windows. Built on top of the powerful `go-ole` library, it introduces **Immutability** and the **Arena (Context) pattern** to help you write clean code without worrying about resource leaks.

## Key Features

- **Standard Execution Pattern (`Do`/`Go`):** Automatically handles thread locking (`LockOSThread`) and COM initialization (`CoInitialize`).
- **Immutable Chain:** All operations (`Get`, `Call`, etc.) return a new `Chain` (Interface) instance, preventing side effects on original objects.
- **Automatic Resource Management (Arena):** All COM objects created within a context are automatically released in reverse order when the block completes.
- **Standard `context.Context` Integration:** Leverage Go's standard context features for cancellation, timeouts, and value passing.
- **Expression-Based Automation:** Navigate complex object hierarchies using a single string expression.
- **Application Specific Subpackages:** Use type-safe wrappers for popular applications like Excel.

## Installation

```sh
go get -u github.com/xll-gen/sugar
```

## Quick Start (Generic)

A simple example using `sugar.Do` to launch Excel. Resource cleanup is handled automatically.

```go
package main

import (
	"log"
	"github.com/xll-gen/sugar"
)

func main() {
	// sugar.Do guarantees COM initialization and automatic resource cleanup.
	err := sugar.Do(func(ctx sugar.Context) error {
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
            
		return nil
	})

	if err != nil {
		log.Fatalf("Automation failed: %v", err)
	}
}
```

## Excel Subpackage (Type-Safe)

For common applications, `sugar` provides subpackages with friendly methods.

```go
import "github.com/xll-gen/sugar/excel"

sugar.Do(func(ctx sugar.Context) error {
    app := excel.NewApplication(ctx)
    defer app.Quit()

    app.Put("Visible", true)
    
    wb := app.Workbooks().Add()
    sheet := wb.ActiveSheet()
    
    // Type-safe Range manipulation
    sheet.Range("A1").SetValue("Hello from Sugar!")
    
    return nil
})
```

## Core Concepts

### 1. Standard Execution (`sugar.Do` & `sugar.Go`)

COM is sensitive to the execution thread. `sugar` provides safe entry points to manage this.

- **`sugar.Do`**: Locks the current goroutine to an OS thread and executes synchronously.
- **`sugar.Go`**: Starts a new goroutine (new OS thread) and independently initializes the COM environment for asynchronous work.

### 2. Immutable Chain

Methods like `Get`, `Call`, and `ForEach` always return a **NEW `Chain` instance**. `Chain` is now an **interface**, allowing for custom wrappers like the `excel` package.

```go
workbooks := excel.Get("Workbooks") // 'excel' still points to Application
wb := workbooks.Call("Add")         // 'workbooks' still points to the Workbooks collection
```

### 3. Iteration with `ForEach`

You can iterate over COM collections using the `ForEach` method. Each item is provided as a `Chain` instance. Returning `sugar.ErrForEachBreak` stops the iteration and the error is recorded in the Chain.

```go
sugar.Do(func(ctx sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    workbooks := excel.Get("Workbooks")

    // Iterate through all open workbooks
    err := workbooks.ForEach(func(wb sugar.Chain) error {
        name, _ := wb.Get("Name").Value()
        fmt.Printf("Workbook: %v\n", name)
        
        // Stop after first item if needed
        return sugar.ErrForEachBreak
    }).Err()

    if errors.Is(err, sugar.ErrForEachBreak) {
        // Handled break
    }
    return nil
})
```

### 4. Arena Context

The `sugar.Context` acts as a resource collector (Arena). Any object created via `ctx.Create`, `ctx.From`, or derived from a chain is automatically registered with that context and cleaned up when the `Do` block ends.

**Manual `Release()` calls are no longer necessary.**

### 5. Nested Scopes

Use `ctx.Do` to create a nested arena for early resource cleanup.

```go
sugar.Do(func(ctx sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    
    ctx.Do(func(innerCtx sugar.Context) error {
        // Objects created in this block are released immediately when it ends.
        wb := excel.Get("Workbooks").Call("Add")
        return nil
    }) 
    // 'wb' is released here, while 'excel' remains valid.
    return nil
})
```

## Expression-Based Automation (Subpackage)

The `expression` package allows you to manipulate complex hierarchies with a single line of code.

```go
import "github.com/xll-gen/sugar/expression"

sugar.Do(func(ctx sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    
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