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
The `excel` package mirrors [xlwings](https://docs.xlwings.org/en/stable/api.html)
naming and behavior — see [AGENTS.md §2](./AGENTS.md) for the parity roadmap.

```go
import "github.com/xll-gen/sugar/excel"

sugar.Do(func(ctx sugar.Context) error {
    app := excel.NewApplication(ctx)
    defer app.Quit()

    // xlwings-parity boolean properties on App.
    // Setters return Application for fluent chaining; getters return a
    // sugar.Chain — call .Value() to materialize the bool.
    app.SetVisible(true).
        SetDisplayAlerts(false).
        SetScreenUpdating(false)

    wb := app.Workbooks().Add()
    sheet := wb.ActiveSheet()

    // Type-safe Range manipulation
    sheet.Range("A1").SetValue("Hello from Sugar!")

    // Re-enable screen updates before exit.
    app.SetScreenUpdating(true)
    return nil
})
```

### Excel object coverage

| xlwings Object | sugar type            | Status                                                                                                                 |
| -------------- | --------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| `App`          | `excel.Application`   | `Visible`, `DisplayAlerts`, `ScreenUpdating`, `Calculation` (get/set), `Version`, `PID`, `Hwnd`, `Workbooks`/`Books`, `ActiveWorkbook`, `Quit`, `Kill` |
| `Books`        | `excel.Workbooks`     | `Add`, `Open` (with `OpenReadOnly`/`OpenPassword`/`OpenUpdateLinks`), `Item`, `Count`, `Active`                        |
| `Book`         | `excel.Workbook`      | `Worksheets`/`Sheets`, `ActiveSheet`, `App`, `Names`, `Name`, `FullName`, `Path`, `Saved`/`SetSaved`, `Activate`, `Save`, `SaveAs`, `Close` |
| `Sheets`       | `excel.Worksheets`    | `Add` (before/after/name), `Item`, `Count`, `Active`                                                                   |
| `Sheet`        | `excel.Worksheet`     | `Range`, `Cells`, `UsedRange`, `Names`, `Name`/`SetName`, `Index`, `Visible`/`SetVisible`, `Activate`, `Delete`, `Clear`, `ClearContents`, `AutoFit` |
| `Range`        | `excel.Range`         | `Value` / `SetValue` (2-D SAFEARRAY decode *and* encode — write whole blocks with `[][]interface{}` or typed slices in one call), `Address`, `Formula`/`SetFormula`, `Formula2`/`SetFormula2`, `NumberFormat`/`SetNumberFormat`, `Cells`, `Offset`, `Resize`, `Rows`, `Columns`, `Row`, `Column`, `Count`, `Width`/`Height`, `ColumnWidth`/`RowHeight`, `End`, `Color`/`SetColor`, `Font()`, `Insert`, `Find`, `Clear`, `ClearContents`, `Delete`, `Copy`, `Merge`/`UnMerge`/`MergeCells`, `AutoFit`, `Options(...)` |
| `Names`/`Name` | `excel.Names`, `excel.Name` | `Add` (formula string or `Range`), `Item`, `Count`, `Contains`, `Name`/`SetName`, `RefersTo`/`SetRefersTo`, `RefersToRange`, `Delete` — via `Workbook.Names()` / `Worksheet.Names()` |
| `Charts`/`Chart` | `excel.Charts`, `excel.Chart` | `Add(left, top, w, h)`, `Item`, `Count`; `SetSourceData(Range)`, `ChartType`/`SetChartType`, `Name`/`SetName`, `Left/Top/Width/Height`, `SetPosition`, `ToPNG`, `ToPDF`, `Delete` — via `Worksheet.Charts()` |
| `Pictures`/`Picture` | `excel.Pictures`, `excel.Picture` | `Add(filename, PictureAt/PictureSize/PictureName...)`, `Item`, `Count`; `Name`/`SetName`, geometry get/set, `Delete` — via `Worksheet.Pictures()` |
| `Shapes`/`Shape` | `excel.Shapes`, `excel.Shape` | `Item`, `Count`, typed `ForEachShape`; `Name`/`SetName`, `Type`, geometry get/set, `SetPosition`, `Delete` — via `Worksheet.Shapes()` |
| `Font`         | `excel.Font`          | `Name`, `Size`, `Bold`, `Italic`, `Color` (get/set each) — via `Range.Font()`; pack colors with `excel.RGB(r, g, b)` |

The xlwings §2.1 object-model roadmap in [AGENTS.md](./AGENTS.md) is now
fully shipped through P2.

### Range.Options — xlwings-style value conversion

`Range.Options(...)` is the Go analogue of xlwings' `Range.options(...)`. It
returns an `OptionedRange` that decodes the range lazily on `.Value()` or
`.Get(&dst)`, applying any combination of shape forcing, range expansion,
header-driven struct decode, empty-cell substitution, or a custom converter.

```go
import "github.com/xll-gen/sugar/excel"

sugar.Do(func(ctx sugar.Context) error {
    app := excel.NewApplication(ctx)
    defer app.Quit()
    app.SetVisible(false).SetDisplayAlerts(false)

    wb := app.Workbooks().Add()
    sheet := wb.ActiveSheet()

    // Seed a small table — one block write, one COM round trip.
    sheet.Range("A1", "B3").SetValue([][]interface{}{
        {"Name", "Age"},
        {"alice", 30.0},
        {"bob", 25.0},
    })

    // 1. Force a scalar read.
    var price float64
    sheet.Range("B2").Options(excel.Scalar()).Get(&price) // -> 30.0

    // 2. Auto-grow from the anchor and decode rows into a struct slice.
    type Person struct {
        Name string
        Age  int
    }
    var people []Person
    err := sheet.Range("A1").Options(
        excel.Expand("table"),
        excel.Header(true),
    ).Get(&people)
    if err != nil {
        return err
    }
    // people -> [{alice 30} {bob 25}]

    // 3. Custom Convert escape hatch.
    sum, _ := sheet.Range("B2", "B3").Options(
        excel.Convert(func(raw [][]interface{}) (interface{}, error) {
            total := 0.0
            for _, row := range raw {
                if v, ok := row[0].(float64); ok {
                    total += v
                }
            }
            return total, nil
        }),
    ).Value()
    _ = sum // -> 55.0

    return nil
})
```

Available options: `Scalar()`, `Vector()` (alias `Vector1D`), `Grid()` (alias
`Vector2D`), `Header(bool)`, `Empty(value)`, `DateFormat(layout)`,
`Expand("table"|"down"|"right")`, and `Convert(fn)`. See
[options.go](./excel/options.go) for full docs.

## Core Concepts

### 1. Standard Execution (`sugar.Do` & `sugar.Go`)

COM is sensitive to the execution thread. `sugar` provides safe entry points to manage this.

- **`sugar.Do`**: Locks the current goroutine to an OS thread and executes synchronously.
- **`sugar.Go`**: Starts a new goroutine (new OS thread) and independently initializes the COM environment for asynchronous work. Returns a buffered `<-chan error` that delivers the goroutine's terminal error — ignore it for fire-and-forget, or receive it to know when the work finished and whether it failed.

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