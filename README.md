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
    // Setters return Application for fluent chaining; getters return
    // (bool, error) like every other typed getter.
    app.SetVisible(true).
        SetDisplayAlerts(false).
        SetScreenUpdating(false)
    visible, _ := app.Visible() // true
    _ = visible

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
| `Book`         | `excel.Workbook`      | `Worksheets`/`Sheets`, `ActiveSheet`, `App`, `Names`, `Name`, `FullName`, `Path`, `Saved`/`SetSaved`, `Activate`, `Save`, `SaveAs` (with `SaveFileFormat`/`SavePassword`), `Close` (with `CloseSaveChanges`) |
| `Sheets`       | `excel.Worksheets`    | `Add` (before/after/name), `Item`, `Count`, `Active`                                                                   |
| `Sheet`        | `excel.Worksheet`     | `Range`, `Cells`, `UsedRange`, `Names`, `Name`/`SetName`, `Index`, `Visible`/`SetVisible`, `Activate`, `Delete`, `Clear`, `ClearContents`, `AutoFit` |
| `Range`        | `excel.Range`         | `Value` / `SetValue` (2-D SAFEARRAY decode *and* encode — write whole blocks with `[][]interface{}` or typed slices in one call), `Address`, `Formula`/`SetFormula`, `Formula2`/`SetFormula2`, `NumberFormat`/`SetNumberFormat`, `Cells`, `Offset`, `Resize`, `Rows`, `Columns`, `Row`, `Column`, `Count`, `Width`/`Height`, `ColumnWidth`/`RowHeight`, `End`, `Color`/`SetColor`, `Font()`, `Insert`, `Find`, `Clear`, `ClearContents`, `Delete`, `Copy`, `Merge`/`Unmerge`/`MergeCells`, `AutoFit` (column width + row height), `Options(...)` |
| `Names`/`Name` | `excel.Names`, `excel.Name` | `Add` (formula string or `Range`), `Item`, `Count`, `Contains`, `Name`/`SetName`, `RefersTo`/`SetRefersTo`, `RefersToRange`, `Delete` — via `Workbook.Names()` / `Worksheet.Names()` |
| `Charts`/`Chart` | `excel.Charts`, `excel.Chart` | `Add(ChartAt/ChartSize...)`, `Item`, `Count`; `SetSourceData(Range)`, `ChartType`/`SetChartType`, `Name`/`SetName`, `Left/Top/Width/Height`, `SetPosition`, `ToPNG`, `ToPDF`, `Delete` — via `Worksheet.Charts()` |
| `Pictures`/`Picture` | `excel.Pictures`, `excel.Picture` | `Add(filename, PictureAt/PictureSize/PictureName...)`, `Item`, `Count`; `Name`/`SetName`, geometry get/set, `Delete` — via `Worksheet.Pictures()` |
| `Shapes`/`Shape` | `excel.Shapes`, `excel.Shape` | `Item`, `Count`, typed `ForEachShape`; `Name`/`SetName`, `Type`, geometry get/set, `SetPosition`, `Delete` — via `Worksheet.Shapes()` |
| `Font`         | `excel.Font`          | `Name`, `Size`, `Bold`, `Italic`, `Color` (get/set each) — via `Range.Font()`; pack colors with `excel.RGB(r, g, b)` |

The xlwings §2.1 object-model roadmap in [AGENTS.md](./AGENTS.md) is now
fully shipped through P2.

### Opening workbooks

`Workbooks.Open` takes an absolute path plus xlwings-style options. Opening a
protected file with the wrong password fails fast with an error (no modal
prompt as long as you always pass *some* password).

```go
sugar.Do(func(ctx sugar.Context) error {
    app := excel.NewApplication(ctx)
    defer app.Quit()
    app.SetVisible(false).SetDisplayAlerts(false)

    wb := app.Workbooks().Open(`C:\data\report.xlsx`,
        excel.OpenReadOnly(),          // xlwings read_only=True
        excel.OpenPassword("secret"),  // xlwings password=...
        excel.OpenUpdateLinks(0),      // 0 = don't refresh external links
    )
    if err := wb.Err(); err != nil {
        return err
    }
    defer wb.Close()

    name, _ := wb.Name() // "report.xlsx"
    _ = name
    return nil
})
```

### Range essentials: block I/O, navigation, formatting

```go
sheet := wb.ActiveSheet()

// Block writes are one COM round trip. Typed Go slices ([][]float64, []int,
// [][]string, ...) encode to SAFEARRAYs natively — no []interface{} required.
sheet.Range("A1", "C1").SetValue([]interface{}{"q1", "q2", "q3"})
sheet.Range("A2", "C3").SetValue([][]float64{
    {1.5, 2.5, 3.5},
    {4.0, 5.0, 6.0},
})

// Ctrl+Arrow navigation and Ctrl+F search.
last := sheet.Range("A1").End("down")       // "up" | "down" | "left" | "right"
addr, _ := last.Address()                   // "$A$3"
cell, found, err := sheet.UsedRange().Find("q2")
if err == nil && found {
    addr, _ = cell.Address()                // a miss is found=false, not an error
}

// Geometry and layout.
hdr := sheet.Range("A1").Resize(1, 3)       // A1:C1
_ = sheet.Range("A1").Offset(1, 0)          // A2
hdr.SetRowHeight(24)
hdr.SetColumnWidth(14)
_ = hdr.AutoFit() // fits both column width and row height (xlwings parity)

// Colors and fonts. excel.RGB packs the OLE &HBBGGRR color int.
hdr.SetColor(excel.RGB(255, 255, 0)) // Interior fill
hdr.Font().SetBold(true).SetSize(12).SetColor(excel.RGB(180, 0, 0))

// Insert cells, shifting the rest away ("down", "right", or "" to let
// Excel decide from the range shape).
_ = sheet.Range("A2", "C2").Insert("down")
```

### Charts

`Worksheet.Charts()` manages embedded charts. Like xlwings, `excel.Chart`
fuses COM's ChartObject (geometry) and Chart (data) into one object.

```go
sheet.Range("A1", "B4").SetValue([][]interface{}{
    {"Month", "Sales"},
    {"Jan", 10.0},
    {"Feb", 20.0},
    {"Mar", 15.0},
})

// Functional options (defaults: ChartAt(0,0), ChartSize(355,211) points).
ch := sheet.Charts().Add(excel.ChartAt(10, 10), excel.ChartSize(360, 220))
if err := ch.SetSourceData(sheet.Range("A1", "B4")); err != nil {
    return err
}
ch.SetChartType(excel.ChartColumnClustered).SetName("Sales")

if err := ch.ToPNG(`C:\out\sales.png`); err != nil { // chart.to_png()
    return err
}
_ = ch.ToPDF(`C:\out\sales.pdf`) // ExportAsFixedFormat

byName := sheet.Charts().Item("Sales") // or 1-based index
_ = byName.Delete()
```

### Pictures and shapes

```go
// Insert an image with optional placement, size, and name.
pic := sheet.Pictures().Add(`C:\out\sales.png`,
    excel.PictureAt(30, 40),      // left, top — defaults to (0, 0)
    excel.PictureSize(120, 90),   // omit to keep the image's natural size
    excel.PictureName("Logo"))
if err := pic.Err(); err != nil {
    return err
}

// Everything on the drawing layer (pictures, charts, ...) is a Shape.
res := sheet.Shapes().ForEachShape(func(s excel.Shape) error {
    st, err := s.Type()
    if err != nil {
        return err
    }
    if st == excel.ShapeTypePicture {
        s.SetPosition(0, 0, 60, 45) // left, top, width, height
    }
    return nil
})
if err := res.Err(); err != nil {
    return err
}
```

### Defined names

```go
// Workbook-scoped: add by Range or by an A1-notation formula string.
n := wb.Names().Add("inputs", sheet.Range("A1", "B4"))
if err := n.Err(); err != nil {
    return err
}
rng := wb.Names().Item("inputs").RefersToRange()
addr, _ := rng.Address() // "$A$1:$B$4"

ok, _ := wb.Names().Contains("inputs") // true
_ = wb.Names().Item("inputs").Delete()

// Sheet-scoped names come back qualified ("Sheet1!local"), like xlwings.
sheet.Names().Add("local", sheet.Range("D1"))
```

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

    // 4. Set is the write-direction mirror of Get: anchor at one cell and
    //    it auto-resizes to fit. With Header(true) a header row of field
    //    names is written first; without it rows are written positionally.
    if err := sheet.Range("D1").Options(excel.Header(true)).Set(people); err != nil {
        return err
    }
    // D1:E3 now holds: Name | Age / alice | 30 / bob | 25

    return nil
})
```

Available options: `Scalar()`, `Vector()`, `Grid()`, `Header(bool)`,
`Index(n)`, `Empty(value)`, `DateFormat(layout)`,
`Expand("table"|"down"|"right")`, `Convert(fn)`, and the compile-time-checked
`ConvertTo[T](fn)`. See [options.go](./excel/options.go) for full docs.

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

**Manual `Release()` calls are not necessary.**

### 5. Skipping Optional COM Parameters (`sugar.Missing`)

Many COM methods take long positional parameter lists where you only want to
set a late parameter. `sugar.Missing()` produces the canonical "parameter not
supplied" VARIANT (`VT_ERROR` / `DISP_E_PARAMNOTFOUND`) so middle optionals
can be skipped:

```go
// Any positional COM method with middle optionals — skip the second arg:
obj.Call("SomeMethod", "first", sugar.Missing(), "third")
```

(For the common Excel cases, prefer the typed wrappers, e.g.
`wb.SaveAs(path, excel.SaveFileFormat(...), excel.SavePassword(...))`, which
handle the Missing() bookkeeping internally.)

Relatedly, `sugar.Error(err)` creates a chain that carries only an error —
useful in typed wrappers that must surface a validation failure through the
fluent chain contract before any COM call happens.

### 6. Nested Scopes

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
- **Testing:** `go test ./...` runs only Excel-free unit tests (lightweight scripting COM objects at most). The live-Excel integration suite is opt-in: `go test -tags=excel_integration ./...` (requires installed Excel; boots real instances).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.