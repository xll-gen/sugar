# Agent Instructions for the `sugar` Repository

This document provides guidance for developers and AI agents working on the `sugar` codebase.

## 0. Mission Statement

**`sugar` aims to be the Go equivalent of [xlwings](https://www.xlwings.org/) — bringing xlwings' main features (Excel automation, range I/O, plotting, reporting, and runtime callbacks) to Go on Windows.**

Where xlwings is "Python in Excel", `sugar` is "Go in Excel". The COM/Chain/Arena foundation already in this repo is the substrate; the `sugar/excel` subpackage is where the xlwings-style ergonomic API surface lives.

When in doubt about API shape, default to **xlwings naming and behavior** (Pythonic but consistent). Translate Pythonic idioms to idiomatic Go (methods over keyword args, `error` returns, options pattern for variadic settings).

## 1. Project Overview

`sugar` is a Go library that **brings xlwings-style Excel automation to Go on Windows**. It provides a fluent, immutable API for Component Object Model (COM) automation, using an **Arena (Context) pattern** for automatic resource management.

**Platform Constraint:** Windows-specific, x86 / x86-64 (Intel/AMD) only. No ARM (incl. Windows-on-ARM, Apple Silicon). All source files must include `//go:build windows`. The COM `IDispatch` automation surface is Windows-bound; no cross-platform abstraction is planned.

**Two layers:**

1. **Core (`sugar`)** — generic COM automation primitives (`Do`/`Go`, `Chain`, `Context`, `expression`). Application-agnostic. Stable.
2. **Excel layer (`sugar/excel`)** — xlwings-parity, type-safe Excel API. This is the active growth area. Modeled after xlwings' object model: `App`/`Books`/`Sheets`/`Range`/`Chart`/`Picture`/`Name`/`Shape`.

## 2. xlwings Feature-Parity Roadmap

Track progress against the xlwings reference API: <https://docs.xlwings.org/en/stable/api.html>.
File new work under `sugar/excel/<object>.go` (one file per top-level object).

### 2.1 Object Model Targets

Implement these in priority order. Each must support method chaining via `sugar.Chain` and live under its own `.go` file.

| xlwings Object | sugar target type     | Status     | Priority | Notes                                                                |
| -------------- | --------------------- | ---------- | -------- | -------------------------------------------------------------------- |
| `App`          | `excel.Application`   | partial    | P0       | Has `NewApplication`, `GetApplication`, `Quit`, `Visible`/`SetVisible`, `DisplayAlerts`/`SetDisplayAlerts`, `ScreenUpdating`/`SetScreenUpdating`. Missing: `Calculation`, `Version`, `PID`, `Hwnd`, `Kill`, `Books`/`Books()` (alias for `Workbooks()`). |
| `Books`        | `excel.Workbooks`     | partial    | P0       | Has `Add`, `Item`. Missing: `Open(path, ...opts)`, `Count`, iteration (`ForEach`), `Active`. |
| `Book`         | `excel.Workbook`      | partial    | P0       | Has `Worksheets`, `ActiveSheet`, `Save`, `Close`. Missing: `SaveAs(path, ...)`, `FullName`, `Name`, `Path`, `Sheets` (alias for `Worksheets`), `Names`, `App`, `Activate`. |
| `Sheets`       | `excel.Worksheets`    | partial    | P0       | Has `Item`. Missing: `Add(before, after, name)`, `Count`, iteration, `Active`. |
| `Sheet`        | `excel.Worksheet`     | partial    | P0       | Has `Range`, `Cells`. Missing: `Name`/`SetName`, `Index`, `Activate`, `Delete`, `Clear`, `ClearContents`, `UsedRange`, `Visible`, `Names`, `Charts`, `Pictures`, `Shapes`, `AutoFit`. |
| `Range`        | `excel.Range`         | partial    | P0       | Has `SetValue`, `Cells`. Missing: `Value()` getter with 2D unmarshal, `Address`, `Formula`, `Formula2`, `NumberFormat`, `Font`, `Color`, `Resize`, `Offset`, `Expand`, `End`, `Rows`, `Columns`, `Count`, `Row`, `Column`, `Width`, `Height`, `MergeCells`, `Merge`, `UnMerge`, `Clear`, `ClearContents`, `Copy`, `Delete`, `Insert`, `AutoFit`, `Sort`, `Find`. |
| `Name`/`Names` | `excel.Name`, `excel.Names` | absent | P1       | Workbook/sheet-scoped named ranges: `Add(name, refersTo)`, `Item`, iteration, `Delete`. |
| `Chart`/`Charts` | `excel.Chart`, `excel.Charts` | absent | P1     | `Add(left, top, width, height)`, `SetSourceData(range)`, `ChartType`, `Name`, `Delete`, `ToPDF`, `ToPNG`. |
| `Picture`/`Pictures` | `excel.Picture`, `excel.Pictures` | absent | P1 | `Add(filename, ...)`, `Update`, `Delete`. Used by xlwings' matplotlib bridge. |
| `Shape`/`Shapes` | `excel.Shape`, `excel.Shapes` | absent | P2  | `Item`, iteration, `Delete`, position/size getters & setters. |
| `Font`           | `excel.Font`               | absent | P2     | `Name`, `Size`, `Bold`, `Italic`, `Color`. Reached via `Range.Font()`. |

### 2.2 Range Value Conversion (xlwings `.options()` analogue)

xlwings' `Range.options(...)` is its defining feature: it converts to/from numpy arrays, pandas DataFrames, dicts, scalars, with configurable dimension/empty/date handling.

In Go we lack pandas/numpy, but the **conversion abstraction is still essential**:

```go
// Target API:
var v float64
rng.Options(excel.Scalar()).Get(&v)

var grid [][]any
rng.Options(excel.Expand("table")).Get(&grid)

var rows []MyStruct
rng.Options(excel.Header(true)).Get(&rows)   // struct unmarshal by header
```

Options to implement (start with the first three):

* `Scalar()` — force 1×1 read as a scalar.
* `Vector()` / `Vector1D()` — force 1-D slice even on `n×1` / `1×n`.
* `Grid()` / `Vector2D()` — always `[][]T`.
* `Header(bool)` — first row as struct field names (for `[]T` decode).
* `Index(int)` — first N columns as index (logging/skip).
* `Empty(value)` — replacement for empty cells (default: zero value / `nil`).
* `DateFormat(layout)` — xlwings uses Python `datetime`; we use `time.Time`.
* `Expand("table"|"down"|"right")` — auto-expand from anchor before reading.
* `Convert(func(raw [][]any) (T, error))` — escape hatch for custom converters.

Implementation note: `Range.Value()` should go through `IDispatch.Value2` (returns COM `VARIANT`); use `go-ole`'s `SafeArrayConversion` and translate dates from OLE date doubles (`1899-12-30` epoch) to `time.Time`.

### 2.3 Runtime Callbacks (xlwings `RunPython` analogue)

xlwings lets Excel call back into Python. For `sugar`, the equivalent path is **not** the responsibility of this repo — it is handled by `xll-gen` (XLL → spawned Go server) and by an eventual VBA `RunGo` macro. This repo should expose just enough to make in-process automation pleasant; do **not** add IPC, RPC, or process-spawning here.

### 2.4 Reporting / Templating (xlwings `Reports` analogue)

xlwings PRO ships a Jinja-style report engine that fills templates with data. Out of scope for v1.x of this repo. Re-evaluate in v2.x once the object model is complete.

### 2.5 UDFs

xlwings registers Python UDFs through `xlwings.xlam`. Go UDFs in Excel are the domain of `xll-gen`, not `sugar`. Do not implement UDF registration here.

## 3. Core Concepts & Usage Patterns

### 3.1 Standard Execution (`sugar.Do` and `sugar.Go`)

The library enforces a standard way to execute COM operations to ensure `runtime.LockOSThread()` and `ole.CoInitialize()` are handled correctly.

```go
sugar.Do(func(ctx sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    // ... work with excel ...
    return nil
})
```

*   **`sugar.Do`**: Executes synchronously in the current goroutine.
*   **`sugar.Go`**: Executes in a new goroutine (new OS thread).
*   **Nested Scopes**: Use `ctx.Do(func(innerCtx sugar.Context) error { ... })` to create local cleanup scopes.

### 3.2 The Immutable `Chain`

All methods on `*sugar.Chain` (`Get`, `Call`, `ForEach`, etc.) return a **NEW** `Chain` instance. The original instance remains unchanged.

```go
excel := ctx.Create("Excel.Application")
workbooks := excel.Get("Workbooks") // 'excel' still points to Application
wb := workbooks.Call("Add")         // 'workbooks' still points to Workbooks collection
```

### 3.3 Automatic Resource Management (Arena)

Every `Chain` created via a `sugar.Context` (or derived from one) is automatically tracked by that context. When the `sugar.Do` block completes, all tracked COM objects are released in reverse order.

**Manual `Release()` is unnecessary** within a `Do/Go` block.

### 3.4 Integration with `context.Context`

`sugar.Context` implements the standard `context.Context` interface. You can use it for cancellation, timeouts, and passing values.

```go
sugar.With(parentCtx).Do(func(ctx sugar.Context) error {
    select {
    case <-ctx.Done():
        return ctx.Err()
    default:
        // ...
        return nil
    }
})
```

## 4. Expression Subpackage

The `expression` package allows navigating COM objects using string expressions (e.g., `"Workbooks.Add().ActiveSheet"`).

*   It uses `sugar.Context` under the hood if a tracked `Chain` is passed.
*   Intermediate objects created during expression evaluation are automatically managed by the chain's context.

## 5. Development Rules

1.  **Always use `sugar.Do`** for entry points.
2.  **Never manually call `CoInitialize`** unless implementing a low-level runner.
3.  **Prefer `ctx.Create`** over `sugar.Create` to ensure automatic tracking.
4.  **Immutable behavior**: Do not expect a `Chain` variable to change its internal state after a method call.
5.  **Thread Safety**: Remember that `Go` routines start fresh threads; do not share raw `IDispatch` pointers across threads without proper COM marshaling (though `sugar.Go` makes creating thread-local objects easy).
6.  **Language Requirement**: All documentation (including README and AGENTS.md) and code comments must be written in English.
7.  **xlwings naming**: When adding an Excel-layer method, mirror the xlwings name and semantics first; deviate only when Go idioms demand it (always document the deviation).
8.  **One object per file** in `sugar/excel/`: `application.go`, `workbook.go`, `worksheet.go`, `range.go`, `chart.go`, etc. The current single-file `excel.go` should be split as the surface grows.
9.  **Tests are mandatory** for every new Excel-layer method. Use the COM-test harness in `sugar_com_test.go` as a template; gate with `//go:build windows && excel_integration` if Excel must be installed.

## 6. Known Improvement Backlog

These items came out of a code review on 2026-05-16. Address them as part of normal work; do not require a separate epic.

* `runner.go`: `Start()`-style panics in production code should become returned errors.
* `excel/excel.go` is a single 162-line file holding every Excel object. Split per object as Section 5.8 directs, before expanding the API surface in §2.1.
* `Range.SetValue` exists but there is no `Value()` getter that round-trips a 2-D `SAFEARRAY` into Go. Add it together with the `.Options()` framework (§2.2) — they are a single coherent change.
* No goroutine-safety tests for `sugar.Go`. Add a regression test that launches two Excel instances on two OS threads and verifies they do not interfere.
* `go.sum` has `golang.org/x/sys v0.1.0` (Jan 2022) as indirect. Bump when `go-ole` is updated.

## 7. Documentation Standards

*   **Timestamps**: When recording timestamped notes (changelog, design records), use the current absolute date — do not rely on relative phrases that age poorly.
*   **READMEs**: Keep README.md examples copy-paste-runnable. If you change a public API in §2.1, update both README.md and any docstring referencing the old form.
