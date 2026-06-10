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
| `App`          | `excel.Application`   | mostly done | P0       | Has `NewApplication`, `GetApplication`, `Quit`, `Kill`, `Visible`/`SetVisible`, `DisplayAlerts`/`SetDisplayAlerts`, `ScreenUpdating`/`SetScreenUpdating`, `Calculation`/`SetCalculation`, `Version`, `PID`, `Hwnd`, `Workbooks`/`Books` alias, `ActiveWorkbook`. |
| `Books`        | `excel.Workbooks`     | done (2026-06-10) | P0  | `Add`, `Open(path, OpenReadOnly/OpenPassword/OpenUpdateLinks...)`, `Item`, `Count`, `Active`. ForEach inherited from `sugar.Chain`. Middle optional COM params are skipped with `sugar.Missing()` (VT_ERROR/DISP_E_PARAMNOTFOUND). `Format` (text-file delimiter) intentionally not exposed. |
| `Book`         | `excel.Workbook`      | mostly done | P0       | Has `Worksheets`/`Sheets` alias, `ActiveSheet`, `App`, `Names`, `Name`, `FullName`, `Path`, `Saved`/`SetSaved`, `Activate`, `Save`, `SaveAs`, `Close`. |
| `Sheets`       | `excel.Worksheets`    | mostly done | P0       | Has `Add(AddBefore/AddAfter/AddName)`, `Item`, `Count`, `Active`. |
| `Sheet`        | `excel.Worksheet`     | mostly done | P0       | Has `Range`, `Cells`, `UsedRange`, `Names`, `Name`/`SetName`, `Index`, `Visible`/`SetVisible`, `Activate`, `Delete`, `Clear`, `ClearContents`, `AutoFit`. Missing: `Charts`, `Pictures`, `Shapes` (those collections live on their own roadmap rows). |
| `Range`        | `excel.Range`         | done (2026-06-10) | P0  | `Value`/`SetValue` (2-D SAFEARRAY decode+encode), `Address`, `Formula`(`2`)/setters, `NumberFormat`, `Cells`, `Offset`, `Resize`, `Rows`/`Columns`/`Row`/`Column`/`Count`, `Width`/`Height` (points), `ColumnWidth`/`RowHeight` (get/set), `End("up"\|"down"\|"left"\|"right")`, `Color`/`SetColor` (Interior), `Font()`, `Insert("down"\|"right"\|"")`, `Find(what)` (returns `found bool` — COM Nothing is a miss, not an error), `Clear`/`ClearContents`/`Delete`/`Copy`, `Merge`/`UnMerge`/`MergeCells`, `AutoFit`, `Options(...)` (§2.2). `Sort` deliberately omitted: xlwings has no `Range.sort`; COM `Range.Sort` is reachable via the raw chain (`rng.Call("Sort", ...)`). |
| `Name`/`Names` | `excel.Name`, `excel.Names` | done (2026-06-10) | P1 | `name.go`/`names.go`: `Add(name, refersTo)` (string formula or Range), `Item` (by name/index), `Count`, `Contains`, `Name`/`SetName`, `RefersTo`/`SetRefersTo`, `RefersToRange`, `Delete`. Reached via `Workbook.Names()` and `Worksheet.Names()`. Note: `Names.Item` is a *method* in Excel's type library (unlike `Sheets.Item`, a property) — it must be invoked with `Call`, not `Get`. |
| `Chart`/`Charts` | `excel.Chart`, `excel.Charts` | done (2026-06-10) | P1 | `chart.go`/`charts.go`: `Charts.Add(left, top, width, height)`, `Item` (by name/index — a *method*, use `Call`), `Count`; `Chart` fuses COM's ChartObject+Chart like xlwings: `Name`/`SetName`, `ChartType`/`SetChartType` (typed `ChartType` consts), `SetSourceData(Range)`, `Left/Top/Width/Height`, `SetPosition`, `ToPNG` (Chart.Export), `ToPDF` (ExportAsFixedFormat), `Delete`. Via `Worksheet.Charts()` (`ChartObjects` is a method — `Call`, not `Get`). |
| `Picture`/`Pictures` | `excel.Picture`, `excel.Pictures` | done (2026-06-10) | P1 | `picture.go`/`pictures.go`: `Add(filename, PictureAt/PictureSize/PictureName...)` via `Shapes.AddPicture`; `Item`/`Count` via the legacy `Worksheet.Pictures` collection — which is a **snapshot** (its Count never grows), so every lookup re-calls `Pictures()` like xlwings' `api` property. `Name`/`SetName`, geometry get/set, `Delete`. |
| `Shape`/`Shapes` | `excel.Shape`, `excel.Shapes` | done (2026-06-10) | P2  | `shape.go`/`shapes.go`: `Item` (method — `Call`), `Count`, typed `ForEachShape`, `Name`/`SetName`, `Type()` (`ShapeType` MsoShapeType consts), geometry get/set, `SetPosition`, `Delete`. Via `Worksheet.Shapes()`. |
| `Font`           | `excel.Font`               | done (2026-06-10) | P2     | `font.go`: `Name`, `Size`, `Bold`, `Italic`, `Color` (get/set each) via `Range.Font()`. `excel.RGB(r, g, b)` packs the OLE `&HBBGGRR` color int. |

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
    *Package layout (decided 2026-06-10):* `sugar/excel` stays a **single flat package** — do not split the object model into subpackages. The Excel object graph is cyclic (`Worksheet` → `Range` → `Worksheet`, `Name` → `Range`, future `Range` → `Font`), so Go's no-import-cycle rule forces one package, and a single `excel.` namespace mirrors xlwings' single `xw.` namespace. Subpackages are reserved for future *consumers* of the object model with acyclic dependencies (e.g. a v2.x `excel/reports` templating engine, a plotting bridge with heavy deps).
9.  **Tests are mandatory** for every new Excel-layer method. Use the COM-test harness in `sugar_com_test.go` as a template; gate with `//go:build windows && excel_integration` if Excel must be installed.

## 6. Known Improvement Backlog

These items came out of a code review on 2026-05-16. Address them as part of normal work; do not require a separate epic.

Resolved in v0.7.0 (2026-05-16):

* ~~`runner.go`: `Start()`-style panics in production code should become returned errors.~~ — `sugar.Go` / `Runner.Go` / `Context.Go` now return `<-chan error` so async COM errors surface to the caller instead of being silently dropped.
* ~~`excel/excel.go` is a single 162-line file holding every Excel object.~~ — Split into `application.go`, `workbooks.go`, `workbook.go`, `worksheets.go`, `worksheet.go`, `range.go` plus shared `helpers.go` / `win32.go` / `safearray.go` (sugar root).
* ~~`Range.SetValue` exists but there is no `Value()` getter that round-trips a 2-D `SAFEARRAY` into Go.~~ — `sugar.Chain.Value()` now decodes `VT_ARRAY|VT_VARIANT` into `[]interface{}` (1-D) or `[][]interface{}` (2-D) via direct `oleaut32.SafeArrayGetElement` calls; `Range.Value()` is the typed entry point. The `.Options(...)` framework remains future work (§2.2).
* ~~`go.sum` has `golang.org/x/sys v0.1.0` (Jan 2022) as indirect.~~ — Bumped to `v0.30.0` (latest version still compatible with `go 1.22`).

Resolved in v0.7.1 (2026-05-17):

* ~~`Range.Options(...)` conversion framework (§2.2) — scalar/vector/grid forcing, `Expand("table"|"down"|"right")`, struct-by-header decode, custom `Convert(func)`.~~ — Shipped in `excel/options.go` as `Range.Options(opts ...RangeOption) OptionedRange` with `OptionedRange.Value()` and `OptionedRange.Get(dst)`. Option helpers: `Scalar`, `Vector` (alias `Vector1D`), `Grid` (alias `Vector2D`), `Header`, `Empty`, `DateFormat`, `Expand`, `Convert`. Header-driven struct-slice decode (case-insensitive, lenient on unknown columns) is included. Unit tests live in `options_test.go`; integration coverage (Expand + struct decode against real Excel) in `options_integration_test.go`.

Resolved in v0.8.0 (2026-06-10):

* ~~Passing a `sugar.Chain` as a COM argument panicked (`go-ole` `panic("unknown type")`), breaking `Worksheets.Add(AddBefore(...))` and any object-valued argument.~~ — `chain.Get/Call/Put` now normalize arguments via `normalizeParams`: `Chain` → AddRef'd `*ole.IDispatch` (released after the call). go-ole panics are additionally converted to chain errors by `invokeGuarded`.
* ~~`Range.SetValue([][]interface{})` panicked for the same reason — there was no SAFEARRAY *encode* path, only decode.~~ — `safearray.go` now has `encodeVariantArray` (1-D `[]interface{}` and 2-D `[][]interface{}` → `VT_ARRAY|VT_VARIANT`), with cell support for nil/bool/string/all int widths/floats/`time.Time` (VT_DATE, wall-clock). Block writes work in one COM round trip; ragged 2-D input is a chain error. COM tests: `normalize_com_test.go` (core), `TestRange_SetValue2D` (excel).
* ~~`Name`/`Names` absent (P1).~~ — Shipped; see §2.1 row.

Resolved in v0.8.0 (2026-06-10, continued):

* ~~Value-result chains leaked their VARIANTs (BSTR per string property read).~~ — `handleResult` now arena-tracks value chains; `Release()` VariantClears them at scope end. Regression test: `TestValueChainsAreTracked`.
* ~~`sugar.Do` failed on threads the host already CoInitialize'd (S_FALSE surfaced as error by go-ole); RPC_E_CHANGED_MODE (MTA thread) also unhandled.~~ — `initializeCOM` in runner.go treats S_FALSE as success-owing-CoUninitialize and RPC_E_CHANGED_MODE as success-without. Critical for xll-gen hosts. Test: `TestDo_OnPreInitializedSTAThread`.
* ~~No goroutine-safety tests for `sugar.Go`.~~ — `TestGo_TwoExcelInstancesIsolated` drives two Excel instances on two OS threads concurrently (distinct Hwnds, independent values).
* ~~`Range.Options(...)` extensions.~~ — Shipped: `Index(n)` leading-column skip, positional struct decode when `Header(false)` (column order = exported field order), and generic `ConvertTo[T]` for compile-time-checked converters.

Resolved in v0.9.0 (2026-06-10):

* ~~Object collections absent: `Picture`/`Pictures`, `Shape`/`Shapes`, `Font`.~~ — All shipped; see §2.1 rows. API-breaking rename that came with it: the `Range.Options` dimension knob `Shape`/`ShapeAuto`/`ShapeScalar`/`ShapeVector`/`ShapeGrid` became **`NDim`/`NDimAuto`/`NDimScalar`/`NDimVector`/`NDimGrid`** (xlwings' `ndim`), freeing the `Shape` identifier for the drawing object per the xlwings-naming rule. The `Scalar()`/`Vector()`/`Grid()` option helpers are unchanged.
* ~~Typed slices need an encode path.~~ — `encodeVariantArray` now widens any Go slice via reflection (`[][]float64`, `[]int`, `[][]string`, …); `[]byte`/`[]string` stay on go-ole's native VT_UI1/VT_BSTR paths to avoid changing behavior for non-Excel COM servers. Encode→decode round-trip unit tests in `safearray_test.go` run without Excel (SafeArray APIs need no CoInitialize). Still open below: `[]T` struct rows.

Resolved in v0.9.0 (2026-06-10, continued):

* ~~COM `Nothing` results crashed: `handleResult` called `AddRef()` on a nil dispatch (e.g. `Range.Find` miss, `ActiveWorkbook` with no book).~~ — Nothing now becomes an empty value chain (`Value()` nil, `IsDispatch()` false).
* ~~No way to skip middle optional COM parameters.~~ — `sugar.Missing()` returns the VT_ERROR/DISP_E_PARAMNOTFOUND placeholder; `sugar.Error(err)` creates an error-only chain for typed-wrapper validation failures.
* ~~`Range` missing members.~~ — `End`, `Color`/`SetColor`, `Width`/`Height`, `ColumnWidth`/`RowHeight`, `Insert`, `Find` shipped; `Sort` deliberately dropped (not in xlwings — use the raw chain).
* ~~Rich `Books.Open` options.~~ — `OpenReadOnly`/`OpenPassword`/`OpenUpdateLinks`. Testing gotcha recorded in `workbooks_test.go`: opening a protected book *without* a password pops a modal prompt that DisplayAlerts does **not** suppress — always probe the failure path with an explicit wrong password.
* ~~`[]T` struct-row write support.~~ — `OptionedRange.Set(src)` writes struct slices (header row with `Header(true)`, positional without), 1-D/2-D slices (auto-resized from the anchor), or scalars. Mirror of `Get`.

Still open:

* Integration tests boot a fresh Excel per test (~2–3 s each, full tagged suite ≈183 s). A shared instance must respect COM apartment rules: tests run on arbitrary goroutines, so sharing raw IDispatch is unsafe — the workable design is one Excel process + per-test `GetActive` attach (ROT marshaling). Worth doing once suite time hurts.

## 7. Documentation Standards

*   **Timestamps**: When recording timestamped notes (changelog, design records), use the current absolute date — do not rely on relative phrases that age poorly.
*   **READMEs**: Keep README.md examples copy-paste-runnable. If you change a public API in §2.1, update both README.md and any docstring referencing the old form.
