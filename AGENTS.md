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
| `App`          | `excel.Application`   | mostly done | P0       | Has `NewApplication`, `GetApplication`, `Quit`, `Kill`, `Visible`/`SetVisible`, `DisplayAlerts`/`SetDisplayAlerts`, `ScreenUpdating`/`SetScreenUpdating`, `Calculation`/`SetCalculation`, `Version`, `PID`, `Hwnd`, `Workbooks`/`Books` alias, `ActiveWorkbook`. (v1.0) The three bool getters return `(bool, error)` like every other typed getter; `Hwnd` returns `uintptr` (handle-sized) and `PID` returns `uint32` (Windows DWORD). |
| `Books`        | `excel.Workbooks`     | done (2026-06-10) | P0  | `Add`, `Open(path, OpenReadOnly/OpenPassword/OpenUpdateLinks...)`, `Item`, `Count`, `Active`. ForEach inherited from `sugar.Chain`. Middle optional COM params are skipped with `sugar.Missing()` (VT_ERROR/DISP_E_PARAMNOTFOUND). `Format` (text-file delimiter) intentionally not exposed. |
| `Book`         | `excel.Workbook`      | mostly done | P0       | Has `Worksheets`/`Sheets` alias, `ActiveSheet`, `App`, `Names`, `Name`, `FullName`, `Path`, `Saved`/`SetSaved`, `Activate`, `Save`, `SaveAs(path, SaveFileFormat/SavePassword...)`, `Close(CloseSaveChanges...)`. (v1.0) SaveAs/Close take variadic functional options following the Books.Open pattern. |
| `Sheets`       | `excel.Worksheets`    | mostly done | P0       | Has `Add(AddBefore/AddAfter/AddName)`, `Item`, `Count`, `Active`. |
| `Sheet`        | `excel.Worksheet`     | mostly done | P0       | Has `Range`, `Cells`, `UsedRange`, `Names`, `Name`/`SetName`, `Index`, `Visible`/`SetVisible`, `Activate`, `Delete`, `Clear`, `ClearContents`, `AutoFit`. Missing: `Charts`, `Pictures`, `Shapes` (those collections live on their own roadmap rows). `Clear`/`ClearContents` reach the whole sheet through the `Cells` **property** (`Get`, not `Call` — they were broken until 2026-07-26; see §6); integration coverage lives in `worksheet_test.go`. |
| `Range`        | `excel.Range`         | done (2026-06-10) | P0  | `Value`/`SetValue` (2-D SAFEARRAY decode+encode), `Address`, `Formula`(`2`)/setters, `NumberFormat`, `Cells`, `Offset`, `Resize`, `Rows`/`Columns`/`Row`/`Column`/`Count`, `Width`/`Height` (points), `ColumnWidth`/`RowHeight` (get/set), `End("up"\|"down"\|"left"\|"right")`, `Color`/`SetColor` (Interior), `Font()`, `Insert("down"\|"right"\|"")`, `Find(what)` (returns `found bool` — COM Nothing is a miss, not an error; a sugar extension with no xlwings analogue, so it pins Excel's session-persisted search settings `LookIn`/`LookAt`/`SearchOrder`/`MatchByte` explicitly instead of inheriting the last search — see §6), `Clear`/`ClearContents`/`Delete`/`Copy`, `Merge`/`Unmerge`/`MergeCells`, `AutoFit` (column width + row height, v1.0), `Options(...)` (§2.2). `SetFormulaSpill(formula)` (sugar-specific, no direct xlwings analogue — documented deviation) writes via the DA-native `Formula2` property and falls back to legacy `Formula` if Formula2 is absent (pre-DA Excel 2016-); use it for any spill-expected formula so DA Excel does not rewrite a UDF call into the implicit-intersection `=@Fn(...)` form that suppresses spilling. `Sort` deliberately omitted: xlwings has no `Range.sort`; COM `Range.Sort` is reachable via the raw chain (`rng.Call("Sort", ...)`). Value decode covers `VT_CY`/`VT_DECIMAL`/`VT_ERROR` (currency & error cells no longer silently decode to nil — see §2.2 note; error cells become `sugar.CellError`). `Options(Expand(...))` is evaluated **lazily at read time** (matching xlwings "options are only evaluated when accessing the values"): a stored `OptionedRange` re-discovers the current block on every `Value()`/`Get()`, so data that grows after `Options()` is captured is included. Only the direction string is validated eagerly. `Expand("down")`/`Expand("right")` preserve a multi-cell anchor's span on the axis perpendicular to the growth direction (`Range("A1:C1")` + `down` reads A1:C\<end\>, matching xlwings' `VerticalExpander`) — they used to collapse it to one column/row and truncate the read silently (see §6, 2026-07-29). The endpoint guard is xlwings' **three-rung ladder** (neighbor blank → origin; second neighbor blank → neighbor; else `neighbor.End(dir)`), so `End()` is only ever called from a cell already proven non-empty — a table with an empty top-left corner expands over the whole block instead of stopping at its second cell (see §6, 2026-08-03). |
| `Name`/`Names` | `excel.Name`, `excel.Names` | done (2026-06-10) | P1 | `name.go`/`names.go`: `Add(name, refersTo)` (string formula or Range), `Item` (by name/index), `Count`, `Contains`, `Name`/`SetName`, `RefersTo`/`SetRefersTo`, `RefersToRange`, `Delete`. Reached via `Workbook.Names()` and `Worksheet.Names()`. Note: `Names.Item` is a *method* in Excel's type library (unlike `Sheets.Item`, a property) — it must be invoked with `Call`, not `Get`. |
| `Chart`/`Charts` | `excel.Chart`, `excel.Charts` | done (2026-06-10) | P1 | `chart.go`/`charts.go`: `Charts.Add(ChartAt/ChartSize...)` (v1.0 functional options; defaults 0,0,355,211), `Item` (by name/index — a *method*, use `Call`), `Count`; `Chart` fuses COM's ChartObject+Chart like xlwings: `Name`/`SetName`, `ChartType`/`SetChartType` (typed `ChartType` consts), `SetSourceData(Range)`, `Left/Top/Width/Height`, `SetPosition`, `ToPNG` (Chart.Export), `ToPDF` (ExportAsFixedFormat), `Delete`. Via `Worksheet.Charts()` (`ChartObjects` is a method — `Call`, not `Get`). |
| `Picture`/`Pictures` | `excel.Picture`, `excel.Pictures` | done (2026-06-10) | P1 | `picture.go`/`pictures.go`: `Add(filename, PictureAt/PictureSize/PictureName...)` via `Shapes.AddPicture`; `Item`/`Count` via the legacy `Worksheet.Pictures` collection — which is a **snapshot** (its Count never grows), so every lookup re-calls `Pictures()` like xlwings' `api` property. `Name`/`SetName`, geometry get/set, `Delete`. |
| `Shape`/`Shapes` | `excel.Shape`, `excel.Shapes` | done (2026-06-10) | P2  | `shape.go`/`shapes.go`: `Item` (method — `Call`), `Count`, typed `ForEachShape`, `Name`/`SetName`, `Type()` (`ShapeType` MsoShapeType consts), geometry get/set, `SetPosition`, `Delete`. Via `Worksheet.Shapes()`. |
| `Font`           | `excel.Font`               | done (2026-06-10) | P2     | `font.go`: `Name`, `Size`, `Bold`, `Italic`, `Color` (get/set each) via `Range.Font()`. `excel.RGB(r, g, b)` packs the OLE `&HBBGGRR` color int. |
| App (multi-instance) | excel.GetApplicationByPID | done (2026-06-12) | P2 | `application.go` + `win32.go`: attaches to a specific running Excel by PID via the `XLMAIN -> XLDESK -> EXCEL7` window walk + `AccessibleObjectFromWindow(OBJID_NATIVEOM)` — NOT the ROT (which is process-ambiguous and fails with `MK_E_UNAVAILABLE` from a child server process). Consumer: xll-gen command handlers receive `CommandContext.ExcelPID`. Fixed the showcase "ribbon click does nothing — cannot attach to Excel" bug: the ribbon→Invoke→SendCommandInvoke→handler chain worked, but the handler's ROT-based `GetApplication` failed so it wrote nowhere. |

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
* `Vector()` — force 1-D slice even on `n×1` / `1×n`.
* `Grid()` — always `[][]T`.
* `Header(bool)` — first row as struct field names (for `[]T` decode).
* `Index(int)` — first N columns as index (logging/skip).
* `Empty(value)` — replacement for empty cells (default: zero value / `nil`).
* `DateFormat(layout)` — xlwings uses Python `datetime`; we use `time.Time`.
* `Expand("table"|"down"|"right")` — auto-expand from anchor before reading.
* `Convert(func(raw [][]any) (T, error))` — escape hatch for custom converters.

Implementation note: `Range.Value()` reads the COM **`Value`** property (not `Value2`), decoding the returned `VARIANT`. `Value2` was rejected: it would drop `VT_DATE`→`time.Time` parity (dates come back as raw OLE serial doubles) and still would not fix the `VT_ERROR` case. Instead the VARIANT decode layer (`decodeVariantScalar` in `safearray.go`, shared by scalar `chain.Value()` and the per-cell `getElement`) fills the gaps go-ole v1.3.0's `(*VARIANT).Value()` switch leaves nil: `VT_CY` (currency → `float64`, /1e-4 scale), `VT_DECIMAL` (via `VarR8FromDec`), and `VT_ERROR` (worksheet error cells → typed `sugar.CellError`, so `#DIV/0!` is distinguishable from a blank cell; `DISP_E_PARAMNOTFOUND` stays nil). 2-D reads drive `oleaut32` directly (not go-ole's 1-D-only `SafeArrayConversion`) and transfer the whole grid in bulk via `SafeArrayAccessData` rather than per cell — see §6, "Resolved post-v0.8.9"; dates translate from OLE serial doubles (`1899-12-30` epoch) to `time.Time`.

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
10. **Removing an internal guard requires proof, not a hunch (over-defensive-logic audit, 2026-06-25).** A 2026-06-25 cross-repo audit found **no** over-defensive logic in `sugar` — the defensive posture (nil/COM-pointer guards in chain steps, Arena liveness checks) is well-placed. Keep it that way: only delete an internal validation when an *earlier check in the same function* provably subsumes it AND the value is not external/COM input. The Arena/Chain invariant makes a live `IDispatch` likely, not guaranteed (a `Do` can fail mid-chain), so a nil-guard on a COM pointer is usually load-bearing, not redundant.

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

Resolved in v0.9.1 (2026-06-10, test infrastructure):

* ~~`go test ./...` booted Excel: the core-chain COM tests (`sugar_com_test.go`, `normalize_com_test.go`, the `sugar.Go` isolation test), `excel/excel_test.go`, and the context/expression tests were only `windows`-tagged.~~ — Everything Excel-bound is now behind `windows && excel_integration` (run via `task integration`, which covers `./...`, not just `./excel/...`). Context- and expression-mechanics tests were converted to lightweight `Scripting.Dictionary` / `Scripting.FileSystemObject` servers so they still run Excel-free in `go test ./...`.
* ~~Excel teardown was a bare `defer Quit()` — a hung Quit (modal dialog) leaked invisible EXCEL.EXE processes.~~ — Cleanup is now **two-tier** everywhere: (1) graceful `DisplayAlerts(false)` + `Quit` deferred inside the `sugar.Do` block, then (2) a PID-based force-kill registered via `t.Cleanup` (`internal/testutil.EnsureProcessExited` — waits up to 5 s, then `TerminateProcess`). Every new Excel-spawning test must follow this contract; the shared harnesses (`withApp`/`withBook`/`withSheet` in `excel/harness_test.go`, `setupExcel` in `sugar_com_test.go`) implement it already — prefer them over hand-rolled setup.

Resolved in v0.9.2 (2026-06-18, getter refactor / R34):

* ~~~84 scalar property getters each hand-rolled the `c.Get(prop).Value()` + coerce pattern, and several bool getters bypassed `toBool` with a bare `v.(bool)` assert — silently returning `false` for legacy 0/-1 VARIANT shapes.~~ — `excel/helpers.go` now exposes generic `getInt32`/`getFloat64`/`getBool`/`getString(c sugar.Chain, prop string, params ...interface{})` and every scalar getter delegates to them in one line. The bool getters (`Font.Bold`/`Font.Italic`, `Range.MergeCells`, `Workbook.Saved`) now route through `getBool`/`toBool`, fixing the 0/-1 coercion. The local `shape.shapeFloat` and `chart.getFloat` helpers were removed in favor of `getFloat64`. Output-neutral for every non-bool getter; no public signatures changed.

Resolved in v0.9.3 (2026-06-18, wrap-constructor refactor / R35):

* ~~~77 hand-written `&T{...}` wrapper-construction literals were open-coded across `excel/*.go` setters and child-object accessors, so the "chain -> typed wrapper" convention was duplicated at every site.~~ — Each wrapper type now owns a one-line `wrapT(c sugar.Chain) T` constructor returning the interface type (`wrapRange`, `wrapWorksheet`, `wrapChart`, ...); `pictures` keeps its parent-sheet field via `wrapPictures(c, sheet sugar.Chain)`. All 77 literal sites route through the matching `wrapT(...)`. The two sites that mutated a concrete `*T` after construction (`Worksheets.Add`, `Pictures.Add` set `Name` post-build) were restructured to finalize the chain before wrapping once. Behavior-neutral; no public signatures changed.

Resolved in v0.9.4 (2026-06-19, optional-arg refactor / R37):

* ~~The "concat required + optional args, trim trailing `Missing`, then `Call`" idiom for positional-optional COM methods was hand-rolled at each site (`SaveAs`, `Workbooks.Open`, `Close`), so trailing-Missing handling lived in three places.~~ — `excel/helpers.go` now exposes `callOptional(c sugar.Chain, method string, leading []interface{}, optional ...interface{}) sugar.Chain` (required+optional concat → `trimTrailingMissing` → `Call`); `trimTrailingMissing` moved from `workbook.go` to `helpers.go` with its guard generalized (`last>=0`) so an all-`Missing` call trims to an empty arg list (supports no-required methods like `Close`). **New positional-optional COM wrappers (PasteSpecial/Run/Copy parity work) should call `callOptional` rather than re-rolling the trim.** Behavior-neutral; `go-ole` import dropped from `workbook.go`.

Resolved post-v0.8.7 (2026-07-24, review LOW/NIT batch):

* ~~A `VT_UNKNOWN` Invoke result was classified as a value chain: `Value()` returned the arena-owned raw `*IUnknown` with no AddRef (use-after-free after Release), and `Store()` could not recover it.~~ — `handleResult` now promotes `VT_UNKNOWN` via `QueryInterface(IID_IDispatch)` (same resolution `ForEach` uses); a non-IDispatch IUnknown degrades to an empty chain, mirroring the `Nothing` convention. `Value()` additionally demotes any stray `VT_UNKNOWN` to nil (getElement's object-cell convention). Regression: `TestVTUnknownResultNotLeakedAsValue`.
* ~~`excel.Names.Contains` folded *every* COM failure into `(false, nil)`, so a disconnected server or access error masqueraded as "name absent".~~ — Only the not-found classes (`DISP_E_BADINDEX`, Excel `0x800A03EC`, including the `DISP_E_EXCEPTION`-wrapped EXCEPINFO delivery) return `(false, nil)`; any other error propagates. Helper `isNameNotFound` with unit test `TestIsNameNotFound`.
* ~~`Chain.IsDispatch()` inspected only `lastResult`, so a chain from `From`/`Create`/`Fork`/`ForEach` (which set `disp` but not `lastResult`) reported `false`.~~ — Widened to `disp != nil || lastResult is VT_DISPATCH`; Nothing/scalar chains still report `false`. Regressions: `TestIsDispatch_ObjectChains`, `TestIsDispatch_NilChain`.
* ~~The runner's nested-scope flag was a bare bool in the context value, which propagates across goroutine boundaries — a captured `Context` could silently authorize skipping COM init on a different, un-initialized thread.~~ — The flag now stores the initializing thread's `GetCurrentThreadId()`; a scope is nested only when the current thread id matches, so a cross-thread `Context.Do` re-initializes. `Context.Go` is unaffected (it sets `forceInit`). Regressions: `TestNestedScope_CrossThreadReinitializes`, `TestNestedScope_SameThreadStillSkipsInit`.
* NITs: `evalBinary`'s unsupported-op message unified to `%T %s %T` (output-neutral; `TestEval_UnsupportedBinaryMessage`); `Chain.Put` now `VariantClear`s the propput result VARIANT (defensive against servers returning an allocating type); `Worksheet.Range` errors on >2 cell arguments instead of silently dropping them (`TestWorksheetRange_TooManyArgs`); the `win32.go` `NewCallback` "~2000 cap" comment refreshed (the cap figure is stale — Go 1.26.3 absorbs 200k+; the thunk leak, not a crash, is what the package-var hoist prevents).

Resolved post-v0.8.8 (2026-07-26, review MED batch):

* ~~`Worksheet.Clear` / `Worksheet.ClearContents` failed 100% of the time since release: they read the `Cells` **property** with `Call` (`DISPATCH_METHOD`), which Excel rejects with `DISP_E_MEMBERNOTFOUND`.~~ — Both now use `Get("Cells").Call("Clear"/"ClearContents")`. Note the asymmetry: `Range.Clear`/`Range.ClearContents` (`range.go`) are genuine methods and correctly stay on `Call`. Regression: `excel/worksheet_test.go` (`TestWorksheet_Clear`, `TestWorksheet_ClearIsWholeSheet`) — the file did not exist before, so `Worksheet` had no integration coverage at all (§5 rule 9 gap, now closed with `Name`/`Index`/`Visible`/`AutoFit` tests too).
  **Recurrence barrier:** `excel/dispatch_kind_test.go` freezes the *member × DISPATCH kind* table (`dispatchKinds`) for every Excel member the package names, and statically scans the package's own non-test sources (`go/ast`) for `Get`/`Put`/`Call` and for the member-name-taking helpers (`getInt32`/`getFloat64`/`getBool`/`getString`/`callOptional`). A wrong verb, or an unclassified new member, fails `go test ./...` with no Excel installed. The traps it pins: `Cells`, `Range`, `Offset`, `Resize`, `End` are argumented *properties*; `ChartObjects`, `Pictures` are collection accessors that are *methods*; `Item` is ambiguous (property on `Sheets`/`Workbooks`, method on `Names`/`Charts`/`Shapes`/`Pictures`) and is exempt by name.
* ~~`getString` ran every VARIANT through `fmt.Sprint`, so a scalar string getter on a multi-cell object forged Go-syntax text: `Range("A1:B1").Formula()` returned the string `"[[=1+1 =2+2]]"`.~~ — `stringFromVariant` now rejects SAFEARRAY results (`[]interface{}` / `[][]interface{}`) with an explicit error naming the property and the array shape. No false positives: no Excel string property can legitimately decode to a slice. Tests: `helpers_test.go` (Excel-free) + `TestRange_FormulaMultiCellErrors`.
  Deliberately **not** done: promoting a `nil` result to an error. The `helpers.go` getters only see an already decoded `interface{}`, where `VT_NULL` (a mixed multi-cell read) and `VT_EMPTY` (unset property, a legitimate empty string) are both `nil` — the distinction only exists at the `decodeVariantScalar` layer in the root `safearray.go`. Doing it properly means a `VT_NULL` sentinel in the core decoder, which changes `Range.Value()` output for every consumer; it needs its own design pass and sign-off.
* ~~`Range.Find(what)` passed only `What`, so it inherited Excel's session-wide search state.~~ — Excel saves `LookIn`, `LookAt`, `SearchOrder` and `MatchByte` for the life of the Excel session (the Find dialog's sticky settings), so omitting them means "reuse the last search", not "use the default": the same call could match whole cells in one session and substrings in the next. All four are now pinned to the pristine-session values (`xlFormulas` / `xlPart` / `xlByRows` / `MatchByte=False`), so behavior is unchanged for a fresh Excel and no longer inheritable. `SearchDirection` and `MatchCase` are *not* persisted, but measurement against live Excel shows `Find` rejects the `DISP_E_PARAMNOTFOUND` marker with `DISP_E_TYPEMISMATCH` in every slot after `SearchOrder` (only `After` tolerates it), so they are passed at their per-call defaults (`xlNext`, `False`) — which also keeps each argument in its own slot, the positional hazard being that a dropped argument slides `MatchByte` into `SearchDirection`. xlwings parity note: xlwings has **no** `Range.find` (verified against the API docs and `xlwings/main.py`+`_xlwindows.py` on main), so there is no xlwings default to mirror; the doc comment records this as a sugar extension. Regressions: `TestRange_FindIgnoresSessionSearchState`, `TestRange_FindArgumentSlots`.

Resolved post-v0.8.9 (2026-07-26, SAFEARRAY bulk transfer):

* ~~2-D SAFEARRAY encode/decode called `SafeArrayGetElement`/`SafeArrayPutElement` **per cell** — a `syscall.LazyProc.Call` plus a VARIANT deep copy for every cell of a `Range.Value` block (~350 ns/cell).~~ — `safearray.go` now locks the array once with `SafeArrayAccessData` and walks the element buffer as a `[]ole.VARIANT` view (`accessVariantData` + `decodeArrayCell` / `fill1D` / `fill2D`). `scalarToVariant` also stopped calling `ole.VariantInit` (an oleaut32 call) per cell — a zeroed VARIANT *is* VT_EMPTY. Bulk encode additionally hands its BSTRs to the array by value instead of letting `SafeArrayPutElement` deep-copy and then freeing the temporary.

  **The size threshold matters.** A 1x1 or few-hundred-cell read was always rounding error; the change only shows up from ~10k cells. Measured (Ryzen 9 3900X, real Excel, cross-process attach):

  | grid | `Range.Value()` e2e | decode share | `SetValue` e2e |
  | --- | --- | --- | --- |
  | 100x100 numeric | 8.9 ms → **4.3 ms** | 40% → 6% | 89 ms → 39 ms |
  | 500x500 numeric | 172 ms → **100 ms** | 52% → 6% | 414 ms → 308 ms |
  | 100x100 string  | 11.6 ms → **7.5 ms** | 50% → 33% | 80 ms → 48 ms |
  | 500x500 string  | 328 ms → **214 ms** | 56% → 35% | 1.66 s → 1.46 s |

  Microbenchmarks (`BenchmarkEncode2D` / `BenchmarkDecode2D` in `safearray_bench_test.go`): numeric decode 89 ms → 5.3 ms (**17x**, allocs 1.50M → 250k), numeric encode 77 ms → 3.1 ms (**25x**, allocs 1.25M → 10). String grids gain only ~2x because the Go `string` allocation in `UTF16ToString`, not the COM call, dominates them — that is the floor, not a missed optimization.

  **Layout hazard, pinned by tests.** A SAFEARRAY element buffer is column-major: the dimension-1 index (rows) varies fastest, so cell (r, c) sits at `c*rows + r`, **not** `r*cols + c`. Getting it backwards transposes every block read/write and is invisible on square grids. `TestSafeArrayDataLayout` asserts the order against oleaut32; `TestBulkMatchesPerElement`(`1D`) cross-checks both directions against the `SafeArrayGetElement`/`PutElement` results (the OS's own index arithmetic) on asymmetric shapes; `TestRange_LargeGridRoundTrip` (integration) pins it against live Excel with a 200x37 grid whose cells encode their own coordinates.

  Every pre-existing guard is intact: `VT_BYREF` array rejection, typed-SAFEARRAY rejection (`SafeArrayGetVartype != VT_VARIANT`), ragged-row rejection, the `VT_DISPATCH`/`VT_UNKNOWN` → nil degradation (now in `decodeArrayCell`), and the `decodeVariantScalar` route for `VT_CY`/`VT_DECIMAL`/`VT_ERROR`. New lifetime rules: read cells are **not** cleared (the array owns them, unlike the copies `SafeArrayGetElement` returns), and every `accessVariantData` lock is released by a `defer`ed unlock in a scope that closes *before* the caller's `SafeArrayDestroy` (a locked array fails destroy with `DISP_E_ARRAYISLOCKED`).

Resolved post-v0.8.9 (2026-07-29, ecosystem review batch):

* ~~`Expand("down")` / `Expand("right")` collapsed a multi-cell anchor's opposite axis to 1, silently truncating the read.~~ — `expandFromEnd` (`excel/options.go`) built its rectangle from two addresses that both sat in the anchor's origin column (down) or origin row (right), because it reduced the anchor to `Cells(1,1)` and never read its `Rows`/`Columns` span. `sheet.Range("A1:C1").Options(Expand("down")).Value()` therefore returned A1:A10 (10x1) instead of A1:C10, with `err == nil` — and `NDimAuto` flattened the 1-column result to a `[]interface{}`, so a caller had no way to detect the loss. The canonical idiom `Range("A1:C1").Options(Header(true), Expand("down")).Get(&rows)` produced structs whose 2nd and 3rd fields were all zero (measured against live Excel: `[{alice 0 } {bob 0 }]`). Both directions now build the same two-corner bounding box the `"table"` branch always used: the `End(direction)` cell is shifted by `crossSpan-1` on the perpendicular axis (`crossSpan` = `Columns().Count()` for down, `Rows().Count()` for right) before its address is resolved, so the rectangle keeps the anchor's width/height. This is also the upstream behavior — xlwings' `VerticalExpander` ends its range at `(end_row, rng.column + rng.shape[1] - 1)`. Corner addresses are still passed to `Worksheet.Range` as **strings**, per the existing "don't marshal chains as COM parameters" note in `applyExpand`. A 1x1 anchor is unchanged (`expandCornerOffset` returns no shift, so no extra COM call). Regressions: `excel/options_expand_test.go` — Excel-free, driving the real expansion through a recording fake `sugar.Chain` over an in-memory grid, so it asserts the exact `Worksheet.Range(cell1, cell2)` corners and pins "down"/"right" against "table" on a rectangular block — plus four live-Excel cases in `options_integration_test.go` (the pre-existing nine Expand cases all used single-cell anchors and could not see the defect).
* ~~`ForEach` called `AddRef` on a `VT_DISPATCH` item whose pointer was NULL, panicking on a nil dereference.~~ — go-ole's `ToIDispatch` returns a nil `*IDispatch` for a `VT_DISPATCH` VARIANT with `Val == 0` (COM `Nothing`), and `AddRef` then dereferences `RawVTable`. The VT_UNKNOWN branch immediately below already had the guard, and `handleResult` already handles a Nothing *result* "instead of panicking on AddRef(nil)" — the asymmetry sat inside one function. A nil item now falls through to the existing "not an object" branch: the VARIANT is Cleared and iteration continues. Regression: `foreach_nil_item_test.go` — a hand-built COM server (fake `IDispatch` exposing `_NewEnum` + fake `IEnumVARIANT`) scripts a null-dispatch item, a real object and a scalar, and asserts only the real object reaches the callback and that every reference the fake handed out came back. It needs neither Excel nor CoInitialize (direct vtable dispatch + `VariantClear`). Note for future COM-server fakes: the `uintptr` -> typed pointer conversions a `syscall.NewCallback` thunk forces trip `go vet`'s `unsafeptr` check, so the file funnels them through one documented `fakePtr[T]` helper to keep `go vet ./...` clean.

Resolved 2026-08-03:

* ~~`Expand` truncated any block whose top-left cell was empty — the commonest table layout there is.~~ — `endpointCell` (`excel/options.go`) probed the neighbor for blankness and then called `End(direction)` from the **origin**. That is safe only while the origin is non-empty: Excel's `End()` from an *empty* cell lands on the first non-empty cell rather than the last cell of a run. So a sheet with an empty A1 corner, headers along row 1 and labels down column A read **A1:B2 — 4 cells out of 12, `err == nil`** (measured against live Excel: `[[<nil> Jan] [North 1]]`). Upstream xlwings uses a **three-rung ladder** in all three expanders (`expansion.py`, read verbatim rather than from memory): neighbor blank → origin is its own endpoint; **second** neighbor blank → the neighbor is the endpoint; otherwise `neighbor.end(dir)`. `endpointCell` now mirrors it exactly, which both fixes the blank origin (`End()` is only ever called from a proven-non-empty cell) and needs the middle rung to stay correct on a two-cell block (`End()` from the neighbor would sail past it to row 1048576 or the next data island). For a non-empty origin every rung agrees with the old single probe, so no case that already worked moved. Regressions: `options_expand_test.go` `TestExpand_BlankOriginDoesNotTruncate` / `TestExpand_TwoCellBlockStopsAtTheBlock` plus live-Excel `TestOptions_ExpandTable_BlankCorner` / `TestOptions_ExpandDown_TwoCellBlock`; both mutations (drop rung 3's start cell, drop rung 2) were verified to fail named tests, and the blank-corner case was also failed against real Excel before the fix.
  * **The fake harness had to be fixed before it could fail.** `fakeCell.end` implemented only Excel's contiguous-run case and its comment justified that with "the only case the blank-neighbor guard lets through" — which was precisely the false assumption the bug lived in. A fake that cannot reproduce End()-from-blank would have shown the *old* code as correct. It now models all three of Excel's cases and parks on the real sheet edge (XFD1048576) so "End() ran off into empty space" is an observable landing spot, and `TestFakeEnd_ModelsExcelsThreeCases` guards the harness itself.
  * **The backlog premise was wrong; the defect next to it was worse.** The queued item claimed upstream probes `rng(rng.shape[0]+1, 1)` (the cell below the anchor's *last* row) and therefore that `Expand("down")` on `A1:A3` with data only in A1:A2 should return A1:A3. Upstream does no such thing — it probes `rng(2, 1)`, origin-relative, exactly as sugar did, and returns A1:A2 too. **Anchor-span-relative probing must not be re-proposed as parity work**; there is no upstream behavior behind it. Reading the upstream source verbatim (instead of trusting the item's paraphrase) is what turned a non-defect into a real one.
  * Not adopted: xlwings' `origin.has_array` branch (legacy CSE array formulas take `end()` unguarded). sugar has no `has_array` notion, and a spilled dynamic array is a different object; out of scope until someone needs CSE anchors.

* ~~`ForEach` reported success after the enumeration FAILED, so a truncated collection looked complete.~~ — `sugar.go`'s loop was `if err != nil || fetched == 0 { break }`, collapsing "no more items" and "the enumeration broke" into one exit; `ForEach` then returned the RECEIVER, so `Err()` stayed nil and the caller processed a short set believing it had everything. Every other failure in that function propagates (the `_NewEnum` acquisitions, the callback's own error), making this the single silent exit. Fixed by splitting the branch.
  * **The obvious fix is wrong, and that is the part worth remembering.** `if err != nil { return error }` breaks EVERY healthy call: go-ole builds its error from `if hr != 0`, and `IEnumVARIANT::Next` signals normal exhaustion with **S_FALSE (hr == 1) — a SUCCESS code**. A non-nil go-ole error is therefore not evidence of failure, and the original lumped branch was defensible for exactly that reason. The predicate is COM's `FAILED()`: new helper `comFailed(err)` unwraps `*ole.OleError` and tests `int32(code) < 0`. Three healthy-input tests failed the moment the naive version was tried, which is how this was caught instead of shipped.
  * **Applies beyond ForEach:** any go-ole call whose HRESULT can be a non-S_OK SUCCESS must go through `comFailed`, never `err != nil`.
  * Regressions in `foreach_nil_item_test.go`: `TestForEach_PropagatesEnumNextError` (scripted enumerator fails at item 2 of 3; asserts the callback saw exactly 2, `Err() != nil`, and no refcount leak on the error path) and `TestForEach_CleanExhaustionStillSucceeds` (so "propagate" cannot be satisfied by failing everything). The fake `IEnumVARIANT` gained a `failAtPos` field — it could previously only produce clean exhaustion, so the distinction being fixed was **untestable with the old harness**. Mutation to the pre-fix shape fails the named test.

Still open:

* Integration tests boot a fresh Excel per test (~2–3 s each, full tagged suite ≈183 s). A shared instance must respect COM apartment rules: tests run on arbitrary goroutines, so sharing raw IDispatch is unsafe — the workable design is one Excel process + per-test `GetActive` attach (ROT marshaling). Worth doing once suite time hurts.

## 7. Documentation Standards

*   **Timestamps**: When recording timestamped notes (changelog, design records), use the current absolute date — do not rely on relative phrases that age poorly.
*   **READMEs**: Keep README.md examples copy-paste-runnable. If you change a public API in §2.1, update both README.md and any docstring referencing the old form.
