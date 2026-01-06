# Agent Instructions for the `sugar` Repository

This document provides guidance for developers and AI agents working on the `sugar` codebase.

## 1. Project Overview

`sugar` is a Go library designed to **sweeten your Windows automation**. It provides a fluent, immutable API for Component Object Model (COM) automation, using an **Arena (Context) pattern** for automatic resource management.

**Platform Constraint:** Windows-specific. All source files must include `//go:build windows`.

## 2. Core Concepts & Usage Patterns

### 2.1 Standard Execution (`sugar.Do` and `sugar.Go`)

The library enforces a standard way to execute COM operations to ensure `runtime.LockOSThread()` and `ole.CoInitialize()` are handled correctly.

```go
sugar.Do(func(ctx *sugar.Context) error {
    excel := ctx.Create("Excel.Application")
    // ... work with excel ...
    return nil
})
```

*   **`sugar.Do`**: Executes synchronously in the current goroutine.
*   **`sugar.Go`**: Executes in a new goroutine (new OS thread).
*   **Nested Scopes**: Use `ctx.Do(func(innerCtx *sugar.Context) error { ... })` to create local cleanup scopes.

### 2.2 The Immutable `Chain`

All methods on `*sugar.Chain` (`Get`, `Call`, `ForEach`, etc.) return a **NEW** `Chain` instance. The original instance remains unchanged.

```go
excel := ctx.Create("Excel.Application")
workbooks := excel.Get("Workbooks") // 'excel' still points to Application
wb := workbooks.Call("Add")         // 'workbooks' still points to Workbooks collection
```

### 2.3 Automatic Resource Management (Arena)

Every `Chain` created via a `sugar.Context` (or derived from one) is automatically tracked by that context. When the `sugar.Do` block completes, all tracked COM objects are released in reverse order.

**Manual `Release()` is unnecessary** within a `Do/Go` block.

### 2.4 Integration with `context.Context`

`sugar.Context` implements the standard `context.Context` interface. You can use it for cancellation, timeouts, and passing values.

```go
sugar.With(parentCtx).Do(func(ctx *sugar.Context) error {
    select {
    case <-ctx.Done():
        return ctx.Err()
    default:
        // ...
        return nil
    }
})
```

## 3. Expression Subpackage

The `expression` package allows navigating COM objects using string expressions (e.g., `"Workbooks.Add().ActiveSheet"`).

*   It uses `sugar.Context` under the hood if a tracked `Chain` is passed.
*   Intermediate objects created during expression evaluation are automatically managed by the chain's context.

## 4. Development Rules

1.  **Always use `sugar.Do`** for entry points.
2.  **Never manually call `CoInitialize`** unless implementing a low-level runner.
3.  **Prefer `ctx.Create`** over `sugar.Create` to ensure automatic tracking.
4.  **Immutable behavior**: Do not expect a `Chain` variable to change its internal state after a method call.
5. **Thread Safety**: Remember that `Go` routines start fresh threads; do not share raw `IDispatch` pointers across threads without proper COM marshaling (though `sugar.Go` makes creating thread-local objects easy).
6. **Language Requirement**: All documentation (including README and AGENTS.md) and code comments must be written in English.