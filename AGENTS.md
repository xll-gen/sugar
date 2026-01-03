# Agent Instructions for the `sugar` Repository

This document provides guidance for developers and AI agents working on the `sugar` codebase.

## 1. Project Overview

`sugar` is a Go library that provides a fluent, chainable API for Component Object Model (COM) automation on Windows. It acts as a "syntactic sugar" wrapper around the `github.com/go-ole/go-ole` library, simplifying common tasks and improving code readability by reducing boilerplate.

**Core Features:**
- A fluent, chainable interface for COM operations.
- Simplified resource management for COM objects.
- Wrappers for common `oleutil` functions like `GetProperty`, `CallMethod`, and `PutProperty`.
- An optional automatic resource management mode using Go's garbage collector.

**Platform Constraint:** This library is Windows-specific due to its dependency on `go-ole`. All Go source files **must** include the `//go:build windows` build constraint at the top.

## 2. Key API Concepts & Rules

Understanding the design patterns of the `sugar` API is crucial for using and extending it correctly.

### 2.1 The `Chain` Struct

The `*Chain` is the central object of the library. All operations start with a `Chain` and all method calls (except terminal methods) return the `Chain` to allow for fluent chaining.

### 2.2 Initiating a Chain

There are three ways to start a chain, each with different ownership semantics:

1.  **`Create(progID string)`**:
    - Creates a new COM object (e.g., "Excel.Application").
    - **The chain takes ownership** of the created object and is responsible for its release.

2.  **`GetActive(progID string)`**:
    - Attaches to a running instance of a COM object.
    - **The chain takes ownership** of the object and is responsible for its release.

3.  **`From(disp *ole.IDispatch)`**:
    - Starts a chain from a pre-existing `IDispatch` object.
    - **The caller retains ownership** of the initial `disp` object and is responsible for releasing it separately. The chain only manages objects created *during* the chain.

### 2.3 Terminal Methods (Crucial for Resource Management)

To prevent memory leaks, every chain of operations **must** end with one of the following terminal methods. These methods handle the release of all COM objects acquired *by the chain* during its operations (unless in `AutoRelease` mode).

1.  **`Release()`**:
    - Releases all intermediate COM objects acquired by the chain.
    - Returns the first error that occurred, if any.
    - Use this when you are not interested in a return value from the chain.

2.  **`Value()`**:
    - Retrieves the result of the final `Get` or `Call` operation as a Go `interface{}`.
    - Releases all intermediate COM objects.
    - Returns the value and any error that occurred during the chain.

3.  **`Err()`**:
    - A synonym for `Release()`. It makes code more readable when the primary purpose is to check for an error at the end of a chain that doesn't return a value.

**Example (Correct Usage):**
```go
// Correct: Chain is terminated with Release()
err := Create("Excel.Application").Put("Visible", true).Release()
```
**Example (Incorrect Usage - LEAKS MEMORY):**
```go
// Incorrect: No terminal method is called. The Excel.Application object is leaked.
Create("Excel.Application").Put("Visible", true)
```

### 2.4 Resource Management Modes

1.  **Manual Mode (Default)**:
    - The user **must** call a terminal method (`Release`, `Value`, `Err`) to free resources.
    - Intermediate `IDispatch` objects returned by `Get` or `Call` are tracked in an internal `releaseChain` slice and released in reverse order of creation by the terminal method.

2.  **Automatic Mode (`AutoRelease()`)**:
    - This is an **opt-in** mode, enabled by calling `.AutoRelease()` on a chain.
    - In this mode, the Go garbage collector is responsible for releasing `IDispatch` objects acquired during the chain via `runtime.SetFinalizer`.
    - While a terminal method is not strictly required for resource cleanup in this mode, it is still best practice to use `Value()` or `Err()` to retrieve results and check for errors.

### 2.5 Storing Intermediate Objects (`Store()`)

The `Store(target **ole.IDispatch)` method allows you to extract a COM object from the middle of a chain for later reuse.

**OWNERSHIP RULE:** When an object is extracted via `Store`, **the user becomes responsible for calling `.Release()` on that object**, regardless of whether the chain is in manual or `AutoRelease` mode. The library removes the object from its internal resource management.

## 3. Development and Testing

### 3.1 COM Lifecycle

The `sugar` library does **not** handle the global COM library initialization. The user of the library is responsible for calling `ole.CoInitialize(0)` at the start of their program and `ole.CoUninitialize()` at the end. This is a fundamental requirement of `go-ole`.

### 3.2 Testing

- Tests are located in `example_test.go` and serve as both integration tests and usage examples.
- **Requirement:** Running the tests requires **Microsoft Excel to be installed** on the Windows machine.
- The `excelTestSetup()` helper function in the test file is a good reference for the correct setup and teardown procedure, including `CoInitialize` and `CoUninitialize`.

## 4. Directory Structure

- **`sugar.go`**: Contains the core logic for the `Chain` struct and all its methods.
- **`excel/`**: A placeholder directory intended for future Excel-specific helper functions that build on the generic `sugar` chain (e.g., `Workbooks.Add()`).
- **`example_test.go`**: Contains all tests and usage examples, demonstrating the API's functionality.
