//go:build windows

package sugar_test

import (
	"testing"

	"github.com/xll-gen/sugar"
)

func TestContext_Lifecycle(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	// 1. Initialize Context
	ctx := sugar.NewContext()
	
	// 2. Create chains using Context
	// Root chain
	root := ctx.From(disp)
	
	// Add a workbook (creates intermediate object)
	// We handle 'wb' manually just to prove we can mix, but ideally we fork or track.
	// But wait, root.Get().Call() returns the SAME root chain modified.
	// So intermediates are handled by 'root'.
	// Let's create a NEW object that needs tracking, e.g., via Fork.
	
	err := root.Get("Workbooks").Call("Add").Err()
	if err != nil {
		t.Fatalf("failed to add workbook: %v", err)
	}

	// Fork off a chain for the active sheet.
	// This creates a NEW chain object that must be released.
	sheetChain := ctx.Track(root.Get("ActiveSheet").Fork())
	
	// Use the forked chain
	if err := sheetChain.Put("Name", "ContextTest").Err(); err != nil {
		t.Errorf("failed to set sheet name: %v", err)
	}

	// 3. Release Context
	// This should release 'sheetChain' AND 'root'.
	// Since 'root' holds intermediates (like the added workbook result if we stored it, 
	// but here we just called methods), it releases them.
	// 'sheetChain' holds the ActiveSheet reference (via AddRef in Fork), so it releases that.
	if err := ctx.Release(); err != nil {
		t.Errorf("Context release failed: %v", err)
	}
	
	// 4. Verify (Implicit)
	// If double-free occurs, we panic. If leak occurs, it's hard to catch without checking ref counts explicitly,
	// but successful execution suggests basic logic is sound.
}

func TestContext_Create(t *testing.T) {
	// Simple test for ctx.Create (tracking Create)
	// Note: We skip if Excel not present, handled by Create error check usually, 
	// but here we assume environment is okay or we check err.
	
	ctx := sugar.NewContext()
	defer ctx.Release() // Cleanup even if test fails

	// We use a safe progID check or reuse the mock idea if possible, 
	// but integration test with Excel is standard here.
	// We'll just skip if Create fails (e.g. CI without Excel).
	
	// Intentionally create invalid to check safe handling
	c := ctx.Create("Invalid.ProgID")
	if c.Err() == nil {
		t.Error("expected error for invalid progID")
	}
	// ctx.Release should handle the error-state chain gracefully
}

func TestContext_Do(t *testing.T) {
	disp, cleanup := initExcel(t)
	defer cleanup()

	// Use sugar.Do to manage lifecycle automatically
	sugar.Do(func(ctx *sugar.Context) {
		root := ctx.From(disp)
		
		// Create a new workbook. 
		// We use Fork() to get a new chain for the workbook, and ctx.Track() to ensure it's released.
		// Note: root.Get(...).Call(...) modifies 'root'. To keep 'root' usable or separate, we Fork.
		// Even if we don't care about 'root', Fork returns a new Chain we can pass to Track.
		
		wb := ctx.Track(root.Get("Workbooks").Call("Add").Fork())
		
		if err := wb.Put("Saved", true).Err(); err != nil {
			t.Errorf("failed to set Saved property inside Do: %v", err)
		}
		
		// No need to defer wb.Release() or root.Release()
	})
	
	// If we reached here without panic, Do() worked and Release() was called.
}
