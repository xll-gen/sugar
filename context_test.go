//go:build windows

package sugar_test

import (
	"context"
	"testing"

	"github.com/xll-gen/sugar"
)

func TestContext_Lifecycle(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) {
		// Test manual context creation within Do
		subCtx := sugar.NewContext(ctx)
		defer subCtx.Release()

		excel := subCtx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			return
		}
		defer excel.Call("Quit")

		if err := excel.Put("Visible", false).Err(); err != nil {
			t.Errorf("failed: %v", err)
		}
	})
}

func TestContext_NestedDo(t *testing.T) {
	// Outer Do
	err := sugar.Do(func(ctx *sugar.Context) {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			return
		}
		defer excel.Call("Quit")

		// Inner Do using parent context's Do method for proper nesting
		err := ctx.Do(func(innerCtx *sugar.Context) {
			// This should be safe and share the thread/COM init
			wb := innerCtx.Track(excel.Get("Workbooks").Call("Add").Fork())
			if err := wb.Err(); err != nil {
				t.Errorf("inner Do failed: %v", err)
			}
			// wb is released when inner Do returns
		})
		
		if err != nil {
			t.Errorf("nested Do returned error: %v", err)
		}
	})
	
	if err != nil {
		t.Errorf("outer Do returned error: %v (type %T)", err, err)
	}
}

func TestContext_WithCancel(t *testing.T) {
	// Test standard context integration
	stdCtx, cancel := context.WithCancel(context.Background())
	cancel() // cancel immediately

	sugar.With(stdCtx).Do(func(ctx *sugar.Context) {
		select {
		case <-ctx.Done():
			// Success: context was correctly passed and is cancelled
		default:
			t.Error("context should have been cancelled")
		}
	})
}