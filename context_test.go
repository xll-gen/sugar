//go:build windows

package sugar_test

import (
	"context"
	"sync"
	"testing"

	"github.com/xll-gen/sugar"
)

func TestContext_Lifecycle(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) error {
		// Test manual context creation within Do
		subCtx := sugar.NewContext(ctx)
		defer subCtx.Release()

		excel := subCtx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			return nil
		}
		defer excel.Call("Quit")

		if err := excel.Put("Visible", false).Err(); err != nil {
			t.Errorf("failed: %v", err)
		}
		return nil
	})
}

func TestContext_NestedDo(t *testing.T) {
	// Outer Do
	err := sugar.Do(func(ctx *sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			return nil
		}
		defer excel.Call("Quit")

		// Inner Do using parent context's Do method for proper nesting
		err := ctx.Do(func(innerCtx *sugar.Context) error {
			// This should be safe and share the thread/COM init
			wb := innerCtx.Track(excel.Get("Workbooks").Call("Add").Fork())
			if err := wb.Err(); err != nil {
				t.Errorf("inner Do failed: %v", err)
			}
			// wb is released when inner Do returns
			return nil
		})
		
		if err != nil {
			t.Errorf("nested Do returned error: %v", err)
		}
		return nil
	})
	
	if err != nil {
		t.Errorf("outer Do returned error: %v (type %T)", err, err)
	}
}

func TestContext_AsyncGo(t *testing.T) {
	var wg sync.WaitGroup
	wg.Add(1)

	// Outer Do
	sugar.Do(func(ctx *sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			wg.Done()
			return nil
		}
		defer excel.Call("Quit")

		// Start an async task using Go
		ctx.Go(func(asyncCtx *sugar.Context) error {
			defer wg.Done()
			
			// Let's create a new object in the new thread to be safe.
			asyncExcel := asyncCtx.Create("Excel.Application")
			if err := asyncExcel.Err(); err != nil {
				t.Errorf("Async Excel creation failed: %v", err)
				return err
			}
			asyncExcel.Call("Quit")
			return nil
		})
		return nil
	})

	wg.Wait()
}

func TestContext_WithCancel(t *testing.T) {
	// Test standard context integration
	stdCtx, cancel := context.WithCancel(context.Background())
	cancel() // cancel immediately

	sugar.With(stdCtx).Do(func(ctx *sugar.Context) error {
		select {
		case <-ctx.Done():
			// Success: context was correctly passed and is cancelled
		default:
			t.Error("context should have been cancelled")
		}
		return nil
	})
}
