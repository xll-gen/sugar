//go:build windows

package sugar_test

import (
	"context"
	"errors"
	"sync"
	"testing"
	"time"

	"github.com/xll-gen/sugar"
)

func TestContext_Lifecycle(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
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
	err := sugar.Do(func(ctx sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			return nil
		}
		defer excel.Call("Quit")

		err := ctx.Do(func(innerCtx sugar.Context) error {
			wb := innerCtx.Track(excel.Get("Workbooks").Call("Add").Fork())
			if err := wb.Err(); err != nil {
				t.Errorf("inner Do failed: %v", err)
			}
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

	sugar.Do(func(ctx sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			t.Skip("Excel not available")
			wg.Done()
			return nil
		}
		defer excel.Call("Quit")

		ctx.Go(func(asyncCtx sugar.Context) error {
			defer wg.Done()
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
	stdCtx, cancel := context.WithCancel(context.Background())
	cancel()

	sugar.With(stdCtx).Do(func(ctx sugar.Context) error {
		select {
		case <-ctx.Done():
		default:
			t.Error("context should have been cancelled")
		}
		return nil
	})
}

// TestGo_ReturnsErrorChannel verifies that the previously fire-and-forget
// sugar.Go now surfaces the goroutine's terminal error through a returned
// channel. Regression for v0.7.0 which made the signature
//
//	func Go(...) <-chan error
//
// Prior to this version the goroutine's error was silently dropped, masking
// COM init failures and panics-wrapped-as-errors.
func TestGo_ReturnsErrorChannel(t *testing.T) {
	wantErr := errors.New("intentional")
	done := sugar.Go(func(ctx sugar.Context) error { return wantErr })

	select {
	case got := <-done:
		if got == nil || got.Error() != wantErr.Error() {
			t.Errorf("expected %v, got %v", wantErr, got)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("sugar.Go did not signal completion within 5s")
	}

	// Channel must be closed after the single value.
	if _, ok := <-done; ok {
		t.Error("expected channel to be closed after delivering the error")
	}
}

// TestGo_NilOnSuccess covers the success path: a clean return must produce
// nil on the channel and then close it. This catches accidental double-send
// or close-before-send refactors.
func TestGo_NilOnSuccess(t *testing.T) {
	done := sugar.Go(func(ctx sugar.Context) error { return nil })
	select {
	case got := <-done:
		if got != nil {
			t.Errorf("expected nil, got %v", got)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("sugar.Go did not signal completion within 5s")
	}
}