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

// The context-mechanics tests below use lightweight scripting COM servers
// (Scripting.Dictionary / Scripting.FileSystemObject) instead of Excel: they
// exercise arena lifecycle, nesting, and async dispatch — none of which is
// Excel-specific — so `go test ./...` runs them on any Windows host without
// spawning Office processes.

func TestContext_Lifecycle(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		subCtx := sugar.NewContext(ctx)
		defer subCtx.Release()

		dict := subCtx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			t.Fatalf("Scripting.Dictionary create failed: %v", err)
		}

		if err := dict.Call("Add", "k", "v").Err(); err != nil {
			t.Errorf("failed: %v", err)
		}
		return nil
	})
}

func TestContext_NestedDo(t *testing.T) {
	err := sugar.Do(func(ctx sugar.Context) error {
		fso := ctx.Create("Scripting.FileSystemObject")
		if err := fso.Err(); err != nil {
			t.Fatalf("Scripting.FileSystemObject create failed: %v", err)
		}

		err := ctx.Do(func(innerCtx sugar.Context) error {
			// GetSpecialFolder(2) = TemporaryFolder — an IDispatch result the
			// inner arena takes ownership of via Track(Fork()).
			folder := innerCtx.Track(fso.Call("GetSpecialFolder", 2).Fork())
			if err := folder.Err(); err != nil {
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
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			wg.Done()
			t.Fatalf("Scripting.Dictionary create failed: %v", err)
		}

		ctx.Go(func(asyncCtx sugar.Context) error {
			defer wg.Done()
			asyncDict := asyncCtx.Create("Scripting.Dictionary")
			if err := asyncDict.Err(); err != nil {
				t.Errorf("Async COM creation failed: %v", err)
				return err
			}
			return asyncDict.Call("Add", "k", "v").Err()
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