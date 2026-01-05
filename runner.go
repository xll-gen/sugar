//go:build windows

package sugar

import (
	"context"
	"runtime"

	"github.com/go-ole/go-ole"
)

// Runner configures the execution environment for COM operations.
type Runner struct {
	parent context.Context
}

// With returns a new Runner with the specified parent context.
func With(ctx context.Context) *Runner {
	return &Runner{parent: ctx}
}

// Do executes the provided function in the current goroutine.
// It ensures proper COM initialization and thread locking only if not already initialized.
// A new sugar.Context (arena) is created for this scope.
func (r *Runner) Do(fn func(ctx *Context)) (err error) {
	if r.parent == nil {
		r.parent = context.Background()
	}

	// Check if we are already inside a sugar.Do block on this call stack/context
	isNested := r.parent.Value(activeSugarKey) != nil

	if !isNested {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()

		if err := ole.CoInitialize(0); err != nil {
			return err
		}
		defer ole.CoUninitialize()
	}

	// Create a sub-context marked as active, and a new arena
	innerStdCtx := context.WithValue(r.parent, activeSugarKey, true)
	ctx := NewContext(innerStdCtx)
	
	defer func() {
		releaseErr := ctx.Release()
		if err == nil {
			err = releaseErr
		}
	}()

	fn(ctx)
	return nil
}

// Go executes the provided function in a new goroutine.
// It always performs full initialization as it's a new thread.
func (r *Runner) Go(fn func(ctx *Context)) {
	go func() {
		// New goroutine always starts fresh, ignoring parent's isNested status 
		// but using the parent context for cancellation/values.
		_ = With(r.parent).Do(fn)
	}()
}

// Do executes the provided function with a Background context.
func Do(fn func(ctx *Context)) error {
	return With(context.Background()).Do(fn)
}

// Go executes the provided function in a new goroutine with a Background context.
func Go(fn func(ctx *Context)) {
	With(context.Background()).Go(fn)
}