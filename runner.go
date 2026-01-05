//go:build windows

package sugar

import (
	"context"
	"runtime"

	"github.com/go-ole/go-ole"
)

// Runner configures the execution environment for COM operations.
type Runner struct {
	parent    context.Context
	forceInit bool
}

// With returns a new Runner with the specified parent context.
func With(ctx context.Context) *Runner {
	return &Runner{parent: ctx}
}

// Do executes the provided function in the current goroutine.
// It ensures proper COM initialization and thread locking only if not already initialized.
func (r *Runner) Do(fn func(ctx *Context) error) (err error) {
	if r.parent == nil {
		r.parent = context.Background()
	}

	// Only skip initialization if we are NOT forced and already inside a sugar.Do block
	isNested := !r.forceInit && r.parent.Value(activeSugarKey) != nil

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

	return fn(ctx)
}

// Go executes the provided function in a new goroutine.
// It always performs full initialization as it's a new thread.
func (r *Runner) Go(fn func(ctx *Context) error) {
	go func() {
		// Create a new runner that forces initialization for the new goroutine
		runner := &Runner{
			parent:    r.parent,
			forceInit: true,
		}
		_ = runner.Do(fn)
	}()
}

// Do executes the provided function with a Background context.
func Do(fn func(ctx *Context) error) error {
	return With(context.Background()).Do(fn)
}

// Go executes the provided function in a new goroutine with a Background context.
func Go(fn func(ctx *Context) error) {
	With(context.Background()).Go(fn)
}
