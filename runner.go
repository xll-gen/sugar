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
func (r *Runner) Do(fn func(ctx Context) error) (err error) {
	if r.parent == nil {
		r.parent = context.Background()
	}

	isNested := !r.forceInit && r.parent.Value(activeSugarKey) != nil

	if !isNested {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()

		if err := ole.CoInitialize(0); err != nil {
			return err
		}
		defer ole.CoUninitialize()
	}

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

// Go executes the provided function in a new goroutine and returns a
// buffered, closed-after-send channel that delivers the goroutine's terminal
// error (nil on success).
//
// Callers may ignore the returned channel for fire-and-forget semantics; the
// goroutine never blocks on the channel because it is buffered with cap 1.
// Use the channel when you need to know whether the async COM work
// succeeded — earlier versions of this library silently dropped the error.
func (r *Runner) Go(fn func(ctx Context) error) <-chan error {
	done := make(chan error, 1)
	go func() {
		defer close(done)
		runner := &Runner{
			parent:    r.parent,
			forceInit: true,
		}
		done <- runner.Do(fn)
	}()
	return done
}

// Do executes the function with a Background context.
func Do(fn func(ctx Context) error) error {
	return With(context.Background()).Do(fn)
}

// Go executes the function in a new goroutine with a Background context. The
// returned channel reports the goroutine's terminal error; see Runner.Go.
func Go(fn func(ctx Context) error) <-chan error {
	return With(context.Background()).Go(fn)
}