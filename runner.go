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

// COM HRESULTs that CoInitialize can return without the thread being unusable.
const (
	hrSFalse           = 0x00000001 // S_FALSE: already initialized on this thread
	hrRPCEChangedMode  = 0x80010106 // RPC_E_CHANGED_MODE: thread is already in a different apartment model
)

// initializeCOM calls CoInitialize and reports whether a matching
// CoUninitialize is owed. go-ole surfaces *any* non-zero HRESULT as an
// error, including two benign cases this library must tolerate:
//
//   - S_FALSE — the thread is already STA-initialized (common when the host
//     process, e.g. an XLL or GUI app, initialized COM first). The init
//     count was still incremented, so the caller owes a CoUninitialize.
//   - RPC_E_CHANGED_MODE — the thread is already initialized as MTA. COM
//     calls still work via implicit marshaling; no CoUninitialize is owed
//     because the call did not take a reference.
func initializeCOM() (needUninit bool, err error) {
	if err := ole.CoInitialize(0); err != nil {
		oleErr, ok := err.(*ole.OleError)
		if !ok {
			return false, err
		}
		switch oleErr.Code() {
		case hrSFalse:
			return true, nil
		case hrRPCEChangedMode:
			return false, nil
		}
		return false, err
	}
	return true, nil
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

		needUninit, err := initializeCOM()
		if err != nil {
			return err
		}
		if needUninit {
			defer ole.CoUninitialize()
		}
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