//go:build windows

package sugar

import (
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Chain provides a fluent interface for chaining OLE operations.
type Chain struct {
	disp         *ole.IDispatch
	err          error
	lastResult   *ole.VARIANT
	autoRelease  bool
	releaseChain []*ole.IDispatch
}

// From starts a new chain with the given IDispatch. The user is responsible
// for releasing the initial dispatch object.
func From(disp *ole.IDispatch) *Chain {
	return &Chain{
		disp: disp,
	}
}

// Create starts a new chain by creating a new COM object from the given
// ProgID. The user is responsible for calling ole.CoInitialize and
// ole.CoUninitialize. The chain takes ownership of the created object and
// will release it when a terminal method (Release, Value, Err) is called.
func Create(progID string) *Chain {
	unknown, err := oleutil.CreateObject(progID)
	if err != nil {
		return &Chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		return &Chain{err: err}
	}
	unknown.Release()

	// The chain is responsible for releasing the created object.
	return &Chain{
		disp:         disp,
		releaseChain: []*ole.IDispatch{disp},
	}
}

// GetActive starts a new chain by attaching to a running COM object from the
// given ProgID. The user is responsible for calling ole.CoInitialize and
// ole.CoUninitialize. The chain takes ownership of the object and will
// release it when a terminal method (Release, Value, Err) is called.
func GetActive(progID string) *Chain {
	unknown, err := oleutil.GetActiveObject(progID)
	if err != nil {
		return &Chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		return &Chain{err: err}
	}
	unknown.Release()

	// The chain is responsible for releasing the object.
	return &Chain{
		disp:         disp,
		releaseChain: []*ole.IDispatch{disp},
	}
}

// AutoRelease switches the chain to automatic resource management mode.
// Any subsequent IDispatch objects created by Get or Call will be released
// automatically by the garbage collector.
func (c *Chain) AutoRelease() *Chain {
	if c.err != nil {
		return c
	}
	c.autoRelease = true
	return c
}

// handleResult is a helper function to process the result of GetProperty and CallMethod.
func (c *Chain) handleResult(result *ole.VARIANT, err error) *Chain {
	if err != nil {
		c.err = err
		return c
	}

	if c.lastResult != nil {
		c.lastResult.Clear()
	}
	c.lastResult = result

	if result.VT == ole.VT_DISPATCH {
		newDisp := result.ToIDispatch()
		if c.autoRelease {
			runtime.SetFinalizer(newDisp, func(disp *ole.IDispatch) {
				disp.Release()
			})
		} else {
			c.releaseChain = append(c.releaseChain, newDisp)
		}
		c.disp = newDisp
	}

	return c
}

// Get is a wrapper for oleutil.GetProperty.
func (c *Chain) Get(prop string, params ...interface{}) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}
	result, err := oleutil.GetProperty(c.disp, prop, params...)
	return c.handleResult(result, err)
}

// Call is a wrapper for oleutil.CallMethod.
func (c *Chain) Call(method string, params ...interface{}) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}
	result, err := oleutil.CallMethod(c.disp, method, params...)
	return c.handleResult(result, err)
}

// Put is a wrapper for oleutil.PutProperty.
func (c *Chain) Put(prop string, params ...interface{}) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	_, err := oleutil.PutProperty(c.disp, prop, params...)
	if err != nil {
		c.err = err
	}
	// PutProperty returns no result to store.
	if c.lastResult != nil {
		c.lastResult.Clear()
		c.lastResult = nil
	}

	return c
}

// Store transfers ownership of the current IDispatch object to an external
// variable. The user becomes responsible for calling Release on the stored
// object, regardless of whether the chain is in AutoRelease mode.
func (c *Chain) Store(target **ole.IDispatch) *Chain {
	if c.err != nil {
		return c
	}
	if c.disp == nil {
		return c
	}

	// Create a new reference for the user.
	c.disp.AddRef()
	*target = c.disp

	// Decouple the object from the chain's lifecycle management.
	if c.autoRelease {
		// In auto-release mode, remove the finalizer as the user is now responsible.
		runtime.SetFinalizer(c.disp, nil)
	} else {
		// In manual mode, remove the object from the release chain.
		// The current disp is always the last one added.
		if len(c.releaseChain) > 0 {
			lastIndex := len(c.releaseChain) - 1
			if c.releaseChain[lastIndex] == c.disp {
				c.releaseChain = c.releaseChain[:lastIndex]
			}
		}
	}

	return c
}

// Release releases all intermediate IDispatch objects created during the chain
// in the reverse order of their creation. It should be used when the chain is
// in manual resource management mode (the default).
func (c *Chain) Release() error {
	for i := len(c.releaseChain) - 1; i >= 0; i-- {
		c.releaseChain[i].Release()
	}
	c.releaseChain = nil // Clear the slice

	// Also clear the last result if it holds a variant
	if c.lastResult != nil {
		c.lastResult.Clear()
	}

	return c.err
}

// Value retrieves the value from the last operation (Get or Call). It also
// releases any intermediate IDispatch objects, similar to Release().
// The returned value is the Go equivalent of the VARIANT result.
func (c *Chain) Value() (interface{}, error) {
	if c.err != nil {
		// Release resources even if there was an error during the chain.
		c.Release()
		return nil, c.err
	}

	if c.lastResult == nil {
		c.Release()
		return nil, c.err // No value to return
	}

	val := c.lastResult.Value()
	err := c.Release() // Release resources after getting the value
	return val, err
}

// Err returns the first error that occurred during the chain. It ensures that
// all intermediate resources are released, similar to Release().
func (c *Chain) Err() error {
	// To prevent resource leaks, Err() must also release resources.
	// The behavior is now consistent with Value() and Release().
	// In AutoRelease mode, this will do nothing as releaseChain will be empty.
	return c.Release()
}
