//go:build windows

package sugar

import (
	"errors"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Chain provides a fluent interface for chaining OLE operations.
// It uses an immutable pattern: methods like Get and Call return a NEW Chain instance,
// leaving the original Chain unmodified.
type Chain struct {
	disp       *ole.IDispatch
	err        error
	lastResult *ole.VARIANT
	ctx        *Context // Optional context for automatic resource tracking
}

// From starts a new chain with the given IDispatch.
// If used without a Context, the caller is responsible for releasing the chain.
func From(disp *ole.IDispatch) *Chain {
	if disp != nil {
		disp.AddRef()
	}
	return &Chain{
		disp: disp,
	}
}

// Create starts a new chain by creating a new COM object from the given ProgID.
func Create(progID string) *Chain {
	unknown, err := oleutil.CreateObject(progID)
	if err != nil {
		return &Chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		return &Chain{err: err}
	}

	return &Chain{
		disp: disp,
	}
}

// GetActive starts a new chain by attaching to a running COM object.
func GetActive(progID string) *Chain {
	unknown, err := oleutil.GetActiveObject(progID)
	if err != nil {
		return &Chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		return &Chain{err: err}
	}

	return &Chain{
		disp: disp,
	}
}

// handleResult processes the result of GetProperty and CallMethod.
// It returns a NEW Chain instance if the result is a dispatch object.
func (c *Chain) handleResult(result *ole.VARIANT, err error) *Chain {
	if err != nil {
		return &Chain{err: err, ctx: c.ctx}
	}

	// Create a new chain for the result
	newChain := &Chain{
		disp:       c.disp, // Default to current disp if result is not object
		lastResult: result,
		ctx:        c.ctx,
	}

	if result.VT == ole.VT_DISPATCH {
		newDisp := result.ToIDispatch()
		newDisp.AddRef() // AddRef because the new chain owns this reference
		newChain.disp = newDisp
		
		// Auto-track if context is present
		if c.ctx != nil {
			c.ctx.Track(newChain)
		}
	} else if result.VT == ole.VT_UNKNOWN {
        // Handle IUnknown if returned (rare but possible)
        // For now, treat as non-dispatch or error?
        // Let's stick to IDispatch support.
    }

	return newChain
}

// Get retrieves a property and returns a NEW Chain.
// The original chain is unaffected.
func (c *Chain) Get(prop string, params ...interface{}) *Chain {
	if c.err != nil {
		return &Chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &Chain{err: errors.New("dispatch is nil"), ctx: c.ctx}
	}
	result, err := oleutil.GetProperty(c.disp, prop, params...)
	return c.handleResult(result, err)
}

// Call executes a method and returns a NEW Chain.
// The original chain is unaffected.
func (c *Chain) Call(method string, params ...interface{}) *Chain {
	if c.err != nil {
		return &Chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &Chain{err: errors.New("dispatch is nil"), ctx: c.ctx}
	}
	result, err := oleutil.CallMethod(c.disp, method, params...)
	return c.handleResult(result, err)
}

// Put sets a property.
// Unlike Get/Call, Put is typically a terminal operation or returns the SAME chain
// because it doesn't produce a new object to traverse.
// However, to maintain consistency, we return 'c' (self).
func (c *Chain) Put(prop string, params ...interface{}) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	_, err := oleutil.PutProperty(c.disp, prop, params...)
	if err != nil {
		// Return a new chain with error, or modify self?
		// Since Put is a side-effect, modifying self's error state is acceptable,
		// OR we return a new chain with error.
		// Let's return a new chain with error to be safe with immutability,
		// although the user might ignore the return value.
		return &Chain{err: err, ctx: c.ctx, disp: c.disp}
	}
	
	// Put doesn't return a value, so we clear lastResult in the returned chain?
	// Or we just return 'c'. Returning 'c' is fine for side-effects.
	return c
}

// ForEach iterates over a collection.
func (c *Chain) ForEach(callback func(item *Chain) bool) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	enumVar, err := oleutil.GetProperty(c.disp, "_NewEnum")
	if err != nil {
		return &Chain{err: err, ctx: c.ctx}
	}
	defer enumVar.Clear()

	if enumVar.VT != ole.VT_UNKNOWN && enumVar.VT != ole.VT_DISPATCH {
		return &Chain{err: errors.New("_NewEnum is not object"), ctx: c.ctx}
	}
	
	unknown := enumVar.ToIUnknown()
	if unknown == nil {
		return &Chain{err: errors.New("_NewEnum nil"), ctx: c.ctx}
	}

	iid, err := ole.IIDFromString("{00020404-0000-0000-C000-000000000046}")
	if err != nil {
		return &Chain{err: err, ctx: c.ctx}
	}

	enumRaw, err := unknown.QueryInterface(iid)
	if err != nil {
		return &Chain{err: err, ctx: c.ctx}
	}
	defer enumRaw.Release()

	enum := (*ole.IEnumVARIANT)(unsafe.Pointer(enumRaw))

	for {
		itemVar, fetched, err := enum.Next(1)
		if err != nil || fetched == 0 {
			break
		}
		
		if itemVar.VT == ole.VT_DISPATCH {
			itemDisp := itemVar.ToIDispatch()
			itemDisp.AddRef()
			
			itemChain := &Chain{
				disp: itemDisp,
				ctx:  c.ctx,
			}
			// Track item if context exists
			if c.ctx != nil {
				c.ctx.Track(itemChain)
			}

			cont := callback(itemChain)
			
			// If not tracked by context, we should release it?
			// If tracked, context releases it later.
			// BUT, ForEach creates many items. If we track ALL of them, we might bloat memory 
			// until Context.Release is called.
			// Ideally, users should use 'itemChain' locally.
			// If we track, we are safe.
			
			if c.ctx == nil {
				itemChain.Release()
			}
			
			if !cont {
				itemVar.Clear()
				break
			}
		}
		itemVar.Clear()
	}

	return c
}

// Fork is now redundant but kept for API compatibility or explicit branching.
// It simply returns the chain itself (or a clone) because chains are immutable.
// Actually, Fork meant "branch off independent ref". Get/Call do that now.
// So Fork can just return 'c' with AddRef? 
// No, Get/Call creates new Refs. Fork creates a clone of CURRENT ref.
func (c *Chain) Fork() *Chain {
	if c.err != nil {
		return &Chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &Chain{err: errors.New("nil dispatch"), ctx: c.ctx}
	}
	c.disp.AddRef()
	newChain := &Chain{disp: c.disp, ctx: c.ctx}
	if c.ctx != nil {
		c.ctx.Track(newChain)
	}
	return newChain
}

// Store transfers ownership to caller.
// In immutable pattern, we just return the disp and detach from Context if needed?
// Or we just return AddRef'd disp.
func (c *Chain) Store() (*ole.IDispatch, error) {
	if c.err != nil {
		return nil, c.err
	}
	if c.disp == nil {
		return nil, errors.New("nil dispatch")
	}

	// Return a new reference
	c.disp.AddRef()
	return c.disp, nil
}

// Release releases the *current* chain's held dispatch object.
// It does not release "parent" objects because chains are now independent.
func (c *Chain) Release() error {
	if c.disp != nil {
		c.disp.Release()
		// c.disp = nil // Cannot zero out if we want to be safe against double free? 
		// Actually Release is destructive.
	}
	if c.lastResult != nil {
		c.lastResult.Clear()
		c.lastResult = nil
	}
	return c.err
}

func (c *Chain) IsDispatch() bool {
	return c.lastResult != nil && c.lastResult.VT == ole.VT_DISPATCH
}

func (c *Chain) Value() (interface{}, error) {
	if c.err != nil {
		return nil, c.err
	}
	if c.lastResult == nil {
		return nil, nil
	}
	if c.lastResult.VT == ole.VT_DISPATCH {
		return nil, errors.New("result is IDispatch, use Store")
	}
	return c.lastResult.Value(), nil
}

func (c *Chain) Err() error {
	return c.err
}