//go:build windows

package sugar

import (
	"errors"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Chain provides a fluent interface for chaining OLE operations.
type Chain struct {
	disp       *ole.IDispatch
	err        error
	lastResult *ole.VARIANT
	ctx        *Context
}

// From starts a new chain with the given IDispatch.
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

func (c *Chain) handleResult(result *ole.VARIANT, err error) *Chain {
	if err != nil {
		return &Chain{err: err, ctx: c.ctx}
	}

	newChain := &Chain{
		disp:       c.disp,
		lastResult: result,
		ctx:        c.ctx,
	}

	if result.VT == ole.VT_DISPATCH {
		newDisp := result.ToIDispatch()
		newDisp.AddRef()
		newChain.disp = newDisp
		
		if c.ctx != nil {
			c.ctx.Track(newChain)
		}
	}

	return newChain
}

// Get retrieves a property and returns a NEW Chain.
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

// Put sets a property and returns the chain.
func (c *Chain) Put(prop string, params ...interface{}) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	_, err := oleutil.PutProperty(c.disp, prop, params...)
	if err != nil {
		return &Chain{err: err, ctx: c.ctx, disp: c.disp}
	}
	
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

	iid, _ := ole.IIDFromString("{00020404-0000-0000-C000-000000000046}")
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
			if c.ctx != nil {
				c.ctx.Track(itemChain)
			}

			cont := callback(itemChain)
			
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

// Fork creates a new independent reference to the current object.
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

// Store transfers ownership of the current dispatch object to the caller.
func (c *Chain) Store() (*ole.IDispatch, error) {
	if c.err != nil {
		return nil, c.err
	}
	if c.disp == nil {
		return nil, errors.New("nil dispatch")
	}

	c.disp.AddRef()
	return c.disp, nil
}

// Release releases the held dispatch object and captures errors.
func (c *Chain) Release() error {
	if c.disp != nil {
		c.disp.Release()
		c.disp = nil
	}
	if c.lastResult != nil {
		c.lastResult.Clear()
		c.lastResult = nil
	}
	err := c.err
	c.err = nil
	return err
}

// IsDispatch returns true if the last result is a dispatch object.
func (c *Chain) IsDispatch() bool {
	return c.lastResult != nil && c.lastResult.VT == ole.VT_DISPATCH
}

// Value retrieves the Go value of the last operation result.
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

// Err returns the first error encountered in the chain.
func (c *Chain) Err() error {
	return c.err
}
