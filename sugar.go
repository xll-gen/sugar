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
	disp         *ole.IDispatch
	err          error
	lastResult   *ole.VARIANT
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
		newDisp.AddRef()
		c.releaseChain = append(c.releaseChain, newDisp)
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

// ForEach iterates over a collection. The current object in the chain must be
// a collection that supports the _NewEnum property. The callback function is
// executed for each item in the collection. The item is passed as a new Chain.
// If the callback returns false, the iteration stops.
//
// Note: This method only supports collections of objects (IDispatch). Items
// that are not objects are skipped.
func (c *Chain) ForEach(callback func(item *Chain) bool) *Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	// Get _NewEnum property
	// -4 is the standard DISPID for _NewEnum
	enumVar, err := oleutil.GetProperty(c.disp, "_NewEnum")
	if err != nil {
		c.err = err
		return c
	}
	defer enumVar.Clear()

	// Get IEnumVARIANT interface
	if enumVar.VT != ole.VT_UNKNOWN && enumVar.VT != ole.VT_DISPATCH {
		c.err = errors.New("property _NewEnum is not an object")
		return c
	}
	
	// IID_IEnumVARIANT
	unknown := enumVar.ToIUnknown()
	if unknown == nil {
		c.err = errors.New("_NewEnum returned nil IUnknown")
		return c
	}

	// IID_IEnumVARIANT
	iid, err := ole.IIDFromString("{00020404-0000-0000-C000-000000000046}")
	if err != nil {
		c.err = err
		return c
	}

	enumRaw, err := unknown.QueryInterface(iid)
	if err != nil {
		c.err = err
		return c
	}
	defer enumRaw.Release()

	enum := (*ole.IEnumVARIANT)(unsafe.Pointer(enumRaw))

	// Iterate
	for {
		// Next returns (VARIANT, uint, error) in some go-ole versions
		itemVar, fetched, err := enum.Next(1)
		if err != nil || fetched == 0 {
			break
		}
		
		if itemVar.VT == ole.VT_DISPATCH {
			// Create a new Chain for the item
			itemDisp := itemVar.ToIDispatch()
			// ToIDispatch does not AddRef, but Next() implementation does AddRef on the variant content?
			// Yes, VariantClear (called by itemVar.Clear()) releases it.
			// So for the Chain, we should probably AddRef if we want the Chain to manage it independently,
			// OR we rely on the fact that we will clear itemVar after callback.
			// The Chain implementation assumes it owns the objects in releaseChain.
			// Let's AddRef for the Chain to be safe and consistent with other methods.
			itemDisp.AddRef()
			
			itemChain := &Chain{
				disp:         itemDisp,
				releaseChain: []*ole.IDispatch{itemDisp},
			}

			cont := callback(itemChain)
			
			// Release the chain created for the item. This cleans up the AddRef we just did.
			itemChain.Release()
			
			if !cont {
				itemVar.Clear()
				break
			}
		}
		
		itemVar.Clear()
	}

	return c
}

// Fork creates a new independent Chain starting from the current object.
// The new chain takes ownership of a new reference to the current object (AddRef is called).
// The caller is responsible for calling Release() on the returned chain.
// This is useful when you want to branch off a new chain from an intermediate result
// without breaking the original chain or manually handling Store/From.
func (c *Chain) Fork() *Chain {
	if c.err != nil {
		return &Chain{err: c.err}
	}
	if c.disp == nil {
		return &Chain{err: errors.New("no object to fork")}
	}

	// AddRef because the new chain will own this object too.
	c.disp.AddRef()

	return &Chain{
		disp:         c.disp,
		releaseChain: []*ole.IDispatch{c.disp},
	}
}

// Store is a terminal method that transfers ownership of the current IDispatch
// object to the caller. The user is responsible for calling Release on the
// returned object. This method also releases all other resources managed by the
// chain.
func (c *Chain) Store() (*ole.IDispatch, error) {
	if c.err != nil {
		c.Release() // Release other resources
		return nil, c.err
	}
	if c.disp == nil {
		c.Release()
		return nil, errors.New("no IDispatch object to store")
	}

	// The object to be returned to the user.
	storedDisp := c.disp

	// Decouple the stored object from the chain's lifecycle management.
	// Remove the object from the release chain.
	// The current disp is always the last one added.
	if len(c.releaseChain) > 0 {
		lastIndex := len(c.releaseChain) - 1
		if c.releaseChain[lastIndex] == storedDisp {
			c.releaseChain = c.releaseChain[:lastIndex]
		}
	}

	// Release any other resources held by the chain.
	c.Release()

	return storedDisp, nil
}

// Release releases all intermediate IDispatch objects created during the chain
// in the reverse order of their creation.
func (c *Chain) Release() error {
	for i := len(c.releaseChain) - 1; i >= 0; i-- {
		c.releaseChain[i].Release()
	}
	c.releaseChain = nil // Clear the slice

	// Also clear the last result if it holds a variant
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

// Value retrieves the value from the last operation (Get or Call). It also
// releases any intermediate IDispatch objects, similar to Release().
// The returned value is the Go equivalent of the VARIANT result.
func (c *Chain) Value() (interface{}, error) {
	if c.err != nil {
		return nil, c.Release()
	}

	if c.lastResult == nil {
		c.Release()
		return nil, nil
	}

	if c.lastResult.VT == ole.VT_DISPATCH {
		c.Release()
		return nil, errors.New("value cannot return IDispatch, use Store() instead")
	}

	val := c.lastResult.Value()
	err := c.Release() // Release resources after getting the value
	return val, err
}

// Err returns the first error that occurred during the chain.
func (c *Chain) Err() error {
	return c.err
}
