//go:build windows

package sugar

import (
	"errors"
	"fmt"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Chain provides a fluent interface for chaining OLE operations.
// It handles error propagation, allowing you to call multiple methods
// and check the error once at the end via Err().
type Chain interface {
	// Get retrieves a property from the current COM object and returns a NEW Chain
	// representing the property value. If the property is a COM object, it will
	// be automatically tracked if a Context is present.
	Get(prop string, params ...interface{}) Chain

	// Call executes a method on the current COM object and returns a NEW Chain
	// representing the return value. If the value is a COM object, it will
	// be automatically tracked if a Context is present.
	Call(method string, params ...interface{}) Chain

	// Put sets a property on the current COM object. It returns the same Chain
	// instance (or an error-carrying Chain) to allow further operations.
	Put(prop string, params ...interface{}) Chain

	// ForEach iterates over a COM collection (any object that implements IEnumVARIANT).
	// For each item, the callback is executed with a new Chain instance.
	//
	// To stop iteration:
	//   - Return nil to continue to the next item.
	//   - Return ErrForEachBreak (or an error wrapping it) to stop iteration.
	//   - Return any other error to stop and propagate the error to the parent Chain.
	//
	// NOTE: The break error is recorded in the Chain and should be checked manually
	// by the caller via Err() if they need to distinguish it from other errors.
	ForEach(callback func(item Chain) error) Chain

	// Fork creates a new independent reference to the current COM object.
	// Both the original and the forked Chain will point to the same object
	// but are managed as separate entries in the Context's arena.
	Fork() Chain

	// Store increases the reference count and returns the raw *ole.IDispatch.
	// The caller is responsible for calling Release() on the returned object
	// if it's not managed by sugar.Context.
	Store() (*ole.IDispatch, error)

	// Release manually releases the held COM object. Usually, this is handled
	// automatically by the sugar.Context, but can be used for early cleanup.
	Release() error

	// IsDispatch returns true if the last operation's result is a COM object (IDispatch).
	IsDispatch() bool

	// Value retrieves the underlying Go value of the last operation's result.
	// Returns an error if the result is a COM object (use Store() instead).
	Value() (interface{}, error)

	// Err returns the first error encountered in the chain of operations.
	Err() error
}

type chain struct {
	disp       *ole.IDispatch
	err        error
	lastResult *ole.VARIANT
	ctx        Context
}

// From starts a new chain with the given IDispatch.
func From(disp *ole.IDispatch) Chain {
	if disp != nil {
		disp.AddRef()
	}
	return &chain{
		disp: disp,
	}
}

// Create starts a new chain by creating a new COM object from the given ProgID.
func Create(progID string) Chain {
	unknown, err := oleutil.CreateObject(progID)
	if err != nil {
		return &chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		return &chain{err: err}
	}

	return &chain{
		disp: disp,
	}
}

// GetActive starts a new chain by attaching to a running COM object.
func GetActive(progID string) Chain {
	unknown, err := oleutil.GetActiveObject(progID)
	if err != nil {
		return &chain{err: err}
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		return &chain{err: err}
	}

	return &chain{
		disp: disp,
	}
}

func (c *chain) handleResult(result *ole.VARIANT, err error) Chain {
	if err != nil {
		if result != nil {
			result.Clear()
		}
		return &chain{err: err, ctx: c.ctx}
	}

	if result.VT == ole.VT_DISPATCH {
		newDisp := result.ToIDispatch()
		newDisp.AddRef()
		newChain := &chain{
			disp:       newDisp,
			lastResult: result,
			ctx:        c.ctx,
		}
		if c.ctx != nil {
			c.ctx.Track(newChain)
		}
		return newChain
	}

	// Value result: no IDispatch ownership. The new chain carries only the
	// VARIANT; sharing parent's disp here would let Release() double-free it.
	return &chain{
		lastResult: result,
		ctx:        c.ctx,
	}
}

// normalizeParams rewrites argument types that go-ole's Invoke cannot
// marshal (it panics on unknown types) into COM-compatible forms:
//
//   - Chain → *ole.IDispatch (AddRef'd for the call; released by cleanup).
//     This makes `wb.Call("Add", sheetChain)` and typed wrappers passing
//     Range/Worksheet values work.
//   - []interface{} / [][]interface{} → *ole.VARIANT carrying a
//     VT_ARRAY|VT_VARIANT SAFEARRAY (destroyed by cleanup). This is the
//     write path for `Range.Value` blocks.
//
// The returned cleanup func must run after the COM call completes; it is
// always non-nil.
func normalizeParams(params []interface{}) ([]interface{}, func(), error) {
	var cleanups []func()
	cleanup := func() {
		for i := len(cleanups) - 1; i >= 0; i-- {
			cleanups[i]()
		}
	}
	out := make([]interface{}, len(params))
	for i, p := range params {
		switch v := p.(type) {
		case Chain:
			disp, err := v.Store()
			if err != nil {
				cleanup()
				return nil, func() {}, fmt.Errorf("sugar: chain argument %d: %w", i, err)
			}
			cleanups = append(cleanups, func() { disp.Release() })
			out[i] = disp
		case []interface{}, [][]interface{}:
			va, err := encodeVariantArray(v)
			if err != nil {
				cleanup()
				return nil, func() {}, fmt.Errorf("sugar: array argument %d: %w", i, err)
			}
			cleanups = append(cleanups, func() { va.Clear() })
			out[i] = va
		default:
			out[i] = p
		}
	}
	return out, cleanup, nil
}

// invokeGuarded runs a go-ole call and converts its panics (go-ole panics on
// argument types it cannot marshal) into chain errors.
func invokeGuarded(fn func() (*ole.VARIANT, error)) (result *ole.VARIANT, err error) {
	defer func() {
		if r := recover(); r != nil {
			result = nil
			err = fmt.Errorf("sugar: COM invoke panicked: %v", r)
		}
	}()
	return fn()
}

// Get retrieves a property and returns a NEW Chain.
func (c *chain) Get(prop string, params ...interface{}) Chain {
	if c.err != nil {
		return &chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &chain{err: errors.New("dispatch is nil"), ctx: c.ctx}
	}
	args, cleanup, err := normalizeParams(params)
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
	}
	defer cleanup()
	result, err := invokeGuarded(func() (*ole.VARIANT, error) {
		return oleutil.GetProperty(c.disp, prop, args...)
	})
	return c.handleResult(result, err)
}

// Call executes a method and returns a NEW Chain.
func (c *chain) Call(method string, params ...interface{}) Chain {
	if c.err != nil {
		return &chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &chain{err: errors.New("dispatch is nil"), ctx: c.ctx}
	}
	args, cleanup, err := normalizeParams(params)
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
	}
	defer cleanup()
	result, err := invokeGuarded(func() (*ole.VARIANT, error) {
		return oleutil.CallMethod(c.disp, method, args...)
	})
	return c.handleResult(result, err)
}

// Put sets a property and returns the chain.
//
// On success the same chain is returned so callers can fluent-chain further
// operations on the parent object (e.g. `app.Put("Visible", true).Get(...)`).
// On error, a fresh error-only chain is returned — it does *not* share the
// parent's IDispatch, so manually Release()ing it cannot double-free.
func (c *chain) Put(prop string, params ...interface{}) Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	args, cleanup, err := normalizeParams(params)
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
	}
	defer cleanup()
	_, err = invokeGuarded(func() (*ole.VARIANT, error) {
		return oleutil.PutProperty(c.disp, prop, args...)
	})
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
	}

	return c
}

// ForEachBreak is returned when ForEach iteration is explicitly broken.
type ForEachBreak struct {
	Value interface{}
}

func (e *ForEachBreak) Error() string {
	return "foreach break"
}

func (e *ForEachBreak) Is(target error) bool {
	_, ok := target.(*ForEachBreak)
	return ok
}

var (
	// ErrForEachBreak is used to break out of a ForEach loop.
	ErrForEachBreak error = &ForEachBreak{}
)

// ForEach executes a callback for each item in a COM collection.
// If the callback returns a non-nil error, the iteration stops and the error
// is recorded in the returned Chain.
func (c *chain) ForEach(callback func(item Chain) error) Chain {
	if c.err != nil || c.disp == nil {
		return c
	}

	enumVar, err := oleutil.GetProperty(c.disp, "_NewEnum")
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
	}
	defer enumVar.Clear()

	if enumVar.VT != ole.VT_UNKNOWN && enumVar.VT != ole.VT_DISPATCH {
		return &chain{err: errors.New("_NewEnum is not object"), ctx: c.ctx}
	}

	unknown := enumVar.ToIUnknown()
	if unknown == nil {
		return &chain{err: errors.New("_NewEnum nil"), ctx: c.ctx}
	}

	iid, _ := ole.IIDFromString("{00020404-0000-0000-C000-000000000046}")
	enumRaw, err := unknown.QueryInterface(iid)
	if err != nil {
		return &chain{err: err, ctx: c.ctx}
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

			itemChain := &chain{
				disp: itemDisp,
				ctx:  c.ctx,
			}
			if c.ctx != nil {
				c.ctx.Track(itemChain)
			}

			cbErr := callback(itemChain)

			if c.ctx == nil {
				itemChain.Release()
			}

			if cbErr != nil {
				itemVar.Clear()
				return &chain{err: cbErr, ctx: c.ctx}
			}
		}
		itemVar.Clear()
	}
	return c
}

// Fork creates a new independent reference to the current object.
func (c *chain) Fork() Chain {
	if c.err != nil {
		return &chain{err: c.err, ctx: c.ctx}
	}
	if c.disp == nil {
		return &chain{err: errors.New("nil dispatch"), ctx: c.ctx}
	}
	c.disp.AddRef()
	newChain := &chain{disp: c.disp, ctx: c.ctx}
	if c.ctx != nil {
		c.ctx.Track(newChain)
	}
	return newChain
}

// Store transfers ownership of the current dispatch object to the caller.
func (c *chain) Store() (*ole.IDispatch, error) {
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
func (c *chain) Release() error {
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
func (c *chain) IsDispatch() bool {
	return c.lastResult != nil && c.lastResult.VT == ole.VT_DISPATCH
}

// Value retrieves the Go value of the last operation result.
//
// SAFEARRAY-of-VARIANT results (commonly returned by `Range.Value` /
// `Range.Value2` in Excel automation) are decoded to Go slices:
//
//   - 1-D SAFEARRAYs become `[]interface{}`.
//   - 2-D SAFEARRAYs become `[][]interface{}` indexed `[row][col]`.
//
// go-ole's built-in `VARIANT.Value()` returns nil for these types, so we
// decode them here. IDispatch results are not values — use Store().
func (c *chain) Value() (interface{}, error) {
	if c.err != nil {
		return nil, c.err
	}
	if c.lastResult == nil {
		return nil, nil
	}
	if c.lastResult.VT == ole.VT_DISPATCH {
		return nil, errors.New("result is IDispatch, use Store")
	}
	if c.lastResult.VT&ole.VT_ARRAY != 0 {
		return decodeVariantArray(c.lastResult)
	}
	return c.lastResult.Value(), nil
}

// Err returns the first error encountered in the chain.
func (c *chain) Err() error {
	return c.err
}

