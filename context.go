//go:build windows

package sugar

import "github.com/go-ole/go-ole"

// Context manages the lifecycle of multiple Chains.
// It acts as an "arena" or "pool", collecting chains and releasing them all at once.
// This simplifies resource management by allowing a single defer statement to clean up
// multiple objects created within a scope.
type Context struct {
	chains []*Chain
}

// Do creates a new Context, executes the provided function within that context,
// and ensures that all tracked resources are released when the function completes.
func Do(fn func(ctx *Context)) {
	ctx := NewContext()
	defer ctx.Release()
	fn(ctx)
}

// NewContext creates a new Context.
func NewContext() *Context {
	return &Context{
		chains: make([]*Chain, 0, 4),
	}
}

// Track registers an existing Chain with the Context.
// When the Context is released, this chain will also be released.
// Returns the passed chain for fluent usage.
// Example: newChain := ctx.Track(oldChain.Fork())
func (c *Context) Track(chain *Chain) *Chain {
	chain.ctx = c
	c.chains = append(c.chains, chain)
	return chain
}

// Create is a wrapper around sugar.Create that automatically tracks the created chain.
func (c *Context) Create(progID string) *Chain {
	return c.Track(Create(progID))
}

// GetActive is a wrapper around sugar.GetActive that automatically tracks the created chain.
func (c *Context) GetActive(progID string) *Chain {
	return c.Track(GetActive(progID))
}

// From is a wrapper around sugar.From that automatically tracks the created chain.
// Note: As with sugar.From, the Context does NOT take ownership of the input 'disp' object itself,
// but it tracks the Chain wrapper which manages intermediate objects produced by it.
func (c *Context) From(disp *ole.IDispatch) *Chain {
	return c.Track(From(disp))
}

// Release releases all tracked chains in reverse order of registration (LIFO).
// It returns the first error encountered, but attempts to release all chains.
// It is safe to call Release multiple times.
func (c *Context) Release() error {
	if c.chains == nil {
		return nil
	}
	var firstErr error
	for i := len(c.chains) - 1; i >= 0; i-- {
		if err := c.chains[i].Release(); err != nil && firstErr == nil {
			firstErr = err
		}
	}
	c.chains = nil
	return firstErr
}
