//go:build windows

package sugar

import (
	"context"

	"github.com/go-ole/go-ole"
)

type sugarCtxKey struct{}
var activeSugarKey = sugarCtxKey{}

// Context manages the lifecycle of multiple Chains and implements context.Context.
type Context struct {
	context.Context
	chains []*Chain
}

// NewContext creates a new Context with the given parent.
func NewContext(parent context.Context) *Context {
	if parent == nil {
		parent = context.Background()
	}
	return &Context{
		Context: parent,
		chains:  make([]*Chain, 0, 4),
	}
}

// Track registers a Chain with the Context for automatic release.
func (c *Context) Track(chain *Chain) *Chain {
	chain.ctx = c
	c.chains = append(c.chains, chain)
	return chain
}

// Create is a wrapper around sugar.Create that automatically tracks the chain.
func (c *Context) Create(progID string) *Chain {
	return c.Track(Create(progID))
}

// GetActive is a wrapper around sugar.GetActive that automatically tracks the chain.
func (c *Context) GetActive(progID string) *Chain {
	return c.Track(GetActive(progID))
}

// From is a wrapper around sugar.From that automatically tracks the chain.
func (c *Context) From(disp *ole.IDispatch) *Chain {
	return c.Track(From(disp))
}

// Release releases all tracked chains in LIFO order.
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

// Do executes the function within a nested scope of this context.
func (c *Context) Do(fn func(ctx *Context) error) error {
	return With(c).Do(fn)
}

// Go executes the function in a new goroutine branching from this context.
func (c *Context) Go(fn func(ctx *Context) error) {
	With(c).Go(fn)
}