//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// collection is the shared base for the six Excel object collections
// (Charts, Workbooks, Worksheets, Names, Pictures, Shapes). It captures the
// trio every collection re-derived by hand: Count (always
// getInt32(source, "Count")), Item (wrap a COM Item lookup), and Add (wrap a
// COM-call result into the element wrapper).
//
// Type parameter T is the element wrapper interface (Chart, Workbook, ...);
// wrap is the chain -> T constructor (wrapChart, wrapWorkbook, ...). source
// resolves the chain that the trio operates on. For five collections that is
// the embedded chain itself; Pictures overrides it to re-fetch the legacy
// snapshot collection on every lookup (see pictures.go).
//
// Each concrete collection embeds a collection[T] and supplies only its wrap
// func and (optionally) a custom source. Public method signatures, return
// types, and the exact COM calls are unchanged from the open-coded versions —
// this is a behavior-preserving deduplication, not a redesign.
type collection[T any] struct {
	sugar.Chain
	wrap   func(sugar.Chain) T
	source func() sugar.Chain
}

// newCollection builds a base whose trio operates on the embedded chain
// itself. Used by every collection except Pictures.
func newCollection[T any](c sugar.Chain, wrap func(sugar.Chain) T) collection[T] {
	base := collection[T]{Chain: c, wrap: wrap}
	base.source = func() sugar.Chain { return base.Chain }
	return base
}

// newCollectionFrom builds a base whose trio operates on a chain resolved by
// src on every call, while the embedded chain stays as the Err/ForEach anchor.
// Used by Pictures, whose Item/Count must re-fetch the snapshot collection.
func newCollectionFrom[T any](anchor sugar.Chain, src func() sugar.Chain, wrap func(sugar.Chain) T) collection[T] {
	return collection[T]{Chain: anchor, wrap: wrap, source: src}
}

// count returns the element count. It is exactly getInt32(source, "Count"),
// the form every collection previously repeated.
func (b collection[T]) count() (int32, error) {
	return getInt32(b.source(), "Count")
}

// itemByCall resolves an element via the COM `Item` *method*
// (DISPATCH_METHOD). Charts, Names, Pictures, and Shapes use this because
// their type-library `Item` is a method, not a parameterized property.
func (b collection[T]) itemByCall(index interface{}) T {
	return b.wrap(b.source().Call("Item", index))
}

// itemByGet resolves an element via the COM `Item` parameterized *property*
// (DISPATCH_PROPERTYGET). Workbooks and Worksheets use this because their
// type-library `Item` is a property.
func (b collection[T]) itemByGet(index interface{}) T {
	return b.wrap(b.source().Get("Item", index))
}

// add wraps the already-built result chain of an `Add` COM call into the
// element type. The per-collection Add methods keep their distinct option
// plumbing and only funnel the final chain through here.
func (b collection[T]) add(result sugar.Chain) T {
	return b.wrap(result)
}
