//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Name is a defined name (named range) — the Go equivalent of xlwings'
// `Name`.
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/name.html
type Name interface {
	sugar.Chain
	// Name returns the name string. For sheet-scoped names this includes the
	// sheet qualifier (e.g. "Sheet1!local_name"), mirroring xlwings.
	Name() (string, error)
	// SetName renames the defined name.
	SetName(name string) Name
	// RefersTo returns the formula the name refers to, in A1 notation and
	// prefixed with "=" (e.g. "=Sheet1!$A$1:$B$2").
	RefersTo() (string, error)
	// SetRefersTo repoints the name at a new formula. Pass A1 notation
	// prefixed with "=" (e.g. "=Sheet1!$A$1").
	SetRefersTo(refersTo string) Name
	// RefersToRange returns the Range the name refers to. Errors surface on
	// the returned Range's chain if the name does not refer to a range.
	RefersToRange() Range
	// Delete removes the defined name from its collection.
	Delete() error
}

type name struct {
	sugar.Chain
}

func (n *name) Name() (string, error) {
	v, err := n.Get("Name").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (n *name) SetName(s string) Name {
	return &name{n.Put("Name", s)}
}

func (n *name) RefersTo() (string, error) {
	v, err := n.Get("RefersTo").Value()
	if err != nil {
		return "", err
	}
	return toString(v), nil
}

func (n *name) SetRefersTo(refersTo string) Name {
	return &name{n.Put("RefersTo", refersTo)}
}

func (n *name) RefersToRange() Range {
	return &excelRange{n.Get("RefersToRange")}
}

func (n *name) Delete() error {
	return n.Call("Delete").Err()
}
