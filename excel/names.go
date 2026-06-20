//go:build windows

package excel

import (
	"github.com/xll-gen/sugar"
)

// Names is a collection of defined names — the Go equivalent of xlwings'
// `Names`. Reached via Workbook.Names() (workbook scope) or
// Worksheet.Names() (sheet scope).
//
// xlwings reference: https://docs.xlwings.org/en/stable/api/names.html
type Names interface {
	sugar.Chain
	// Add defines a new name. refersTo is either an A1-notation formula
	// string prefixed with "=" (e.g. "=Sheet1!$A$1:$B$2") or a Range.
	Add(name string, refersTo interface{}) Name
	// Item returns a defined name by 1-based index or by name string.
	Item(index interface{}) Name
	// Count returns the number of defined names in the collection.
	Count() (int32, error)
	// Contains reports whether a name with the given name string exists.
	// xlwings analogue: `names.contains(name_or_index)`.
	Contains(name string) (bool, error)
}

type names struct {
	collection[Name]
}

// wrapNames wraps a chain in the Names typed wrapper. It is the single
// construction point for the chain -> Names convention.
func wrapNames(c sugar.Chain) Names { return &names{newCollection(c, wrapName)} }

func (n *names) Add(nameStr string, refersTo interface{}) Name {
	// Range values pass through as-is: sugar.Chain arguments are normalized
	// to raw IDispatch by the core Call.
	return n.add(n.Call("Add", nameStr, refersTo))
}

func (n *names) Item(index interface{}) Name {
	// Names.Item is a method in Excel's type library (unlike Sheets.Item,
	// which is a parameterized property), so DISPATCH_METHOD is required.
	return n.itemByCall(index)
}

func (n *names) Count() (int32, error) {
	return n.count()
}

func (n *names) Contains(nameStr string) (bool, error) {
	if err := n.Err(); err != nil {
		return false, err
	}
	// Excel's Names collection has no membership test; probing Item and
	// checking the chain error is the canonical COM idiom.
	return n.Call("Item", nameStr).Err() == nil, nil
}
