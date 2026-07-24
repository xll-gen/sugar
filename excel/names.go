//go:build windows

package excel

import (
	"errors"

	ole "github.com/go-ole/go-ole"

	"github.com/xll-gen/sugar"
)

// COM HRESULTs that mean "the requested name is not in the collection" — the
// only failures Contains is allowed to fold into (false, nil). Everything else
// (permission errors, disconnected server, marshaling faults) is a real error
// and must propagate.
const (
	hrDispEException = 0x80020009 // DISP_E_EXCEPTION (wraps the Excel scode below)
	hrDispEBadIndex  = 0x8002000B // DISP_E_BADINDEX (bad Item index)
	hrXlItemNotFound = 0x800A03EC // Excel automation "item not found"
)

// isNameNotFound reports whether err is a "no such name" miss rather than a
// genuine COM failure. Excel delivers the miss two ways: DISP_E_BADINDEX is
// returned directly by Invoke, while a name lookup surfaces 0x800A03EC nested
// inside a DISP_E_EXCEPTION's EXCEPINFO (go-ole stores it as the OleError's
// SubError). Both shapes are classified as not-found; any other HRESULT is a
// real error.
func isNameNotFound(err error) bool {
	var oleErr *ole.OleError
	if !errors.As(err, &oleErr) {
		return false
	}
	switch uint32(oleErr.Code()) {
	case hrDispEBadIndex, hrXlItemNotFound:
		return true
	case hrDispEException:
		if ex, ok := oleErr.SubError().(ole.EXCEPINFO); ok {
			return ex.SCODE() == hrXlItemNotFound
		}
	}
	return false
}

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
	// checking the chain error is the canonical COM idiom. A not-found miss is
	// a clean (false, nil); any other COM failure (disconnected server, access
	// denied, …) must not be masqueraded as "absent" — propagate it.
	err := n.Call("Item", nameStr).Err()
	if err == nil {
		return true, nil
	}
	if isNameNotFound(err) {
		return false, nil
	}
	return false, err
}
