//go:build windows

package excel

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// callOptional issues a positional COM call where `leading` holds the required
// arguments and `optional` holds the trailing optional arguments in COM
// signature order. Every optional that should be omitted must be sugar.Missing();
// trailing Missing() placeholders are then dropped so a call supplying no
// optionals collapses to just the leading args (e.g. SaveAs(path) stays a 1-arg
// call, Close() stays a 0-arg call). This centralizes the Missing-seed +
// trailing-trim pattern that every positional COM method with optional
// parameters (SaveAs, Workbooks.Open, Workbook.Close, future PasteSpecial/Run/
// Copy/...) otherwise re-implements.
func callOptional(c sugar.Chain, method string, leading []interface{}, optional ...interface{}) sugar.Chain {
	args := append(append([]interface{}{}, leading...), optional...)
	return c.Call(method, trimTrailingMissing(args)...)
}

// trimTrailingMissing drops trailing sugar.Missing() placeholders from a
// positional COM argument list so that simple calls collapse to their shortest
// form (down to and including the empty list when every argument is omitted).
func trimTrailingMissing(args []interface{}) []interface{} {
	last := len(args) - 1
	for last >= 0 {
		if _, isMissing := args[last].(*ole.VARIANT); !isMissing {
			break
		}
		last--
	}
	return args[:last+1]
}

// getInt32 reads an int32-family scalar COM property off a chain and coerces it
// through toInt32. It is the generic form of the repeated
// `c.Get(prop).Value()` + narrow pattern used by every Count/Index/Row/Column
// getter. Extra params are forwarded to Get so indexed reads (e.g.
// Get("Cells", r, c)) work too.
func getInt32(c sugar.Chain, prop string, params ...interface{}) (int32, error) {
	v, err := c.Get(prop, params...).Value()
	if err != nil {
		return 0, err
	}
	return toInt32(v), nil
}

// getFloat64 reads a numeric scalar COM property off a chain and coerces it
// through toFloat64. Generic form of the geometry-getter pattern
// (Left/Top/Width/Height/ColumnWidth/RowHeight).
func getFloat64(c sugar.Chain, prop string, params ...interface{}) (float64, error) {
	v, err := c.Get(prop, params...).Value()
	if err != nil {
		return 0, err
	}
	return toFloat64(v), nil
}

// getBool reads a boolean scalar COM property off a chain and coerces it
// through toBool. Routing every bool getter here is the correctness fix: toBool
// treats legacy 0/-1 VARIANT shapes as false/true, whereas a bare v.(bool)
// type-assert silently returns false for those.
func getBool(c sugar.Chain, prop string, params ...interface{}) (bool, error) {
	v, err := c.Get(prop, params...).Value()
	if err != nil {
		return false, err
	}
	return toBool(v), nil
}

// getString reads a string scalar COM property off a chain and coerces it
// through stringFromVariant. Generic form of the Name/Address/Formula/Path
// getter pattern.
func getString(c sugar.Chain, prop string, params ...interface{}) (string, error) {
	v, err := c.Get(prop, params...).Value()
	if err != nil {
		return "", err
	}
	return stringFromVariant(prop, v)
}

// stringFromVariant converts a decoded VARIANT to the Go string a scalar
// string-property getter promises, rejecting the one shape that has no string
// form: a SAFEARRAY.
//
// Excel returns an array from a *scalar* string property whenever the source
// object spans more than one cell — Range("A1:B1").Formula on a multi-cell
// range decodes to [][]interface{}. Running that through fmt.Sprint fabricated
// Go-syntax text ("[[=1+1 =2+2]]") and handed it back as if it were a real
// Excel value, so callers silently stored or re-wrote garbage. No Excel string
// property can legitimately be a slice, so the guard has no false positives:
// arrays become an explicit error instead of a forged string.
func stringFromVariant(prop string, v interface{}) (string, error) {
	if rows, cols, isArray := variantArrayShape(v); isArray {
		return "", fmt.Errorf(
			"excel: %s returned a %dx%d array, not a string: the object spans "+
				"multiple cells — narrow it to one cell or read the block with Value()",
			prop, rows, cols)
	}
	return toString(v), nil
}

// variantArrayShape reports whether a decoded VARIANT is a SAFEARRAY and, if
// so, its [rows][cols] shape. It matches the two shapes sugar's VARIANT decoder
// produces (safearray.go: 1-D -> []interface{}, 2-D -> [][]interface{}); a 1-D
// array is reported as a single row. Any other Go type is a scalar.
func variantArrayShape(v interface{}) (rows, cols int, isArray bool) {
	switch a := v.(type) {
	case [][]interface{}:
		if len(a) > 0 {
			cols = len(a[0])
		}
		return len(a), cols, true
	case []interface{}:
		return 1, len(a), true
	}
	return 0, 0, false
}

// toInt32 narrows the variety of integer types Excel COM can hand back into a
// single Go int32. Excel emits Count/Index/Row/Column as VT_I4 most of the
// time but VT_I2/VT_R8 turn up for legacy properties. A panic-free fallback
// keeps the typed wrappers from crashing on unexpected shapes.
func toInt32(v interface{}) int32 {
	switch x := v.(type) {
	case int32:
		return x
	case int16:
		return int32(x)
	case int64:
		return int32(x)
	case int:
		return int32(x)
	case uint32:
		return int32(x)
	case float64:
		return int32(x)
	case float32:
		return int32(x)
	}
	return 0
}

// toFloat64 widens the numeric VARIANT shapes Excel emits for geometry
// properties (Left/Top/Width/Height arrive as VT_R8, but VT_I4 turns up on
// some hosts) into a float64. Same panic-free contract as toInt32.
func toFloat64(v interface{}) float64 {
	switch x := v.(type) {
	case float64:
		return x
	case float32:
		return float64(x)
	case int32:
		return float64(x)
	case int16:
		return float64(x)
	case int64:
		return float64(x)
	case int:
		return float64(x)
	case uint32:
		return float64(x)
	}
	return 0
}

// toBool coerces the VARIANT shapes Excel emits for boolean properties into a
// Go bool. COM hands back VT_BOOL as a Go bool, but legacy hosts occasionally
// surface 0/-1 integers; treat any non-zero numeric as true. Same panic-free
// contract as toInt32.
func toBool(v interface{}) bool {
	switch x := v.(type) {
	case bool:
		return x
	case int32:
		return x != 0
	case int16:
		return x != 0
	case int64:
		return x != 0
	case int:
		return x != 0
	case uint32:
		return x != 0
	case float64:
		return x != 0
	case float32:
		return x != 0
	}
	return false
}

// toString coerces any VARIANT scalar value to a Go string, formatting
// numbers/bools via fmt when the COM property is documented as string but
// occasionally arrives typed (e.g. Application.Version on some hosts).
//
// Only scalars reach here: stringFromVariant rejects SAFEARRAY results first,
// because fmt.Sprint on a slice produces Go syntax, never an Excel value.
func toString(v interface{}) string {
	if v == nil {
		return ""
	}
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprint(v)
}
