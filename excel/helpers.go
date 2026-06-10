//go:build windows

package excel

import "fmt"

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
func toString(v interface{}) string {
	if v == nil {
		return ""
	}
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprint(v)
}
