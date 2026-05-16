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
