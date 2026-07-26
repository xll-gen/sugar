//go:build windows

// Excel-free unit tests for the VARIANT coercion helpers. These run under
// plain `go test ./...` because they exercise pure Go logic over already
// decoded VARIANT values.

package excel

import (
	"strings"
	"testing"
	"time"
)

// TestStringFromVariant_RejectsArrays pins the getString fix: a scalar string
// property that comes back as a SAFEARRAY (Formula on a multi-cell range, say)
// must be an explicit error. Before the fix, fmt.Sprint forged Go-syntax text
// such as "[[=1+1 =2+2]]" and returned it as a legitimate Excel value.
func TestStringFromVariant_RejectsArrays(t *testing.T) {
	arrays := []struct {
		name string
		v    interface{}
		want string // shape fragment expected in the message
	}{
		{"2-D grid", [][]interface{}{{"=1+1", "=2+2"}}, "1x2"},
		{"2-D column", [][]interface{}{{"=1+1"}, {"=2+2"}}, "2x1"},
		{"1-D row", []interface{}{"=1+1", "=2+2"}, "1x2"},
		{"empty grid", [][]interface{}{}, "0x0"},
	}
	for _, tc := range arrays {
		t.Run(tc.name, func(t *testing.T) {
			got, err := stringFromVariant("Formula", tc.v)
			if err == nil {
				t.Fatalf("stringFromVariant(%v) = %q, nil; want an error", tc.v, got)
			}
			if got != "" {
				t.Errorf("value on the error path = %q; want the empty string", got)
			}
			msg := err.Error()
			if !strings.Contains(msg, "Formula") || !strings.Contains(msg, tc.want) {
				t.Errorf("error %q should name the property and the %s shape", msg, tc.want)
			}
			// The forged-string symptom must be gone.
			if strings.Contains(msg, "[[") {
				t.Errorf("error %q still leaks the Go-syntax rendering", msg)
			}
		})
	}
}

// TestStringFromVariant_PassesScalars proves the guard has no false positives:
// every scalar VARIANT shape a string property can legitimately carry still
// converts, including nil (VT_EMPTY), which stays the empty string — see the
// helpers.go note on why nil cannot be promoted to an error at this layer.
func TestStringFromVariant_PassesScalars(t *testing.T) {
	when := time.Date(2026, 7, 26, 0, 0, 0, 0, time.UTC)
	cases := []struct {
		v    interface{}
		want string
	}{
		{"=SUM(A1:A2)", "=SUM(A1:A2)"},
		{"", ""},
		{nil, ""},
		{42.5, "42.5"},
		{int32(7), "7"},
		{true, "true"},
		{when, when.String()},
	}
	for _, tc := range cases {
		got, err := stringFromVariant("Formula", tc.v)
		if err != nil {
			t.Errorf("stringFromVariant(%v (%T)) errored: %v", tc.v, tc.v, err)
			continue
		}
		if got != tc.want {
			t.Errorf("stringFromVariant(%v (%T)) = %q; want %q", tc.v, tc.v, got, tc.want)
		}
	}
}

// TestVariantArrayShape covers the shape classifier on its own.
func TestVariantArrayShape(t *testing.T) {
	cases := []struct {
		name       string
		v          interface{}
		rows, cols int
		isArray    bool
	}{
		{"2x3 grid", [][]interface{}{{1, 2, 3}, {4, 5, 6}}, 2, 3, true},
		{"1-D of 4", []interface{}{1, 2, 3, 4}, 1, 4, true},
		{"empty 2-D", [][]interface{}{}, 0, 0, true},
		{"string scalar", "A1", 0, 0, false},
		{"nil scalar", nil, 0, 0, false},
		{"float scalar", 1.5, 0, 0, false},
		{"typed slice is not a VARIANT array", []string{"a"}, 0, 0, false},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			rows, cols, isArray := variantArrayShape(tc.v)
			if rows != tc.rows || cols != tc.cols || isArray != tc.isArray {
				t.Errorf("variantArrayShape(%v) = (%d, %d, %v); want (%d, %d, %v)",
					tc.v, rows, cols, isArray, tc.rows, tc.cols, tc.isArray)
			}
		})
	}
}
