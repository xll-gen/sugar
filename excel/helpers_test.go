//go:build windows

// Excel-free unit tests for the VARIANT coercion helpers. These run under
// plain `go test ./...` because they exercise pure Go logic over already
// decoded VARIANT values.

package excel

import (
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
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

// TestStringFromVariant_RejectsNull covers the converter on its own, which is
// what makes its (deliberately duplicated) Null check load-bearing rather than
// decoration: sugar.Null has a String() method, so the toString fallback would
// render the literal text "Null" and hand it back as if Excel had said the
// number format WAS "Null" — the same forged-string failure the array guard
// exists for.
func TestStringFromVariant_RejectsNull(t *testing.T) {
	got, err := stringFromVariant("NumberFormat", sugar.Null{})
	requireNullError(t, "NumberFormat", err)
	if got != "" {
		t.Errorf("value on the error path = %q; want the empty string", got)
	}
	// The forged rendering must not be what comes back instead.
	if got == "Null" {
		t.Errorf("stringFromVariant forged the sentinel's String() as an Excel value")
	}
}

// TestScalarGetters_RejectNull is the mixed-cell half of the getString array
// fix. Excel answers a scalar property read with VT_NULL — "no single value" —
// whenever the object spans cells that disagree. That decodes to sugar.Null,
// and every one of the four typed getters must refuse it: before this, the
// nil-shaped decode ran through toString/toFloat64/toBool/toInt32 and became
// "" / 0 / false / 0 WITH A NIL ERROR, so `Range("A1:B1").MergeCells()` reported
// a confident "not merged" for a half-merged block and `ColumnWidth()` reported
// a width of 0 points.
func TestScalarGetters_RejectNull(t *testing.T) {
	fc := newFakeChain()
	fc.root.value = sugar.Null{}

	t.Run("getString", func(t *testing.T) {
		got, err := getString(fc, "NumberFormat")
		requireNullError(t, "NumberFormat", err)
		if got != "" {
			t.Errorf("value on the error path = %q; want the empty string", got)
		}
	})
	t.Run("getFloat64", func(t *testing.T) {
		got, err := getFloat64(fc, "ColumnWidth")
		requireNullError(t, "ColumnWidth", err)
		if got != 0 {
			t.Errorf("value on the error path = %v; want 0", got)
		}
	})
	t.Run("getBool", func(t *testing.T) {
		got, err := getBool(fc, "MergeCells")
		requireNullError(t, "MergeCells", err)
		if got {
			t.Errorf("value on the error path = %v; want false", got)
		}
	})
	t.Run("getInt32", func(t *testing.T) {
		got, err := getInt32(fc, "Color")
		requireNullError(t, "Color", err)
		if got != 0 {
			t.Errorf("value on the error path = %v; want 0", got)
		}
	})
}

// TestTypedGetters_RejectNull drives the same VT_NULL through the PUBLIC
// wrappers a user actually calls, so the guard cannot be satisfied by a helper
// that no typed getter routes through. The five properties here are the ones
// Excel documents as Null-on-disagreement; the guard itself is generic, so a
// future getter inherits it.
func TestTypedGetters_RejectNull(t *testing.T) {
	cases := []struct {
		name string
		prop string
		call func(sugar.Chain) error
	}{
		{"Range.NumberFormat", "NumberFormat", func(c sugar.Chain) error { _, err := wrapRange(c).NumberFormat(); return err }},
		{"Range.ColumnWidth", "ColumnWidth", func(c sugar.Chain) error { _, err := wrapRange(c).ColumnWidth(); return err }},
		{"Range.RowHeight", "RowHeight", func(c sugar.Chain) error { _, err := wrapRange(c).RowHeight(); return err }},
		{"Range.MergeCells", "MergeCells", func(c sugar.Chain) error { _, err := wrapRange(c).MergeCells(); return err }},
		{"Range.Color", "Color", func(c sugar.Chain) error { _, err := wrapRange(c).Color(); return err }},
		{"Font.Name", "Name", func(c sugar.Chain) error { _, err := wrapFont(c).Name(); return err }},
		{"Font.Size", "Size", func(c sugar.Chain) error { _, err := wrapFont(c).Size(); return err }},
		{"Font.Bold", "Bold", func(c sugar.Chain) error { _, err := wrapFont(c).Bold(); return err }},
		{"Font.Italic", "Italic", func(c sugar.Chain) error { _, err := wrapFont(c).Italic(); return err }},
		{"Font.Color", "Color", func(c sugar.Chain) error { _, err := wrapFont(c).Color(); return err }},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			fc := newFakeChain()
			fc.root.value = sugar.Null{}
			requireNullError(t, tc.prop, tc.call(fc))
		})
	}
}

// TestScalarGetters_NullGuardHasNoFalsePositives is the mandatory other half:
// the guard must not be satisfiable by failing every read. A VT_EMPTY property
// (nil) keeps its old lenient coercion, because an unset string property IS
// legitimately the empty string and nothing distinguishes it from "absent".
func TestScalarGetters_NullGuardHasNoFalsePositives(t *testing.T) {
	fc := newFakeChain()

	fc.root.value = nil
	if got, err := getString(fc, "NumberFormat"); err != nil || got != "" {
		t.Errorf("nil (VT_EMPTY) through getString = %q, %v; want \"\", nil", got, err)
	}
	if got, err := getBool(fc, "MergeCells"); err != nil || got {
		t.Errorf("nil through getBool = %v, %v; want false, nil", got, err)
	}

	fc.root.value = "0.00"
	if got, err := getString(fc, "NumberFormat"); err != nil || got != "0.00" {
		t.Errorf("real format through getString = %q, %v; want \"0.00\", nil", got, err)
	}
	fc.root.value = 8.43
	if got, err := getFloat64(fc, "ColumnWidth"); err != nil || got != 8.43 {
		t.Errorf("real width through getFloat64 = %v, %v; want 8.43, nil", got, err)
	}
	fc.root.value = true
	if got, err := getBool(fc, "MergeCells"); err != nil || !got {
		t.Errorf("real bool through getBool = %v, %v; want true, nil", got, err)
	}
	fc.root.value = int32(255)
	if got, err := getInt32(fc, "Color"); err != nil || got != 255 {
		t.Errorf("real color through getInt32 = %v, %v; want 255, nil", got, err)
	}
}

// requireNullError asserts err is the Null rejection for prop and that the
// message tells the caller both what went wrong and how to opt out.
func requireNullError(t *testing.T, prop string, err error) {
	t.Helper()
	if err == nil {
		t.Fatalf("%s returned a nil error for a VT_NULL read", prop)
	}
	msg := err.Error()
	for _, want := range []string{prop, "Null", "sugar.IsNull"} {
		if !strings.Contains(msg, want) {
			t.Errorf("error %q should mention %q", msg, want)
		}
	}
}

// TestTrimTrailingMissing pins the optional-argument trim to the VALUE of the
// omitted-optional marker, not to the *ole.VARIANT type. A wrapper that hands
// callOptional a hand-built VARIANT for a picky COM slot must keep that
// argument; a type-only check silently dropped it.
func TestTrimTrailingMissing(t *testing.T) {
	real1 := ole.NewVariant(ole.VT_R8, 0)
	realVariant := &real1
	cases := []struct {
		name string
		args []interface{}
		want []interface{}
	}{
		{"trailing markers dropped", []interface{}{"path", sugar.Missing(), sugar.Missing()}, []interface{}{"path"}},
		{"all markers trim to empty", []interface{}{sugar.Missing(), sugar.Missing(), sugar.Missing()}, []interface{}{}},
		{"real trailing VARIANT is kept", []interface{}{"path", sugar.Missing(), realVariant}, []interface{}{"path", sugar.Missing(), realVariant}},
		{"real VARIANT before a marker survives", []interface{}{"path", realVariant, sugar.Missing()}, []interface{}{"path", realVariant}},
		{"no optionals", []interface{}{"path"}, []interface{}{"path"}},
		{"empty list", []interface{}{}, []interface{}{}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := trimTrailingMissing(tc.args)
			if len(got) != len(tc.want) {
				t.Fatalf("trimTrailingMissing kept %d args, want %d (%v)", len(got), len(tc.want), got)
			}
			for i := range got {
				// Missing() allocates a fresh VARIANT per call, so compare the
				// marker positionally by value and everything else by identity.
				if sugar.IsMissing(tc.want[i]) {
					if !sugar.IsMissing(got[i]) {
						t.Errorf("arg %d = %v, want the Missing marker", i, got[i])
					}
					continue
				}
				if !reflect.DeepEqual(got[i], tc.want[i]) {
					t.Errorf("arg %d = %v, want %v", i, got[i], tc.want[i])
				}
			}
		})
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
