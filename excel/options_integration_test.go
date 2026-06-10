//go:build windows && excel_integration

// Integration tests for the excel.Options conversion framework.
//
// These tests cover the option flags that depend on real COM behaviour and
// cannot be faked: Expand (uses Range.End / Worksheet.Range) and the
// struct-by-header decode driven off Range.Value. Run with:
//
//	go test -tags=excel_integration ./excel/...

package excel_test

import (
	"reflect"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestOptions_ExpandDown grows a one-cell anchor down through a contiguous
// column of values and verifies the resulting range covers the whole block.
// xlwings parity: `rng.options(expand="down").value`.
func TestOptions_ExpandDown(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(1.0)
		sheet.Range("A2").SetValue(2.0)
		sheet.Range("A3").SetValue(3.0)
		// A4 deliberately empty — Expand stops at the gap.

		v, err := sheet.Range("A1").Options(excel.Expand("down")).Value()
		if err != nil {
			t.Fatalf("Options(Expand=down).Value(): %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{} for 3x1 expand, got %T (%v)", v, v)
		}
		want := []interface{}{1.0, 2.0, 3.0}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_ExpandRight is the row analogue of ExpandDown.
func TestOptions_ExpandRight(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(10.0)
		sheet.Range("B1").SetValue(20.0)
		sheet.Range("C1").SetValue(30.0)

		v, err := sheet.Range("A1").Options(excel.Expand("right")).Value()
		if err != nil {
			t.Fatalf("Expand=right: %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{}, got %T (%v)", v, v)
		}
		want := []interface{}{10.0, 20.0, 30.0}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_ExpandTable seeds a 2x3 block and verifies expand="table"
// walks both directions from the anchor.
func TestOptions_ExpandTable(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(1.0)
		sheet.Range("B1").SetValue(2.0)
		sheet.Range("C1").SetValue(3.0)
		sheet.Range("A2").SetValue(4.0)
		sheet.Range("B2").SetValue(5.0)
		sheet.Range("C2").SetValue(6.0)

		v, err := sheet.Range("A1").Options(excel.Expand("table")).Value()
		if err != nil {
			t.Fatalf("Expand=table: %v", err)
		}
		got, ok := v.([][]interface{})
		if !ok {
			t.Fatalf("expected [][]interface{}, got %T (%v)", v, v)
		}
		want := [][]interface{}{
			{1.0, 2.0, 3.0},
			{4.0, 5.0, 6.0},
		}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}

// TestOptions_StructDecode is the end-to-end happy path for the struct-by-
// header decode driven directly from real Excel data.
func TestOptions_StructDecode(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("Name")
		sheet.Range("B1").SetValue("Age")
		sheet.Range("A2").SetValue("alice")
		sheet.Range("B2").SetValue(30.0)
		sheet.Range("A3").SetValue("bob")
		sheet.Range("B3").SetValue(25.0)

		var rows []Row
		err := sheet.Range("A1", "B3").Options(excel.Header(true)).Get(&rows)
		if err != nil {
			t.Fatalf("Get(&rows): %v", err)
		}
		want := []Row{{Name: "alice", Age: 30}, {Name: "bob", Age: 25}}
		if !reflect.DeepEqual(rows, want) {
			t.Errorf("got %+v, want %+v", rows, want)
		}
	})
}

// TestOptions_ExpandTableAndStruct chains Expand + Header(true) — the
// typical xlwings idiom `rng.options(MyType, header=1, expand="table")`.
func TestOptions_ExpandTableAndStruct(t *testing.T) {
	type Row struct {
		Name string
		Age  int
	}
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("Name")
		sheet.Range("B1").SetValue("Age")
		sheet.Range("A2").SetValue("alice")
		sheet.Range("B2").SetValue(30.0)
		sheet.Range("A3").SetValue("bob")
		sheet.Range("B3").SetValue(25.0)
		// A4 empty -> expand=table stops here.

		var rows []Row
		err := sheet.Range("A1").Options(
			excel.Expand("table"),
			excel.Header(true),
		).Get(&rows)
		if err != nil {
			t.Fatalf("Expand+Header: %v", err)
		}
		want := []Row{{Name: "alice", Age: 30}, {Name: "bob", Age: 25}}
		if !reflect.DeepEqual(rows, want) {
			t.Errorf("got %+v, want %+v", rows, want)
		}
	})
}

// TestOptions_ScalarForcing exercises the .options(Scalar()) shape-forcing
// path against real Excel: a 1x1 read should hand back the bare scalar.
func TestOptions_ScalarForcing(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue(123.5)
		v, err := sheet.Range("A1").Options(excel.Scalar()).Value()
		if err != nil {
			t.Fatalf("Scalar.Value: %v", err)
		}
		if v != 123.5 {
			t.Errorf("got %v (%T), want 123.5", v, v)
		}
	})
}

// TestOptions_SetGetStructRoundTrip writes a struct slice with a header row
// and reads it back through the header decode — the full write/read mirror.
func TestOptions_SetGetStructRoundTrip(t *testing.T) {
	type Person struct {
		Name string
		Age  float64
	}
	withSheet(t, func(sheet excel.Worksheet) {
		in := []Person{{"alice", 30}, {"bob", 25}}

		if err := sheet.Range("A1").Options(excel.Header(true)).Set(in); err != nil {
			t.Fatalf("Set: %v", err)
		}

		// The block now spans A1:B3 (header + 2 rows). Expand("table") from
		// the anchor must rediscover it.
		var out []Person
		err := sheet.Range("A1").
			Options(excel.Expand("table"), excel.Header(true)).
			Get(&out)
		if err != nil {
			t.Fatalf("Get: %v", err)
		}
		if !reflect.DeepEqual(out, in) {
			t.Errorf("round trip: got %+v, want %+v", out, in)
		}
	})
}

// TestOptions_EmptyReplacement validates that Empty(value) substitutes nil
// cells in the raw read.
func TestOptions_EmptyReplacement(t *testing.T) {
	withSheet(t, func(sheet excel.Worksheet) {
		sheet.Range("A1").SetValue("x")
		// B1 left blank.
		sheet.Range("C1").SetValue("z")

		v, err := sheet.Range("A1", "C1").
			Options(excel.Empty("N/A"), excel.Vector()).
			Value()
		if err != nil {
			t.Fatalf("Empty+Vector: %v", err)
		}
		got, ok := v.([]interface{})
		if !ok {
			t.Fatalf("expected []interface{}, got %T", v)
		}
		want := []interface{}{"x", "N/A", "z"}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("got %v, want %v", got, want)
		}
	})
}
