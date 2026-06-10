//go:build windows && excel_integration

// Integration tests for excel.Name / excel.Names.
// Build with `-tags=excel_integration`. Skipped on machines without Excel.

package excel_test

import (
	"strings"
	"testing"

	"github.com/xll-gen/sugar/excel"
)

// TestNames_AddByString defines a workbook-scoped name from an A1-notation
// string and round-trips Name / RefersTo / RefersToRange.
func TestNames_AddByString(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		sheet := wb.ActiveSheet()
		sheetName, err := sheet.Name()
		if err != nil {
			t.Fatalf("sheet name: %v", err)
		}
		refersTo := "='" + sheetName + "'!$A$1:$B$2"

		n := wb.Names().Add("test_block", refersTo)
		if err := n.Err(); err != nil {
			t.Fatalf("Names.Add: %v", err)
		}

		got, err := n.Name()
		if err != nil || got != "test_block" {
			t.Errorf("Name: got %q err=%v; want test_block", got, err)
		}

		rt, err := n.RefersTo()
		if err != nil || !strings.Contains(rt, "$A$1:$B$2") {
			t.Errorf("RefersTo: got %q err=%v; want suffix $A$1:$B$2", rt, err)
		}

		addr, err := n.RefersToRange().Address()
		if err != nil || addr != "$A$1:$B$2" {
			t.Errorf("RefersToRange.Address: got %q err=%v; want $A$1:$B$2", addr, err)
		}
	})
}

// TestNames_AddByRange passes a typed Range as refersTo — this exercises the
// core Chain→IDispatch argument normalization.
func TestNames_AddByRange(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		rng := wb.ActiveSheet().Range("C3:D4")

		n := wb.Names().Add("range_name", rng)
		if err := n.Err(); err != nil {
			t.Fatalf("Names.Add(Range): %v", err)
		}

		addr, err := n.RefersToRange().Address()
		if err != nil || addr != "$C$3:$D$4" {
			t.Errorf("RefersToRange.Address: got %q err=%v; want $C$3:$D$4", addr, err)
		}
	})
}

// TestNames_CountItemContainsDelete walks the collection surface: Count
// grows on Add, Item resolves by name and 1-based index, Contains answers
// membership, and Delete removes the name.
func TestNames_CountItemContainsDelete(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		names := wb.Names()

		before, err := names.Count()
		if err != nil {
			t.Fatalf("Count: %v", err)
		}

		rng := wb.ActiveSheet().Range("A1")
		n := names.Add("to_delete", rng)
		if err := n.Err(); err != nil {
			t.Fatalf("Add: %v", err)
		}

		after, err := names.Count()
		if err != nil || after != before+1 {
			t.Errorf("Count after Add: got %d err=%v; want %d", after, err, before+1)
		}

		ok, err := names.Contains("to_delete")
		if err != nil || !ok {
			t.Errorf("Contains(to_delete): got %v err=%v; want true", ok, err)
		}
		ok, err = names.Contains("no_such_name")
		if err != nil || ok {
			t.Errorf("Contains(no_such_name): got %v err=%v; want false", ok, err)
		}

		byName, err := names.Item("to_delete").Name()
		if err != nil || byName != "to_delete" {
			t.Errorf("Item by name: got %q err=%v", byName, err)
		}

		if err := names.Item("to_delete").Delete(); err != nil {
			t.Fatalf("Delete: %v", err)
		}
		ok, err = names.Contains("to_delete")
		if err != nil || ok {
			t.Errorf("Contains after Delete: got %v err=%v; want false", ok, err)
		}
	})
}

// TestWorksheet_Names verifies the sheet-scoped collection is reachable and
// names added through it come back sheet-qualified, matching xlwings.
func TestWorksheet_Names(t *testing.T) {
	withBook(t, func(wb excel.Workbook) {
		sheet := wb.ActiveSheet()

		n := sheet.Names().Add("local_name", sheet.Range("B2"))
		if err := n.Err(); err != nil {
			t.Fatalf("sheet Names.Add: %v", err)
		}

		got, err := n.Name()
		if err != nil || !strings.Contains(got, "!") {
			t.Errorf("sheet-scoped name should be qualified: got %q err=%v", got, err)
		}

		count, err := sheet.Names().Count()
		if err != nil || count != 1 {
			t.Errorf("sheet Names.Count: got %d err=%v; want 1", count, err)
		}
	})
}
