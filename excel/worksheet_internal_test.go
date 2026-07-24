//go:build windows

// Excel-free unit test for Worksheet.Range argument arity. Runs under plain
// `go test ./...`: a nil-dispatch chain never touches COM, so the arity guard
// is checked before any Invoke.

package excel

import (
	"strings"
	"testing"

	"github.com/xll-gen/sugar"
)

// TestWorksheetRange_TooManyArgs pins the item-5c fix: Range accepts at most
// two cell anchors, and extra arguments must surface as a chain error rather
// than being silently discarded. Before the fix, Range("A1","B2","C3") dropped
// "C3" and produced only a generic "dispatch is nil" error (or, on a live
// sheet, a wrong two-anchor range).
func TestWorksheetRange_TooManyArgs(t *testing.T) {
	w := wrapWorksheet(sugar.From(nil))

	err := w.Range("A1", "B2", "C3").Err()
	if err == nil {
		t.Fatal("Range with 3 cell arguments should error")
	}
	if !strings.Contains(err.Error(), "at most 2") {
		t.Errorf("expected an arity error mentioning 'at most 2', got: %v", err)
	}
}
