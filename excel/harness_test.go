//go:build windows && excel_integration

// Shared harness for the excel package's COM integration tests. Run with:
//
//	go test -tags=excel_integration ./excel/...
//
// Why a build tag instead of a runtime skip: on hosts without Excel installed
// (or without a real COM server bound to the `Excel.Application` ProgID),
// `ctx.Create` can succeed far enough to return a Chain whose `.Err()` is nil,
// but later `Get`/`Put` calls fail because there is no real COM server to
// dispatch through. Skipping at NewApplication time is therefore insufficient
// — the only reliable signal "Excel really is present" is whether actual
// operations succeed, which is the integration tier. Hence the build tag.
//
// Cleanup is two-tier (see internal/testutil):
//
//  1. Graceful: withApp defers app.Quit() inside the sugar.Do block (with
//     DisplayAlerts already off), while COM is still initialized.
//  2. Force-kill: withApp registers testutil.EnsureProcessExited on the PID
//     via t.Cleanup, so a hung Quit (modal prompt, zombie process) cannot
//     leak an invisible EXCEL.EXE past the test.

package excel_test

import (
	"testing"
	"time"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
	"github.com/xll-gen/sugar/internal/testutil"
)

// withApp launches a hidden Excel instance with alerts suppressed, hands it
// to fn, and guarantees the process is gone afterwards (two-tier cleanup —
// see the package comment above).
func withApp(t *testing.T, fn func(app excel.Application)) {
	t.Helper()
	err := sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		app.SetVisible(false).SetDisplayAlerts(false)

		if pid, err := app.PID(); err == nil && pid != 0 {
			t.Cleanup(func() { testutil.EnsureProcessExited(t, pid, 5*time.Second) })
		} else {
			t.Logf("could not resolve Excel PID for force-kill cleanup: %v", err)
		}
		defer app.Quit() // graceful tier; force-kill tier registered above

		fn(app)
		return nil
	})
	if err != nil {
		t.Fatalf("sugar.Do: %v", err)
	}
}

// withBook runs fn against a fresh workbook in a managed Excel instance.
func withBook(t *testing.T, fn func(wb excel.Workbook)) {
	t.Helper()
	withApp(t, func(app excel.Application) {
		wb := app.Workbooks().Add()
		if err := wb.Err(); err != nil {
			t.Fatalf("Add workbook failed: %v", err)
		}
		fn(wb)
	})
}

// withSheet runs fn against the active sheet of a fresh workbook.
func withSheet(t *testing.T, fn func(sheet excel.Worksheet)) {
	t.Helper()
	withBook(t, func(wb excel.Workbook) {
		sheet := wb.ActiveSheet()
		if err := sheet.Err(); err != nil {
			t.Fatalf("ActiveSheet failed: %v", err)
		}
		fn(sheet)
	})
}
