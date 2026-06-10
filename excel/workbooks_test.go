//go:build windows && excel_integration

// Integration tests for excel.Workbooks.Open options.
// Build with `-tags=excel_integration`.

package excel_test

import (
	"path/filepath"
	"testing"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

// TestWorkbooks_OpenReadOnly saves a workbook, reopens it read-only, and
// checks Excel agrees via the ReadOnly property.
func TestWorkbooks_OpenReadOnly(t *testing.T) {
	path := filepath.Join(t.TempDir(), "readonly_test.xlsx")

	withApp(t, func(app excel.Application) {
		wb := app.Workbooks().Add()
		if err := wb.SaveAs(path); err != nil {
			t.Fatalf("SaveAs: %v", err)
		}
		if err := wb.Close(); err != nil {
			t.Fatalf("Close: %v", err)
		}

		ro := app.Workbooks().Open(path, excel.OpenReadOnly())
		if err := ro.Err(); err != nil {
			t.Fatalf("Open(ReadOnly): %v", err)
		}
		v, err := ro.Get("ReadOnly").Value()
		if err != nil {
			t.Fatalf("ReadOnly property: %v", err)
		}
		if isRO, _ := v.(bool); !isRO {
			t.Errorf("workbook should be read-only, got %v", v)
		}
		_ = ro.SetSaved(true).Close()
	})
}

// TestWorkbooks_OpenWithPassword saves a password-protected workbook (via
// the raw chain — SaveAs' Password is positional parameter 3) and reopens
// it through OpenPassword. The failure mode is asserted with an explicit
// *wrong* password: omitting the password entirely would pop Excel's modal
// password prompt and hang the suite — DisplayAlerts does not suppress it.
func TestWorkbooks_OpenWithPassword(t *testing.T) {
	path := filepath.Join(t.TempDir(), "password_test.xlsx")
	const pw = "sugar-secret"

	withApp(t, func(app excel.Application) {
		wb := app.Workbooks().Add()
		// SaveAs(Filename, FileFormat, Password, ...) — skip FileFormat.
		if err := wb.Call("SaveAs", path, sugar.Missing(), pw).Err(); err != nil {
			t.Fatalf("SaveAs with password: %v", err)
		}
		if err := wb.Close(); err != nil {
			t.Fatalf("Close: %v", err)
		}

		// A wrong password fails fast with an error; no password would
		// prompt interactively and hang.
		bad := app.Workbooks().Open(path, excel.OpenPassword("wrong-password"))
		if bad.Err() == nil {
			t.Errorf("Open with wrong password should error")
			_ = bad.SetSaved(true).Close()
		}

		good := app.Workbooks().Open(path, excel.OpenPassword(pw))
		if err := good.Err(); err != nil {
			t.Fatalf("Open(Password): %v", err)
		}
		name, err := good.Name()
		if err != nil || name != "password_test.xlsx" {
			t.Errorf("opened workbook name: got %q err=%v", name, err)
		}
		_ = good.SetSaved(true).Close()
	})
}
