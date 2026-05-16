//go:build windows && excel_integration

// This integration test exercises real Excel via COM and must be opted into:
//
//	go test -tags=excel_integration ./excel/...
//
// On hosts without Excel installed (or without a real COM server bound to the
// `Excel.Application` ProgID), `ctx.Create` succeeds far enough to return a
// Chain whose `.Err()` is nil, but later `Get`/`Put` calls fail with "CoInitialize
// not called" because there is no real COM server to dispatch through. Skipping
// at NewApplication time is therefore insufficient — the only reliable signal
// "Excel really is present" is whether actual operations succeed, which is the
// integration tier. Hence the build tag.

package excel_test

import (
	"testing"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

// TestApplication_BoolProperties exercises the three xlwings-parity boolean
// properties on `excel.Application`: Visible, DisplayAlerts, ScreenUpdating.
// It launches a real Excel instance, so it requires Excel to be installed.
//
// IMPORTANT: this test deliberately does NOT use `t.Run` subtests. `t.Run`
// schedules each subtest on a fresh goroutine, which has not had
// CoInitialize called on its OS thread — every COM call then fails with
// "CoInitialize was not called". The original v0.5.0 form of this test
// hit exactly that trap; we now drive all three properties from a single
// loop inside the COM-initialized sugar.Do callback.
func TestApplication_BoolProperties(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}

		// Ensure cleanup: silence prompts then quit regardless of test outcome.
		defer app.SetDisplayAlerts(false).Quit()

		cases := []struct {
			name string
			set  func(v bool) excel.Application
			get  func() sugar.Chain
		}{
			{name: "Visible", set: app.SetVisible, get: app.Visible},
			{name: "DisplayAlerts", set: app.SetDisplayAlerts, get: app.DisplayAlerts},
			{name: "ScreenUpdating", set: app.SetScreenUpdating, get: app.ScreenUpdating},
		}

		for _, tc := range cases {
			for _, want := range []bool{true, false} {
				a := tc.set(want)
				if err := a.Err(); err != nil {
					t.Fatalf("Set%s(%v) failed: %v", tc.name, want, err)
				}
				got, err := tc.get().Value()
				if err != nil {
					t.Fatalf("%s getter failed: %v", tc.name, err)
				}
				if got != want {
					t.Errorf("%s: set %v, got %v", tc.name, want, got)
				}
			}
		}

		return nil
	})
}

// TestApplication_Identity exercises the read-only identity properties
// added in v0.7.0: Version, PID, Hwnd. All three should return non-zero
// values for a real Excel instance; the test self-skips if Excel is missing.
func TestApplication_Identity(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		defer app.SetDisplayAlerts(false).Quit()

		ver, err := app.Version()
		if err != nil || ver == "" {
			t.Errorf("Version: got %q, err=%v", ver, err)
		}

		hwnd, err := app.Hwnd()
		if err != nil || hwnd == 0 {
			t.Errorf("Hwnd: got %d, err=%v", hwnd, err)
		}

		pid, err := app.PID()
		if err != nil || pid == 0 {
			t.Errorf("PID: got %d, err=%v", pid, err)
		}
		return nil
	})
}

// TestApplication_Calculation round-trips the Calculation property through
// every defined mode and verifies the value persists.
func TestApplication_Calculation(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		app := excel.NewApplication(ctx)
		if err := app.Err(); err != nil {
			t.Skip("Excel not installed:", err)
			return nil
		}
		defer app.SetDisplayAlerts(false).Quit()

		// Calculation requires at least one open workbook to read/write.
		app.Workbooks().Add()

		for _, want := range []excel.Calculation{
			excel.CalculationManual,
			excel.CalculationAutomatic,
		} {
			if err := app.SetCalculation(want).Err(); err != nil {
				t.Fatalf("SetCalculation(%d) failed: %v", want, err)
			}
			got, err := app.Calculation()
			if err != nil {
				t.Fatalf("Calculation getter failed: %v", err)
			}
			if got != want {
				t.Errorf("Calculation: set %d, got %d", want, got)
			}
		}
		return nil
	})
}
