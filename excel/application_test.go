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
// The existing excel_test.go file uses t.Skip when Excel is unavailable; we
// follow the same convention here rather than gating with an extra build tag,
// so the test runs out-of-the-box on developer Windows boxes and self-skips
// on CI hosts without Excel.
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
			{
				name: "Visible",
				set:  app.SetVisible,
				get:  app.Visible,
			},
			{
				name: "DisplayAlerts",
				set:  app.SetDisplayAlerts,
				get:  app.DisplayAlerts,
			},
			{
				name: "ScreenUpdating",
				set:  app.SetScreenUpdating,
				get:  app.ScreenUpdating,
			},
		}

		for _, tc := range cases {
			t.Run(tc.name, func(t *testing.T) {
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
			})
		}

		return nil
	})
}
