//go:build windows

// Non-integration unit test for GetApplicationByPID's failure path. It needs
// no Excel: a PID with no XLMAIN window must surface an error on the returned
// Application (not panic, not attach to the wrong instance). This pins the
// multi-instance attach helper that fixes the "ribbon click does nothing —
// cannot attach to Excel: 작업을 사용할 수 없습니다" bug (the Go server is a
// separate process, so the ROT-based GetApplication fails for it).

package excel_test

import (
	"context"
	"testing"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
)

func TestGetApplicationByPID_NoWindow(t *testing.T) {
	// A PID that owns no XLMAIN window. We use an impossible/última PID value;
	// even if it happened to be live, it is not an Excel frame, so the window
	// walk returns no XLMAIN and the helper must report an error rather than
	// attach to some other Excel via the ROT.
	const bogusPID = uint32(0xFFFFFFFE)

	ctx := sugar.NewContext(context.Background())
	defer ctx.Release()

	app := excel.GetApplicationByPID(ctx, bogusPID)
	if app == nil {
		t.Fatal("GetApplicationByPID returned nil Application")
	}
	if err := app.Err(); err == nil {
		t.Fatal("GetApplicationByPID(bogus PID) must return an error chain, got nil error")
	}
}
