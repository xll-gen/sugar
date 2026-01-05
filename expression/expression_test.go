//go:build windows

package expression

import (
	"testing"

	"github.com/xll-gen/sugar"
)

// setupExcel creates an Excel instance using the context.
func setupExcel(t *testing.T, ctx *sugar.Context) *sugar.Chain {
	excel := ctx.Create("Excel.Application")
	if err := excel.Err(); err != nil {
		t.Skip("Excel not installed or failed to create:", err)
		return nil
	}
	return excel
}

func TestEval_Basic(t *testing.T) {
	// Test without COM
	res, err := Eval("2 + 2", nil)
	if err != nil {
		t.Fatalf("Eval failed: %v", err)
	}
	if res.(float64) != 4 {
		t.Errorf("Expected 4, got %v", res)
	}

	res, err = Eval("'Hello ' + 'Sugar'", nil)
	if err != nil {
		t.Fatalf("Eval failed: %v", err)
	}
	if res.(string) != "Hello Sugar" {
		t.Errorf("Expected 'Hello Sugar', got %v", res)
	}
}

func TestEval_CompileRun(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		p, err := Compile("Version")
		if err != nil {
			t.Fatalf("Compile failed: %v", err)
		}

		res, err := p.Run(excel)
		if err != nil {
			t.Fatalf("Run failed: %v", err)
		}
		
		if res == nil {
			t.Fatal("Run returned nil")
		}
		return nil
	})
}

func TestEval_EnvMap(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		env := map[string]interface{}{
			"ex": excel,
			"val": 10,
		}

		res, err := Eval("ex.Version", env)
		if err != nil {
			t.Fatalf("Eval with map failed: %v", err)
		}
		if res == nil {
			t.Fatal("result is nil")
		}

		res, err = Eval("val * 2", env)
		if err != nil {
			t.Fatalf("arithmetic with map failed: %v", err)
		}
		if res.(float64) != 20 {
			t.Errorf("expected 20, got %v", res)
		}
		return nil
	})
}

func TestGet_Legacy(t *testing.T) {
	sugar.Do(func(ctx *sugar.Context) error {
		excel := setupExcel(t, ctx)
		if excel == nil { return nil }
		defer excel.Put("DisplayAlerts", false).Call("Quit")

		version, err := Get(excel, "Version")
		if err != nil {
			t.Fatalf("Get failed: %v", err)
		}
		if version == nil {
			t.Fatal("Get returned nil")
		}
		return nil
	})
}