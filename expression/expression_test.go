//go:build windows

// The expression engine is COM-server-agnostic, so these tests drive
// Scripting.Dictionary (present on every Windows host) instead of Excel —
// `go test ./...` runs them without spawning Office processes.

package expression

import (
	"fmt"
	"strings"
	"testing"

	"github.com/xll-gen/sugar"
)

// setupDict creates a Scripting.Dictionary seeded with one entry ("k" -> "v")
// so property reads like Count have a non-trivial value to return.
func setupDict(t *testing.T, ctx sugar.Context) sugar.Chain {
	t.Helper()
	dict := ctx.Create("Scripting.Dictionary")
	if err := dict.Err(); err != nil {
		t.Fatalf("Scripting.Dictionary create failed: %v", err)
	}
	if err := dict.Call("Add", "k", "v").Err(); err != nil {
		t.Fatalf("Dictionary.Add failed: %v", err)
	}
	return dict
}

func TestEval_Basic(t *testing.T) {
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
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)

		p, err := Compile("Count")
		if err != nil {
			t.Fatalf("Compile failed: %v", err)
		}

		res, err := p.Run(dict)
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
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)

		env := map[string]interface{}{
			"d":   dict,
			"val": 10,
		}

		res, err := Eval("d.Count", env)
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
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)

		count, err := Get(dict, "Count")
		if err != nil {
			t.Fatalf("Get failed: %v", err)
		}
		if count == nil {
			t.Fatal("Get returned nil")
		}
		return nil
	})
}

func TestEval_UnaryOperators(t *testing.T) {
	cases := []struct {
		expr string
		want interface{}
	}{
		{"-1", float64(-1)},
		{"+5", float64(5)},
		{"2 + -3", float64(-1)},
		{"-(1 + 2)", float64(-3)},
		{"!true", false},
		{"not false", true},
	}
	for _, tc := range cases {
		res, err := Eval(tc.expr, nil)
		if err != nil {
			t.Errorf("Eval(%q) failed: %v", tc.expr, err)
			continue
		}
		if res != tc.want {
			t.Errorf("Eval(%q) = %v (%T), want %v (%T)", tc.expr, res, res, tc.want, tc.want)
		}
	}
}

// TestEval_IndexSyntax exercises the COM default-member mapping: d[1] must
// become d.Get("Item", 1), not Get("") as it silently did before.
func TestEval_IndexSyntax(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Call("Add", 1, "one").Err(); err != nil {
			t.Fatalf("Dictionary.Add failed: %v", err)
		}
		env := map[string]interface{}{"d": dict}

		res, err := Get(env, "d[1]")
		if err != nil {
			t.Fatalf("Get(d[1]) failed: %v", err)
		}
		if res != "one" {
			t.Errorf("d[1] = %v, want \"one\"", res)
		}

		// Unsupported (non-string, non-numeric) index expressions must
		// produce a clear error instead of Get("").
		if _, err := Get(env, "d[true]"); err == nil ||
			!strings.Contains(err.Error(), "unsupported property expression") {
			t.Errorf("d[true]: want unsupported-property error, got %v", err)
		}
		return nil
	})
}

// TestCallNode_ChainArgPassedAsObject verifies that a COM object argument is
// passed through as an object instead of being flattened to nil via Value().
func TestCallNode_ChainArgPassedAsObject(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		d := ctx.Create("Scripting.Dictionary")
		if err := d.Err(); err != nil {
			t.Fatalf("create failed: %v", err)
		}
		e := setupDict(t, ctx) // holds one entry
		env := map[string]interface{}{"d": d, "e": e}

		if _, err := Eval("d.Add(2, e)", env); err != nil {
			t.Fatalf("d.Add(2, e) failed: %v", err)
		}

		// If e had been dropped to nil, d[2] would be empty and the Count
		// access would fail. (Dictionary.Item is a propget, so retrieval
		// must use index syntax — Get("Item", 2) — not a method call.)
		res, err := Get(env, "d[2].Count")
		if err != nil {
			t.Fatalf("d[2].Count failed: %v", err)
		}
		if fmt.Sprint(res) != "1" {
			t.Errorf("stored object Count = %v, want 1", res)
		}
		return nil
	})
}

func TestEnvMap_CollisionError(t *testing.T) {
	env := map[string]interface{}{"d": 42}
	_, err := Eval("d(1)", env)
	if err == nil || !strings.Contains(err.Error(), "not callable") {
		t.Errorf("want not-callable error for env collision, got %v", err)
	}
}

func TestEnvMap_Function(t *testing.T) {
	env := map[string]interface{}{
		"double": func(args ...interface{}) (interface{}, error) {
			return float64(args[0].(int)) * 2, nil
		},
	}
	res, err := Eval("double(21)", env)
	if err != nil {
		t.Fatalf("Eval failed: %v", err)
	}
	if res != float64(42) {
		t.Errorf("double(21) = %v, want 42", res)
	}
}

func TestUnsupportedEnvType(t *testing.T) {
	if _, err := Eval("x", 42); err == nil ||
		!strings.Contains(err.Error(), "unsupported env type") {
		t.Errorf("Eval: want unsupported-env-type error, got %v", err)
	}
	if err := Put(42, "x.y", 1); err == nil ||
		!strings.Contains(err.Error(), "unsupported env type") {
		t.Errorf("Put: want unsupported-env-type error, got %v", err)
	}
}

// TestEval_DeferredChainError verifies that a Chain carrying a deferred COM
// error is unwrapped by Run instead of being returned as a "successful" Chain.
func TestEval_DeferredChainError(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)
		res, err := Eval("NoSuchProperty123", dict)
		if err == nil {
			t.Errorf("want deferred chain error, got result %v", res)
		}
		return nil
	})
}

func TestPut_RootIdentifier(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		// CompareMode can only be set while the dictionary is empty.
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			t.Fatalf("create failed: %v", err)
		}
		if err := Put(dict, "CompareMode", 1); err != nil {
			t.Fatalf("Put(CompareMode) failed: %v", err)
		}
		res, err := Get(dict, "CompareMode")
		if err != nil {
			t.Fatalf("Get(CompareMode) failed: %v", err)
		}
		if fmt.Sprint(res) != "1" {
			t.Errorf("CompareMode = %v, want 1", res)
		}

		// Root-identifier Put must also work through the raw-IDispatch
		// arena path.
		raw, err := dict.Store()
		if err != nil {
			t.Fatalf("Store failed: %v", err)
		}
		defer raw.Release()
		if err := Put(raw, "CompareMode", 0); err != nil {
			t.Fatalf("Put(raw, CompareMode) failed: %v", err)
		}
		res, err = Get(raw, "CompareMode")
		if err != nil {
			t.Fatalf("Get(raw, CompareMode) failed: %v", err)
		}
		if fmt.Sprint(res) != "0" {
			t.Errorf("CompareMode = %v, want 0", res)
		}
		return nil
	})
}

func TestPut_IndexedProperty(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Call("Add", 1, "one").Err(); err != nil {
			t.Fatalf("Dictionary.Add failed: %v", err)
		}
		env := map[string]interface{}{"d": dict}

		if err := Put(env, "d[1]", "uno"); err != nil {
			t.Fatalf("Put(d[1]) failed: %v", err)
		}
		res, err := Get(env, "d[1]")
		if err != nil {
			t.Fatalf("Get(d[1]) failed: %v", err)
		}
		if res != "uno" {
			t.Errorf("d[1] = %v, want \"uno\"", res)
		}
		return nil
	})
}

// TestIDispatchEnv_ArenaNoLeak drives the raw-IDispatch env path several
// times and asserts the root object's refcount is unchanged afterwards.
// Scripting.Dictionary is in-process, so the count returned by AddRef is
// deterministic here.
func TestIDispatchEnv_ArenaNoLeak(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)
		raw, err := dict.Store()
		if err != nil {
			t.Fatalf("Store failed: %v", err)
		}
		defer raw.Release()

		base := raw.AddRef()
		raw.Release()

		for i := 0; i < 3; i++ {
			res, err := Get(raw, "Count")
			if err != nil {
				t.Fatalf("Get(raw, Count) failed: %v", err)
			}
			if fmt.Sprint(res) != "1" {
				t.Errorf("Count = %v, want 1", res)
			}
		}

		after := raw.AddRef()
		raw.Release()
		if after != base {
			t.Errorf("refcount leak on raw IDispatch env path: %d -> %d", base, after)
		}
		return nil
	})
}

// TestIDispatchEnv_IntermediatesReleased evaluates an expression that creates
// an intermediate COM object (an FSO Folder) inside the arena and extracts a
// scalar; the intermediate must be released before Get returns.
func TestIDispatchEnv_IntermediatesReleased(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		fso := ctx.Create("Scripting.FileSystemObject")
		if err := fso.Err(); err != nil {
			t.Fatalf("FSO create failed: %v", err)
		}
		raw, err := fso.Store()
		if err != nil {
			t.Fatalf("Store failed: %v", err)
		}
		defer raw.Release()

		// GetSpecialFolder(0) is the Windows folder.
		res, err := Get(raw, "GetSpecialFolder(0).Name")
		if err != nil {
			t.Fatalf("Get failed: %v", err)
		}
		name, ok := res.(string)
		if !ok || name == "" {
			t.Errorf("GetSpecialFolder(0).Name = %v (%T), want non-empty string", res, res)
		}
		return nil
	})
}

// TestIDispatchEnv_StoreEscapesArena verifies that a COM-object result of the
// raw-IDispatch env path escapes the internal arena alive: the caller owns
// one reference and can keep using the object after Eval/Store return.
func TestIDispatchEnv_StoreEscapesArena(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		fso := ctx.Create("Scripting.FileSystemObject")
		if err := fso.Err(); err != nil {
			t.Fatalf("FSO create failed: %v", err)
		}
		raw, err := fso.Store()
		if err != nil {
			t.Fatalf("Store failed: %v", err)
		}
		defer raw.Release()

		folderDisp, err := Store(raw, "GetSpecialFolder(1)") // system folder
		if err != nil {
			t.Fatalf("Store(GetSpecialFolder) failed: %v", err)
		}
		if folderDisp == nil {
			t.Fatal("Store returned nil dispatch")
		}
		defer folderDisp.Release()

		// The arena released its references already; the escaped reference
		// must keep the object alive.
		folder := ctx.From(folderDisp)
		name, err := folder.Get("Name").Value()
		if err != nil {
			t.Fatalf("escaped folder unusable: %v", err)
		}
		if s, ok := name.(string); !ok || s == "" {
			t.Errorf("folder Name = %v (%T), want non-empty string", name, name)
		}

		// Get must refuse COM-object results on this path (and release the
		// escaped reference itself).
		if _, err := Get(raw, "GetSpecialFolder(1)"); err == nil ||
			!strings.Contains(err.Error(), "use Store") {
			t.Errorf("Get on object result: want use-Store error, got %v", err)
		}
		return nil
	})
}
