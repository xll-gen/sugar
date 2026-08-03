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

// TestEval_UnsupportedBinaryMessage documents the item-5a format: the
// unsupported-binary-operation error names both operand types symmetrically via
// %T. Subtracting a string from a number reaches evalBinary's fallthrough.
func TestEval_UnsupportedBinaryMessage(t *testing.T) {
	_, err := Eval("2 - 'a'", nil)
	if err == nil {
		t.Fatal("expected an error for number - string")
	}
	msg := err.Error()
	if !strings.Contains(msg, "unsupported binary operation:") {
		t.Fatalf("unexpected error: %v", msg)
	}
	// Both operand types must appear (the %T substitutions), not just one side:
	// the left operand is numeric and the right is a string.
	if !strings.Contains(msg, "int") || !strings.Contains(msg, "string") {
		t.Errorf("error should name both operand types, got: %v", msg)
	}
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

// TestEvalBinary_Comparisons covers the six comparison operators the expression
// grammar gained. Numeric operands go through the same toFloat coercion the
// arithmetic operators use (so every number is compared as a float64), strings
// compare lexicographically, bools support only equality, and nil is only
// comparable for equality.
//
// The mixed-kind rows are as load-bearing as the true/false ones: a permissive
// fallback (comparing fmt.Sprint of both sides, say) would make every one of
// them "work" and silently answer wrong.
func TestEvalBinary_Comparisons(t *testing.T) {
	ok := []struct {
		expr string
		want interface{}
	}{
		{"1 == 1", true},
		{"1 == 2", false},
		{"1 != 2", true},
		{"1 < 2", true},
		{"2 < 2", false},
		{"2 <= 2", true},
		{"3 > 2", true},
		{"2 >= 3", false},
		// BOUNDARY-TRUE rows for the two inclusive operators. These are the only rows
		// that separate >= from > and <= from <, so without them a >= implemented as a
		// plain > passes the whole suite -- verified: that mutant survived until these
		// were added. Every other >= row here is false-direction, which a bare > also
		// satisfies.
		{"2 >= 2", true},
		{"2 <= 1", false},
		// Mixed numeric widths still compare as numbers (int vs float).
		{"1 == 1.0", true},
		{"1.5 > 1", true},
		{"-1 < 0", true},
		// Strings compare lexicographically.
		{"'a' == 'a'", true},
		{"'a' == 'b'", false},
		{"'a' < 'b'", true},
		{"'b' <= 'a'", false},
		// String boundary-true, same reason as the numeric pair above.
		{"'b' >= 'b'", true},
		{"'a' <= 'a'", true},
		{"'abc' > 'abb'", true},
		// Bools: equality only.
		{"true == true", true},
		{"true != false", true},
		// nil equality, including the "is this cell empty" idiom.
		{"nil == nil", true},
		{"nil != nil", false},
		{"1 == nil", false},
		{"'a' != nil", true},
		{"nil == 'a'", false},
	}
	for _, tc := range ok {
		res, err := Eval(tc.expr, nil)
		if err != nil {
			t.Errorf("Eval(%q) failed: %v", tc.expr, err)
			continue
		}
		if res != tc.want {
			t.Errorf("Eval(%q) = %v (%T), want %v (%T)", tc.expr, res, res, tc.want, tc.want)
		}
	}

	// Comparisons that must stay ERRORS rather than silently answering.
	bad := []string{
		"1 < 'a'",     // number vs string ordering
		"'a' > true",  // string vs bool
		"true < true", // bools are not ordered
		"1 == true",   // number vs bool
		"nil < 1",     // nil is not ordered
		"nil > nil",
	}
	for _, expr := range bad {
		res, err := Eval(expr, nil)
		if err == nil {
			t.Errorf("Eval(%q) = %v, want an unsupported-operation error", expr, res)
			continue
		}
		if !strings.Contains(err.Error(), "unsupported binary operation:") {
			t.Errorf("Eval(%q) error = %v; want the unsupported-binary-operation message", expr, err)
		}
	}
}

// TestEval_LogicalShortCircuit is the load-bearing test for && / || / and / or.
//
// Without it the operators can be "implemented" in evalBinary and every
// value-level assertion still passes — eval evaluates BOTH sides of a
// BinaryNode before calling evalBinary, so the un-taken branch would still
// issue its COM round trips and could surface an error that short-circuit
// semantics avoid entirely. The recorder is what distinguishes the two
// implementations.
func TestEval_LogicalShortCircuit(t *testing.T) {
	calls := 0
	env := map[string]interface{}{
		"t": true,
		"f": false,
		"boom": func(...interface{}) (interface{}, error) {
			calls++
			return nil, fmt.Errorf("right-hand side must not be evaluated")
		},
		"yes": func(...interface{}) (interface{}, error) {
			calls++
			return true, nil
		},
	}

	shorted := []struct {
		expr string
		want interface{}
	}{
		{"f && boom()", false},
		{"false and boom()", false},
		{"t || boom()", true},
		{"true or boom()", true},
	}
	for _, tc := range shorted {
		calls = 0
		res, err := Eval(tc.expr, env)
		if err != nil {
			t.Errorf("Eval(%q) failed: %v", tc.expr, err)
			continue
		}
		if res != tc.want {
			t.Errorf("Eval(%q) = %v, want %v", tc.expr, res, tc.want)
		}
		if calls != 0 {
			t.Errorf("Eval(%q) evaluated the right-hand side %d time(s); it must short-circuit", tc.expr, calls)
		}
	}

	// The non-shorted direction must still evaluate the right-hand side, so
	// "short-circuit" cannot be satisfied by never evaluating it.
	for _, tc := range []struct {
		expr string
		want interface{}
	}{
		{"t && yes()", true},
		{"f || yes()", true},
		{"t and yes()", true},
	} {
		calls = 0
		res, err := Eval(tc.expr, env)
		if err != nil {
			t.Errorf("Eval(%q) failed: %v", tc.expr, err)
			continue
		}
		if res != tc.want {
			t.Errorf("Eval(%q) = %v, want %v", tc.expr, res, tc.want)
		}
		if calls != 1 {
			t.Errorf("Eval(%q) called the right-hand side %d time(s), want 1", tc.expr, calls)
		}
	}

	// A non-bool operand is an error, not a truthiness guess.
	for _, expr := range []string{"1 && true", "true && 1", "'a' || true"} {
		if res, err := Eval(expr, env); err == nil {
			t.Errorf("Eval(%q) = %v, want an error (no truthiness coercion)", expr, res)
		}
	}
}

// TestEval_ComparisonOnChainValues drives the comparison operators over a real
// COM value (Scripting.Dictionary.Count), so the Chain-unwrap preamble is
// exercised rather than just Go literals.
func TestEval_ComparisonOnChainValues(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx) // one entry, so Count == 1
		env := map[string]interface{}{"d": dict}

		cases := []struct {
			expr string
			want interface{}
		}{
			{"d.Count == 1", true},
			{"d.Count > 0", true},
			{"d.Count >= 2", false},
			// Boundary-true through a real COM property read, not just a literal.
			{"d.Count >= 1", true},
			{"d.Count <= 1", true},
			{"d.Count != 1", false},
			{"d.Count > 0 && d.Count < 5", true},
			{"d.Count > 9 || d.Count == 1", true},
		}
		for _, tc := range cases {
			res, err := Eval(tc.expr, env)
			if err != nil {
				t.Errorf("Eval(%q) failed: %v", tc.expr, err)
				continue
			}
			if res != tc.want {
				t.Errorf("Eval(%q) = %v (%T), want %v", tc.expr, res, res, tc.want)
			}
		}

		// An object-valued operand is NOT comparable: Chain.Value() on a
		// VT_DISPATCH result is an error ("use Store"), and object identity
		// comparison is deliberately out of scope. It must surface as an error,
		// never as a silent false.
		if res, err := Eval("d == nil", env); err == nil {
			t.Errorf("Eval(\"d == nil\") = %v, want an error (object operands are not comparable)", res)
		}
		return nil
	})
}

// TestEnvTypes_RunAndPutAgree is the deliverable of the visitorFor extraction:
// Run and Put must classify every env shape the same way. Before the extraction
// each had its own copy of the four-case type switch (with a byte-identical
// default error), so adding a fifth env type to one and forgetting the other
// would leave the forgotten side answering "unsupported env type" — which reads
// like a user error rather than a missing case.
//
// The assertion is specifically about ENV-TYPE classification, not about the
// expression succeeding: Put("Count") on a nil env legitimately fails ("env has
// no root COM object") while Run fails with "identifier not found". Both accept
// the env type, which is what this pins.
func TestEnvTypes_RunAndPutAgree(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := setupDict(t, ctx)
		disp, err := dict.Store()
		if err != nil {
			t.Fatalf("Store: %v", err)
		}
		defer disp.Release()

		const envTypeErr = "unsupported env type"
		cases := []struct {
			name      string
			env       interface{}
			supported bool
		}{
			{"sugar.Chain", dict, true},
			{"*ole.IDispatch", disp, true},
			{"map[string]interface{}", map[string]interface{}{"Count": 1}, true},
			{"nil", nil, true},
			{"int", 1, false},
			{"[]string", []string{"x"}, false},
		}
		for _, tc := range cases {
			_, runErr := Eval("Count", tc.env)
			putErr := Put(tc.env, "Count", 1)

			runRejected := runErr != nil && strings.Contains(runErr.Error(), envTypeErr)
			putRejected := putErr != nil && strings.Contains(putErr.Error(), envTypeErr)

			if runRejected != putRejected {
				t.Errorf("%s: Run rejected=%v (%v) but Put rejected=%v (%v) — the two env "+
					"classifications have drifted apart", tc.name, runRejected, runErr, putRejected, putErr)
			}
			if runRejected == tc.supported {
				t.Errorf("%s: env-type rejected=%v, want supported=%v (Run err: %v)",
					tc.name, runRejected, tc.supported, runErr)
			}
		}
		return nil
	})
}

// TestPut_RawIDispatchArenaNoLeak pins the refcount on Put's raw-IDispatch env
// path (2026-08-03).
//
// The gap it closes: the env-type switch shared by Run and Put was factored into
// visitorFor, which returns the arena teardown as a CLOSURE. Only the
// *ole.IDispatch arm returns a real one -- every other arm returns noCleanup --
// so replacing that arm's `func() { ctx.Release() }` with noCleanup is a genuine
// COM leak on every Put(rawIDispatch, ...). Verified: that mutation left the
// ENTIRE suite green. TestPut_RootIdentifier drives this path but asserts only
// behavior, and behavior is unaffected by a leak; TestIDispatchEnv_ArenaNoLeak
// has the right shape but covers Get/Run, not Put.
//
// A leak is invisible to every functional assertion by construction, so the only
// thing that can catch it is counting. Same base-vs-after AddRef pattern as
// TestIDispatchEnv_ArenaNoLeak: Scripting.Dictionary is in-process, so the count
// AddRef returns is deterministic.
func TestPut_RawIDispatchArenaNoLeak(t *testing.T) {
	sugar.Do(func(ctx sugar.Context) error {
		dict := ctx.Create("Scripting.Dictionary")
		if err := dict.Err(); err != nil {
			t.Fatalf("create failed: %v", err)
		}
		raw, err := dict.Store()
		if err != nil {
			t.Fatalf("Store failed: %v", err)
		}
		defer raw.Release()

		base := raw.AddRef()
		raw.Release()

		// CompareMode is settable only while the dictionary is empty, so repeat
		// with a value that stays legal: writing the same value is still a full
		// trip through visitorFor's IDispatch arm.
		for i := 0; i < 3; i++ {
			if err := Put(raw, "CompareMode", 1); err != nil {
				t.Fatalf("Put(raw, CompareMode) failed on iteration %d: %v", i, err)
			}
		}

		after := raw.AddRef()
		raw.Release()
		if after != base {
			t.Errorf("refcount leak on Put's raw-IDispatch env path: %d -> %d after 3 Puts. "+
				"visitorFor's *ole.IDispatch arm must return a cleanup that releases the arena; "+
				"a noCleanup there leaks one reference per call and no functional assertion can see it",
				base, after)
		}
		return nil
	})
}
