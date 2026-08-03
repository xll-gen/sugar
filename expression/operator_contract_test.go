//go:build windows

// Executable form of the operator contract documented in expression.go's package
// doc and AGENTS.md §4.
//
// The point of this file is that the "supported" and "deliberately absent"
// operator lists stop being prose. A doc-only claim about which operators error
// is indistinguishable from no claim at all: nothing in the build breaks when
// the engine gains an operator the doc says it does not have, or loses one the
// doc says it has.

package expression

import (
	"go/parser"
	"go/token"
	"regexp"
	"strings"
	"testing"

	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// supportedOps / unsupportedOps are the SINGLE inventory of operators this
// package takes a position on. Three tests read them: the two execution tests
// below, and TestOperatorContract_PackageDocNamesEveryOperator, which pins that
// the package doc still names each one. `op` is the token as it must appear in
// the doc's operator listing.
var supportedOps = []struct {
	op   string
	expr string
	want interface{}
}{
	// Arithmetic: the headline "only + - * /".
	{"+", "7 + 2", float64(9)},
	{"-", "7 - 2", float64(5)},
	{"*", "7 * 2", float64(14)},
	{"/", "7 / 2", float64(3.5)},
	// Comparison.
	{"==", "7 == 7", true},
	{"!=", "7 != 2", true},
	{"<", "7 < 2", false},
	{"<=", "7 <= 7", true},
	{">", "7 > 2", true},
	{">=", "7 >= 8", false},
	// Logical, both spellings.
	{"&&", "true && false", false},
	{"||", "true || false", true},
	{"and", "true and false", false},
	{"or", "true or false", true},
	// Unary.
	{"-", "-7", float64(-7)},
	{"+", "+7", float64(7)},
	{"!", "!true", false},
	{"not", "not true", false},
}

// TestOperatorContract_SupportedOnesEvaluate is the control half. Without it the
// unsupported-operator table below could be satisfied by an engine that rejects
// EVERYTHING, which would be a passing test suite and a dead package.
func TestOperatorContract_SupportedOnesEvaluate(t *testing.T) {
	cases := supportedOps
	for _, tc := range cases {
		got, err := Eval(tc.expr, nil)
		if err != nil {
			t.Errorf("Eval(%q) failed: %v — this operator is documented as SUPPORTED", tc.expr, err)
			continue
		}
		if got != tc.want {
			t.Errorf("Eval(%q) = %v (%T), want %v (%T)", tc.expr, got, got, tc.want, tc.want)
		}
	}
}

// TestOperatorContract_UnsupportedOnesError enumerates every operator the expr
// grammar PARSES but this engine deliberately does not evaluate. Each one must
// be an error, never a silent value.
//
// wantBinaryMsg marks the expressions that must reach evalBinary's shared
// fallthrough, i.e. the operator is unknown to the engine rather than the node
// type being unsupported. Keeping them separated stops a future change that
// turns "%" into a parse error from looking like a pass.
var unsupportedOps = []struct {
	op            string
	expr          string
	wantBinaryMsg bool
}{
	// Arithmetic operators expr parses and this engine does not implement.
	{"%", "7 % 2", true},
	{"**", "7 ** 2", true},
	{"^", "7 ^ 2", true},
	// String / collection operators.
	{"matches", "'abc' matches 'a'", true},
	{"contains", "'abc' contains 'a'", true},
	{"startsWith", "'abc' startsWith 'a'", true},
	{"endsWith", "'abc' endsWith 'c'", true},
	{"in", "1 in [1, 2]", false}, // the ArrayNode operand is rejected first
	{"..", "1 .. 3", true},
	// Nil coalescing and the ternary (a distinct ConditionalNode).
	{"??", "nil ?? 1", true},
	{"?:", "true ? 1 : 2", false},
}

// The three arithmetic entries (%, **, ^) are the reason this test exists: the
// contract's "deliberately absent" list named only the string/collection/
// conditional operators, so the three that look most like the supported four
// were the ones nobody had written down. They are also the ones a reader is most
// likely to assume work, because the headline says "arithmetic is + - * /" and a
// modulo or a power reads like arithmetic.
func TestOperatorContract_UnsupportedOnesError(t *testing.T) {
	for _, tc := range unsupportedOps {
		got, err := Eval(tc.expr, nil)
		if err == nil {
			t.Errorf("Eval(%q) = %v (%T), want an error — this operator is documented as NOT supported", tc.expr, got, got)
			continue
		}
		if !tc.wantBinaryMsg {
			continue
		}
		if !strings.Contains(err.Error(), "unsupported binary operation:") {
			t.Errorf("Eval(%q) error = %v; want the shared unsupported-binary-operation message", tc.expr, err)
		}
	}
}

// --- the doc pin ---------------------------------------------------------------

var reWordOp = regexp.MustCompile(`^[\pL]+$`)

// packageDocInventoryLines returns the INDENTED lines of expression.go's package
// doc — the bullet lists and the indented literal block, i.e. exactly the places
// the operator inventory is written down. Flush-left prose is excluded so that a
// stray English "or"/"in"/"not" in a paragraph cannot satisfy a word operator.
func packageDocInventoryLines(t *testing.T) []string {
	t.Helper()
	fset := token.NewFileSet()
	f, err := parser.ParseFile(fset, "expression.go", nil, parser.ParseComments)
	if err != nil {
		t.Fatalf("parse expression.go: %v", err)
	}
	if f.Doc == nil {
		t.Fatal("expression.go has no package doc comment at all")
	}
	var out []string
	for _, line := range strings.Split(f.Doc.Text(), "\n") {
		if line == "" || (line[0] != ' ' && line[0] != '\t') {
			continue
		}
		out = append(out, line)
	}
	if len(out) == 0 {
		t.Fatal("expression.go's package doc has no indented operator listing at all")
	}
	return out
}

// TestOperatorContract_PackageDocNamesEveryOperator closes the last prose gap in
// the operator contract.
//
// The execution tables above pin what the ENGINE does. Nothing pinned what the
// package DOC says, and that is the half that actually failed once: the doc read
// "Binary operators: + - * / only (no comparison or logical operators)" for the
// whole life of the comparison and logical support, and no test anywhere went
// red. Reverting the doc to that sentence today fails THIS test and nothing else.
//
// Strength, stated honestly: the symbol operators (== != <= >= && || % ** ^ ..
// ?? ?: ...) are pinned tightly, because those character sequences occur nowhere
// else. The four WORD operators (and, or, not, in) are pinned only to "appears
// as a standalone word on an inventory line" — they are ordinary English words,
// so a doc line that happened to use one in prose inside a bullet would satisfy
// them. That is the residual weakness; it is not worth a grammar to close.
func TestOperatorContract_PackageDocNamesEveryOperator(t *testing.T) {
	lines := packageDocInventoryLines(t)

	named := func(op string) bool {
		if reWordOp.MatchString(op) {
			re := regexp.MustCompile(`(^|[^\pL])` + regexp.QuoteMeta(op) + `([^\pL]|$)`)
			for _, l := range lines {
				if re.MatchString(l) {
					return true
				}
			}
			return false
		}
		for _, l := range lines {
			if strings.Contains(l, op) {
				return true
			}
		}
		return false
	}

	for _, tc := range supportedOps {
		if !named(tc.op) {
			t.Errorf("operator %q evaluates (%q) but is not named in expression.go's package doc — "+
				"a supported operator nobody documented is a feature only the tests know about", tc.op, tc.expr)
		}
	}
	for _, tc := range unsupportedOps {
		if !named(tc.op) {
			t.Errorf("operator %q is deliberately REFUSED (%q) but is not named in expression.go's "+
				"package doc — an undocumented refusal reads as a bug to whoever hits it", tc.op, tc.expr)
		}
	}
}

// --- chained comparison ---------------------------------------------------------

// countChain counts the COM-shaped operations an evaluation issues. Get and Call
// each return a FRESH chain so a re-walk of the same AST subtree cannot be
// mistaken for a cached one.
type countChain struct {
	gets  *int
	calls *int
	value interface{}
}

func (c *countChain) Get(string, ...interface{}) sugar.Chain {
	*c.gets++
	return &countChain{gets: c.gets, calls: c.calls, value: c.value}
}
func (c *countChain) Call(string, ...interface{}) sugar.Chain {
	*c.calls++
	return &countChain{gets: c.gets, calls: c.calls, value: c.value}
}
func (c *countChain) Put(string, ...interface{}) sugar.Chain      { return c }
func (c *countChain) ForEach(func(sugar.Chain) error) sugar.Chain { return c }
func (c *countChain) Fork() sugar.Chain                           { return c }
func (c *countChain) Store() (*ole.IDispatch, error)              { return nil, errNilDispatch{} }
func (c *countChain) Release() error                              { return nil }
func (c *countChain) IsDispatch() bool                            { return false }
func (c *countChain) Value() (interface{}, error)                 { return c.value, nil }
func (c *countChain) Err() error                                  { return nil }

// TestComparison_ChainedComparisonWalksTheMiddleOperandTwice pins the one
// comparison behaviour that costs REAL COM round trips and is invisible in the
// source expression.
//
// expr folds `a < b < c` into the conjunction `a < b && b < c` — b appears in
// BOTH conjuncts — and this evaluator walks each conjunct independently. So the
// middle operand is evaluated twice: `1 < obj.P < 9` issues two property reads,
// and `1 < obj.f() < 9` CALLS f twice (side effects included). Measured, not
// inferred.
//
// It sits next to the short-circuit rule, which exists precisely to avoid COM
// round trips the reader did not ask for; chained comparison quietly adds one.
// If someone ever makes the evaluator hoist the shared operand, this test fails
// and the doc must be rewritten with it.
func TestComparison_ChainedComparisonWalksTheMiddleOperandTwice(t *testing.T) {
	cases := []struct {
		expr      string
		wantGets  int
		wantCalls int
	}{
		{"1 < x.P", 1, 0},             // control: one conjunct, one round trip
		{"1 < x.P < 9", 2, 0},         // the fold: the middle operand is re-walked
		{"1 < x.f() < 9", 0, 2},       // …and a METHOD operand runs twice
		{"x.P > 1 and x.P < 9", 2, 0}, // what the user would have written by hand
	}
	for _, tc := range cases {
		gets, calls := 0, 0
		root := &countChain{gets: &gets, calls: &calls, value: float64(5)}
		got, err := Eval(tc.expr, map[string]interface{}{"x": root})
		if err != nil {
			t.Errorf("Eval(%q) failed: %v", tc.expr, err)
			continue
		}
		if got != true {
			t.Errorf("Eval(%q) = %v, want true", tc.expr, got)
		}
		if gets != tc.wantGets || calls != tc.wantCalls {
			t.Errorf("Eval(%q): %d property reads / %d method calls, want %d / %d",
				tc.expr, gets, calls, tc.wantGets, tc.wantCalls)
		}
	}
}

// --- the nil-comparison hole -------------------------------------------------

// stubChain is a sugar.Chain whose observable shape is fixed by the test. It
// exists to reproduce, from OUTSIDE package sugar, the exact triple a chain
// referencing a non-IDispatch COM object presents:
//
//	IsDispatch() == false, Store() == error, Value() == (nil, nil)
//
// What this fake CANNOT show is that a real sugar chain ever has that shape —
// the `chain` struct is unexported and handleResult never builds one from a COM
// call. That half is pinned inside package sugar by
// TestVTUnknownChain_IsIndistinguishableFromNothing (chain_unknown_test.go).
// Read the two together; either alone proves nothing.
type stubChain struct {
	isDispatch bool
	value      interface{}
}

func (s *stubChain) Get(string, ...interface{}) sugar.Chain  { return s }
func (s *stubChain) Call(string, ...interface{}) sugar.Chain { return s }
func (s *stubChain) Put(string, ...interface{}) sugar.Chain  { return s }
func (s *stubChain) ForEach(func(sugar.Chain) error) sugar.Chain {
	return s
}
func (s *stubChain) Fork() sugar.Chain { return s }
func (s *stubChain) Store() (*ole.IDispatch, error) {
	// Mirrors chain.Store: it reads `disp`, which a VT_UNKNOWN / degraded chain
	// does not have. So Store is NOT an escape hatch the engine could use to
	// tell such a chain apart from COM Nothing.
	if !s.isDispatch {
		return nil, errNilDispatch{}
	}
	return nil, nil
}
func (s *stubChain) Release() error              { return nil }
func (s *stubChain) IsDispatch() bool            { return s.isDispatch }
func (s *stubChain) Value() (interface{}, error) { return s.value, nil }
func (s *stubChain) Err() error                  { return nil }

type errNilDispatch struct{}

func (errNilDispatch) Error() string { return "nil dispatch" }

// TestComparison_NonDispatchObjectChainComparesEqualToNil pins the documented
// HOLE in the comparison contract, on purpose and as a hole.
//
// The v0.8.13 object-operand guard tests IsDispatch(), so it catches a live
// IDispatch and refuses to compare it. It does not catch — and cannot catch — a
// chain that references a COM object the engine cannot reach as an IDispatch: a
// raw VT_UNKNOWN result, and the empty chain handleResult degrades a
// non-IDispatch-capable IUnknown to. All three of IsDispatch/Store/Value answer
// exactly as they do for COM `Nothing`, and `Nothing == nil` must stay TRUE —
// it is the "is this object absent" idiom the whole nil arm exists for.
//
// So `x == nil` is TRUE for those chains. That is the contract, not an accident,
// and this test is what stops the contract from drifting away from the doc: if
// anyone ever implements the "reject" alternative, this test fails and the doc
// has to be rewritten with it.
func TestComparison_NonDispatchObjectChainComparesEqualToNil(t *testing.T) {
	// The hole: IsDispatch() false + Value() (nil, nil) == nil is TRUE.
	nonDispatch := &stubChain{isDispatch: false, value: nil}
	env := map[string]interface{}{"x": nonDispatch}

	got, err := Eval("x == nil", env)
	if err != nil {
		t.Fatalf(`Eval("x == nil") failed: %v — the documented contract is that this ANSWERS, not errors`, err)
	}
	if got != true {
		t.Errorf(`Eval("x == nil") = %v (%T), want true (the documented hole)`, got, got)
	}
	got, err = Eval("x != nil", env)
	if err != nil {
		t.Fatalf(`Eval("x != nil") failed: %v`, err)
	}
	if got != false {
		t.Errorf(`Eval("x != nil") = %v (%T), want false`, got, got)
	}

	// The control: a chain that DOES report IsDispatch() is refused. Without
	// this row the test above would also pass against a guard that never fires,
	// i.e. against a removed guard.
	dispatch := &stubChain{isDispatch: true, value: nil}
	res, err := Eval("y == nil", map[string]interface{}{"y": dispatch})
	if err == nil {
		t.Fatalf(`Eval("y == nil") on a dispatch chain = %v, want the object-operand error`, res)
	}
	if !strings.Contains(err.Error(), "object operands are not comparable") {
		t.Errorf("dispatch-operand error = %v; want the object-operand refusal", err)
	}
}
