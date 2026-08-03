//go:build windows

// Package expression evaluates string expressions against COM automation
// objects, navigating properties and calling methods through sugar.Chain:
//
//	name, err := expression.Get(excelChain, "Workbooks.Add().ActiveSheet.Name")
//
// # Supported grammar
//
// Expressions are parsed with the expr-lang parser, but only a subset of the
// language is evaluated:
//
//   - Property access: Workbooks.Count
//   - Method calls: Workbooks.Add(), d.Add('k', 'v')
//   - Index access: Sheets[1], d['Key']. Numeric indexes are sent to the
//     conventional COM default member, so Sheets[1] becomes Sheets.Item(1);
//     string indexes are equivalent to dotted access (d['Key'] == d.Key).
//   - Literals: integers, floats, strings, booleans, nil
//   - Unary operators: -x, +x, !x / not x
//   - Arithmetic: + - * / only. Numeric operands are coerced to float64, so
//     "2 + 2" yields float64(4) and division by zero follows float64
//     semantics (+Inf). "+" with a string operand concatenates.
//   - Comparison: == != < <= > >= (see the contract below).
//   - Logical: && || and or. These SHORT-CIRCUIT — the right-hand side is not
//     evaluated, and therefore issues no COM round trips, when the left side
//     already decides the result. Operands must be real booleans; there is no
//     truthiness coercion, because "0 is false" over COM values (VT_BOOL vs
//     0/-1 vs "") has no single right answer.
//
// Every other operator the expr grammar parses is an ERROR, never a silent
// value. The full list, each pinned by TestOperatorContract_UnsupportedOnesError:
//
//	%   **   ..   ??   ?:
//	^   in   matches   contains   startsWith   endsWith
//
// Note the three arithmetic ones. "Arithmetic is + - * /" is literally true,
// but modulo and the two exponentiation spellings read like arithmetic and are
// the ones most likely to be assumed present.
//
// # Comparison contract
//
// The type rules are deliberately narrow, because every permissive alternative
// is silently wrong on spreadsheet data:
//
//   - numbers compare as float64, via the same coercion arithmetic uses, so
//     int-vs-float mixes work — and 0.1+0.2 == 0.3 is false, and a uint64
//     above 2^53 loses precision;
//   - strings compare lexicographically;
//   - bools and nil support EQUALITY ONLY;
//   - "x == nil" is the "is this empty / absent" idiom;
//   - a VT_NULL operand (sugar.Null — "the cells disagree", or SQL NULL) is
//     REFUSED for every comparison including ==, because VBA and SQL both make
//     a NULL comparison itself NULL, i.e. neither true nor false. It is NOT
//     nil: nil means VT_EMPTY, a genuinely empty cell. Before the sentinel
//     existed a Null decoded to nil and "x == nil" answered TRUE for it, which
//     is the wrong answer to the question the idiom asks. Pinned by
//     TestComparison_NullOperandIsRefused;
//   - every cross-kind pairing (number vs string, string vs bool, ...) is an
//     error, not a fmt.Sprint comparison.
//
// CHAINED COMPARISON COSTS AN EXTRA ROUND TRIP. The parser folds "a < b < c"
// into the conjunction "a < b && b < c", so the middle operand appears in both
// conjuncts and this evaluator walks each one independently:
//
//   - "1 < obj.Prop < 9" issues TWO property reads of obj.Prop, not one;
//   - "1 < obj.f() < 9" CALLS f TWICE, side effects included.
//
// Nothing caches between the conjuncts. This is the exact cost the logical
// operators short-circuit to avoid, arriving through a spelling that does not
// look like it has two operands at all. Read the property into the env, or write
// the conjunction yourself, when the extra call matters. Measured and pinned by
// TestComparison_ChainedComparisonWalksTheMiddleOperandTwice.
//
// A COM OBJECT operand is refused outright: a chain holding a dispatch but no
// VARIANT result answers (nil, nil) from Value(), so "someObject == nil" would
// otherwise report TRUE for a live object. Compare a property of it instead.
// Object identity comparison is out of scope.
//
// KNOWN HOLE, deliberate and not fixable here: that refusal tests IsDispatch(),
// which is true only for a reachable IDispatch. A chain that references a COM
// object the engine cannot reach as an IDispatch — a bare VT_UNKNOWN result, or
// the empty chain sugar degrades a non-IDispatch-capable IUnknown to — is NOT
// refused, and "x == nil" answers TRUE for it. There is no observable that
// separates those from COM Nothing: IsDispatch() is false, Store() reports "nil
// dispatch" and Value() is (nil, nil) for all three (pinned by sugar's
// TestVTUnknownChain_IsIndistinguishableFromNothing). Refusing them would mean
// refusing Nothing too, and "Nothing == nil" must stay TRUE — that is what the
// nil arm is for. Note also that sugar itself never hands this package a
// VT_UNKNOWN chain: handleResult promotes a QI-able IUnknown to a dispatch
// chain and degrades the rest, so the state is reachable only from a
// hand-built VARIANT or a third-party Chain implementation.
//
// # Argumented properties are unreachable
//
// A call node is issued as DISPATCH_METHOD, so Excel members that take
// arguments and are PROPERTIES — Range("A1"), Cells(1, 1), Offset, Resize,
// End — answer DISP_E_MEMBERNOTFOUND through this package. Reach them with the
// chain (sheet.Get("Range", "A1").Put("Value", v)) or the typed excel package.
// A Call-then-Get fallback would silently change which COM verb a caller's
// expression uses; it needs its own design pass, not a drive-by.
//
// # Environments and ownership
//
// The env argument of Run/Eval/Get/Store/Put may be:
//
//   - sugar.Chain — the root object for identifier lookups. Pass a chain
//     tracked by a sugar.Context: intermediate COM objects created during
//     evaluation inherit the chain's context and are released with it. An
//     untracked chain (package-level sugar.From) leaks every intermediate.
//   - *ole.IDispatch — evaluation runs inside an internal arena: every
//     intermediate COM object is released before returning. Scalar results
//     come back as plain Go values. If the result is a COM object, Run/Eval
//     return a *ole.IDispatch that the caller owns and must Release; Store
//     returns it under the same contract; Get rejects it with an error.
//   - map[string]interface{} — named values: sugar.Chain entries (COM
//     objects), plain Go values, and callables of the exact type
//     func(...interface{}) (interface{}, error).
//   - nil — only literal/operator expressions can be evaluated.
//
// Any other env type is an error.
package expression

import (
	"fmt"
	"reflect"

	"github.com/expr-lang/expr/ast"
	"github.com/expr-lang/expr/parser"
	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// Program represents a compiled expression.
type Program struct {
	node ast.Node
}

// Compile parses an expression.
func Compile(expression string) (*Program, error) {
	tree, err := parser.Parse(expression)
	if err != nil {
		return nil, err
	}
	return &Program{node: tree.Node}, nil
}

// Run executes a compiled Program against an environment. See the package
// documentation for the supported env types and the ownership contract of
// the *ole.IDispatch env path.
func (p *Program) Run(env interface{}) (interface{}, error) {
	// The *ole.IDispatch arm stays SEPARATE and does not use visitorFor: its result
	// may be a COM object that must outlive the arena, so the Store/Value choice
	// has to sit textually next to the release (see runInArena). Every other arm
	// shares the one env-type switch.
	if d, ok := env.(*ole.IDispatch); ok {
		return runInArena(d, p.node)
	}
	v, cleanup, err := visitorFor(env)
	if err != nil {
		return nil, err
	}
	defer cleanup()
	return finishEval(v, p.node)
}

// visitorFor builds the comVisitor for an env value, together with the cleanup
// that must run when evaluation is done. It is the single definition of "which
// env types are supported"; Run and Put both went through their own copy of this
// switch, so a new env type had to be added twice and the two could disagree.
//
// OWNERSHIP WARNING: a caller that lets a RESULT escape the arena must
// materialize it BEFORE cleanup runs. Run does that in runInArena -- which
// deliberately does NOT go through this function, so its Store/Value ordering
// stays textually intact next to the release. Put has no escaping result: its
// arena exists only for intermediates, so `defer cleanup()` is correct there.
//
// Only the *ole.IDispatch arm returns a cleanup that does anything. That asymmetry
// is the whole risk surface: replacing that one closure with noCleanup is a silent
// COM refcount leak that no behavioural assertion can see, which is why
// TestPut_RawIDispatchArenaNoLeak counts references rather than checking results.
func visitorFor(env interface{}) (*comVisitor, func(), error) {
	noCleanup := func() {}
	switch v := env.(type) {
	case sugar.Chain:
		return &comVisitor{initialChain: v}, noCleanup, nil
	case *ole.IDispatch:
		ctx := sugar.NewContext(nil)
		return &comVisitor{initialChain: ctx.From(v)}, func() { ctx.Release() }, nil
	case map[string]interface{}:
		return &comVisitor{envMap: v}, noCleanup, nil
	case nil:
		return &comVisitor{}, noCleanup, nil
	default:
		return nil, nil, fmt.Errorf("expression: unsupported env type %T", env)
	}
}

// finishEval evaluates node and unwraps a deferred chain error so callers of
// Run/Eval never receive a Chain whose Err() is already set.
func finishEval(v *comVisitor, node ast.Node) (interface{}, error) {
	res, err := v.eval(node)
	if err != nil {
		return nil, err
	}
	if ch, ok := res.(sugar.Chain); ok {
		if err := ch.Err(); err != nil {
			return nil, err
		}
	}
	return res, nil
}

// runInArena evaluates node against a raw, untracked IDispatch. All chains
// created during evaluation (including the root sugar.From) live in an
// internal arena that is released before returning, so nothing leaks. The
// final result escapes the arena either as a plain Go value or — for COM
// object results — as an AddRef'd *ole.IDispatch the caller must Release.
func runInArena(disp *ole.IDispatch, node ast.Node) (interface{}, error) {
	ctx := sugar.NewContext(nil)
	defer ctx.Release()

	visitor := &comVisitor{initialChain: ctx.From(disp)}
	res, err := visitor.eval(node)
	if err != nil {
		return nil, err
	}
	ch, ok := res.(sugar.Chain)
	if !ok {
		return res, nil
	}
	if err := ch.Err(); err != nil {
		return nil, err
	}
	if d, err := ch.Store(); err == nil {
		// COM object result: Store() AddRefs, so this reference survives
		// the arena release below. The caller owns it.
		return d, nil
	}
	// Value result: materialize the Go value (strings/arrays are copied)
	// before the arena clears the underlying VARIANT.
	return ch.Value()
}

// Eval parses and executes an expression.
func Eval(expression string, env interface{}) (interface{}, error) {
	p, err := Compile(expression)
	if err != nil {
		return nil, err
	}
	return p.Run(env)
}

// Get retrieves a property or calls a method using an expression and returns
// the result as a plain Go value. COM object results are an error — use
// Store for those.
func Get(obj interface{}, expression string) (interface{}, error) {
	result, err := Eval(expression, obj)
	if err != nil {
		return nil, err
	}

	switch r := result.(type) {
	case sugar.Chain:
		if err := r.Err(); err != nil {
			return nil, err
		}
		return r.Value()
	case *ole.IDispatch:
		// Arena path (*ole.IDispatch env): the object escaped as a
		// caller-owned reference, but Get promises a plain value.
		r.Release()
		return nil, fmt.Errorf("expression: result is a COM object, use Store")
	default:
		return result, nil
	}
}

// Store retrieves a COM object (IDispatch) using an expression. The returned
// dispatch is AddRef'd; the caller is responsible for Release unless it is
// handed to a sugar.Context.
func Store(obj interface{}, expression string) (*ole.IDispatch, error) {
	result, err := Eval(expression, obj)
	if err != nil {
		return nil, err
	}

	switch r := result.(type) {
	case sugar.Chain:
		if err := r.Err(); err != nil {
			return nil, err
		}
		return r.Store()
	case *ole.IDispatch:
		return r, nil
	default:
		return nil, fmt.Errorf("expression did not evaluate to a COM object")
	}
}

// Put sets a property using an expression. The expression must be a property
// access: either a member expression ("ActiveSheet.Name", "d[1]") or — when
// env carries a root object — a bare identifier ("Visible").
func Put(obj interface{}, expression string, value interface{}) error {
	p, err := Compile(expression)
	if err != nil {
		return err
	}

	v, cleanup, err := visitorFor(obj)
	if err != nil {
		return err
	}
	// Put has no escaping result, so releasing the arena on return is safe.
	defer cleanup()
	return putNode(v, p.node, value)
}

func putNode(v *comVisitor, node ast.Node, value interface{}) error {
	switch n := node.(type) {
	case *ast.IdentifierNode:
		// Root-level property, e.g. Put(app, "Visible", true).
		if v.initialChain == nil {
			return fmt.Errorf("expression: cannot Put %q: env has no root COM object", n.Value)
		}
		return v.initialChain.Put(n.Value, value).Err()

	case *ast.MemberNode:
		parentObj, err := v.eval(n.Node)
		if err != nil {
			return err
		}
		parentChain, ok := parentObj.(sugar.Chain)
		if !ok {
			return fmt.Errorf("parent is not COM object: %T", parentObj)
		}
		name, idxArgs, err := v.resolveProperty(n.Property)
		if err != nil {
			return err
		}
		return parentChain.Put(name, append(idxArgs, value)...).Err()

	default:
		return fmt.Errorf("invalid Put expression: must be property access, got %T", node)
	}
}

type comVisitor struct {
	initialChain sugar.Chain
	envMap       map[string]interface{}
}

func (v *comVisitor) eval(node ast.Node) (interface{}, error) {
	switch n := node.(type) {
	case *ast.IdentifierNode:
		if v.envMap != nil {
			if val, ok := v.envMap[n.Value]; ok {
				return val, nil
			}
		}
		if v.initialChain != nil {
			return v.initialChain.Get(n.Value), nil
		}
		return nil, fmt.Errorf("identifier not found: %s", n.Value)

	case *ast.MemberNode:
		left, err := v.eval(n.Node)
		if err != nil {
			return nil, err
		}
		chain, ok := left.(sugar.Chain)
		if !ok {
			return nil, fmt.Errorf("cannot access property on type %T", left)
		}
		name, idxArgs, err := v.resolveProperty(n.Property)
		if err != nil {
			return nil, err
		}
		return chain.Get(name, idxArgs...), nil

	case *ast.CallNode:
		args := make([]interface{}, len(n.Arguments))
		for i, argNode := range n.Arguments {
			argVal, err := v.eval(argNode)
			if err != nil {
				return nil, err
			}
			argChain, ok := argVal.(sugar.Chain)
			if !ok {
				args[i] = argVal
				continue
			}
			if err := argChain.Err(); err != nil {
				return nil, fmt.Errorf("arg %d error: %w", i, err)
			}
			if chainHoldsDispatch(argChain) {
				// Pass COM objects through unchanged; sugar's
				// normalizeParams (v0.8.0+) marshals Chain arguments
				// natively as IDispatch.
				args[i] = argChain
				continue
			}
			val, err := argChain.Value()
			if err != nil {
				return nil, fmt.Errorf("arg %d error: %w", i, err)
			}
			args[i] = val
		}

		switch callee := n.Callee.(type) {
		case *ast.MemberNode:
			obj, err := v.eval(callee.Node)
			if err != nil {
				return nil, err
			}
			chain, ok := obj.(sugar.Chain)
			if !ok {
				return nil, fmt.Errorf("cannot call method on type %T", obj)
			}
			name, idxArgs, err := v.resolveProperty(callee.Property)
			if err != nil {
				return nil, err
			}
			return chain.Call(name, append(idxArgs, args...)...), nil

		case *ast.IdentifierNode:
			if v.envMap != nil {
				if val, ok := v.envMap[callee.Value]; ok {
					fn, isFn := val.(func(...interface{}) (interface{}, error))
					if !isFn {
						// Never fall through to a COM call when the name is
						// bound in the env map — a typo'd binding must not
						// silently become a COM method invocation.
						return nil, fmt.Errorf("expression: env entry %q is not callable (have %T, want func(...interface{}) (interface{}, error))", callee.Value, val)
					}
					return fn(args...)
				}
			}
			if v.initialChain != nil {
				return v.initialChain.Call(callee.Value, args...), nil
			}
			return nil, fmt.Errorf("method not found: %s", callee.Value)

		default:
			return nil, fmt.Errorf("unsupported call on %T", callee)
		}

	case *ast.UnaryNode:
		val, err := v.eval(n.Node)
		if err != nil {
			return nil, err
		}
		return evalUnary(n.Operator, val)

	case *ast.BinaryNode:
		// && / || / and / or are handled HERE, before either side is evaluated,
		// because they must SHORT-CIRCUIT. evalBinary only ever sees two operands
		// that have already been evaluated, so implementing them there would give
		// the right values while still issuing the un-taken branch's COM round
		// trips -- and surfacing any error they raise. See evalLogical.
		if isLogicalOp(n.Operator) {
			return v.evalLogical(n.Operator, n.Left, n.Right)
		}
		left, err := v.eval(n.Left)
		if err != nil {
			return nil, err
		}
		right, err := v.eval(n.Right)
		if err != nil {
			return nil, err
		}
		return evalBinary(n.Operator, left, right)

	case *ast.IntegerNode:
		return n.Value, nil
	case *ast.StringNode:
		return n.Value, nil
	case *ast.BoolNode:
		return n.Value, nil
	case *ast.FloatNode:
		return n.Value, nil
	case *ast.NilNode:
		return nil, nil
	default:
		return nil, fmt.Errorf("unsupported node: %T", node)
	}
}

// resolveProperty turns a member/callee property expression into a COM member
// name plus optional index arguments. Dotted access (obj.Name) and string
// keys (obj['Name']) resolve to the name itself; numeric indexes resolve to
// the conventional COM default member, so Sheets[1] becomes Get("Item", 1).
// Shared by member access, method calls, and Put.
func (v *comVisitor) resolveProperty(prop ast.Node) (string, []interface{}, error) {
	switch p := prop.(type) {
	case *ast.StringNode:
		return p.Value, nil, nil
	case *ast.IdentifierNode:
		return p.Value, nil, nil
	}

	// Index expression (obj[expr]): evaluate it.
	idx, err := v.eval(prop)
	if err != nil {
		return "", nil, err
	}
	if c, ok := idx.(sugar.Chain); ok {
		idx, err = c.Value()
		if err != nil {
			return "", nil, err
		}
	}
	if s, ok := idx.(string); ok {
		return s, nil, nil
	}
	if idx != nil && isNumber(reflect.ValueOf(idx)) {
		return "Item", []interface{}{idx}, nil
	}
	return "", nil, fmt.Errorf("unsupported property expression %T", prop)
}

// chainHoldsDispatch reports whether ch currently wraps a live COM object.
// Store() AddRefs on success, so the probe reference is released immediately.
func chainHoldsDispatch(ch sugar.Chain) bool {
	d, err := ch.Store()
	if err != nil {
		return false
	}
	d.Release()
	return true
}

func evalUnary(op string, val interface{}) (interface{}, error) {
	if c, ok := val.(sugar.Chain); ok {
		var err error
		val, err = c.Value()
		if err != nil {
			return nil, err
		}
	}

	switch op {
	case "-":
		if rv := reflect.ValueOf(val); val != nil && isNumber(rv) {
			return -toFloat(rv), nil
		}
	case "+":
		if rv := reflect.ValueOf(val); val != nil && isNumber(rv) {
			return toFloat(rv), nil
		}
	case "!", "not":
		if b, ok := val.(bool); ok {
			return !b, nil
		}
	}

	return nil, fmt.Errorf("unsupported unary operation: %s %T", op, val)
}

func evalBinary(op string, left, right interface{}) (interface{}, error) {
	// An OBJECT operand is not comparable, and this has to be decided BEFORE the
	// Value() unwrap below. A dispatch chain's Value() answers (nil, nil) rather
	// than an error, so `obj == nil` would unwrap to nil and cheerfully report
	// TRUE -- i.e. "this live COM object is empty". Callers testing a cell for
	// emptiness would get a confident wrong answer.
	//
	// The guard covers exactly what IsDispatch() covers: a REACHABLE IDispatch.
	// A chain referencing a COM object the engine cannot reach as an IDispatch
	// (bare VT_UNKNOWN, or the empty chain sugar degrades a non-IDispatch-capable
	// IUnknown to) slips through and compares equal to nil. That is a documented
	// hole, not an oversight: those chains are observationally identical to COM
	// Nothing through the Chain interface (IsDispatch false / Store "nil
	// dispatch" / Value (nil, nil)), and `Nothing == nil` must stay TRUE. See the
	// package doc's "Comparison contract" section and
	// TestComparison_NonDispatchObjectChainComparesEqualToNil.
	leftIsObject := false
	rightIsObject := false
	if lc, ok := left.(sugar.Chain); ok {
		leftIsObject = lc.IsDispatch()
		var err error
		left, err = lc.Value()
		if err != nil {
			return nil, err
		}
	}
	if rc, ok := right.(sugar.Chain); ok {
		rightIsObject = rc.IsDispatch()
		var err error
		right, err = rc.Value()
		if err != nil {
			return nil, err
		}
	}
	if leftIsObject || rightIsObject {
		return nil, fmt.Errorf("unsupported binary operation: COM object %s operand (object operands are not comparable; compare a property of it instead)", op)
	}

	lv := reflect.ValueOf(left)
	rv := reflect.ValueOf(right)

	switch op {
	case "+":
		if lv.Kind() == reflect.String || rv.Kind() == reflect.String {
			return fmt.Sprintf("%v%v", left, right), nil
		}
		if isNumber(lv) && isNumber(rv) {
			return toFloat(lv) + toFloat(rv), nil
		}
	case "-":
		if isNumber(lv) && isNumber(rv) {
			return toFloat(lv) - toFloat(rv), nil
		}
	case "*":
		if isNumber(lv) && isNumber(rv) {
			return toFloat(lv) * toFloat(rv), nil
		}
	case "/":
		if isNumber(lv) && isNumber(rv) {
			return toFloat(lv) / toFloat(rv), nil
		}
	}

	if isComparisonOp(op) {
		if res, ok := evalComparison(op, left, right, lv, rv); ok {
			return res, nil
		}
		// Fall through to the shared unsupported-operation error below, so an
		// unorderable pairing reads the same way as an unknown operator.
	}

	return nil, fmt.Errorf("unsupported binary operation: %T %s %T", left, op, right)
}

// isLogicalOp reports whether op is one of the short-circuiting connectives.
// Both spellings of each are accepted, matching the expr grammar this package
// parses with.
func isLogicalOp(op string) bool {
	switch op {
	case "&&", "||", "and", "or":
		return true
	}
	return false
}

// evalLogical implements && / || / and / or with real short-circuit semantics.
//
// It MUST live here, called from eval's BinaryNode arm before the right-hand side
// is touched -- not in evalBinary, which only ever sees two already-evaluated
// operands. For a COM expression the difference is not academic: the un-taken
// branch would still issue its COM round trips, and an error it raised would
// propagate even though short-circuit semantics say that branch never ran.
//
// No truthiness coercion. A non-bool operand is an error, deliberately: this
// package evaluates expressions over spreadsheet data, where "1" and "non-empty
// string" are exactly the values a user would be surprised to see silently
// treated as true.
func (v *comVisitor) evalLogical(op string, leftNode, rightNode ast.Node) (interface{}, error) {
	lv, err := v.eval(leftNode)
	if err != nil {
		return nil, err
	}
	lb, err := asBool(op, lv)
	if err != nil {
		return nil, err
	}

	// The short circuit itself: && with a false left and || with a true left
	// answer WITHOUT evaluating the right node at all.
	switch op {
	case "&&", "and":
		if !lb {
			return false, nil
		}
	case "||", "or":
		if lb {
			return true, nil
		}
	}

	rv, err := v.eval(rightNode)
	if err != nil {
		return nil, err
	}
	rb, err := asBool(op, rv)
	if err != nil {
		return nil, err
	}
	return rb, nil
}

// asBool unwraps a Chain result and requires a real bool.
func asBool(op string, val interface{}) (bool, error) {
	if ch, ok := val.(sugar.Chain); ok {
		var err error
		val, err = ch.Value()
		if err != nil {
			return false, err
		}
	}
	b, ok := val.(bool)
	if !ok {
		return false, fmt.Errorf("unsupported binary operation: %T %s bool (no truthiness coercion)", val, op)
	}
	return b, nil
}

// compareOrdered applies one of the six operators to two values of the same
// ordered type.
//
// The two INCLUSIVE operators are the ones worth reading twice: >= differs from >
// only at equality, so a test suite whose >= rows are all false-direction cannot
// tell the two apart. (That was the case here until boundary-true rows were
// added -- the mutant survived.)
func compareOrdered[T float64 | string](op string, a, b T) bool {
	switch op {
	case "==":
		return a == b
	case "!=":
		return a != b
	case "<":
		return a < b
	case "<=":
		return a <= b
	case ">":
		return a > b
	default: // ">="
		return a >= b
	}
}

// isComparisonOp reports whether op is one of the six comparisons.
func isComparisonOp(op string) bool {
	switch op {
	case "==", "!=", "<", "<=", ">", ">=":
		return true
	}
	return false
}

// evalComparison applies a comparison to two already-unwrapped operands, or
// reports that the operand TYPES do not support it.
//
// The type rules are deliberately narrow, because the permissive alternatives are
// all silently wrong on spreadsheet data:
//   - numbers compare as float64 (the same toFloat coercion arithmetic uses), so
//     int-vs-float mixes work;
//   - strings compare lexicographically;
//   - bools support EQUALITY ONLY -- "true < false" is a question with no answer;
//   - nil supports equality only, which is the "is this cell empty" idiom;
//   - every cross-type pairing is an ERROR, not a fmt.Sprint comparison.
func evalComparison(op string, left, right interface{}, lv, rv reflect.Value) (interface{}, bool) {
	eqOnly := op == "==" || op == "!="

	// A VT_NULL COM value ("the cells disagree, there is no single value" /
	// SQL NULL) is not comparable, and the check must come BEFORE the nil arm.
	// It used to decode to a bare nil, so `x == nil` answered TRUE for it — the
	// "is this cell empty" idiom giving a confident wrong answer. With the
	// sugar.Null sentinel the untouched code would silently flip that to FALSE,
	// which is no better. VBA and SQL both say a comparison against NULL is
	// itself NULL; in a (value, ok) evaluator the honest spelling of "neither
	// true nor false" is to refuse, and the caller then sees the shared
	// unsupported-binary-operation error naming sugar.Null.
	if sugar.IsNull(left) || sugar.IsNull(right) {
		return nil, false
	}

	// nil: equality only, and nil equals only nil.
	if left == nil || right == nil {
		if !eqOnly {
			return nil, false
		}
		both := left == nil && right == nil
		if op == "==" {
			return both, true
		}
		return !both, true
	}

	if isNumber(lv) && isNumber(rv) {
		return compareOrdered(op, toFloat(lv), toFloat(rv)), true
	}
	if lv.Kind() == reflect.String && rv.Kind() == reflect.String {
		return compareOrdered(op, lv.String(), rv.String()), true
	}
	if lv.Kind() == reflect.Bool && rv.Kind() == reflect.Bool {
		if !eqOnly {
			return nil, false // bools are not ordered
		}
		if op == "==" {
			return lv.Bool() == rv.Bool(), true
		}
		return lv.Bool() != rv.Bool(), true
	}
	// Mixed kinds (number vs string, string vs bool, number vs bool, ...).
	return nil, false
}

func isNumber(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64,
		reflect.Float32, reflect.Float64:
		return true
	}
	return false
}

func toFloat(v reflect.Value) float64 {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return float64(v.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return float64(v.Uint())
	case reflect.Float32, reflect.Float64:
		return v.Float()
	}
	return 0
}
