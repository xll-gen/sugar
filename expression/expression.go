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
//   - Binary operators: + - * / only (no comparison or logical operators).
//     Numeric operands are coerced to float64, so "2 + 2" yields float64(4)
//     and division by zero follows float64 semantics (+Inf). "+" with a
//     string operand concatenates.
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
	switch v := env.(type) {
	case sugar.Chain:
		return finishEval(&comVisitor{initialChain: v}, p.node)
	case *ole.IDispatch:
		return runInArena(v, p.node)
	case map[string]interface{}:
		return finishEval(&comVisitor{envMap: v}, p.node)
	case nil:
		return finishEval(&comVisitor{}, p.node)
	default:
		return nil, fmt.Errorf("expression: unsupported env type %T", env)
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

	switch v := obj.(type) {
	case sugar.Chain:
		return putNode(&comVisitor{initialChain: v}, p.node, value)
	case *ole.IDispatch:
		// Same arena contract as Run: intermediates are released on return.
		ctx := sugar.NewContext(nil)
		defer ctx.Release()
		return putNode(&comVisitor{initialChain: ctx.From(v)}, p.node, value)
	case map[string]interface{}:
		return putNode(&comVisitor{envMap: v}, p.node, value)
	case nil:
		return putNode(&comVisitor{}, p.node, value)
	default:
		return fmt.Errorf("expression: unsupported env type %T", obj)
	}
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
	if lc, ok := left.(sugar.Chain); ok {
		var err error
		left, err = lc.Value()
		if err != nil {
			return nil, err
		}
	}
	if rc, ok := right.(sugar.Chain); ok {
		var err error
		right, err = rc.Value()
		if err != nil {
			return nil, err
		}
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

	return nil, fmt.Errorf("unsupported binary operation: %T %s %T", left, op, right)
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
