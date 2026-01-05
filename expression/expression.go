//go:build windows

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

// Compile parses an expression and returns a Program.
func Compile(expression string) (*Program, error) {
	tree, err := parser.Parse(expression)
	if err != nil {
		return nil, err
	}
	return &Program{node: tree.Node}, nil
}

// Run executes a compiled Program against an environment.
// The env can be a *sugar.Chain, *ole.IDispatch, or a map[string]interface{}.
func (p *Program) Run(env interface{}) (interface{}, error) {
	var chain *sugar.Chain
	var envMap map[string]interface{}

	switch v := env.(type) {
	case *sugar.Chain:
		chain = v
	case *ole.IDispatch:
		chain = sugar.From(v)
	case map[string]interface{}:
		envMap = v
	}

	visitor := &comVisitor{initialChain: chain, envMap: envMap}
	return visitor.eval(p.node)
}

// Eval parses and executes an expression in one step.
func Eval(expression string, env interface{}) (interface{}, error) {
	p, err := Compile(expression)
	if err != nil {
		return nil, err
	}
	return p.Run(env)
}

// Get retrieves a property or calls a method using an expression.
func Get(obj interface{}, expression string) (interface{}, error) {
	result, err := Eval(expression, obj)
	if err != nil {
		return nil, err
	}

	if finalChain, ok := result.(*sugar.Chain); ok {
		if err := finalChain.Err(); err != nil {
			return nil, err
		}
		return finalChain.Value()
	}

	return result, nil
}

// Store retrieves a COM object (IDispatch) using an expression.
func Store(obj interface{}, expression string) (*ole.IDispatch, error) {
	result, err := Eval(expression, obj)
	if err != nil {
		return nil, err
	}

	if finalChain, ok := result.(*sugar.Chain); ok {
		if err := finalChain.Err(); err != nil {
			return nil, err
		}
		return finalChain.Store()
	}

	return nil, fmt.Errorf("expression did not evaluate to a COM object")
}

// Put sets a property using an expression.
func Put(obj interface{}, expression string, value interface{}) error {
	p, err := Compile(expression)
	if err != nil {
		return err
	}

	memberNode, ok := p.node.(*ast.MemberNode)
	if !ok {
		return fmt.Errorf("invalid Put expression: must be property access")
	}

	var chain *sugar.Chain
	var envMap map[string]interface{}
	switch v := obj.(type) {
	case *sugar.Chain:
		chain = v
	case *ole.IDispatch:
		chain = sugar.From(v)
	case map[string]interface{}:
		envMap = v
	}

	visitor := &comVisitor{initialChain: chain, envMap: envMap}
	parentObj, err := visitor.eval(memberNode.Node)
	if err != nil {
		return err
	}

	parentChain, ok := parentObj.(*sugar.Chain)
	if !ok {
		return fmt.Errorf("parent is not COM object: %T", parentObj)
	}

	propName := ""
	if id, ok := memberNode.Property.(*ast.StringNode); ok {
		propName = id.Value
	} else if id, ok := memberNode.Property.(*ast.IdentifierNode); ok {
		propName = id.Value
	}

	return parentChain.Put(propName, value).Err()
}

type comVisitor struct {
	initialChain *sugar.Chain
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
		chain, ok := left.(*sugar.Chain)
		if !ok {
			// Fallback to reflection for non-COM objects if needed? 
			// For now, return error.
			return nil, fmt.Errorf("cannot access property on type %T", left)
		}

		propName := ""
		if id, ok := n.Property.(*ast.StringNode); ok {
			propName = id.Value
		} else if id, ok := n.Property.(*ast.IdentifierNode); ok {
			propName = id.Value
		}
		return chain.Get(propName), nil

	case *ast.CallNode:
		args := make([]interface{}, len(n.Arguments))
		for i, argNode := range n.Arguments {
			argVal, err := v.eval(argNode)
			if err != nil {
				return nil, err
			}
			if argChain, ok := argVal.(*sugar.Chain); ok {
				val, err := argChain.Value()
				if err != nil {
					return nil, fmt.Errorf("arg %d error: %w", i, err)
				}
				args[i] = val
			} else {
				args[i] = argVal
			}
		}

		switch callee := n.Callee.(type) {
		case *ast.MemberNode:
			obj, err := v.eval(callee.Node)
			if err != nil {
				return nil, err
			}
			chain, ok := obj.(*sugar.Chain)
			if !ok {
				return nil, fmt.Errorf("cannot call method on type %T", obj)
			}

			methodName := ""
			if id, ok := callee.Property.(*ast.StringNode); ok {
				methodName = id.Value
			} else if id, ok := callee.Property.(*ast.IdentifierNode); ok {
				methodName = id.Value
			}
			return chain.Call(methodName, args...), nil

		case *ast.IdentifierNode:
			if v.envMap != nil {
				if val, ok := v.envMap[callee.Value]; ok {
					// If it's a function in map, we could call it here.
					// For now, assume it's a COM method on the implicit root if not in map.
					if fn, ok := val.(func(...interface{}) (interface{}, error)); ok {
						return fn(args...)
					}
				}
			}
			if v.initialChain != nil {
				return v.initialChain.Call(callee.Value, args...), nil
			}
			return nil, fmt.Errorf("method not found: %s", callee.Value)
		default:
			return nil, fmt.Errorf("unsupported call on %T", callee)
		}

	case *ast.BinaryNode:
		left, err := v.eval(n.Left)
		if err != nil {
			return nil, err
		}
		right, err := v.eval(n.Right)
		if err != nil {
			return nil, err
		}

		// Handle basic arithmetic
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

func evalBinary(op string, left, right interface{}) (interface{}, error) {
	// Unwrap Chains if necessary
	if lc, ok := left.(*sugar.Chain); ok {
		var err error
		left, err = lc.Value()
		if err != nil {
			return nil, err
		}
	}
	if rc, ok := right.(*sugar.Chain); ok {
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

	return nil, fmt.Errorf("unsupported binary operation: %v %s %v", reflect.TypeOf(left), op, reflect.TypeOf(right))
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