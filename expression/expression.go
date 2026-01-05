//go:build windows

package expression

import (
	"fmt"

	"github.com/expr-lang/expr/ast"
	"github.com/expr-lang/expr/parser"
	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// evaluate parses and evaluates the expression within the context of the object.
// It returns the raw result (usually a *sugar.Chain for objects).
func evaluate(obj interface{}, expression string) (interface{}, error) {
	var chain *sugar.Chain
	switch v := obj.(type) {
	case *sugar.Chain:
		chain = v
	case *ole.IDispatch:
		// Note: Creating a chain from IDispatch without a context means 
		// it won't be auto-tracked unless the caller manages it.
		// However, inside sugar.Do, the user should ideally pass a *sugar.Chain.
		chain = sugar.From(v)
	default:
		return nil, fmt.Errorf("unsupported type for evaluate: %T", obj)
	}

	tree, err := parser.Parse(expression)
	if err != nil {
		return nil, err
	}

	visitor := &comVisitor{initialChain: chain}
	return visitor.eval(tree.Node)
}

// Get retrieves a property or calls a method using an expression.
// It returns the Go value of the result. If the result is a COM object,
// it returns an error (use Store or just use the expression to navigate).
func Get(obj interface{}, expression string) (interface{}, error) {
	result, err := evaluate(obj, expression)
	if err != nil {
		return nil, err
	}

	if finalChain, ok := result.(*sugar.Chain); ok {
		if err := finalChain.Err(); err != nil {
			return nil, err
		}
		// Returns the Go value. For IDispatch, this will return an error.
		return finalChain.Value()
	}

	return result, nil
}

// Store is kept for compatibility but in the Context/Arena model, 
// you can often just use the Chain returned by evaluate-like operations.
// It returns the raw *ole.IDispatch with an increased reference count.
func Store(obj interface{}, expression string) (*ole.IDispatch, error) {
	result, err := evaluate(obj, expression)
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

type comVisitor struct {
	initialChain *sugar.Chain
}

func (v *comVisitor) eval(node ast.Node) (interface{}, error) {
	switch n := node.(type) {
	case *ast.IdentifierNode:
		return v.initialChain.Get(n.Value), nil

	case *ast.MemberNode:
		left, err := v.eval(n.Node)
		if err != nil {
			return nil, err
		}
		chain, ok := left.(*sugar.Chain)
		if !ok {
			return nil, fmt.Errorf("cannot access property on type %T", left)
		}
		
		propName := ""
		if id, ok := n.Property.(*ast.StringNode); ok {
			propName = id.Value
		} else if id, ok := n.Property.(*ast.IdentifierNode); ok {
			propName = id.Value
		} else {
			return nil, fmt.Errorf("unsupported property node: %T", n.Property)
		}
		return chain.Get(propName), nil

	case *ast.CallNode:
		args := make([]interface{}, len(n.Arguments))
		for i, argNode := range n.Arguments {
			argVal, err := v.eval(argNode)
			if err != nil {
				return nil, err
			}
			// If argument is a chain, get its value
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
			return v.initialChain.Call(callee.Value, args...), nil
		default:
			return nil, fmt.Errorf("unsupported call on %T", callee)
		}

	case *ast.IntegerNode: return n.Value, nil
	case *ast.StringNode:  return n.Value, nil
	case *ast.BoolNode:    return n.Value, nil
	case *ast.FloatNode:   return n.Value, nil
	case *ast.NilNode:     return nil, nil
	default:
		return nil, fmt.Errorf("unsupported node: %T", n)
	}
}

// Put sets a property using an expression.
func Put(obj interface{}, expression string, value interface{}) error {
	var chain *sugar.Chain
	switch v := obj.(type) {
	case *sugar.Chain:
		chain = v
	case *ole.IDispatch:
		chain = sugar.From(v)
	default:
		return fmt.Errorf("unsupported type for Put: %T", obj)
	}

	tree, err := parser.Parse(expression)
	if err != nil {
		return err
	}

	memberNode, ok := tree.Node.(*ast.MemberNode)
	if !ok {
		return fmt.Errorf("invalid Put expression: must be property access")
	}

	visitor := &comVisitor{initialChain: chain}
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
