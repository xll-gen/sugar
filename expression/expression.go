//go:build windows

package expression

import (
	"fmt"
	"github.com/expr-lang/expr/ast"
	"github.com/expr-lang/expr/parser"
	"github.com/go-ole/go-ole"
	"github.com/xll-gen/sugar"
)

// evaluate parses and evaluates the expression, returning the raw result.
func evaluate(obj interface{}, expression string) (interface{}, error) {
	var chain *sugar.Chain
	switch v := obj.(type) {
	case *sugar.Chain:
		chain = v
	case *ole.IDispatch:
		chain = sugar.From(v)
	default:
		return nil, fmt.Errorf("unsupported type for evaluate: %T", obj)
	}

	// We do not check chain.Err() here because it would release the chain.
	// We rely on the chain to carry the error.

	tree, err := parser.Parse(expression)
	if err != nil {
		return nil, err
	}

	visitor := &comVisitor{initialChain: chain}
	return visitor.eval(tree.Node)
}

// Get retrieves a property or calls a method on a COM object using an expression.
// It returns the value of the property or method result.
// If the result is a COM object (IDispatch), this function will return an error.
// Use Store() to retrieve COM objects.
func Get(obj interface{}, expression string) (interface{}, error) {
	result, err := evaluate(obj, expression)
	if err != nil {
		return nil, err
	}

	if result == nil {
		return nil, nil
	}

	if finalChain, ok := result.(*sugar.Chain); ok {
		if finalChain.Err() != nil {
			return nil, finalChain.Err()
		}
		// Directly call Value(). If it's an IDispatch, Value() will error.
		return finalChain.Value()
	}

	return result, nil
}

// Store retrieves a COM object (IDispatch) using an expression and transfers ownership to the caller.
// The caller is responsible for releasing the returned object.
func Store(obj interface{}, expression string) (*ole.IDispatch, error) {
	result, err := evaluate(obj, expression)
	if err != nil {
		return nil, err
	}

	if finalChain, ok := result.(*sugar.Chain); ok {
		if finalChain.Err() != nil {
			return nil, finalChain.Err()
		}
		return finalChain.Store()
	}

	return nil, fmt.Errorf("expression did not evaluate to a COM object")
}

type comVisitor struct {
	initialChain *sugar.Chain
}

func (v *comVisitor) eval(node ast.Node) (interface{}, error) {
	// Do not check v.initialChain.Err() here.

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
			return nil, fmt.Errorf("cannot access property on a non-COM object type: %T", left)
		}
		
		propName := ""
		if id, ok := n.Property.(*ast.StringNode); ok {
			propName = id.Value
		} else if id, ok := n.Property.(*ast.IdentifierNode); ok {
			propName = id.Value
		} else {
			return nil, fmt.Errorf("unsupported property node type: %T", n.Property)
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
					return nil, fmt.Errorf("error evaluating argument %d: %w", i, err)
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
				return nil, fmt.Errorf("cannot call method on a non-COM object type: %T", obj)
			}
			
			methodName := ""
			if id, ok := callee.Property.(*ast.StringNode); ok {
				methodName = id.Value
			} else if id, ok := callee.Property.(*ast.IdentifierNode); ok {
				methodName = id.Value
			} else {
				return nil, fmt.Errorf("unsupported method property node type: %T", callee.Property)
			}

			// Use Call for methods
			return chain.Call(methodName, args...), nil

		case *ast.IdentifierNode:
			return v.initialChain.Call(callee.Value, args...), nil
		default:
			return nil, fmt.Errorf("unsupported call on type: %T", callee)
		}

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
		return nil, fmt.Errorf("unsupported expression node type: %T", n)
	}
}

// Put sets a property on a COM object using an expression.
// obj can be a *sugar.Chain or *ole.IDispatch.
// expression is a dot-separated property path like "ActiveCell.Value".
// value is the value to assign to the property.
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

	// Do not check chain.Err() here.

	tree, err := parser.Parse(expression)
	if err != nil {
		return err
	}

	memberNode, ok := tree.Node.(*ast.MemberNode)
	if !ok {
		return fmt.Errorf("invalid expression for Put: must be a property access, e.g., 'Application.Version'")
	}

	// Evaluate the object part of the expression (e.g., 'Application' in 'Application.Version')
	visitor := &comVisitor{initialChain: chain}
	parentObj, err := visitor.eval(memberNode.Node)
	if err != nil {
		return fmt.Errorf("could not retrieve parent object: %w", err)
	}

	parentChain, ok := parentObj.(*sugar.Chain)
	if !ok {
		return fmt.Errorf("parent path did not resolve to a COM object, but to %T", parentObj)
	}

	// Now, call Put on the parent object with the final property name.
	propName := ""
	if id, ok := memberNode.Property.(*ast.StringNode); ok {
		propName = id.Value
	} else if id, ok := memberNode.Property.(*ast.IdentifierNode); ok {
		propName = id.Value
	} else {
		return fmt.Errorf("unsupported property node type for Put: %T", memberNode.Property)
	}

	return parentChain.Put(propName, value).Err()
}