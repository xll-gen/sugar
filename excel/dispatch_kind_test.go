//go:build windows

// Excel-free guard against the "propget invoked as DISPATCH_METHOD" class of
// bug (and its mirror, "method read as a property").
//
// Background: `chain.Get` invokes with DISPATCH_PROPERTYGET, `chain.Put` with
// DISPATCH_PROPERTYPUT, `chain.Call` with DISPATCH_METHOD. Excel's IDispatch
// rejects the wrong flag with DISP_E_MEMBERNOTFOUND, so picking the wrong verb
// makes the wrapper fail 100% of the time — but only at runtime, against a live
// Excel, which plain `go test ./...` never sees. That is exactly how
// Worksheet.Clear/ClearContents shipped broken: they read the Cells *property*
// with Call.
//
// The confusion is real because Excel is not consistent: `Worksheet.Cells`,
// `Worksheet.UsedRange` and `Range.EntireColumn` are properties, while
// `Worksheet.ChartObjects`, `Worksheet.Pictures` and `Names.Item` are methods
// even though they read like collection properties.
//
// This test freezes the "member x DISPATCH kind" table below and statically
// checks every literal member name used in the package's non-test sources
// against it. It needs no Excel, so it runs on every `go test ./...`.

package excel

import (
	"go/ast"
	"go/parser"
	"go/token"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// dispatchKind is how a COM member must be invoked.
type dispatchKind int

const (
	// kindProperty members are propget/propput: use chain.Get / chain.Put
	// (or the getInt32/getFloat64/getBool/getString helpers, which Get).
	kindProperty dispatchKind = iota
	// kindMethod members are real methods: use chain.Call (or callOptional).
	kindMethod
)

func (k dispatchKind) String() string {
	if k == kindMethod {
		return "method (use Call)"
	}
	return "property (use Get/Put)"
}

// dispatchKinds is the authoritative "member x DISPATCH kind" table for every
// Excel COM member this package names with a string literal. Adding a new
// member to the object model requires adding a row here — the test fails on any
// unclassified member so the classification decision cannot be skipped.
//
// Sources: the Excel object model reference. Members marked "trap" read like a
// property but are methods (or vice versa) and are the ones worth double
// checking against the type library before editing.
var dispatchKinds = map[string]dispatchKind{
	// --- Object/collection accessors (properties) ---
	"ActiveSheet":    kindProperty,
	"ActiveWorkbook": kindProperty,
	"Application":    kindProperty,
	// trap: reads like a method, is a propget — and unlike End/Offset/Resize it
	// declares NO parameters. Get("Cells", r, c) works because Invoke forwards
	// the surplus arguments to the returned Range's default member (same
	// mechanism as VBScript's xl.Cells(1,1)).
	"Cells":         kindProperty,
	"Chart":         kindProperty, // ChartObject.Chart
	"Columns":       kindProperty,
	"End":           kindProperty, // trap: takes an argument, still a propget
	"EntireColumn":  kindProperty,
	"EntireRow":     kindProperty,
	"Font":          kindProperty,
	"Interior":      kindProperty,
	"Names":         kindProperty,
	"Offset":        kindProperty, // trap: argumented propget
	"Parent":        kindProperty,
	"Range":         kindProperty, // trap: argumented propget
	"RefersToRange": kindProperty,
	"Resize":        kindProperty, // trap: argumented propget
	"Rows":          kindProperty,
	"Shapes":        kindProperty,
	"Sheets":        kindProperty,
	"UsedRange":     kindProperty,
	"Workbooks":     kindProperty,
	"Worksheet":     kindProperty,
	"Worksheets":    kindProperty,

	// --- Scalar properties ---
	"Address":        kindProperty,
	"Bold":           kindProperty,
	"Calculation":    kindProperty,
	"ChartType":      kindProperty,
	"Color":          kindProperty,
	"Column":         kindProperty,
	"ColumnWidth":    kindProperty,
	"Count":          kindProperty,
	"DisplayAlerts":  kindProperty,
	"Formula":        kindProperty,
	"Formula2":       kindProperty,
	"FullName":       kindProperty,
	"Height":         kindProperty,
	"Hwnd":           kindProperty,
	"Index":          kindProperty,
	"Italic":         kindProperty,
	"Left":           kindProperty,
	"MergeCells":     kindProperty,
	"Name":           kindProperty,
	"NumberFormat":   kindProperty,
	"Path":           kindProperty,
	"RefersTo":       kindProperty,
	"Row":            kindProperty,
	"RowHeight":      kindProperty,
	"Saved":          kindProperty,
	"ScreenUpdating": kindProperty,
	"Size":           kindProperty,
	"Top":            kindProperty,
	"Type":           kindProperty,
	"Value":          kindProperty,
	"Version":        kindProperty,
	"Visible":        kindProperty,
	"Width":          kindProperty,

	// --- Methods ---
	"Activate":            kindMethod,
	"Add":                 kindMethod,
	"AddPicture":          kindMethod,
	"AutoFit":             kindMethod,
	"ChartObjects":        kindMethod, // trap: collection accessor that is a method
	"Clear":               kindMethod,
	"ClearContents":       kindMethod,
	"Close":               kindMethod,
	"Copy":                kindMethod,
	"Delete":              kindMethod,
	"Export":              kindMethod,
	"ExportAsFixedFormat": kindMethod,
	"Find":                kindMethod,
	"Insert":              kindMethod,
	"Merge":               kindMethod,
	"Open":                kindMethod,
	"Pictures":            kindMethod, // trap: legacy collection accessor, a method
	"Quit":                kindMethod,
	"Save":                kindMethod,
	"SaveAs":              kindMethod,
	"SetSourceData":       kindMethod,
	"UnMerge":             kindMethod,
}

// ambiguousMembers are member names whose kind depends on the object they are
// read from, so a single global classification would be wrong. They are exempt
// from the check; the call site must get it right and say why in a comment.
var ambiguousMembers = map[string]string{
	// Sheets.Item / Workbooks.Item are properties; Names.Item,
	// ChartObjects.Item, Shapes.Item and Pictures.Item are methods.
	"Item": "property on Sheets/Workbooks, method on Names/Charts/Shapes/Pictures",
}

// verbKinds maps the chain verbs to the member kind they require.
var verbKinds = map[string]dispatchKind{
	"Get":  kindProperty,
	"Put":  kindProperty,
	"Call": kindMethod,
}

// helperArgKinds maps this package's member-name-taking helpers to the kind
// they imply and the argument index holding the member name. getInt32 and
// friends all route through chain.Get; callOptional routes through chain.Call.
var helperArgKinds = map[string]struct {
	kind dispatchKind
	arg  int
}{
	"getInt32":     {kindProperty, 1},
	"getFloat64":   {kindProperty, 1},
	"getBool":      {kindProperty, 1},
	"getString":    {kindProperty, 1},
	"callOptional": {kindMethod, 1},
}

type memberUse struct {
	member string
	want   dispatchKind
	via    string
	pos    string
}

// TestDispatchKinds_MatchRegistry walks the package's own non-test sources and
// asserts that every COM member named by a string literal is invoked with the
// verb its DISPATCH kind requires. It is the static regression barrier for the
// Worksheet.Clear bug (Call("Cells")) and for its mirror, calling a real method
// with Get.
func TestDispatchKinds_MatchRegistry(t *testing.T) {
	uses := collectMemberUses(t)
	if len(uses) < 50 {
		t.Fatalf("scanner found only %d member uses; it is probably broken", len(uses))
	}

	for _, u := range uses {
		if _, ok := ambiguousMembers[u.member]; ok {
			continue
		}
		got, known := dispatchKinds[u.member]
		if !known {
			t.Errorf("%s: COM member %q used via %s is not classified.\n"+
				"Add it to dispatchKinds in dispatch_kind_test.go (property vs method),\n"+
				"or to ambiguousMembers if its kind depends on the parent object.",
				u.pos, u.member, u.via)
			continue
		}
		if got != u.want {
			t.Errorf("%s: %q is a %s but is invoked via %s.\n"+
				"Excel rejects the wrong DISPATCH flag with DISP_E_MEMBERNOTFOUND, so this call always fails.",
				u.pos, u.member, got, u.via)
		}
	}
}

// TestDispatchKinds_RegistryIsLive fails if the table accumulates rows for
// members the package no longer uses, so the table stays a description of the
// code rather than folklore.
func TestDispatchKinds_RegistryIsLive(t *testing.T) {
	used := map[string]bool{}
	for _, u := range collectMemberUses(t) {
		used[u.member] = true
	}
	// Escape hatch for members reachable only through a variable member name
	// (so the scanner cannot see them). Empty today; add a reason when used.
	allowUnused := map[string]bool{}

	for member := range dispatchKinds {
		if !used[member] && !allowUnused[member] {
			t.Errorf("dispatchKinds has a row for %q but no source uses it; drop the row "+
				"or add it to allowUnused with a reason", member)
		}
	}
}

// collectMemberUses parses every non-test .go file in the package directory and
// returns each (member name, required kind) pair implied by a chain verb or a
// member-name-taking helper.
func collectMemberUses(t *testing.T) []memberUse {
	t.Helper()

	entries, err := os.ReadDir(".")
	if err != nil {
		t.Fatalf("read package dir: %v", err)
	}
	fset := token.NewFileSet()
	var uses []memberUse

	for _, e := range entries {
		name := e.Name()
		if e.IsDir() || !strings.HasSuffix(name, ".go") || strings.HasSuffix(name, "_test.go") {
			continue
		}
		file, err := parser.ParseFile(fset, filepath.Join(".", name), nil, 0)
		if err != nil {
			t.Fatalf("parse %s: %v", name, err)
		}
		ast.Inspect(file, func(n ast.Node) bool {
			call, ok := n.(*ast.CallExpr)
			if !ok {
				return true
			}
			switch fn := call.Fun.(type) {
			case *ast.SelectorExpr: // x.Get("Prop") / x.Call("Method") / x.Put("Prop", v)
				kind, ok := verbKinds[fn.Sel.Name]
				if !ok {
					return true
				}
				if member, ok := stringArg(call, 0); ok {
					uses = append(uses, memberUse{member, kind, fn.Sel.Name + "()",
						fset.Position(call.Pos()).String()})
				}
			case *ast.Ident: // getString(c, "Prop") / callOptional(c, "Method", ...)
				spec, ok := helperArgKinds[fn.Name]
				if !ok {
					return true
				}
				if member, ok := stringArg(call, spec.arg); ok {
					uses = append(uses, memberUse{member, spec.kind, fn.Name + "()",
						fset.Position(call.Pos()).String()})
				}
			}
			return true
		})
	}
	return uses
}

// stringArg returns the value of call's i-th argument when it is an untagged
// string literal. Non-literal member names (dynamic property lookups) are
// skipped: the static check cannot classify them.
func stringArg(call *ast.CallExpr, i int) (string, bool) {
	if i >= len(call.Args) {
		return "", false
	}
	lit, ok := call.Args[i].(*ast.BasicLit)
	if !ok || lit.Kind != token.STRING {
		return "", false
	}
	s, err := strconv.Unquote(lit.Value)
	if err != nil {
		return "", false
	}
	return s, true
}
