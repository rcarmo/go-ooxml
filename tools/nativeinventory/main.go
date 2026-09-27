// Command nativeinventory inventories tracked native Go test declarations and
// parameter groups. Discovery is static and grants no execution credit.
package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"flag"
	"fmt"
	"go/ast"
	"go/parser"
	"go/printer"
	"go/token"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
)

type group struct {
	Line       int      `json:"line"`
	Expression string   `json:"expression"`
	Rows       int      `json:"rows"`
	Names      []string `json:"names,omitempty"`
}
type subtest struct {
	Line       int    `json:"line"`
	Name       string `json:"name,omitempty"`
	Expression string `json:"expression,omitempty"`
	Dynamic    bool   `json:"dynamic"`
}
type declaration struct {
	ID          string    `json:"id"`
	File        string    `json:"file"`
	Line        int       `json:"line"`
	Function    string    `json:"function"`
	Kind        string    `json:"kind"`
	Subtests    []subtest `json:"subtests"`
	Groups      []group   `json:"parameter_groups"`
	Calls       []string  `json:"calls"`
	ScenarioIDs []string  `json:"scenario_ids"`
	Coverage    string    `json:"coverage"`
}
type fileRecord struct {
	Path        string `json:"path"`
	Hash        string `json:"sha256"`
	HelpersOnly bool   `json:"helpers_only"`
}
type inventory struct {
	Schema       int           `json:"schema"`
	Revision     string        `json:"observation_revision"`
	Files        []fileRecord  `json:"files"`
	Declarations []declaration `json:"declarations"`
	Limits       []string      `json:"limits"`
}

func main() {
	root := flag.String("root", ".", "native consumer root")
	out := flag.String("out", "docs/behaviors/native-inventory.json", "output path relative to root")
	flag.Parse()
	i, err := discover(*root)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	b, err := json.MarshalIndent(i, "", "  ")
	if err != nil {
		panic(err)
	}
	path := filepath.Join(*root, *out)
	if err = os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		panic(err)
	}
	if err = os.WriteFile(path, append(b, '\n'), 0644); err != nil {
		panic(err)
	}
	fmt.Printf("%d tracked test files, %d runnable declarations; no execution credit\n", len(i.Files), len(i.Declarations))
}
func git(root string, args ...string) ([]byte, error) {
	c := exec.Command("git", args...)
	c.Dir = root
	return c.Output()
}
func discover(root string) (inventory, error) {
	out := inventory{Schema: 1, Limits: []string{"Static native-source discovery only; no execution credit.", "Dynamic subtest names and runtime-generated loops are parameter groups, not enumerated leaf cases.", "Literal table rows may include setup data; functional mapping requires review.", "Acceptance step registration and helpers-only files are listed but do not create executable cases.", "Fuzz seeds and benchmarks do not establish exploratory fuzz or performance results."}}
	rev, err := git(root, "rev-parse", "HEAD")
	if err != nil {
		return out, err
	}
	out.Revision = strings.TrimSpace(string(rev))
	paths, err := git(root, "ls-files", "-z", "--", "*_test.go")
	if err != nil {
		return out, err
	}
	for _, p := range strings.Split(string(paths), "\x00") {
		if p == "" {
			continue
		}
		b, err := os.ReadFile(filepath.Join(root, p))
		if err != nil {
			return out, err
		}
		f, ds, err := scan(p, b)
		if err != nil {
			return out, err
		}
		out.Files = append(out.Files, f)
		out.Declarations = append(out.Declarations, ds...)
	}
	return out, nil
}
func expression(fs *token.FileSet, n ast.Node) string {
	var b bytes.Buffer
	_ = printer.Fprint(&b, fs, n)
	s := b.String()
	if len(s) > 240 {
		s = s[:240] + "…"
	}
	return s
}
func testKind(name string) string {
	for _, k := range []string{"Test", "Fuzz", "Benchmark", "Example"} {
		if strings.HasPrefix(name, k) {
			if len(name) == len(k) || !(name[len(k)] >= 'a' && name[len(k)] <= 'z') {
				return strings.ToLower(k)
			}
		}
	}
	return ""
}
func scan(path string, b []byte) (fileRecord, []declaration, error) {
	hash := sha256.Sum256(b)
	f := fileRecord{Path: path, Hash: hex.EncodeToString(hash[:]), HelpersOnly: true}
	fs := token.NewFileSet()
	node, err := parser.ParseFile(fs, path, b, 0)
	if err != nil {
		return f, nil, err
	}
	var out []declaration
	for _, n := range node.Decls {
		fn, ok := n.(*ast.FuncDecl)
		if !ok || fn.Recv != nil || fn.Body == nil {
			continue
		}
		kind := testKind(fn.Name.Name)
		if kind == "" {
			continue
		}
		f.HelpersOnly = false
		d := declaration{ID: path + "::" + fn.Name.Name, File: path, Line: fs.Position(fn.Pos()).Line, Function: fn.Name.Name, Kind: kind, Subtests: []subtest{}, Groups: []group{}, Calls: []string{}, ScenarioIDs: []string{}, Coverage: "unreviewed"}
		calls := map[string]bool{}
		ast.Inspect(fn.Body, func(n ast.Node) bool {
			switch x := n.(type) {
			case *ast.CallExpr:
				name := ""
				switch fun := x.Fun.(type) {
				case *ast.SelectorExpr:
					name = fun.Sel.Name
				case *ast.Ident:
					name = fun.Name
				}
				if name != "" {
					calls[name] = true
				}
				if sel, ok := x.Fun.(*ast.SelectorExpr); ok && sel.Sel.Name == "Run" && len(x.Args) >= 2 {
					r := subtest{Line: fs.Position(x.Pos()).Line, Dynamic: true, Expression: expression(fs, x.Args[0])}
					if s, ok := x.Args[0].(*ast.BasicLit); ok && s.Kind == token.STRING {
						r.Name, _ = strconv.Unquote(s.Value)
						r.Expression = ""
						r.Dynamic = false
					}
					d.Subtests = append(d.Subtests, r)
				}
			case *ast.CompositeLit:
				if _, ok := x.Type.(*ast.ArrayType); ok {
					g := group{Line: fs.Position(x.Pos()).Line, Expression: expression(fs, x.Type), Rows: len(x.Elts)}
					for _, e := range x.Elts {
						row, ok := e.(*ast.CompositeLit)
						if !ok {
							continue
						}
						for _, v := range row.Elts {
							if kv, ok := v.(*ast.KeyValueExpr); ok {
								key, ok := kv.Key.(*ast.Ident)
								if !ok || (key.Name != "name" && key.Name != "Name") {
									continue
								}
								v = kv.Value
							}
							if s, ok := v.(*ast.BasicLit); ok && s.Kind == token.STRING {
								value, _ := strconv.Unquote(s.Value)
								if len(value) > 160 {
									value = value[:160] + "…"
								}
								g.Names = append(g.Names, value)
								break
							}
						}
					}
					d.Groups = append(d.Groups, g)
				}
			}
			return true
		})
		for c := range calls {
			d.Calls = append(d.Calls, c)
		}
		sort.Strings(d.Calls)
		out = append(out, d)
	}
	return f, out, nil
}
