package main

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestGoLeafWhitespace(t *testing.T) {
	if got := goName("@id-case [RGB fill]"); got != "@id-case_[RGB_fill]" {
		t.Fatal(got)
	}
}
func TestEventBatchControls(t *testing.T) {
	cases := []struct {
		name        string
		actions     []string
		omitPackage bool
		badJSON     bool
		wantError   bool
	}{
		{name: "passed", actions: []string{"pass"}},
		{name: "skip leaf retained", actions: []string{"skip"}},
		{name: "failed leaf retained", actions: []string{"fail"}},
		{name: "duplicate retained", actions: []string{"pass", "pass"}},
		{name: "incomplete package", actions: []string{"pass"}, omitPackage: true, wantError: true},
		{name: "bad JSON", badJSON: true, wantError: true},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "events.jsonl")
			f, e := os.Create(path)
			if e != nil {
				t.Fatal(e)
			}
			enc := json.NewEncoder(f)
			for _, p := range []string{"pkg/presentation", "pkg/packaging", "internal/losslessxml"} {
				if c.omitPackage && p == "pkg/presentation" {
					continue
				}
				if e = enc.Encode(map[string]string{"Action": "pass", "Package": "github.com/rcarmo/go-ooxml/" + p}); e != nil {
					t.Fatal(e)
				}
			}
			for _, a := range c.actions {
				if e = enc.Encode(map[string]string{"Action": a, "Package": "github.com/rcarmo/go-ooxml/pkg/presentation", "Test": "TestRecipes/@id-case_[row]"}); e != nil {
					t.Fatal(e)
				}
			}
			if c.badJSON {
				_, e = f.WriteString("broken\n")
				if e != nil {
					t.Fatal(e)
				}
			}
			if e = f.Close(); e != nil {
				t.Fatal(e)
			}
			got, e := readEvents(path)
			if (e != nil) != c.wantError {
				t.Fatal(e)
			}
			if !c.wantError && len(got["TestRecipes/@id-case_[row]"]) != len(c.actions) {
				t.Fatal("terminal actions lost")
			}
		})
	}
}

func TestExactlyOneExecutedPass(t *testing.T) {
	for _, c := range []struct {
		name    string
		actions []string
		want    bool
	}{{"one pass", []string{"pass"}, true}, {"missing", nil, false}, {"skipped", []string{"skip"}, false}, {"failed", []string{"fail"}, false}, {"duplicate", []string{"pass", "pass"}, false}, {"retry", []string{"fail", "pass"}, false}} {
		t.Run(c.name, func(t *testing.T) {
			if exactlyOnePass(c.actions) != c.want {
				t.Fatal(c.actions)
			}
		})
	}
}
