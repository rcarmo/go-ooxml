package document

import (
	"bytes"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"os"
	"path/filepath"
	"testing"
)

func TestRetainedTableWordLifecycle(t *testing.T) {
	input := wordTableGuardFixture(t)
	one, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	two, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	target, e := one.FindRetainedTable(0, 0, 0)
	if e != nil {
		t.Fatal(e)
	}
	check := func(err error, kind string) {
		t.Helper()
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != kind {
			t.Fatalf("refusal=%v want %s", err, kind)
		}
	}
	check(two.SetRetainedTable(target, map[string]any{"shading": "A1B2C3"}), "stale_target")
	if e = one.SetRetainedTable(target, map[string]any{"shading": "4472C4"}); e != nil {
		t.Fatalf("same-value edit: %v", e)
	}
	if e = one.SetRetainedTable(target, map[string]any{"shading": "A1B2C3"}); e != nil {
		t.Fatal(e)
	}
	check(one.SetRetainedTable(target, map[string]any{"shading": "334455"}), "stale_target")
	dest := filepath.Join(t.TempDir(), "output.docx")
	if _, e = one.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	saved, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Equal(saved, input) {
		t.Fatal("successful edit lost")
	}
	reopened, e := OpenEditing(saved, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	if _, e = reopened.FindRetainedTable(0, 0, 0); e != nil {
		t.Fatal(e)
	}
}
