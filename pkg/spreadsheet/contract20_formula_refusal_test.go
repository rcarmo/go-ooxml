package spreadsheet

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestContract20DirectFormulaRefusalBatch(t *testing.T) {
	for _, row := range []struct{ address, kind string }{{"B2", "xlsx-shared-formula-edit-unsupported"}, {"D2", "xlsx-array-formula-edit-unsupported"}} {
		t.Run(row.address, func(t *testing.T) {
			source := contract20ReadInput(t, "attributed-formulas")
			original := append([]byte(nil), source...)
			parts := contract20InputMembers(t, source)
			session, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			unaffected, err := session.FindNumber("Calc", "A2")
			if err != nil {
				t.Fatal(err)
			}
			// The operation, not a diagnostic wrapper around FindNumber,
			// classifies the selected attributed formula before any mutation.
			err = session.SetNumberAt("Calc", row.address, 99)
			var refusal *packaging.Refusal
			if !errors.As(err, &refusal) || refusal.Kind != row.kind {
				t.Errorf("direct formula write = %#v, want %s", err, row.kind)
			}
			if !bytes.Equal(source, original) || !reflect.DeepEqual(parts, contract20InputMembers(t, source)) {
				t.Error("refusal changed source/caller bytes")
			}
			refused := filepath.Join(t.TempDir(), "refused.xlsx")
			if _, err := session.SaveAs(refused); err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(refused)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(parts, contract20InputMembers(t, data)) {
				t.Error("refused save changed formula anchor/follower/cache or unrelated member")
			}
			if value, err := session.CurrentNumber(unaffected); err != nil || value != 1 {
				t.Errorf("held unrelated A2 read after formula refusal: %v / %v", value, err)
			}
			if effect, err := session.SetNumberWithAllCachesInvalidated(unaffected, 1); err != nil || effect.ValueChanged {
				t.Errorf("held unrelated A2 no-op after formula refusal: %#v / %v", effect, err)
			}
			if _, err := session.SetNumberWithAllCachesInvalidated(unaffected, 99); err == nil {
				t.Error("changed A2 bypassed attributed topology")
			}
			if value, err := session.CurrentNumber(unaffected); err != nil || value != 1 {
				t.Errorf("changed refusal consumed held A2: %v / %v", value, err)
			}
		})
	}
}
