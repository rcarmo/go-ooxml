package packaging_test

import (
	"bytes"
	"errors"
	"fmt"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestEnvelopeTransactionExecutionModesAndCustody(t *testing.T) {
	original := envelopeProfileSample(profileCT, profileRels, "word/document.xml")
	model, err := packaging.OpenEnvelope(original)
	if err != nil {
		t.Fatal(err)
	}
	originalBytes, err := model.Bytes()
	if err != nil || !bytes.Equal(originalBytes, original) {
		t.Fatal("initial bytes", err)
	}
	ran := false
	value, err := model.Transaction(packaging.TransactionDeferred, func(*packaging.Envelope) (any, error) { ran = true; return "should not run", nil })
	var refusal *packaging.ProfileError
	if ran || value != nil || !errors.As(err, &refusal) || refusal.Reason != "opc-deferred-transaction" {
		t.Fatalf("deferred callback ran/result=%v/error=%v", value, err)
	}
	if _, err := model.Transaction(packaging.TransactionDeferred, nil); !errors.As(err, &refusal) || refusal.Reason != "opc-deferred-transaction" {
		t.Fatal("deferred nil callback was not refused", err)
	}
	unchanged, err := model.Bytes()
	if err != nil || !bytes.Equal(unchanged, original) {
		t.Fatal("deferred changed source", err)
	}
	if !bytes.Equal(originalBytes, original) {
		t.Fatal("returned source changed")
	}

	originalPart, err := model.Part("word/document.xml")
	if err != nil {
		t.Fatal(err)
	}
	originalPart[0] = 'X'
	fresh, err := model.Part("word/document.xml")
	if err != nil || string(fresh) != profileDoc {
		t.Fatal("Part exposed mutable source", err)
	}
	var retained *packaging.Envelope
	opaque := &struct{ calls int }{}
	result, err := model.Transaction(packaging.TransactionImmediate, func(work *packaging.Envelope) (any, error) {
		ran = true
		retained = work
		part, e := work.Part("word/document.xml")
		if e != nil {
			return nil, e
		}
		if string(part) != profileDoc {
			return nil, fmt.Errorf("working copy did not contain original")
		}
		copyData := []byte(strings.Replace(string(part), "Alpha", "Beta", 1))
		if e = work.SetPart("word/document.xml", copyData); e != nil {
			return nil, e
		}
		copyData[0] = 'X' // prove SetPart owns its bytes before commit
		return opaque, nil
	})
	if err != nil || !ran || result != opaque || opaque.calls != 0 {
		t.Fatalf("immediate result=%v err=%v", result, err)
	}
	beta, err := model.Part("word/document.xml")
	if err != nil || !bytes.Contains(beta, []byte("Beta")) || bytes.Contains(beta, []byte("Alpha")) {
		t.Fatal("Beta not committed", err)
	}
	serialized, err := model.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	reopened, err := packaging.OpenEnvelope(serialized)
	if err != nil {
		t.Fatal(err)
	}
	b, err := reopened.Part("word/document.xml")
	if err != nil || !bytes.Equal(b, beta) {
		t.Fatal("reopened Beta missing", err)
	}
	for _, name := range []string{"[Content_Types].xml", "_rels/.rels"} {
		originalModel, e := packaging.OpenEnvelope(original)
		if e != nil {
			t.Fatal(e)
		}
		before, e := originalModel.Part(name)
		if e != nil {
			t.Fatal(e)
		}
		after, e := reopened.Part(name)
		if e != nil || !bytes.Equal(before, after) {
			t.Fatal("unrelated payload changed", name, e)
		}
	}
	if err := retained.SetPart("word/document.xml", []byte("late")); err != nil {
		t.Fatal(err)
	}
	if err := retained.DeletePart("word/document.xml"); err != nil {
		t.Fatal(err)
	}
	afterHeld, err := model.Part("word/document.xml")
	if err != nil || !bytes.Equal(afterHeld, beta) {
		t.Fatal("callback-held state aliased model", err)
	}
	stillCommitted, err := model.Bytes()
	if err != nil || !bytes.Equal(stillCommitted, serialized) || !bytes.Equal(originalBytes, original) {
		t.Fatal("callback-held state changed committed archive or caller source", err)
	}

	callbackFailure := errors.New("callback failure")
	for _, tc := range []struct {
		name     string
		callback func(*packaging.Envelope) (any, error)
	}{
		{"callback error", func(work *packaging.Envelope) (any, error) {
			_ = work.SetPart("word/document.xml", []byte(profileDoc))
			return nil, callbackFailure
		}},
		{"validation failure", func(work *packaging.Envelope) (any, error) { _ = work.DeletePart("word/document.xml"); return nil, nil }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			before, err := model.Bytes()
			if err != nil {
				t.Fatal(err)
			}
			value, err := model.Transaction(packaging.TransactionImmediate, tc.callback)
			if err == nil || value != nil {
				t.Fatal("transaction unexpectedly committed", value, err)
			}
			after, e := model.Bytes()
			if e != nil || !bytes.Equal(before, after) {
				t.Fatal("rollback changed archive", e)
			}
		})
	}
	if !bytes.Equal(originalBytes, original) {
		t.Fatal("caller original altered")
	}
}
