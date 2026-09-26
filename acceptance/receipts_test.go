package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func receiptSteps(sc *godog.ScenarioContext) {
	var p *packaging.Preserved
	var source []byte
	var dir, out string
	var receipt packaging.Receipt
	var failure error
	setup := func() error {
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<document/>`))
		_, _ = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte{4, 3, 2, 1})
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		p, err = packaging.OpenPreserved(source, packaging.Limits{})
		if err != nil {
			return err
		}
		_, h, err := p.Part("word/document.xml")
		if err != nil {
			return err
		}
		return p.Replace([]packaging.Replacement{{Part: "word/document.xml", ExpectedSHA256: h, Data: []byte(`<changed/>`)}})
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		dir = ""
		p = nil
		source = nil
		failure = nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if dir != "" {
			return ctx, os.RemoveAll(dir)
		}
		return ctx, nil
	})
	sc.Step(`^a staged main-part edit with an opaque custom part$`, setup)
	sc.Step(`^I save a main-part replacement with a delivery receipt$`, func() error {
		var err error
		dir, err = os.MkdirTemp("", "ooxml-receipt-")
		if err != nil {
			return err
		}
		out = filepath.Join(dir, "out.docx")
		receipt, err = p.SaveAs(out)
		return err
	})
	sc.Step(`^the receipt identifies exactly the changed member and both hashes$`, func() error {
		if receipt.Schema != 1 || len(receipt.Changes) != 1 {
			return fmt.Errorf("receipt: %+v", receipt)
		}
		after, err := os.ReadFile(out)
		if err != nil {
			return err
		}
		a, err := zipPayloads(after)
		if err != nil {
			return err
		}
		b, err := zipPayloads(source)
		if err != nil {
			return err
		}
		c := receipt.Changes[0]
		hash := func(data []byte) string { s := sha256.Sum256(data); return hex.EncodeToString(s[:]) }
		if c.Part != "word/document.xml" || c.BeforeSHA256 != hash(b[c.Part]) || c.AfterSHA256 != hash(a[c.Part]) {
			return fmt.Errorf("receipt disagrees with independent hashes")
		}
		return nil
	})
	sc.Step(`^the delivered archive reopens with unchanged opaque bytes$`, func() error {
		r, err := zip.OpenReader(out)
		if err != nil {
			return err
		}
		_ = r.Close()
		a, err := os.ReadFile(out)
		if err != nil {
			return err
		}
		parts, err := zipPayloads(a)
		if err != nil {
			return err
		}
		if !bytes.Equal(parts["custom/opaque.bin"], []byte{4, 3, 2, 1}) {
			return fmt.Errorf("opaque part changed")
		}
		return nil
	})
	sc.Step(`^I deliver staged changes to a directory destination$`, func() error {
		var err error
		dir, err = os.MkdirTemp("", "ooxml-receipt-")
		if err != nil {
			return err
		}
		_, failure = p.SaveAs(dir)
		return nil
	})
	sc.Step(`^delivery fails without changing the directory$`, func() error {
		if failure == nil {
			return fmt.Errorf("directory destination accepted")
		}
		entries, err := os.ReadDir(dir)
		if err != nil {
			return err
		}
		if len(entries) != 0 {
			return fmt.Errorf("directory changed")
		}
		return nil
	})
	sc.Step(`^the staged change remains available for another save$`, func() error {
		b, _, err := p.Part("word/document.xml")
		if err != nil {
			return err
		}
		if string(b) != `<changed/>` {
			return fmt.Errorf("staged edit lost")
		}
		_, err = p.SaveAs(filepath.Join(dir, "retry.docx"))
		return err
	})
}
