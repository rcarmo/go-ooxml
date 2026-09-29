package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// This case checks the saved document part and reopened direct getter, not styles or rendering.
func directFontSizeSteps(sc *godog.ScenarioContext, outputDir string) {
	var doc document.Document
	var run document.Run
	var savedPath string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, run, savedPath = nil, nil, ""
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word document with one body paragraph and one run containing "Size sample"$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		p := doc.AddParagraph()
		if p == nil {
			return fmt.Errorf("new Word paragraph missing")
		}
		run = p.AddRun()
		if run == nil {
			return fmt.Errorf("new Word run missing")
		}
		run.SetText("Size sample")
		return nil
	})
	sc.Step(`^that run's direct font size is set to 10\.5 points$`, func() error {
		if run == nil {
			return fmt.Errorf("new Word run missing")
		}
		run.SetFontSize(10.5)
		return nil
	})
	sc.Step(`^the Word document is saved and reopened$`, func() error {
		if doc == nil {
			return fmt.Errorf("new Word document missing")
		}
		savedPath = filepath.Join(outputDir, "direct-half-point.docx")
		if err := doc.SaveAs(savedPath); err != nil {
			return err
		}
		if err := doc.Close(); err != nil {
			return err
		}
		doc, run = nil, nil
		var err error
		doc, err = document.Open(savedPath)
		return err
	})
	sc.Step(`^the paragraph text is "Size sample"$`, func() error {
		_, err := checkDirectSizeReopened(doc, "Size sample")
		return err
	})
	sc.Step(`^the run has exactly one direct WordprocessingML w:sz element with w:val "21"$`, func() error {
		if doc == nil || savedPath == "" {
			return fmt.Errorf("Word document not reopened")
		}
		part, err := savedWordDocumentXML(savedPath)
		if err != nil {
			return err
		}
		return checkDirectWordSize(part, "21")
	})
	sc.Step(`^the reopened run's direct font size is 10\.5 points$`, func() error {
		r, err := checkDirectSizeReopened(doc, "Size sample")
		if err != nil {
			return err
		}
		return checkDirectSizeGetter(r, 10.5)
	})
}

func savedWordDocumentXML(path string) ([]byte, error) {
	pkg, err := packaging.Open(path)
	if err != nil {
		return nil, err
	}
	defer pkg.Close()
	part, err := pkg.GetPart(packaging.WordDocumentPath)
	if err != nil {
		return nil, err
	}
	return part.Content()
}

func checkDirectSizeReopened(doc document.Document, wantText string) (document.Run, error) {
	if doc == nil {
		return nil, fmt.Errorf("Word document not reopened")
	}
	paragraphs := doc.Body().Paragraphs()
	if len(paragraphs) != 1 {
		return nil, fmt.Errorf("reopened body paragraphs = %d, want 1", len(paragraphs))
	}
	if got := paragraphs[0].Text(); got != wantText {
		return nil, fmt.Errorf("reopened paragraph text = %q, want %q", got, wantText)
	}
	runs := paragraphs[0].Runs()
	if len(runs) != 1 {
		return nil, fmt.Errorf("reopened paragraph runs = %d, want 1", len(runs))
	}
	return runs[0], nil
}

func checkDirectSizeGetter(run document.Run, want float64) error {
	if run == nil {
		return fmt.Errorf("reopened run missing")
	}
	if got := run.FontSize(); got != want {
		return fmt.Errorf("reopened direct font size = %v, want %v", got, want)
	}
	return nil
}

// Count only w:sz immediately beneath w:rPr for the single body run. w:szCs
// and style definitions cannot satisfy the direct non-complex-script predicate.
func checkDirectWordSize(data []byte, want string) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	wml := packaging.NSWordprocessingML
	var path []xml.Name
	count, matches := 0, 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("saved document XML: %w", err)
		}
		switch node := token.(type) {
		case xml.StartElement:
			path = append(path, node.Name)
			if len(path) >= 5 {
				last := path[len(path)-5:]
				if last[0] == (xml.Name{Space: wml, Local: "body"}) && last[1] == (xml.Name{Space: wml, Local: "p"}) && last[2] == (xml.Name{Space: wml, Local: "r"}) && last[3] == (xml.Name{Space: wml, Local: "rPr"}) && last[4] == (xml.Name{Space: wml, Local: "sz"}) {
					count++
					for _, attr := range node.Attr {
						if attr.Name == (xml.Name{Space: wml, Local: "val"}) && attr.Value == want {
							matches++
						}
					}
				}
			}
		case xml.EndElement:
			if len(path) == 0 || path[len(path)-1] != node.Name {
				return fmt.Errorf("unbalanced saved document XML")
			}
			path = path[:len(path)-1]
		}
	}
	if count != 1 || matches != 1 {
		return fmt.Errorf("direct w:sz count = %d, value %q matches = %d; want one each", count, want, matches)
	}
	return nil
}

func guardDirectFontSizeCase(name string, steps []string) error {
	const wantName = "A 10.5-point run stores 21 half-points after save and reopen"
	if name != wantName {
		return fmt.Errorf("unexpected direct font-size case %q", name)
	}
	return guardParagraphSteps(name, steps, []string{
		`a new Word document with one body paragraph and one run containing "Size sample"`,
		"that run's direct font size is set to 10.5 points",
		"the Word document is saved and reopened",
		`the paragraph text is "Size sample"`,
		`the run has exactly one direct WordprocessingML w:sz element with w:val "21"`,
		"the reopened run's direct font size is 10.5 points",
	})
}

func TestDirectFontSizeNegativeControls(t *testing.T) {
	steps := []string{`a new Word document with one body paragraph and one run containing "Size sample"`, "that run's direct font size is set to 10.5 points", "the Word document is saved and reopened", `the paragraph text is "Size sample"`, `the run has exactly one direct WordprocessingML w:sz element with w:val "21"`, "the reopened run's direct font size is 10.5 points"}
	const name = "A 10.5-point run stores 21 half-points after save and reopen"
	if err := guardDirectFontSizeCase(name, steps); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name    string
		index   int
		changed string
	}{
		{"sample", 0, `a new Word document with one body paragraph and one run containing "Other"`},
		{"setter", 1, "that run's direct font size is set to 11 points"},
		{"save", 2, "the Word document is saved"},
		{"text", 3, `the paragraph text is "Other"`},
		{"XML", 4, `the run has exactly one direct WordprocessingML w:sz element with w:val "22"`},
		{"getter", 5, "the reopened run's direct font size is 11 points"},
	} {
		changed := append([]string{}, steps...)
		changed[tc.index] = tc.changed
		if err := guardDirectFontSizeCase(name, changed); err == nil {
			t.Errorf("changed %s step passed", tc.name)
		}
	}
	if err := guardDirectFontSizeCase("An inherited font size is 10.5 points", steps); err == nil {
		t.Error("changed case name passed")
	}
	if err := guardDirectFontSizeCase(name, steps[:5]); err == nil {
		t.Error("missing step passed")
	}
	if err := guardDirectFontSizeCase(name, append(append([]string{}, steps...), "Office renders 10.5 points")); err == nil {
		t.Error("extra step passed")
	}
	if err := checkDirectWordSize(nil, "21"); err == nil {
		t.Error("empty XML passed")
	}
	if err := checkDirectSizeGetter(nil, 10.5); err == nil {
		t.Error("missing reopened run passed")
	}

	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	p := doc.AddParagraph()
	run := p.AddRun()
	run.SetText("Size sample")
	run.SetFontSize(10.5)
	path := filepath.Join(t.TempDir(), "size.docx")
	if err := doc.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := doc.Close(); err != nil {
		t.Fatal(err)
	}
	reopened, err := document.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	r, err := checkDirectSizeReopened(reopened, "Size sample")
	if err != nil {
		t.Fatal(err)
	}
	if err := checkDirectSizeGetter(r, 10.5); err != nil {
		t.Fatal(err)
	}
	data, err := savedWordDocumentXML(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := checkDirectWordSize(data, "21"); err != nil {
		t.Fatal(err)
	}
	// Controlled XML variants test the predicate without relying on a particular
	// serializer prefix or attribute order in the saved document above.
	wrap := func(runProperties string) []byte {
		return []byte(`<w:document xmlns:w="` + packaging.NSWordprocessingML + `"><w:body><w:p><w:r><w:rPr>` + runProperties + `</w:rPr><w:t>Size sample</w:t></w:r></w:p></w:body></w:document>`)
	}
	if err := checkDirectWordSize(wrap(`<w:sz w:val="21"/><w:szCs w:val="21"/>`), "21"); err != nil {
		t.Fatalf("one direct size alongside complex-script size: %v", err)
	}
	for _, tc := range []struct {
		name string
		xml  []byte
	}{
		{"wrong half-points", wrap(`<w:sz w:val="22"/>`)},
		{"duplicate direct", wrap(`<w:sz w:val="21"/><w:sz w:val="21"/>`)},
		{"missing direct", wrap(`<w:szCs w:val="21"/>`)},
		{"wrong namespace", wrap(`<x:sz xmlns:x="urn:other" x:val="21"/>`)},
		{"unqualified value", wrap(`<w:sz val="21"/>`)},
	} {
		if err := checkDirectWordSize(tc.xml, "21"); err == nil {
			t.Errorf("%s passed", tc.name)
		}
	}
	if _, err := checkDirectSizeReopened(reopened, "Wrong"); err == nil {
		t.Error("wrong reopened text passed")
	}
	r.SetFontSize(11)
	if err := checkDirectSizeGetter(r, 10.5); err == nil {
		t.Error("changed reopened getter passed")
	}
	if _, err := checkDirectSizeReopened(nil, "Size sample"); err == nil {
		t.Error("missing reopened document passed")
	}
}
