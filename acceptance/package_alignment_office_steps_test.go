package acceptance

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func batch2OfficeSteps(sc *godog.ScenarioContext, s *batch2State) {
	sc.Step(`^Office relationship fixture (fixture-[0-9a-f]+) for (pptx|xlsx) has its main-part officeDocument relationship prefix renamed from r to link without changing the namespace URI$`, func(id, format string) error { return s.officeSource(id, format, false) })
	sc.Step(`^Office relationship fixture (fixture-[0-9a-f]+) for (pptx|xlsx) has only the main-part r namespace URI changed from officeDocument relationships to package relationships$`, func(id, format string) error { return s.officeSource(id, format, true) })
	sc.Step(`(?s)^the production (pptx|xlsx) editor sets (.+) to JSON (.*) and saves then reopens$`, func(format, target, raw string) error {
		if format != s.format {
			return fmt.Errorf("format drift")
		}
		s.target = target
		if err := json.Unmarshal([]byte(batch2ValueJSON(raw)), &s.value); err != nil {
			return err
		}
		dest := filepath.Join(s.temp, "edited."+format)
		switch format {
		case "pptx":
			session, err := presentation.OpenEditing(s.archive, packaging.Limits{})
			if err != nil {
				return err
			}
			selected, err := session.FindText("ppt/slides/slide1.xml", 2, "Original title")
			if err != nil {
				return err
			}
			if err = session.Replace(selected, s.value); err != nil {
				return err
			}
			if _, err = session.SaveAs(dest); err != nil {
				return err
			}
		case "xlsx":
			session, err := spreadsheet.OpenEditing(s.archive, packaging.Limits{})
			if err != nil {
				return err
			}
			if err = session.SetInlineStringWrap("Sheet", "A1", s.value, true); err != nil {
				return err
			}
			if _, err = session.SaveAs(dest); err != nil {
				return err
			}
		}
		var err error
		s.output, err = os.ReadFile(dest)
		if err != nil {
			return err
		}
		s.savedParts, err = opcZipMembers(s.output)
		if err != nil {
			return err
		}
		if !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("caller archive changed")
		}
		return nil
	})
	sc.Step(`(?s)^the exact value JSON (.*) is read at (.+) and every internal relationship target resolves$`, func(raw, target string) error {
		var expected string
		if err := json.Unmarshal([]byte(batch2ValueJSON(raw)), &expected); err != nil {
			return err
		}
		if target != s.target || expected != s.value {
			return fmt.Errorf("Office target/value drift")
		}
		packageView, err := packaging.OpenPreserved(s.output, packaging.Limits{})
		if err != nil {
			return err
		}
		graph, err := packageView.Graph()
		if err != nil {
			return err
		}
		for _, edge := range graph.Edges {
			if !edge.External && edge.ResolvedPart == "" {
				return fmt.Errorf("unresolved relationship %s", edge.ID)
			}
		}
		if s.format == "pptx" {
			reopened, err := presentation.OpenEditing(s.output, packaging.Limits{})
			if err != nil {
				return err
			}
			if _, err = reopened.FindText("ppt/slides/slide1.xml", 2, expected); err != nil {
				return err
			}
		} else {
			reopened, err := spreadsheet.OpenEditing(s.output, packaging.Limits{})
			if err != nil {
				return err
			}
			sheet := s.savedParts["xl/worksheets/sheet1.xml"]
			doc, err := losslessxml.Parse(sheet)
			if err != nil {
				return err
			}
			found := 0
			for _, element := range doc.Elements() {
				if element.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "c"}) || batch2Attr(element, "r") != "A1" {
					continue
				}
				if batch2Attr(element, "s") != "1" {
					return fmt.Errorf("wrap XF not selected")
				}
				for _, text := range doc.Elements() {
					parent, ok := text.Parent()
					if !ok || parent.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "is"}) {
						continue
					}
					owner, ok := parent.Parent()
					if owner != element || !ok {
						continue
					}
					value, leaf := text.Text()
					if leaf && text.Name().Local == "t" && value == expected {
						found++
					}
				}
			}
			if found != 1 {
				return fmt.Errorf("reopened inline string count %d", found)
			}
			style := s.savedParts["xl/styles.xml"]
			styles, err := losslessxml.Parse(style)
			if err != nil {
				return err
			}
			wrap := false
			for _, e := range styles.Elements() {
				if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "alignment"}) && batch2Attr(e, "wrapText") == "true" {
					wrap = true
				}
			}
			if !wrap {
				return fmt.Errorf("wrap alignment not retained")
			}
			_ = reopened
		}
		return nil
	})
	sc.Step(`^the saved main part retains link:id attributes without r:id attributes and the caller's archive bytes remain unchanged$`, func() error {
		main := s.savedParts[s.officeMain()]
		if !bytes.Contains(main, []byte("link:id=")) || bytes.Contains(main, []byte("r:id=")) || !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("relationship alias/caller custody")
		}
		ids, err := batch2OfficeRelationshipIDs(main)
		if err != nil || !reflect.DeepEqual(ids, s.officeIDs) {
			return fmt.Errorf("saved relationship IDs %v, want original %v: %v", ids, s.officeIDs, err)
		}
		return nil
	})
	sc.Step(`^the production (pptx|xlsx) reader attempts to open the namespace-mismatched document$`, func(format string) error {
		if format != s.format {
			return fmt.Errorf("format drift")
		}
		switch format {
		case "pptx":
			opened, err := presentation.OpenEditing(s.archive, packaging.Limits{})
			s.failure = err
			if opened != nil {
				s.result = opened
			}
		case "xlsx":
			opened, err := spreadsheet.OpenEditing(s.archive, packaging.Limits{})
			s.failure = err
			if opened != nil {
				s.result = opened
			}
		}
		return nil
	})
	sc.Step(`^the reader refuses with structured reason (PPTX_PRESENTATION_INVALID|xlsx-workbook-invalid) and no document result$`, func(reason string) error {
		if s.result != nil {
			return fmt.Errorf("reader delivered invalid document")
		}
		return batch2Refusal(s.failure, reason)
	})
	sc.Step(`^the supplied archive bytes remain unchanged$`, func() error {
		if !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("Office caller changed")
		}
		return nil
	})
}

func (s *batch2State) officeMain() string {
	if s.format == "pptx" {
		return "ppt/presentation.xml"
	}
	return "xl/workbook.xml"
}
func (s *batch2State) officeSource(id, format string, wrong bool) error {
	want := map[string]string{"pptx": "fixture-2aec94471f93c300d56ca4789106a974411085d1588f3424155362a06dd043f3", "xlsx": "fixture-38c2ed936696179d3b2359e9107ad2b8d62d71d69296f8f60bdfe1fe8f7f2439"}
	if id != want[format] {
		return fmt.Errorf("Office fixture ID mismatch")
	}
	s.format = format
	path, err := testutil.LookupFixture(id)
	if err != nil {
		return err
	}
	source, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	p, err := packaging.OpenPreserved(source, packaging.Limits{})
	if err != nil {
		return err
	}
	main, hash, err := p.Part(s.officeMain())
	if err != nil {
		return err
	}
	original := []byte(`xmlns:r="` + packaging.NSDocumentRelationships + `"`)
	if bytes.Count(main, original) != 1 {
		return fmt.Errorf("main relationship declaration not unique")
	}
	if !wrong {
		s.officeIDs, err = batch2OfficeRelationshipIDs(main)
		if err != nil || len(s.officeIDs) == 0 {
			return fmt.Errorf("original relationship IDs: %v", err)
		}
	}
	if wrong {
		main = bytes.Replace(main, original, []byte(`xmlns:r="`+packaging.NSRelationships+`"`), 1)
	} else {
		main = bytes.Replace(main, original, []byte(`xmlns:link="`+packaging.NSDocumentRelationships+`"`), 1)
		main = bytes.ReplaceAll(main, []byte("r:id="), []byte("link:id="))
	}
	if err = p.Replace([]packaging.Replacement{{Part: s.officeMain(), ExpectedSHA256: hash, Data: main}}); err != nil {
		return err
	}
	var b bytes.Buffer
	if err = p.WriteTo(&b); err != nil {
		return err
	}
	s.archive = b.Bytes()
	s.initial = bytes.Clone(s.archive)
	return nil
}

// Compare actual relationship IDs by expanded name across the authored input
// and reopened output, independently of the editor's relationship graph.
func batch2OfficeRelationshipIDs(main []byte) ([]string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(main))
	ids := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return nil, err
		}
		start, ok := token.(xml.StartElement)
		if !ok || (start.Name.Local != "sldId" && start.Name.Local != "sheet") {
			continue
		}
		for _, attr := range start.Attr {
			if attr.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				ids = append(ids, attr.Value)
			}
		}
	}
}

// Gherkin Examples expansion decodes JSON \\n to a literal line feed in the
// interpolated step. Reconstruct only that JSON string escape for decoding.
func batch2ValueJSON(raw string) string { return strings.ReplaceAll(raw, "\n", `\n`) }

func batch2Attr(e losslessxml.Element, key string) string {
	for _, a := range e.Attributes() {
		if a.Name == (xml.Name{Local: key}) {
			return a.Value
		}
	}
	return ""
}
