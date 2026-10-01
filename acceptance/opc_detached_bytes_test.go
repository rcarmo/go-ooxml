package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
	"testing"
	"unicode/utf8"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const opcDetachedByteCaseID = "@id-bun-opc-detached-byte-copies"

func opcDetachedRuleLine() int {
	if batch2Candidate() {
		return 28
	}
	return 27
}

func opcDetachedBackgroundLine() int {
	if packageReasonsCandidate() || batch2Candidate() {
		return 35
	}
	return 33
}
func opcDetachedScenarioLine() int {
	if packageReasonsCandidate() || batch2Candidate() {
		return 61
	}
	return 59
}

const opcDetachedMain = `<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>`
const opcDetachedTypesNS = "http://schemas.openxmlformats.org/package/2006/content-types"
const opcDetachedRelsNS = "http://schemas.openxmlformats.org/package/2006/relationships"
const opcDetachedOfficeType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
const opcDetachedMainType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
const opcDetachedRelsType = "application/vnd.openxmlformats-package.relationships+xml"

var opcDetachedBackground = []string{
	"a ZIP contains word/document.xml with UTF-8 XML text " + opcDetachedMain,
	"its content types use namespace " + opcDetachedTypesNS,
	"its content types have these defaults and overrides",
	"_rels/.rels uses namespace " + opcDetachedRelsNS,
	"its root relationship is rId1 of type " + opcDetachedOfficeType + " targeting word/document.xml",
}
var opcDetachedSteps = []string{
	"the package editor opens the base archive bytes",
	"every byte in the caller's original archive array is overwritten with zero",
	"every byte in the array returned by get for word/document.xml is overwritten with zero",
	"a fresh get of word/document.xml contains the UTF-8 text Alpha",
	"serializing the package returns the exact original archive bytes",
}
var opcDetachedTypeRows = [][]string{
	{"kind", "key", "content_type"},
	{"Default", "rels", opcDetachedRelsType},
	{"Default", "xml", "application/xml"},
	{"Override", "/word/document.xml", opcDetachedMainType},
}

func guardOPCDetachedTableValues(rows [][]string) error {
	if len(rows) != len(opcDetachedTypeRows) {
		return fmt.Errorf("OPC detached-byte Background table rows drift")
	}
	for i, row := range rows {
		if len(row) != len(opcDetachedTypeRows[i]) {
			return fmt.Errorf("OPC detached-byte Background row %d width drift", i)
		}
		for j, value := range row {
			if value != opcDetachedTypeRows[i][j] {
				return fmt.Errorf("OPC detached-byte Background cell %d,%d drift", i, j)
			}
		}
	}
	return nil
}

func guardOPCDetachedASTTable(table *messages.DataTable) error {
	if table == nil {
		return fmt.Errorf("missing Background table")
	}
	rows := make([][]string, len(table.Rows))
	for i, row := range table.Rows {
		for _, cell := range row.Cells {
			rows[i] = append(rows[i], cell.Value)
		}
	}
	return guardOPCDetachedTableValues(rows)
}

func guardOPCDetachedPickleTable(table *messages.PickleTable) error {
	if table == nil {
		return fmt.Errorf("missing Pickle table")
	}
	rows := make([][]string, len(table.Rows))
	for i, row := range table.Rows {
		for _, cell := range row.Cells {
			rows[i] = append(rows[i], cell.Value)
		}
	}
	return guardOPCDetachedTableValues(rows)
}

func guardOPCDetachedByteRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "OPC package custody, transactions and save destinations" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("OPC detached-byte feature drift")
	}
	foundRule, foundCase := 0, 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "OPC byte custody, transaction callbacks and safe save destinations" {
			continue
		}
		foundRule++
		if len(child.Rule.Tags) != 0 || int(child.Rule.Location.Line) != opcDetachedRuleLine() || len(child.Rule.Children) == 0 || child.Rule.Children[0].Background == nil {
			return fmt.Errorf("OPC detached-byte Rule structure drift")
		}
		backgrounds := 0
		for _, member := range child.Rule.Children {
			if bg := member.Background; bg != nil {
				backgrounds++
				if int(bg.Location.Line) != opcDetachedBackgroundLine() || len(bg.Steps) != len(opcDetachedBackground) {
					return fmt.Errorf("OPC detached-byte Background structure drift")
				}
				for i, step := range bg.Steps {
					if step.Text != opcDetachedBackground[i] || int(step.Location.Line) != []int{opcDetachedBackgroundLine() + 1, opcDetachedBackgroundLine() + 2, opcDetachedBackgroundLine() + 3, opcDetachedBackgroundLine() + 8, opcDetachedBackgroundLine() + 9}[i] {
						return fmt.Errorf("OPC detached-byte Background step %d drift", i+1)
					}
					if i == 2 {
						if step.DataTable == nil || guardOPCDetachedASTTable(step.DataTable) != nil || step.DocString != nil {
							return fmt.Errorf("OPC detached-byte Background table drift")
						}
					} else if step.DataTable != nil || step.DocString != nil {
						return fmt.Errorf("OPC detached-byte Background gained argument")
					}
				}
			}
			if s := member.Scenario; s != nil && hasScenarioTag(s, opcDetachedByteCaseID) {
				foundCase++
				if int(s.Location.Line) != opcDetachedScenarioLine() || len(s.Tags) != 2 || s.Tags[0].Name != "@profile-opc-byte-custody" || s.Tags[1].Name != opcDetachedByteCaseID || int(s.Tags[0].Location.Line) != opcDetachedScenarioLine()-1 || int(s.Tags[1].Location.Line) != opcDetachedScenarioLine()-1 || s.Name != "Caller and returned byte arrays cannot modify an opened package" || len(s.Examples) != 0 || len(s.Steps) != len(opcDetachedSteps) {
					return fmt.Errorf("OPC detached-byte Scenario structure drift")
				}
				for i, step := range s.Steps {
					if step.Text != opcDetachedSteps[i] || int(step.Location.Line) != opcDetachedScenarioLine()+1+i || step.DataTable != nil || step.DocString != nil {
						return fmt.Errorf("OPC detached-byte Scenario step %d drift", i+1)
					}
				}
			}
		}
		if backgrounds != 1 {
			return fmt.Errorf("OPC detached-byte Background count %d", backgrounds)
		}
	}
	if foundRule != 1 || foundCase != 1 {
		return fmt.Errorf("OPC detached-byte Rule/Scenario count %d/%d", foundRule, foundCase)
	}
	return nil
}

func guardOPCDetachedByteCase(id string, p *messages.Pickle, line int) error {
	if id != opcDetachedByteCaseID || line != opcDetachedScenarioLine() || p.Name != "Caller and returned byte arrays cannot modify an opened package" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(opcDetachedBackground)+len(opcDetachedSteps) {
		return fmt.Errorf("OPC detached-byte Pickle drift")
	}
	if len(p.Tags) != 3 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != "@profile-opc-byte-custody" || p.Tags[2].Name != opcDetachedByteCaseID {
		return fmt.Errorf("OPC detached-byte Pickle tags drift")
	}
	for i, want := range append(append([]string{}, opcDetachedBackground...), opcDetachedSteps...) {
		step := p.Steps[i]
		if step.Text != want {
			return fmt.Errorf("OPC detached-byte Pickle step %d drift", i+1)
		}
		if i == 2 {
			if step.Argument == nil || guardOPCDetachedPickleTable(step.Argument.DataTable) != nil || step.Argument.DocString != nil {
				return fmt.Errorf("OPC detached-byte Pickle table drift")
			}
		} else if step.Argument != nil {
			return fmt.Errorf("OPC detached-byte Pickle gained argument")
		}
	}
	return nil
}

func TestOPCDetachedByteGuardRejectsDrift(t *testing.T) {
	path := packagePreservationFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	if err := guardOPCDetachedByteRule(doc); err != nil {
		t.Fatal(err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == opcDetachedByteCaseID {
				if selected != nil {
					t.Fatal("duplicate detached-byte Pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil {
		t.Fatal("missing detached-byte Pickle")
	}
	if err := guardOPCDetachedByteCase(opcDetachedByteCaseID, selected, opcDetachedScenarioLine()); err != nil {
		t.Fatal(err)
	}
	for i := range selected.Steps {
		t.Run(fmt.Sprint("step", i), func(t *testing.T) {
			clone := *selected
			clone.Steps = append([]*messages.PickleStep(nil), selected.Steps...)
			copyStep := *clone.Steps[i]
			copyStep.Text += " drift"
			clone.Steps[i] = &copyStep
			if guardOPCDetachedByteCase(opcDetachedByteCaseID, &clone, opcDetachedScenarioLine()) == nil {
				t.Fatal("step drift accepted")
			}
		})
	}
	clone := *selected
	clone.Name += " drift"
	if guardOPCDetachedByteCase(opcDetachedByteCaseID, &clone, opcDetachedScenarioLine()) == nil || guardOPCDetachedByteCase(opcPreserveUnrelatedCaseID, selected, opcDetachedScenarioLine()) == nil || guardOPCDetachedByteCase(opcDetachedByteCaseID, selected, opcDetachedScenarioLine()+1) == nil {
		t.Fatal("identity drift accepted")
	}
	clone = *selected
	clone.Steps = append([]*messages.PickleStep(nil), selected.Steps...)
	copyStep := *selected.Steps[2]
	copyArg := *copyStep.Argument
	copyTable := *copyArg.DataTable
	copyTable.Rows = append([]*messages.PickleTableRow(nil), copyTable.Rows...)
	copyRow := *copyTable.Rows[1]
	copyRow.Cells = append([]*messages.PickleTableCell(nil), copyRow.Cells...)
	copyCell := *copyRow.Cells[2]
	copyCell.Value += " drift"
	copyRow.Cells[2] = &copyCell
	copyTable.Rows[1] = &copyRow
	copyArg.DataTable = &copyTable
	copyStep.Argument = &copyArg
	clone.Steps[2] = &copyStep
	if guardOPCDetachedByteCase(opcDetachedByteCaseID, &clone, opcDetachedScenarioLine()) == nil {
		t.Fatal("table drift accepted")
	}
	for _, tc := range []struct{ name, before, after string }{
		{"background text", "its content types use namespace " + opcDetachedTypesNS, "its content types use namespace urn:wrong"},
		{"table header", "| content_type", "| content_typo"},
		{"table value", "/word/document.xml | " + opcDetachedMainType, "/word/document.xml | application/xml"},
		{"rule", "Rule: OPC byte custody, transaction callbacks and safe save destinations", "Rule: OPC bytes"},
		{"profile", "@profile-opc-byte-custody @id-bun-opc-detached-byte-copies", "@profile-opc-error-api @id-bun-opc-detached-byte-copies"},
		{"scenario name", "Scenario: Caller and returned byte arrays cannot modify an opened package", "Scenario: Caller and returned bytes changed"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			original, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if bytes.Count(original, []byte(tc.before)) != 1 {
				t.Fatalf("nonunique mutation %q", tc.before)
			}
			mutated := bytes.Replace(original, []byte(tc.before), []byte(tc.after), 1)
			count := 0
			parsed, err := gherkin.ParseGherkinDocument(bytes.NewReader(mutated), func() string { count++; return fmt.Sprint(count) })
			if err != nil {
				t.Fatal(err)
			}
			if guardOPCDetachedByteRule(parsed) == nil {
				t.Fatal("Background/Rule/profile drift accepted")
			}
		})
	}
}

// Build the exact three-member base using archive/zip, never the package writer.
func opcDetachedArchive() ([]byte, error) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for _, entry := range []struct{ name, content string }{
		{"[Content_Types].xml", `<Types xmlns="` + opcDetachedTypesNS + `"><Default Extension="rels" ContentType="` + opcDetachedRelsType + `"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="` + opcDetachedMainType + `"/></Types>`},
		{"_rels/.rels", `<Relationships xmlns="` + opcDetachedRelsNS + `"><Relationship Id="rId1" Type="` + opcDetachedOfficeType + `" Target="word/document.xml"/></Relationships>`},
		{"word/document.xml", opcDetachedMain},
	} {
		writer, err := z.Create(entry.name)
		if err != nil {
			return nil, err
		}
		if _, err := io.WriteString(writer, entry.content); err != nil {
			return nil, err
		}
	}
	if err := z.Close(); err != nil {
		return nil, err
	}
	return bytes.Clone(b.Bytes()), nil
}

func opcDetachedMembers(archive []byte) (map[string][]byte, error) {
	z, err := zip.NewReader(bytes.NewReader(archive), int64(len(archive)))
	if err != nil {
		return nil, err
	}
	out := map[string][]byte{}
	for _, f := range z.File {
		if _, ok := out[f.Name]; ok {
			return nil, fmt.Errorf("duplicate member %s", f.Name)
		}
		r, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, readErr := io.ReadAll(r)
		closeErr := r.Close()
		if readErr != nil {
			return nil, readErr
		}
		if closeErr != nil {
			return nil, closeErr
		}
		out[f.Name] = data
	}
	return out, nil
}

func opcDetachedXML(data []byte, root xml.Name) ([]xml.Attr, string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	depth, roots := 0, 0
	var attrs []xml.Attr
	var text strings.Builder
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, "", err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			if depth == 0 {
				if roots != 0 || v.Name != root {
					return nil, "", fmt.Errorf("unexpected XML root")
				}
				roots = 1
				attrs = v.Attr
			} else if root.Local == "document" || depth != 1 || v.Name.Space != root.Space {
				return nil, "", fmt.Errorf("unexpected XML child")
			}
			depth++
		case xml.CharData:
			if depth != 1 {
				if strings.TrimSpace(string(v)) != "" {
					return nil, "", fmt.Errorf("text outside root")
				}
				break
			}
			text.Write(v)
		case xml.EndElement:
			if depth == 0 || (depth == 1 && v.Name != root) || (depth == 2 && (v.Name.Space != root.Space || root.Local == "document")) {
				return nil, "", fmt.Errorf("unexpected XML end")
			}
			depth--
		}
	}
	if roots != 1 || depth != 0 {
		return nil, "", fmt.Errorf("incomplete XML")
	}
	return attrs, text.String(), nil
}

func opcDetachedAssertTypes(data []byte) error {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var entries [][3]string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if start, ok := tok.(xml.StartElement); ok {
			if start.Name == (xml.Name{Space: opcDetachedTypesNS, Local: "Types"}) {
				continue
			}
			if start.Name.Space != opcDetachedTypesNS {
				return fmt.Errorf("unexpected type namespace")
			}
			fields := map[string]string{}
			for _, a := range start.Attr {
				if a.Name.Space != "" {
					return fmt.Errorf("unexpected type attribute")
				}
				fields[a.Name.Local] = a.Value
			}
			switch start.Name.Local {
			case "Default":
				if len(fields) != 2 {
					return fmt.Errorf("default width")
				}
				entries = append(entries, [3]string{"Default", fields["Extension"], fields["ContentType"]})
			case "Override":
				if len(fields) != 2 {
					return fmt.Errorf("override width")
				}
				entries = append(entries, [3]string{"Override", fields["PartName"], fields["ContentType"]})
			default:
				return fmt.Errorf("unexpected type entry")
			}
		}
	}
	if len(entries) != 3 {
		return fmt.Errorf("content type entry count")
	}
	for i, e := range entries {
		if e[0] != opcDetachedTypeRows[i+1][0] || e[1] != opcDetachedTypeRows[i+1][1] || e[2] != opcDetachedTypeRows[i+1][2] {
			return fmt.Errorf("content type row %d", i)
		}
	}
	return nil
}

func opcDetachedByteSteps(sc *godog.ScenarioContext) {
	var source, frozen []byte
	var opened *packaging.Preserved
	var members map[string][]byte
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, frozen, opened, members = nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^a ZIP contains word/document.xml with UTF-8 XML text (.*)$`, func(text string) error {
		if text != opcDetachedMain {
			return fmt.Errorf("base XML drift")
		}
		var err error
		source, err = opcDetachedArchive()
		if err != nil {
			return err
		}
		frozen = bytes.Clone(source)
		members, err = opcDetachedMembers(source)
		if err != nil || len(members) != 3 || !bytes.Equal(members["word/document.xml"], []byte(opcDetachedMain)) {
			return fmt.Errorf("base ZIP mismatch: %v", err)
		}
		return nil
	})
	sc.Step(`^its content types use namespace (.*)$`, func(uri string) error {
		if uri != opcDetachedTypesNS {
			return fmt.Errorf("type namespace drift")
		}
		_, _, err := opcDetachedXML(members["[Content_Types].xml"], xml.Name{Space: uri, Local: "Types"})
		return err
	})
	sc.Step(`^its content types have these defaults and overrides$`, func(table *godog.Table) error {
		if len(table.Rows) != len(opcDetachedTypeRows) {
			return fmt.Errorf("type row count")
		}
		for i, row := range table.Rows {
			if len(row.Cells) != 3 {
				return fmt.Errorf("type row width")
			}
			for j, cell := range row.Cells {
				if cell.Value != opcDetachedTypeRows[i][j] {
					return fmt.Errorf("type row drift")
				}
			}
		}
		return opcDetachedAssertTypes(members["[Content_Types].xml"])
	})
	sc.Step(`^_rels/\.rels uses namespace (.*)$`, func(uri string) error {
		if uri != opcDetachedRelsNS {
			return fmt.Errorf("relationship namespace drift")
		}
		_, _, err := opcDetachedXML(members["_rels/.rels"], xml.Name{Space: uri, Local: "Relationships"})
		return err
	})
	sc.Step(`^its root relationship is rId1 of type (.*) targeting word/document\.xml$`, func(typ string) error {
		if typ != opcDetachedOfficeType {
			return fmt.Errorf("relationship type drift")
		}
		dec := xml.NewDecoder(bytes.NewReader(members["_rels/.rels"]))
		count := 0
		for {
			tok, err := dec.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return err
			}
			if s, ok := tok.(xml.StartElement); ok && s.Name.Local == "Relationship" {
				count++
				if s.Name.Space != opcDetachedRelsNS || len(s.Attr) != 3 {
					return fmt.Errorf("root relationship shape")
				}
				got := map[string]string{}
				for _, a := range s.Attr {
					if a.Name.Space != "" {
						return fmt.Errorf("root relationship attr namespace")
					}
					got[a.Name.Local] = a.Value
				}
				if got["Id"] != "rId1" || got["Type"] != typ || got["Target"] != "word/document.xml" {
					return fmt.Errorf("root relationship values")
				}
			}
		}
		if count != 1 {
			return fmt.Errorf("root relationship count %d", count)
		}
		return nil
	})
	sc.Step(`^the package editor opens the base archive bytes$`, func() error {
		if source == nil || !bytes.Equal(source, frozen) {
			return fmt.Errorf("missing base ZIP")
		}
		var err error
		opened, err = packaging.OpenPreserved(source, packaging.Limits{})
		if err != nil || opened == nil {
			return fmt.Errorf("open preserved: %v", err)
		}
		return nil
	})
	sc.Step(`^every byte in the caller's original archive array is overwritten with zero$`, func() error {
		if opened == nil || len(source) == 0 {
			return fmt.Errorf("missing opened archive")
		}
		clear(source)
		for _, b := range source {
			if b != 0 {
				return fmt.Errorf("source not fully zeroed")
			}
		}
		return nil
	})
	sc.Step(`^every byte in the array returned by get for word/document\.xml is overwritten with zero$`, func() error {
		if opened == nil {
			return fmt.Errorf("missing session")
		}
		part, _, err := opened.Part("word/document.xml")
		if err != nil || !bytes.Equal(part, members["word/document.xml"]) {
			return fmt.Errorf("first get mismatch: %v", err)
		}
		clear(part)
		for _, b := range part {
			if b != 0 {
				return fmt.Errorf("part not fully zeroed")
			}
		}
		return nil
	})
	sc.Step(`^a fresh get of word/document\.xml contains the UTF-8 text Alpha$`, func() error {
		if opened == nil {
			return fmt.Errorf("missing session")
		}
		part, _, err := opened.Part("word/document.xml")
		if err != nil || !utf8.Valid(part) || !bytes.Equal(part, members["word/document.xml"]) {
			return fmt.Errorf("fresh get bytes/UTF-8: %v", err)
		}
		if !bytes.HasPrefix(part, []byte(`<?xml version="1.0" encoding="UTF-8"?>`)) {
			return fmt.Errorf("UTF-8 declaration lost")
		}
		attrs, text, err := opcDetachedXML(part, xml.Name{Local: "document"})
		if err != nil || len(attrs) != 0 || text != "Alpha" {
			return fmt.Errorf("fresh XML text: %q %v", text, err)
		}
		return nil
	})
	sc.Step(`^serializing the package returns the exact original archive bytes$`, func() error {
		if opened == nil || len(frozen) == 0 {
			return fmt.Errorf("missing original archive")
		}
		var output bytes.Buffer
		if err := opened.WriteTo(&output); err != nil {
			return err
		}
		if !bytes.Equal(output.Bytes(), frozen) {
			return fmt.Errorf("no-op archive changed after caller/result mutation")
		}
		for _, b := range source {
			if b != 0 {
				return fmt.Errorf("caller source revived")
			}
		}
		return nil
	})
}
