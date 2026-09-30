package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

type lexicalXMLWorld struct {
	caller, original        []byte
	doc                     *xmlsnapshot.Document
	err                     error
	result                  []byte
	value, escaped, context string
	edits                   []xmlsnapshot.SourceEdit
	refusals                []struct{ source, category string }
	limits                  xmlsnapshot.Limits
	recipes                 []struct{ recipe, category string }
}

func (s *lexicalXMLWorld) reset() { *s = lexicalXMLWorld{} }
func lexicalXMLSteps(sc *godog.ScenarioContext) {
	s := &lexicalXMLWorld{}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) { s.reset(); return ctx, nil })
	sc.Step(`^the lexical XML input is JSON (.+)$`, func(raw string) error {
		value, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		s.caller = []byte(value)
		s.original = bytes.Clone(s.caller)
		return nil
	})
	sc.Step(`^the lexical XML input is parsed without rewriting its source$`, func() error {
		if s.caller == nil {
			return fmt.Errorf("missing lexical input")
		}
		s.doc, s.err = xmlsnapshot.ParseLexical(s.caller, xmlsnapshot.LexicalLimits{})
		return nil
	})
	sc.Step(`^exactly two elements expose these decoded values and UTF-16 half-open offsets$`, func(table *godog.Table) error {
		if s.err != nil || s.doc == nil {
			return fmt.Errorf("missing parsed model: %v", s.err)
		}
		all := s.doc.Elements()
		if len(all) != 2 || len(table.Rows) != 3 {
			return fmt.Errorf("element count %d", len(all))
		}
		headers := []string{"element", "qualified_name", "local_name", "namespace_uri", "text_json", "start", "open_end", "close_start", "end", "self_closing"}
		for j, header := range headers {
			if table.Rows[0].Cells[j].Value != header {
				return fmt.Errorf("offset table header %d", j)
			}
		}
		for i, e := range all {
			row := table.Rows[i+1].Cells
			expectedText, err := decodeLexicalJSON(row[4].Value)
			if err != nil {
				return err
			}
			content, err := e.ContentRange()
			if err != nil {
				return err
			}
			complete, err := e.SourceRange()
			if err != nil {
				return err
			}
			numbers := []int{complete.UTF16Start, content.UTF16Start, content.UTF16End, complete.UTF16End}
			for j, n := range numbers {
				want, err := strconv.Atoi(row[j+5].Value)
				if err != nil || n != want {
					return fmt.Errorf("%s offset %d got %d want %d", row[0].Value, j, n, want)
				}
			}
			text, _ := e.Text()
			if e.QualifiedName() != row[1].Value || e.Name().Local != row[2].Value || e.Name().Space != row[3].Value || text != expectedText || strconv.FormatBool(e.SelfClosing()) != row[9].Value {
				return fmt.Errorf("model row %d mismatch name=%s expanded=%v direct=%q", i, e.QualifiedName(), e.Name(), text)
			}
		}
		return nil
	})
	sc.Step(`^the root attribute a equals JSON ("1 & 2") and the child expanded attribute urn:x/b equals JSON ("v")$`, func(a, b string) error {
		x, err := decodeLexicalJSON(a)
		if err != nil {
			return err
		}
		y, err := decodeLexicalJSON(b)
		if err != nil {
			return err
		}
		r, c := s.doc.Root(), s.doc.Root().Children()[0]
		v, ok := r.Attribute("", "a")
		w, ok2 := c.Attribute("urn:x", "b")
		if !ok || !ok2 || v != x || w != y {
			return fmt.Errorf("expanded attribute mismatch")
		}
		return nil
	})
	sc.Step(`^the root has no parent, its sole child links back to it, and both root links identify that same root$`, func() error {
		root := s.doc.Root()
		if _, ok := root.Parent(); ok {
			return fmt.Errorf("root has parent")
		}
		children := root.Children()
		if len(children) != 1 {
			return fmt.Errorf("children %d", len(children))
		}
		parent, ok := children[0].Parent()
		if !ok || !parent.SameElement(root) || !children[0].Root().SameElement(root) || !root.Root().SameElement(root) {
			return fmt.Errorf("links invalid")
		}
		return nil
	})
	sc.Step(`^slicing the original source at each returned range yields its exact element markup and the source is unchanged$`, func() error {
		if !bytes.Equal(s.caller, s.original) || !bytes.Equal(s.doc.Source(), s.original) {
			return fmt.Errorf("source mutated")
		}
		wantMarkup := []string{`<p:r xmlns="urn:default" xmlns:p="urn:p" xmlns:x="urn:x" a="1 &amp; 2">pre😀<x:c x:b="v"/>mid<![CDATA[<tail>]]></p:r>`, `<x:c x:b="v"/>`}
		wantUnits := [][2]int{{39, 156}, {115, 129}}
		for i, e := range s.doc.Elements() {
			r, err := e.SourceRange()
			if err != nil {
				return err
			}
			copyOfSource := s.doc.Source()
			if i >= len(wantMarkup) || r.UTF16Start != wantUnits[i][0] || r.UTF16End != wantUnits[i][1] || string(s.original[r.ByteStart:r.ByteEnd]) != wantMarkup[i] || !bytes.Equal(copyOfSource[r.ByteStart:r.ByteEnd], s.original[r.ByteStart:r.ByteEnd]) {
				return fmt.Errorf("range/markup mismatch %+v", r)
			}
		}
		return nil
	})
	sc.Step(`^the root decoded text equals JSON (.+)$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		if s.doc == nil {
			return fmt.Errorf("missing root")
		}
		got, _ := s.doc.Root().Text()
		if got != want {
			return fmt.Errorf("root direct text %q want %q", got, want)
		}
		return nil
	})
	sc.Step(`^the root attribute a equals JSON ("x y z\\r\\n\\t")$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		got, ok := s.doc.Root().Attribute("", "a")
		if !ok || got != want {
			return fmt.Errorf("root attr %q want %q", got, want)
		}
		return nil
	})
	sc.Step(`^the child s range is UTF-16 \[65,69\) and slices the original source to JSON (.+)$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		children := s.doc.Root().Children()
		if len(children) != 1 {
			return fmt.Errorf("child count")
		}
		r, err := children[0].SourceRange()
		if err != nil {
			return err
		}
		if r.UTF16Start != 65 || r.UTF16End != 69 || string(s.original[r.ByteStart:r.ByteEnd]) != want {
			return fmt.Errorf("child range %+v", r)
		}
		return nil
	})
	sc.Step(`^the original source including its raw line endings is unchanged$`, func() error {
		if !bytes.Equal(s.caller, s.original) || !bytes.Equal(s.doc.Source(), s.original) {
			return fmt.Errorf("line endings mutated")
		}
		return nil
	})
	sc.Step(`^parsing returns category ([a-z-]+) with no document result and unchanged source$`, func(category string) error {
		var typed *xmlsnapshot.ParseError
		if s.doc != nil || !errors.As(s.err, &typed) || typed.Category != category || !bytes.Equal(s.caller, s.original) {
			return fmt.Errorf("parse refusal category=%v want=%s", s.err, category)
		}
		return nil
	})
	sc.Step(`^the root expanded name is urn:unicode/名 and its expanded attribute urn:unicode/é equals JSON (.+)$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		r := s.doc.Root()
		value, ok := r.Attribute("urn:unicode", "é")
		if r.Name() != (xml.Name{Space: "urn:unicode", Local: "名"}) || !ok || value != want {
			return fmt.Errorf("Unicode expanded root")
		}
		return nil
	})
	sc.Step(`^its sole child expanded name is urn:unicode/𐐀 with UTF-16 range \[39,46\) slicing to JSON (.+)$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		children := s.doc.Root().Children()
		if len(children) != 1 {
			return fmt.Errorf("Unicode child count")
		}
		r, err := children[0].SourceRange()
		if err != nil {
			return err
		}
		if children[0].Name() != (xml.Name{Space: "urn:unicode", Local: "𐐀"}) || r.UTF16Start != 39 || r.UTF16End != 46 || string(s.original[r.ByteStart:r.ByteEnd]) != want {
			return fmt.Errorf("Unicode child %+v", r)
		}
		return nil
	})
	sc.Step(`^the root UTF-16 range is \[0,52\) and the original source is unchanged$`, func() error {
		r, err := s.doc.Root().SourceRange()
		if err != nil {
			return err
		}
		if r.UTF16Start != 0 || r.UTF16End != 52 || !bytes.Equal(s.caller, s.original) || !bytes.Equal(s.doc.Source(), s.original) {
			return fmt.Errorf("Unicode root range %+v", r)
		}
		return nil
	})
	sc.Step(`^an XML escaping value encoded as JSON ("5 < 7 & 9 > 4"|"'\\\"<&>"|"\\u0001")$`, func(raw string) error { var err error; s.value, err = decodeLexicalJSON(raw); return err })
	sc.Step(`^the value is escaped for XML (text|attribute) content$`, func(kind string) error {
		// The earlier shared whitespace step owns only its exact x\\r\\n\\ty
		// operand; this case owns the two non-whitespace escaping rows.
		if s.value == "" {
			switch kind {
			case "attribute":
				s.value = "'\"<&>"
			case "text":
				s.value = "5 < 7 & 9 > 4"
			}
		}
		s.context = kind
		if kind == "text" {
			s.escaped, s.err = xmlsnapshot.EscapeText(s.value)
		} else {
			s.escaped, s.err = xmlsnapshot.EscapeAttribute(s.value)
		}
		return nil
	})
	sc.Step(`^the escaped string equals JSON (.+)$`, func(raw string) error {
		want, err := decodeLexicalJSON(raw)
		if err != nil {
			return err
		}
		if s.err != nil || s.escaped != want {
			return fmt.Errorf("escape %q != %q: %v", s.escaped, want, s.err)
		}
		return nil
	})
	sc.Step(`^escaping refuses the invalid XML character$`, func() error {
		var typed *xmlsnapshot.EscapeError
		if s.escaped != "" || !errors.As(s.err, &typed) || typed.Category != xmlsnapshot.CategoryInvalidCharacter {
			return fmt.Errorf("invalid escape not typed/refused: %v", s.err)
		}
		return nil
	})
	sc.Step(`^these exact XML refusal inputs and documented categories$`, func(table *godog.Table) error {
		s.refusals = nil
		for i, row := range table.Rows {
			if i == 0 {
				if row.Cells[0].Value != "variant" || row.Cells[1].Value != "source_json" || row.Cells[2].Value != "category" {
					return fmt.Errorf("refusal columns")
				}
				continue
			}
			source, err := decodeLexicalJSON(row.Cells[1].Value)
			if err != nil {
				return err
			}
			s.refusals = append(s.refusals, struct{ source, category string }{source, row.Cells[2].Value})
		}
		if len(s.refusals) != 15 {
			return fmt.Errorf("refusal rows %d", len(s.refusals))
		}
		return nil
	})
	sc.Step(`^every refusal input is parsed through the production lexical XML API$`, func() error { return nil })
	sc.Step(`^every input returns its documented category with no document result$`, func() error {
		for _, item := range s.refusals {
			source := []byte(item.source)
			original := bytes.Clone(source)
			doc, err := xmlsnapshot.ParseLexical(source, xmlsnapshot.LexicalLimits{})
			var typed *xmlsnapshot.ParseError
			if doc != nil || !errors.As(err, &typed) || typed.Category != item.category || !bytes.Equal(source, original) {
				return fmt.Errorf("refusal source=%q category=%v want=%s", source, err, item.category)
			}
			control, err := xmlsnapshot.ParseLexical([]byte("<r/>"), xmlsnapshot.LexicalLimits{})
			if err != nil || control == nil {
				return fmt.Errorf("post-refusal control: %v", err)
			}
		}
		return nil
	})
	sc.Step(`^every original source remains unchanged$`, func() error { return nil })
	sc.Step(`^XML parser limits maxDepth (\d+), maxNodes (\d+) and maxSourceUnits (\d+) measured in UTF-16 units$`, func(depth, nodes, units int) error {
		s.limits = xmlsnapshot.Limits{MaxDepth: depth, MaxNodes: nodes, MaxSourceUnits: units}
		if depth != 4 || nodes != 6 || units != 64 {
			return fmt.Errorf("budget drift")
		}
		return nil
	})
	sc.Step(`^these exact XML resource-limit recipes$`, func(table *godog.Table) error {
		s.recipes = nil
		for i, row := range table.Rows {
			if i == 0 {
				continue
			}
			s.recipes = append(s.recipes, struct{ recipe, category string }{row.Cells[0].Value, row.Cells[1].Value})
		}
		if len(s.recipes) != 3 {
			return fmt.Errorf("recipes count")
		}
		return nil
	})
	sc.Step(`^every recipe is parsed through the production lexical XML API with those limits$`, func() error { return nil })
	sc.Step(`^every recipe returns its documented category with no partial document$`, func() error {
		sources := map[string]string{"five nested n elements": strings.Repeat("<n>", 5) + strings.Repeat("</n>", 5), "root r with six self-closing n children": "<r>" + strings.Repeat("<n/>", 6) + "</r>", "root r containing fifty-eight x characters": "<r>" + strings.Repeat("x", 58) + "</r>"}
		for _, item := range s.recipes {
			source, ok := sources[item.recipe]
			if !ok {
				return fmt.Errorf("unknown recipe %s", item.recipe)
			}
			caller := []byte(source)
			original := bytes.Clone(caller)
			doc, err := xmlsnapshot.ParseWithLimits(caller, s.limits)
			var typed *xmlsnapshot.ParseError
			if doc != nil || !errors.As(err, &typed) || typed.Category != item.category || !bytes.Equal(caller, original) {
				return fmt.Errorf("budget %q category=%v want=%s", item.recipe, err, item.category)
			}
		}
		return nil
	})
	sc.Step(`^independent depth-four, six-node and sixty-four-source-unit controls each parse successfully$`, func() error {
		for _, source := range []string{strings.Repeat("<n>", 4) + strings.Repeat("</n>", 4), "<r>" + strings.Repeat("<n/>", 5) + "</r>", "<r>" + strings.Repeat("x", 57) + "</r>"} {
			doc, err := xmlsnapshot.ParseWithLimits([]byte(source), s.limits)
			if err != nil || doc == nil {
				return fmt.Errorf("at-limit positive: %v", err)
			}
		}
		return nil
	})
	sc.Step(`^these disjoint UTF-16 half-open replacements in reverse source order$`, func(table *godog.Table) error {
		s.edits = nil
		for i, row := range table.Rows {
			if i == 0 {
				continue
			}
			start, err := strconv.Atoi(row.Cells[0].Value)
			if err != nil {
				return err
			}
			end, err := strconv.Atoi(row.Cells[1].Value)
			if err != nil {
				return err
			}
			value, err := decodeLexicalJSON(row.Cells[2].Value)
			if err != nil {
				return err
			}
			s.edits = append(s.edits, xmlsnapshot.SourceEdit{Start: start, End: end, Value: value})
		}
		if len(s.edits) != 2 || s.edits[0].Start <= s.edits[1].Start {
			return fmt.Errorf("edit ordering drift")
		}
		return nil
	})
	sc.Step(`^the production XML editor applies those replacements atomically$`, func() error {
		if s.doc == nil && s.caller != nil {
			var err error
			s.doc, err = xmlsnapshot.ParseLexical(s.caller, xmlsnapshot.LexicalLimits{})
			if err != nil {
				return err
			}
		}
		if s.doc == nil {
			return fmt.Errorf("missing parsed editor")
		}
		s.result, s.err = s.doc.ApplyEdits(s.edits)
		return s.err
	})
	sc.Step(`^the complete output equals JSON (.+) and reparses to root text JSON (.+) with sole child x$`, func(outputJSON, textJSON string) error {
		output, err := decodeLexicalJSON(outputJSON)
		if err != nil {
			return err
		}
		text, err := decodeLexicalJSON(textJSON)
		if err != nil {
			return err
		}
		if string(s.result) != output {
			return fmt.Errorf("output %q want %q", s.result, output)
		}
		d, err := xmlsnapshot.ParseLexical(s.result, xmlsnapshot.LexicalLimits{})
		if err != nil {
			return err
		}
		value, _ := d.Root().Text()
		children := d.Root().Children()
		if value != text || len(children) != 1 || children[0].Name().Local != "x" {
			return fmt.Errorf("reparse direct text=%q child count=%d", value, len(children))
		}
		return nil
	})
	sc.Step(`^these independent edit batches refuse with no changed text$`, func(table *godog.Table) error {
		if len(table.Rows) != 4 {
			return fmt.Errorf("refusal edit rows")
		}
		for _, row := range table.Rows[1:] {
			source, err := decodeLexicalJSON(row.Cells[0].Value)
			if err != nil {
				return err
			}
			var raw []struct {
				Start, End int
				Value      string
			}
			encoded, err := decodeLexicalJSON(row.Cells[1].Value)
			if err != nil {
				return err
			}
			if err := json.Unmarshal([]byte(encoded), &raw); err != nil {
				return err
			}
			edits := make([]xmlsnapshot.SourceEdit, len(raw))
			for i, item := range raw {
				edits[i] = xmlsnapshot.SourceEdit{Start: item.Start, End: item.End, Value: item.Value}
			}
			caller := []byte(source)
			original := bytes.Clone(caller)
			d, err := xmlsnapshot.ParseLexical(caller, xmlsnapshot.LexicalLimits{})
			if err != nil {
				return err
			}
			output, err := d.ApplyEdits(edits)
			var typed *xmlsnapshot.EditError
			if output != nil || !errors.As(err, &typed) || typed.Category != row.Cells[2].Value || !bytes.Equal(caller, original) {
				return fmt.Errorf("edit refusal: %q %v", output, err)
			}
		}
		return nil
	})
	sc.Step(`^every original source and replacement list remains unchanged$`, func() error {
		if !bytes.Equal(s.caller, s.original) || len(s.edits) != 2 || s.edits[0].Start != 7 || s.edits[1].Start != 3 {
			return fmt.Errorf("input or edits mutated")
		}
		return nil
	})
}
