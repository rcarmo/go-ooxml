package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const unicodeQNameSource = `<r xmlns:名="urn:one"><名:項 名:鍵="值"/><名:項 名:鍵="値"/></r>`
const unicodeQNameChild = `<名:項 名:鍵="值"/>`

func guardUnicodeQNameCase(id string, p *messages.Pickle) error {
	expected := []string{
		"XML with valid Unicode prefix and local name components",
		"the namespace-name fixture is parsed",
		"expanded element and attribute names retain their Unicode identity",
		"source offsets still address the original Unicode element",
	}
	if id != unicodeQNameCaseID || p.Name != "Unicode prefix and local names remain valid" || len(p.Steps) != len(expected) {
		return fmt.Errorf("unexpected Unicode QName canonical case %s %q", id, p.Name)
	}
	for i, text := range expected {
		if p.Steps[i].Text != text {
			return fmt.Errorf("Unicode QName canonical step %d drift", i+1)
		}
	}
	return nil
}

func unicodeQNameSteps(sc *godog.ScenarioContext) {
	var source, snapshot []byte
	var doc *losslessxml.Document
	var parseErr error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, snapshot, doc, parseErr = nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^XML with valid Unicode prefix and local name components$`, func() error {
		source = []byte(unicodeQNameSource)
		snapshot = bytes.Clone(source)
		return nil
	})
	sc.Step(`^the namespace-name fixture is parsed$`, func() error {
		if source == nil {
			return fmt.Errorf("missing Unicode source")
		}
		doc, parseErr = losslessxml.Parse(source)
		return nil
	})
	sc.Step(`^expanded element and attribute names retain their Unicode identity$`, func() error {
		if parseErr != nil || doc == nil {
			return fmt.Errorf("Unicode parse: %v", parseErr)
		}
		elements := doc.Elements()
		if len(elements) != 3 || elements[0].Name() != (xml.Name{Local: "r"}) {
			return fmt.Errorf("unexpected Unicode element count or root")
		}
		for i, value := range []string{"值", "値"} {
			child := elements[i+1]
			if child.Name() != (xml.Name{Space: "urn:one", Local: "項"}) || child.Ordinal() != i+1 {
				return fmt.Errorf("expanded child %d name/ordinal differs", i)
			}
			parent, ok := child.Parent()
			if !ok || parent.Ordinal() != 0 {
				return fmt.Errorf("child %d parent differs", i)
			}
			attrs := child.Attributes()
			if len(attrs) != 1 || attrs[0] != (xml.Attr{Name: xml.Name{Space: "urn:one", Local: "鍵"}, Value: value}) {
				return fmt.Errorf("expanded child %d attribute differs: %+v", i, attrs)
			}
		}
		return nil
	})
	sc.Step(`^source offsets still address the original Unicode element$`, func() error {
		if doc == nil || !bytes.Equal(source, snapshot) {
			return fmt.Errorf("missing document or caller bytes changed")
		}
		first := []byte(unicodeQNameChild)
		start := bytes.Index(snapshot, first)
		if start < 0 || bytes.Index(snapshot[start+1:], first) >= 0 || start <= len([]rune(`<r xmlns:名="urn:one">`)) {
			return fmt.Errorf("Unicode byte geometry not unique/non-ASCII")
		}
		end := start + len(first)
		if start != 23 || end != 47 || len(snapshot) != 75 {
			return fmt.Errorf("independent byte geometry: [%d:%d) source=%d", start, end, len(snapshot))
		}
		elements := doc.Elements()
		gotStart, gotEnd := elements[1].SourceRange()
		if gotStart != start || gotEnd != end || !bytes.Equal(elements[1].Raw(), snapshot[start:end]) {
			return fmt.Errorf("first Unicode source range [%d:%d) differs from [%d:%d)", gotStart, gotEnd, start, end)
		}
		second := []byte(`<名:項 名:鍵="値"/>`)
		secondStart := bytes.Index(snapshot[end:], second)
		if secondStart != 0 {
			return fmt.Errorf("second Unicode element misplaced: %d", secondStart)
		}
		s2, e2 := elements[2].SourceRange()
		if s2 != end || e2 != end+len(second) || !bytes.Equal(elements[2].Raw(), snapshot[s2:e2]) || s2 == gotStart {
			return fmt.Errorf("second Unicode source range [%d:%d) invalid", s2, e2)
		}
		raw := elements[1].Raw()
		raw[0] = '!'
		if !bytes.Equal(elements[1].Raw(), snapshot[start:end]) || !bytes.Equal(source, snapshot) {
			return fmt.Errorf("Raw or source custody lost")
		}
		return nil
	})
}
