package acceptance

import (
	"fmt"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// Check reopened property child positions independently of the preservation
// mask, which deliberately removes selected children before comparing bytes.
func formattingChildOrder(doc *losslessxml.Document, selected losslessxml.Element, kind string) error {
	positions := map[string]int{}
	for _, child := range doc.Elements() {
		if parent, ok := child.Parent(); ok && parent == selected && child.Name().Space == pptxA {
			name := child.Name().Local
			if _, exists := positions[name]; exists {
				return fmt.Errorf("duplicate property child %s", name)
			}
			start, _ := child.SourceRange()
			positions[name] = start
		}
	}
	if strings.HasPrefix(kind, "run-") && kind != "run-bold" && kind != "run-inherit" {
		fill, f := positions["solidFill"]
		latin, l := positions["latin"]
		if !f || !l || fill >= latin {
			return fmt.Errorf("run property child order: solidFill before latin required")
		}
	}
	if strings.HasPrefix(kind, "paragraph-") {
		names := []string{"lnSpc", "spcBef", "spcAft"}
		for i, name := range names {
			at, ok := positions[name]
			if !ok {
				return fmt.Errorf("missing paragraph property child %s", name)
			}
			if i > 0 && positions[names[i-1]] >= at {
				return fmt.Errorf("paragraph property child order: %s before %s required", names[i-1], name)
			}
		}
		for name, at := range positions {
			if strings.HasPrefix(name, "bu") || name == "tabLst" || name == "defRPr" {
				if positions["spcAft"] >= at {
					return fmt.Errorf("paragraph spacing must precede %s", name)
				}
			}
		}
	}
	return nil
}

func TestFormattingChildOrderCorruption(t *testing.T) {
	cases := []struct{ kind, xml string }{
		{"run-color", `<a:rPr xmlns:a="` + pptxA + `"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="A1B2C3"/></a:solidFill></a:rPr>`},
		{"paragraph-spacing", `<a:pPr xmlns:a="` + pptxA + `"><a:spcBef/><a:lnSpc/><a:spcAft/><a:buNone/><a:defRPr/></a:pPr>`},
		{"paragraph-spacing", `<a:pPr xmlns:a="` + pptxA + `"><a:buNone/><a:lnSpc/><a:spcBef/><a:spcAft/><a:defRPr/></a:pPr>`},
		{"paragraph-spacing", `<a:pPr xmlns:a="` + pptxA + `"><a:tabLst/><a:lnSpc/><a:spcBef/><a:spcAft/><a:defRPr/></a:pPr>`},
		{"paragraph-spacing", `<a:pPr xmlns:a="` + pptxA + `"><a:defRPr/><a:lnSpc/><a:spcBef/><a:spcAft/></a:pPr>`},
		{"paragraph-spacing", `<a:pPr xmlns:a="` + pptxA + `"><a:lnSpc/><a:spcBef/><a:spcAft/><a:tabLst/><a:defRPr/></a:pPr>`},
	}
	for i, c := range cases {
		doc, err := losslessxml.Parse([]byte(c.xml))
		if err != nil {
			t.Fatal(err)
		}
		if i == len(cases)-1 {
			if err := formattingChildOrder(doc, doc.Elements()[0], c.kind); err != nil {
				t.Fatalf("valid order rejected: %v", err)
			}
		} else if err := formattingChildOrder(doc, doc.Elements()[0], c.kind); err == nil {
			t.Fatalf("corrupt order %d accepted", i)
		}
	}
}
