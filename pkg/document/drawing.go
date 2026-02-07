package document

import (
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func (d *documentImpl) initDrawingCounters() {
	d.nextImageID = maxPartCounter(d.pkg, "word/media/image", ".")
	d.nextChartID = maxPartCounter(d.pkg, "word/charts/chart", ".xml")
	d.nextDiagramID = maxPartCounter(d.pkg, "word/diagrams/data", ".xml")
	d.nextDrawingID = maxDrawingID(d.document)
}

func maxPartCounter(pkg *packaging.Package, prefix, suffix string) int {
	if pkg == nil {
		return 1
	}
	maxID := 0
	for _, part := range pkg.Parts() {
		uri := part.URI()
		if !strings.HasPrefix(uri, prefix) {
			continue
		}
		name := strings.TrimPrefix(uri, prefix)
		if suffix != "" && strings.HasSuffix(name, suffix) {
			name = strings.TrimSuffix(name, suffix)
		}
		if name == "" {
			continue
		}
		if dot := strings.Index(name, "."); dot >= 0 {
			name = name[:dot]
		}
		if id, err := strconv.Atoi(name); err == nil && id > maxID {
			maxID = id
		}
	}
	if maxID == 0 {
		return 1
	}
	return maxID + 1
}

func maxDrawingID(doc *wml.Document) int {
	if doc == nil || doc.Body == nil {
		return 1
	}
	maxID := 0
	for _, elem := range doc.Body.Content {
		p, ok := elem.(*wml.P)
		if !ok {
			continue
		}
		for _, runElem := range p.Content {
			run, ok := runElem.(*wml.R)
			if !ok {
				continue
			}
			for _, rElem := range run.Content {
				drawing, ok := rElem.(*wml.Drawing)
				if !ok || drawing.Inner == "" {
					continue
				}
				maxID = maxInt(maxID, parseDocPrID(drawing.Inner))
			}
		}
	}
	if maxID == 0 {
		return 1
	}
	return maxID + 1
}

func parseDocPrID(inner string) int {
	idx := strings.Index(inner, "docPr")
	if idx < 0 {
		return 0
	}
	sub := inner[idx:]
	idAttr := `id="`
	attrIdx := strings.Index(sub, idAttr)
	if attrIdx < 0 {
		return 0
	}
	sub = sub[attrIdx+len(idAttr):]
	end := strings.Index(sub, `"`)
	if end < 0 {
		return 0
	}
	if id, err := strconv.Atoi(sub[:end]); err == nil {
		return id
	}
	return 0
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
