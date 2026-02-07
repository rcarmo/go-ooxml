package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// Elements returns all body elements.
func (b *bodyImpl) Elements() []BodyElement {
	var elements []BodyElement
	for i, elem := range b.body().Content {
		switch v := elem.(type) {
		case *wml.P:
			elements = append(elements, &paragraphImpl{doc: b.doc, p: v, index: i})
		case *wml.Tbl:
			elements = append(elements, &tableImpl{doc: b.doc, tbl: v, index: i})
		}
	}
	return elements
}

// body returns the underlying wml.Body.
func (b *bodyImpl) body() *wml.Body {
	if b.doc.document.Body == nil {
		b.doc.document.Body = &wml.Body{}
	}
	return b.doc.document.Body
}

// Paragraphs returns all paragraphs in the body.
func (b *bodyImpl) Paragraphs() []Paragraph {
	var result []Paragraph
	for i, elem := range b.body().Content {
		if p, ok := elem.(*wml.P); ok {
			result = append(result, &paragraphImpl{doc: b.doc, p: p, index: i})
		}
	}
	return result
}

// ContentControls returns all block-level content controls in the body.
func (b *bodyImpl) ContentControls() []*ContentControl {
	var result []*ContentControl
	for _, elem := range b.body().Content {
		if sdt, ok := elem.(*wml.Sdt); ok {
			result = append(result, &ContentControl{doc: b.doc, sdt: sdt})
		}
	}
	return result
}

// Tables returns all tables in the body.
func (b *bodyImpl) Tables() []Table {
	var result []Table
	for i, elem := range b.body().Content {
		if tbl, ok := elem.(*wml.Tbl); ok {
			result = append(result, &tableImpl{doc: b.doc, tbl: tbl, index: i})
		}
	}
	return result
}

// AddParagraph adds a new paragraph at the end of the body.
func (b *bodyImpl) AddParagraph() Paragraph {
	p := &wml.P{}
	b.body().Content = append(b.body().Content, p)
	return &paragraphImpl{doc: b.doc, p: p, index: len(b.body().Content) - 1}
}

// AddTable adds a new table at the end of the body.
func (b *bodyImpl) AddTable(rows, cols int) Table {
	tbl := &wml.Tbl{
		TblPr: &wml.TblPr{
			TblW: &wml.TblWidth{W: 5000, Type: "pct"}, // 100% width
		},
		TblGrid: &wml.TblGrid{},
	}

	// Add grid columns
	colWidth := int64(9576 / cols) // Approx letter width in twips
	for i := 0; i < cols; i++ {
		tbl.TblGrid.GridCol = append(tbl.TblGrid.GridCol, &wml.GridCol{W: colWidth})
	}

	// Add rows
	for i := 0; i < rows; i++ {
		tr := &wml.Tr{}
		for j := 0; j < cols; j++ {
			tc := &wml.Tc{
				TcPr: &wml.TcPr{
					TcW: &wml.TblWidth{W: colWidth, Type: "dxa"},
				},
				Content: []interface{}{&wml.P{}}, // Each cell must have at least one paragraph
			}
			tr.Tc = append(tr.Tc, tc)
		}
		tbl.Tr = append(tbl.Tr, tr)
	}

	b.body().Content = append(b.body().Content, tbl)
	return &tableImpl{doc: b.doc, tbl: tbl, index: len(b.body().Content) - 1}
}

// AddChart adds a chart drawing in a new paragraph.
func (b *bodyImpl) AddChart(widthEMU, heightEMU int64, title string) (Paragraph, error) {
	p := b.AddParagraph()
	if err := p.AddChart(widthEMU, heightEMU, title); err != nil {
		return nil, err
	}
	return p, nil
}

// AddDiagram adds a diagram drawing in a new paragraph.
func (b *bodyImpl) AddDiagram(widthEMU, heightEMU int64, title string) (Paragraph, error) {
	p := b.AddParagraph()
	if err := p.AddDiagram(widthEMU, heightEMU, title); err != nil {
		return nil, err
	}
	return p, nil
}

// AddPicture adds an image drawing in a new paragraph.
func (b *bodyImpl) AddPicture(imagePath string, widthEMU, heightEMU int64) (Paragraph, error) {
	p := b.AddParagraph()
	if err := p.AddPicture(imagePath, widthEMU, heightEMU); err != nil {
		return nil, err
	}
	return p, nil
}

// InsertParagraphBefore inserts a paragraph before a target element.
func (b *bodyImpl) InsertParagraphBefore(target BodyElement) Paragraph {
	return b.insertParagraphRelative(target, 0)
}

// InsertParagraphAfter inserts a paragraph after a target element.
func (b *bodyImpl) InsertParagraphAfter(target BodyElement) Paragraph {
	return b.insertParagraphRelative(target, 1)
}

func (b *bodyImpl) insertParagraphRelative(target BodyElement, offset int) Paragraph {
	if target == nil {
		return b.AddParagraph()
	}
	index := target.Index()
	if index < 0 {
		index = 0
	}
	return b.InsertParagraphAt(index + offset)
}

// InsertParagraphAt inserts a paragraph at the given index.
func (b *bodyImpl) InsertParagraphAt(index int) Paragraph {
	p := &wml.P{}
	content := b.body().Content

	if index >= len(content) {
		content = append(content, p)
	} else {
		content = append(content[:index+1], content[index:]...)
		content[index] = p
	}
	b.body().Content = content

	return &paragraphImpl{doc: b.doc, p: p, index: index}
}

// ElementCount returns the number of elements in the body.
func (b *bodyImpl) ElementCount() int {
	return len(b.body().Content)
}
