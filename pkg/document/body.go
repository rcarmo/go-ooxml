package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// Body represents the document body.
type Body struct {
	doc *Document
}

// body returns the underlying wml.Body.
func (b *Body) body() *wml.Body {
	if b.doc.document.Body == nil {
		b.doc.document.Body = &wml.Body{}
	}
	return b.doc.document.Body
}

// Paragraphs returns all paragraphs in the body.
func (b *Body) Paragraphs() []*Paragraph {
	var result []*Paragraph
	for i, elem := range b.body().Content {
		if p, ok := elem.(*wml.P); ok {
			result = append(result, &Paragraph{doc: b.doc, p: p, index: i})
		}
	}
	return result
}

// ContentControls returns all block-level content controls in the body.
func (b *Body) ContentControls() []*ContentControl {
	var result []*ContentControl
	for _, elem := range b.body().Content {
		if sdt, ok := elem.(*wml.Sdt); ok {
			result = append(result, &ContentControl{doc: b.doc, sdt: sdt})
		}
	}
	return result
}

// Tables returns all tables in the body.
func (b *Body) Tables() []*Table {
	var result []*Table
	for i, elem := range b.body().Content {
		if tbl, ok := elem.(*wml.Tbl); ok {
			result = append(result, &Table{doc: b.doc, tbl: tbl, index: i})
		}
	}
	return result
}

// AddParagraph adds a new paragraph at the end of the body.
func (b *Body) AddParagraph() *Paragraph {
	p := &wml.P{}
	b.body().Content = append(b.body().Content, p)
	return &Paragraph{doc: b.doc, p: p, index: len(b.body().Content) - 1}
}

// AddTable adds a new table at the end of the body.
func (b *Body) AddTable(rows, cols int) *Table {
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
	return &Table{doc: b.doc, tbl: tbl, index: len(b.body().Content) - 1}
}

// InsertParagraphAt inserts a paragraph at the given index.
func (b *Body) InsertParagraphAt(index int) *Paragraph {
	p := &wml.P{}
	content := b.body().Content

	if index >= len(content) {
		content = append(content, p)
	} else {
		content = append(content[:index+1], content[index:]...)
		content[index] = p
	}
	b.body().Content = content

	return &Paragraph{doc: b.doc, p: p, index: index}
}

// ElementCount returns the number of elements in the body.
func (b *Body) ElementCount() int {
	return len(b.body().Content)
}
