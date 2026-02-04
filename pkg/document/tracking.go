// Package document provides track changes functionality.
package document

import (
	"errors"
	"strconv"
	"strings"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// ErrInvalidIndex is returned when an index is out of bounds.
var ErrInvalidIndex = errors.New("invalid index")

// RevisionType represents the type of a tracked change.
type RevisionType int

const (
	// RevisionInsert indicates inserted content.
	RevisionInsert RevisionType = iota
	// RevisionDelete indicates deleted content.
	RevisionDelete
	// RevisionFormat indicates a formatting change.
	RevisionFormat
)

// String returns a string label for the revision type.
func (rt RevisionType) String() string {
	switch rt {
	case RevisionInsert:
		return "insert"
	case RevisionDelete:
		return "delete"
	case RevisionFormat:
		return "format"
	default:
		return "unknown"
	}
}

type revisionImpl struct {
	doc       *documentImpl
	id        int
	revType   RevisionType
	author    string
	date      time.Time
	paragraph *paragraphImpl
	ins       *wml.Ins
	del       *wml.Del
}

// ID returns the revision ID.
func (r *revisionImpl) ID() string {
	return strconv.Itoa(r.id)
}

// IDInt returns the numeric revision ID.
func (r *revisionImpl) IDInt() int {
	return r.id
}

// Type returns the revision type.
func (r *revisionImpl) Type() RevisionType {
	return r.revType
}

// Author returns the author of the change.
func (r *revisionImpl) Author() string {
	return r.author
}

// Date returns when the change was made.
func (r *revisionImpl) Date() time.Time {
	return r.date
}

// Text returns the text content of the revision.
func (r *revisionImpl) Text() string {
	if r.ins != nil {
		return textFromInlineContent(r.ins.Content)
	}
	if r.del != nil {
		var sb strings.Builder
		for _, elem := range r.del.Content {
			if run, ok := elem.(*wml.R); ok {
				for _, runElem := range run.Content {
					if dt, ok := runElem.(*wml.DelText); ok {
						sb.WriteString(dt.Text)
					}
				}
			}
		}
		return sb.String()
	}
	return ""
}

// Location returns the revision location placeholder.
func (r *revisionImpl) Location() RevisionLocation {
	return RevisionLocation{}
}

// Accept accepts this revision, making the change permanent.
func (r *revisionImpl) Accept() error {
	if r.paragraph == nil {
		return nil
	}
	
	switch r.revType {
	case RevisionInsert:
		// Move content from ins to paragraph
		return r.paragraph.acceptInsertion(r.ins)
	case RevisionDelete:
		// Remove the del element and its content
		return r.paragraph.acceptDeletion(r.del)
	}
	return nil
}

// Reject rejects this revision, undoing the change.
func (r *revisionImpl) Reject() error {
	if r.paragraph == nil {
		return nil
	}
	
	switch r.revType {
	case RevisionInsert:
		// Remove the ins element
		return r.paragraph.rejectInsertion(r.ins)
	case RevisionDelete:
		// Convert del content back to normal
		return r.paragraph.rejectDeletion(r.del)
	}
	return nil
}

// =============================================================================
// TrackChanges methods.
// =============================================================================

// AllRevisions returns all tracked changes in the document.
func (d *documentImpl) AllRevisions() []Revision {
	var revisions []Revision
	
	for i, elem := range d.document.Body.Content {
		switch v := elem.(type) {
		case *wml.P:
			para := &paragraphImpl{doc: d, p: v, index: i}
			revisions = append(revisions, para.revisions()...)
		case *wml.Tbl:
			revisions = append(revisions, revisionsFromTable(d, v)...)
		}
	}
	
	return revisions
}

func (d *documentImpl) revisionsByType(revType RevisionType) []Revision {
	var revisions []Revision
	for _, rev := range d.AllRevisions() {
		if rev.Type() == revType {
			revisions = append(revisions, rev)
		}
	}
	return revisions
}

// InsertTrackedText inserts text at the end of a paragraph with tracking.
func (p *paragraphImpl) InsertTrackedText(text string) Run {
	if !p.doc.trackChanges {
		return p.AddRun()
	}
	
	now := time.Now()
	ins := &wml.Ins{
		ID:     p.doc.nextRevID(),
		Author: p.doc.trackAuthor,
		Date:   now.Format(time.RFC3339),
		Content: []interface{}{
			&wml.R{Content: []interface{}{wml.NewT(text)}},
		},
	}
	
	p.p.Content = append(p.p.Content, ins)
	
	// Return a wrapper for the run inside the insertion
	if run, ok := ins.Content[0].(*wml.R); ok {
		return &runImpl{doc: p.doc, r: run}
	}
	return nil
}

// DeleteTrackedText marks text for deletion with tracking.
func (p *paragraphImpl) DeleteTrackedText(runIndex int) error {
	if runIndex < 0 || runIndex >= len(p.p.Content) {
		return ErrInvalidIndex
	}
	
	run, ok := p.p.Content[runIndex].(*wml.R)
	if !ok {
		return ErrInvalidIndex
	}
	
	if !p.doc.trackChanges {
		// Just remove it
		p.p.Content = append(p.p.Content[:runIndex], p.p.Content[runIndex+1:]...)
		return nil
	}
	
	// Convert text to deletion
	now := time.Now()
	del := &wml.Del{
		ID:     p.doc.nextRevID(),
		Author: p.doc.trackAuthor,
		Date:   now.Format(time.RFC3339),
	}
	
	// Convert T elements to DelText
	newRun := &wml.R{RPr: run.RPr}
	for _, elem := range run.Content {
		if t, ok := elem.(*wml.T); ok {
			newRun.Content = append(newRun.Content, &wml.DelText{Text: t.Text})
		}
	}
	del.Content = []interface{}{newRun}
	
	// Replace the run with the deletion
	p.p.Content[runIndex] = del
	return nil
}

// =============================================================================
// Paragraph helper methods for revisions
// =============================================================================

func (p *paragraphImpl) revisions() []Revision {
	var revisions []Revision
	
	for _, elem := range p.p.Content {
		switch v := elem.(type) {
		case *wml.Ins:
			rev := &revisionImpl{
				doc:       p.doc,
				id:        v.ID,
				revType:   RevisionInsert,
				author:    v.Author,
				paragraph: p,
				ins:       v,
			}
			if v.Date != "" {
				rev.date, _ = time.Parse(time.RFC3339, v.Date)
			}
			revisions = append(revisions, rev)
			
		case *wml.Del:
			rev := &revisionImpl{
				doc:       p.doc,
				id:        v.ID,
				revType:   RevisionDelete,
				author:    v.Author,
				paragraph: p,
				del:       v,
			}
			if v.Date != "" {
				rev.date, _ = time.Parse(time.RFC3339, v.Date)
			}
			revisions = append(revisions, rev)
		}
	}
	
	return revisions
}

func revisionsFromTable(doc *documentImpl, tbl *wml.Tbl) []Revision {
	var revisions []Revision
	for _, row := range tbl.Tr {
		for _, cell := range row.Tc {
			for i, elem := range cell.Content {
				if p, ok := elem.(*wml.P); ok {
					para := &paragraphImpl{doc: doc, p: p, index: i}
					revisions = append(revisions, para.revisions()...)
				}
			}
		}
	}
	return revisions
}

func (p *paragraphImpl) acceptInsertion(ins *wml.Ins) error {
	// Find and replace the ins with its content
	for i, elem := range p.p.Content {
		if elem == ins {
			// Remove the ins and insert the runs in its place
			newContent := make([]interface{}, 0, len(p.p.Content)-1+len(ins.Content))
			newContent = append(newContent, p.p.Content[:i]...)
			for _, c := range ins.Content {
				newContent = append(newContent, c)
			}
			newContent = append(newContent, p.p.Content[i+1:]...)
			p.p.Content = newContent
			return nil
		}
	}
	return nil
}

func (p *paragraphImpl) acceptDeletion(del *wml.Del) error {
	// Remove the del element entirely
	for i, elem := range p.p.Content {
		if elem == del {
			p.p.Content = append(p.p.Content[:i], p.p.Content[i+1:]...)
			return nil
		}
	}
	return nil
}

func (p *paragraphImpl) rejectInsertion(ins *wml.Ins) error {
	// Remove the ins element entirely
	for i, elem := range p.p.Content {
		if elem == ins {
			p.p.Content = append(p.p.Content[:i], p.p.Content[i+1:]...)
			return nil
		}
	}
	return nil
}

func (p *paragraphImpl) rejectDeletion(del *wml.Del) error {
	// Convert del back to normal runs with T instead of DelText
	for i, elem := range p.p.Content {
		if elem == del {
			newContent := make([]interface{}, 0, len(p.p.Content)-1+len(del.Content))
			newContent = append(newContent, p.p.Content[:i]...)
			
			for _, c := range del.Content {
				if r, ok := c.(*wml.R); ok {
					newRun := &wml.R{RPr: r.RPr}
					for _, runElem := range r.Content {
						if dt, ok := runElem.(*wml.DelText); ok {
							newRun.Content = append(newRun.Content, wml.NewT(dt.Text))
						}
					}
					newContent = append(newContent, newRun)
				}
			}
			
			newContent = append(newContent, p.p.Content[i+1:]...)
			p.p.Content = newContent
			return nil
		}
	}
	return nil
}
