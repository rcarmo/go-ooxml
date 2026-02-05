// Package document provides comment functionality.
package document

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// ID returns the comment ID.
func (c *commentImpl) ID() string {
	return strconv.Itoa(c.comment.ID)
}

// IDInt returns the numeric comment ID.
func (c *commentImpl) IDInt() int {
	return c.comment.ID
}

// Author returns the comment author.
func (c *commentImpl) Author() string {
	return c.comment.Author
}

// SetAuthor sets the comment author.
func (c *commentImpl) SetAuthor(author string) {
	c.comment.Author = author
}

// Initials returns the author's initials.
func (c *commentImpl) Initials() string {
	return c.comment.Initials
}

// SetInitials sets the author's initials.
func (c *commentImpl) SetInitials(initials string) {
	c.comment.Initials = initials
}

// Date returns when the comment was created.
func (c *commentImpl) Date() time.Time {
	if c.comment.Date == "" {
		return time.Time{}
	}
	t, _ := time.Parse(time.RFC3339, c.comment.Date)
	return t
}

// Text returns the comment text.
func (c *commentImpl) Text() string {
	var text string
	for _, p := range c.comment.Content {
		for _, elem := range p.Content {
			if r, ok := elem.(*wml.R); ok {
				for _, re := range r.Content {
					if t, ok := re.(*wml.T); ok {
						text += t.Text
					}
				}
			}
		}
	}
	return text
}

// AnchoredText returns the text this comment is attached to.
func (c *commentImpl) AnchoredText() string {
	if c == nil || c.doc == nil {
		return ""
	}
	return c.doc.CommentedText(c.comment.ID)
}

// Replies returns replies to this comment.
func (c *commentImpl) Replies() []Comment {
	if c == nil || c.doc == nil || c.doc.commentsExtended == nil {
		return nil
	}
	if c.paraID == "" {
		c.paraID = c.doc.commentParaID(c.comment.ID)
	}
	if c.paraID == "" {
		return nil
	}
	var replies []Comment
	for _, ex := range c.doc.commentsExtended.CommentEx {
		if ex.ParaIDParent == c.paraID {
			if reply := c.doc.commentByParaID(ex.ParaID); reply != nil {
				replies = append(replies, reply)
			}
		}
	}
	return replies
}

// AddReply adds a reply to this comment.
func (c *commentImpl) AddReply(text, author string) (Comment, error) {
	if c == nil || c.doc == nil || c.comment == nil {
		return nil, ErrInvalidIndex
	}
	parentID := c.paraID
	if parentID == "" {
		parentID = c.doc.commentParaID(c.comment.ID)
		c.paraID = parentID
	}
	if parentID == "" {
		return nil, utils.NewValidationError("commentReply", "missing parent paraId", nil)
	}
	reply := c.doc.AddComment(text, author)
	c.doc.ensureCommentsExtended()
	for _, ex := range c.doc.commentsExtended.CommentEx {
		if ex.ParaID == reply.paraID {
			ex.ParaIDParent = parentID
			break
		}
	}
	return reply, nil
}

// SetText sets the comment text.
func (c *commentImpl) SetText(text string) {
	c.comment.Content = []*wml.P{
		{
			Content: []interface{}{
				&wml.R{
					Content: []interface{}{
						wml.NewT(text),
					},
				},
			},
		},
	}
}

// =============================================================================
// Document comment methods
// =============================================================================

// Comments returns the document comments manager.
func (d *documentImpl) Comments() Comments {
	return &commentsImpl{doc: d}
}

// AllComments returns all comments in the document (legacy helper).
func (d *documentImpl) AllComments() []Comment {
	if d.comments == nil {
		return nil
	}

	result := make([]Comment, len(d.comments.Comment))
	for i, c := range d.comments.Comment {
		paraID := ""
		if d.commentsExtended != nil && i < len(d.commentsExtended.CommentEx) {
			paraID = d.commentsExtended.CommentEx[i].ParaID
		}
		result[i] = &commentImpl{doc: d, comment: c, paraID: paraID}
	}
	return result
}

// All returns all comments.
func (c *commentsImpl) All() []Comment {
	if c == nil || c.doc == nil {
		return nil
	}
	comments := c.doc.AllComments()
	result := make([]Comment, len(comments))
	for i, comment := range comments {
		result[i] = comment
	}
	return result
}

// ByID returns a comment by ID.
func (c *commentsImpl) ByID(id string) (Comment, error) {
	if c == nil || c.doc == nil {
		return nil, ErrInvalidIndex
	}
	comment := c.doc.CommentByID(id)
	if comment == nil {
		return nil, ErrInvalidIndex
	}
	return comment, nil
}

// Add adds a new comment and anchors it to text if provided.
func (c *commentsImpl) Add(text, author string, anchorText string) (Comment, error) {
	if c == nil || c.doc == nil {
		return nil, ErrInvalidIndex
	}
	comment := c.doc.AddComment(text, author)
	if anchorText != "" {
		if para, start, end := c.doc.findAnchorRuns(anchorText); para != nil {
			if err := para.AnchorComment(comment.IDInt(), start, end); err != nil {
				return nil, err
			}
		}
	}
	return comment, nil
}

func (d *documentImpl) findAnchorRuns(anchorText string) (*paragraphImpl, int, int) {
	if d == nil || d.document == nil || d.document.Body == nil {
		return nil, 0, 0
	}
	for i, elem := range d.document.Body.Content {
		p, ok := elem.(*wml.P)
		if !ok {
			continue
		}
		para := &paragraphImpl{doc: d, p: p, index: i}
		runs := para.Runs()
		for runIndex, run := range runs {
			if strings.Contains(run.Text(), anchorText) {
				return para, runIndex, runIndex
			}
		}
	}
	return nil, 0, 0
}

// Delete removes a comment by ID.
func (c *commentsImpl) Delete(id string) error {
	if c == nil || c.doc == nil {
		return ErrInvalidIndex
	}
	return c.doc.DeleteComment(id)
}

// AddComment adds a new comment to the document.
// The comment is not anchored to any text until you call AnchorComment.
func (d *documentImpl) AddComment(text, author string) *commentImpl {
	if d.comments == nil {
		d.comments = &wml.Comments{}
	}
	
	now := time.Now()
	id := d.nextCommentID
	d.nextCommentID++
	
	comment := &wml.Comment{
		ID:       id,
		Author:   author,
		Date:     now.Format(time.RFC3339),
		Initials: initials(author),
		Content: []*wml.P{
			{
				Content: []interface{}{
					&wml.R{
						Content: []interface{}{
							wml.NewT(text),
						},
					},
				},
			},
		},
	}

	d.comments.Comment = append(d.comments.Comment, comment)
	paraID := d.ensureCommentParaID(id, "")
	return &commentImpl{doc: d, comment: comment, paraID: paraID}
}

// AnchorComment anchors a comment to text within a paragraph.
// startRun and endRun are the indices of runs that mark the anchor range.
func (p *paragraphImpl) AnchorComment(commentID int, startRun, endRun int) error {
	// Add comment range start before startRun
	rangeStart := &wml.CommentRangeStart{ID: commentID}
	rangeEnd := &wml.CommentRangeEnd{ID: commentID}
	ref := &wml.R{
		Content: []interface{}{
			&wml.CommentReference{ID: commentID},
		},
	}
	
	// Build new content with comment markers
	newContent := make([]interface{}, 0, len(p.p.Content)+3)
	
	for i, elem := range p.p.Content {
		if i == startRun {
			newContent = append(newContent, rangeStart)
		}
		newContent = append(newContent, elem)
		if i == endRun {
			newContent = append(newContent, rangeEnd, ref)
		}
	}
	
	// Handle case where markers go at the end
	if startRun >= len(p.p.Content) {
		newContent = append(newContent, rangeStart)
	}
	if endRun >= len(p.p.Content) {
		newContent = append(newContent, rangeEnd, ref)
	}
	
	p.p.Content = newContent
	return nil
}

// CommentedText returns the text that a comment is attached to.
func (d *documentImpl) CommentedText(commentID int) string {
	var collecting bool
	var text string
	
	for _, elem := range d.document.Body.Content {
		if p, ok := elem.(*wml.P); ok {
			for _, pElem := range p.Content {
				switch v := pElem.(type) {
				case *wml.CommentRangeStart:
					if v.ID == commentID {
						collecting = true
					}
				case *wml.CommentRangeEnd:
					if v.ID == commentID {
						collecting = false
					}
				case *wml.R:
					if collecting {
						for _, rElem := range v.Content {
							if t, ok := rElem.(*wml.T); ok {
								text += t.Text
							}
						}
					}
				}
			}
		}
	}
	
	return text
}

func (d *documentImpl) removeCommentRanges(id int) {
	for _, elem := range d.document.Body.Content {
		if p, ok := elem.(*wml.P); ok {
			newContent := make([]interface{}, 0, len(p.Content))
			for _, pElem := range p.Content {
				switch v := pElem.(type) {
				case *wml.CommentRangeStart:
					if v.ID != id {
						newContent = append(newContent, pElem)
					}
				case *wml.CommentRangeEnd:
					if v.ID != id {
						newContent = append(newContent, pElem)
					}
				case *wml.R:
					// Filter out comment references
					newRContent := make([]interface{}, 0, len(v.Content))
					for _, rElem := range v.Content {
						if ref, ok := rElem.(*wml.CommentReference); ok {
							if ref.ID != id {
								newRContent = append(newRContent, rElem)
							}
						} else {
							newRContent = append(newRContent, rElem)
						}
					}
					if len(newRContent) > 0 {
						v.Content = newRContent
						newContent = append(newContent, v)
					}
				default:
					newContent = append(newContent, pElem)
				}
			}
			p.Content = newContent
		}
	}
}

func (d *documentImpl) ensureCommentsExtended() {
	if d.commentsExtended == nil {
		d.commentsExtended = &wml.CommentsEx{}
	}
}

func (d *documentImpl) ensureCommentParaID(commentID int, parentParaID string) string {
	if d == nil {
		return ""
	}
	d.ensureCommentsExtended()
	if existing := d.commentParaID(commentID); existing != "" {
		return existing
	}
	if len(d.commentsExtended.CommentEx) > len(d.comments.Comment) {
		d.commentsExtended.CommentEx = d.commentsExtended.CommentEx[:len(d.comments.Comment)]
	}
	paraID := formatParaID(d.nextCommentParaID)
	d.nextCommentParaID++
	ex := &wml.CommentEx{
		ParaID:       paraID,
		ParaIDParent: parentParaID,
	}
	d.commentsExtended.CommentEx = append(d.commentsExtended.CommentEx, ex)
	return paraID
}

func (d *documentImpl) commentParaID(commentID int) string {
	if d == nil || d.commentsExtended == nil || d.comments == nil {
		return ""
	}
	if len(d.commentsExtended.CommentEx) == 1 && len(d.comments.Comment) == 1 {
		return d.commentsExtended.CommentEx[0].ParaID
	}
	for i, c := range d.comments.Comment {
		if c.ID == commentID {
			if i < len(d.commentsExtended.CommentEx) {
				return d.commentsExtended.CommentEx[i].ParaID
			}
			break
		}
	}
	return ""
}

func (d *documentImpl) commentByParaID(paraID string) Comment {
	if d == nil || d.commentsExtended == nil || paraID == "" {
		return nil
	}
	for i, ex := range d.commentsExtended.CommentEx {
		if ex.ParaID == paraID {
			if d.comments == nil || i >= len(d.comments.Comment) {
				return nil
			}
			return &commentImpl{doc: d, comment: d.comments.Comment[i], paraID: paraID}
		}
	}
	return nil
}

func formatParaID(id int) string {
	if id < 0 {
		id = 0
	}
	return fmt.Sprintf("%08X", id)
}

// initials extracts initials from a name.
func initials(name string) string {
	if name == "" {
		return ""
	}
	
	var init string
	words := splitName(name)
	for _, w := range words {
		if len(w) > 0 {
			init += string(w[0])
		}
	}
	return init
}

func splitName(name string) []string {
	var words []string
	var word string
	for _, c := range name {
		if c == ' ' || c == '.' || c == '-' {
			if word != "" {
				words = append(words, word)
				word = ""
			}
		} else {
			word += string(c)
		}
	}
	if word != "" {
		words = append(words, word)
	}
	return words
}

// =============================================================================
// Comments marshaling
// =============================================================================

func (d *documentImpl) saveComments() error {
	if d.comments == nil || len(d.comments.Comment) == 0 {
		return nil
	}
	
	data, err := utils.MarshalXMLWithHeader(d.comments)
	if err != nil {
		return err
	}
	
	_, err = d.pkg.AddPart("word/comments.xml", packaging.ContentTypeComments, data)
	if err != nil {
		return err
	}
	d.pkg.AddRelationship(packaging.WordDocumentPath, "comments.xml", packaging.RelTypeComments)
	
	return nil
}
