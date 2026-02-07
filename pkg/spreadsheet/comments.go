package spreadsheet

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Comment represents a worksheet comment.
type commentImpl struct {
	worksheet *worksheetImpl
	comment   *sml.Comment
	author    string
}

// Reference returns the cell reference this comment is attached to.
func (c *commentImpl) Reference() string {
	if c == nil || c.comment == nil {
		return ""
	}
	return c.comment.Ref
}

// Author returns the comment author name.
func (c *commentImpl) Author() string {
	return c.author
}

// Text returns the comment text.
func (c *commentImpl) Text() string {
	if c == nil || c.comment == nil || c.comment.Text == nil {
		return ""
	}
	return c.comment.Text.T
}

// SetText updates the comment text.
func (c *commentImpl) SetText(text string) {
	if c == nil || c.comment == nil {
		return
	}
	if c.comment.Text == nil {
		c.comment.Text = &sml.Text{}
	}
	c.comment.Text.T = text
}

// SheetComments manages comments for a worksheet.
type SheetComments struct {
	path     string
	relID    string
	vmlPath  string
	vmlRelID string
	comments *sml.Comments
}

func newSheetComments(path, relID string, comments *sml.Comments) *SheetComments {
	if comments == nil {
		comments = &sml.Comments{}
	}
	if comments.Authors == nil {
		comments.Authors = &sml.Authors{}
	}
	if comments.CommentList == nil {
		comments.CommentList = &sml.CommentList{}
	}
	return &SheetComments{
		path:     path,
		relID:    relID,
		comments: comments,
	}
}

func (sc *SheetComments) ensure() {
	if sc.comments == nil {
		sc.comments = &sml.Comments{}
	}
	if sc.comments.Authors == nil {
		sc.comments.Authors = &sml.Authors{}
	}
	if sc.comments.CommentList == nil {
		sc.comments.CommentList = &sml.CommentList{}
	}
}

func (sc *SheetComments) authorIndex(author string) int {
	sc.ensure()
	for i, name := range sc.comments.Authors.Author {
		if name == author {
			return i
		}
	}
	sc.comments.Authors.Author = append(sc.comments.Authors.Author, author)
	return len(sc.comments.Authors.Author) - 1
}

func (ws *worksheetImpl) ensureComments() *SheetComments {
	if ws.comments == nil {
		ws.comments = newSheetComments("", "", nil)
	}
	return ws.comments
}

// Comments returns all comments in the worksheet.
func (ws *worksheetImpl) Comments() []Comment {
	if ws.comments == nil || ws.comments.comments == nil || ws.comments.comments.CommentList == nil {
		return nil
	}
	authors := ws.comments.comments.Authors
	result := make([]Comment, len(ws.comments.comments.CommentList.Comment))
	for i, c := range ws.comments.comments.CommentList.Comment {
		result[i] = &commentImpl{
			worksheet: ws,
			comment:   c,
			author:    commentAuthor(authors, c.AuthorID),
		}
	}
	return result
}

// PageMargins returns the worksheet page margins.
func (ws *worksheetImpl) PageMargins() (PageMargins, bool) {
	if ws == nil || ws.worksheet == nil || ws.worksheet.PageMargins == nil {
		return PageMargins{}, false
	}
	return *ws.worksheet.PageMargins, true
}

// SetPageMargins sets the worksheet page margins.
func (ws *worksheetImpl) SetPageMargins(margins PageMargins) {
	if ws == nil {
		return
	}
	if ws.worksheet == nil {
		ws.worksheet = &sml.Worksheet{}
	}
	ws.worksheet.PageMargins = &margins
}

// Comment returns the comment for a cell if present.
func (c *cellImpl) Comment() (Comment, bool) {
	if c == nil || c.worksheet == nil {
		return nil, false
	}
	comments := c.worksheet.Comments()
	for _, comment := range comments {
		if comment.Reference() == c.Reference() {
			if cmt, ok := comment.(*commentImpl); ok {
				return cmt, true
			}
			return nil, false
		}
	}
	return nil, false
}

// SetComment adds or updates a comment for the cell.
func (c *cellImpl) SetComment(text, author string) error {
	if c == nil || c.worksheet == nil {
		return utils.ErrInvalidIndex
	}
	ws := c.worksheet
	comments := ws.ensureComments()
	comments.ensure()
	for _, existing := range comments.comments.CommentList.Comment {
		if existing.Ref == c.Reference() {
			existing.Text = &sml.Text{T: text}
			existing.AuthorID = comments.authorIndex(author)
			existing.ShapeID = "0"
			return nil
		}
	}
	comment := &sml.Comment{
		Ref:      c.Reference(),
		AuthorID: comments.authorIndex(author),
		ShapeID:  "0",
		Text:     &sml.Text{T: text},
	}
	comments.comments.CommentList.Comment = append(comments.comments.CommentList.Comment, comment)
	return nil
}

func commentAuthor(authors *sml.Authors, index int) string {
	if authors == nil || index < 0 || index >= len(authors.Author) {
		return ""
	}
	return authors.Author[index]
}
