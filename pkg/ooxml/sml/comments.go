package sml

import "encoding/xml"

// Comments represents the comments part.
type Comments struct {
	XMLName     xml.Name     `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main comments"`
	Authors     *Authors     `xml:"authors,omitempty"`
	CommentList *CommentList `xml:"commentList,omitempty"`
}

// Authors is a collection of comment authors.
type Authors struct {
	Author []string `xml:"author,omitempty"`
}

// CommentList is a collection of comments.
type CommentList struct {
	Comment []*Comment `xml:"comment,omitempty"`
}

// Comment represents a worksheet comment.
type Comment struct {
	Ref      string `xml:"ref,attr"`
	AuthorID int    `xml:"authorId,attr"`
	ShapeID  string `xml:"shapeId,attr,omitempty"`
	Text     *Text  `xml:"text"`
}

// Text represents the comment text.
type Text struct {
	T string `xml:"t"`
}
