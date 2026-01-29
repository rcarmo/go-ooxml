package wml

import "encoding/xml"

// Comments represents the comments part.
type Comments struct {
	XMLName  xml.Name   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main comments"`
	Comments []*Comment `xml:"comment,omitempty"`
}

// Comment represents a document comment.
type Comment struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main comment"`
	ID       int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author   string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date     string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	Initials string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main initials,attr,omitempty"`
	Content  []interface{} `xml:",any"` // Paragraphs
}

// CommentRangeStart marks the start of a comment range.
type CommentRangeStart struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main commentRangeStart"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}

// CommentRangeEnd marks the end of a comment range.
type CommentRangeEnd struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main commentRangeEnd"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}

// CommentReference is a reference to a comment from within the document.
type CommentReference struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main commentReference"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}
