package pml

import "encoding/xml"

const (
	// PPTXCommentsNS is the namespace for modern PowerPoint comments.
	PPTXCommentsNS = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
)

// CommentList represents modern PowerPoint comments.
type CommentList struct {
	XMLName xml.Name   `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main cmLst"`
	Comment []*Comment `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main cm,omitempty"`
}

// Comment represents a single comment.
type Comment struct {
	ID       string        `xml:"id,attr"`
	AuthorID string        `xml:"authorId,attr"`
	Created  string        `xml:"created,attr,omitempty"`
	Pos      *CommentPos   `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main pos,omitempty"`
	TxBody   *CommentText  `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main txBody,omitempty"`
}

// CommentPos represents the comment position on a slide.
type CommentPos struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// CommentText represents the comment text body.
type CommentText struct {
	BodyPr   *CommentBodyPr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main bodyPr,omitempty"`
	LstStyle *CommentLstStyle `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lstStyle,omitempty"`
	P        []*CommentParagraph `xml:"http://schemas.openxmlformats.org/drawingml/2006/main p,omitempty"`
}

// CommentBodyPr represents body properties for comment text.
type CommentBodyPr struct{}

// CommentLstStyle represents list style for comment text.
type CommentLstStyle struct{}

// CommentParagraph represents a paragraph in a comment.
type CommentParagraph struct {
	R []*CommentRun `xml:"http://schemas.openxmlformats.org/drawingml/2006/main r,omitempty"`
}

// CommentRun represents a run in comment text.
type CommentRun struct {
	RPr *CommentRunProps `xml:"http://schemas.openxmlformats.org/drawingml/2006/main rPr,omitempty"`
	T   string           `xml:"http://schemas.openxmlformats.org/drawingml/2006/main t"`
}

// CommentRunProps represents run properties.
type CommentRunProps struct {
	Lang string `xml:"lang,attr,omitempty"`
}

// AuthorList represents comment authors.
type AuthorList struct {
	XMLName xml.Name  `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main authorLst"`
	Author  []*Author `xml:"http://schemas.microsoft.com/office/powerpoint/2018/8/main author,omitempty"`
}

// Author represents a comment author.
type Author struct {
	ID         string `xml:"id,attr"`
	Name       string `xml:"name,attr,omitempty"`
	Initials   string `xml:"initials,attr,omitempty"`
	UserID     string `xml:"userId,attr,omitempty"`
	ProviderID string `xml:"providerId,attr,omitempty"`
}
