package sml

import "encoding/xml"

// Comments represents the comments part.
type Comments struct {
	XMLName     xml.Name     `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main comments"`
	XMLNS_MC    string       `xml:"xmlns:mc,attr,omitempty"`
	MCIgnorable string       `xml:"mc:Ignorable,attr,omitempty"`
	XMLNS_XR    string       `xml:"xmlns:xr,attr,omitempty"`
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
	XRUID    string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2014/revision uid,attr,omitempty"`
	Text     *Text  `xml:"text"`
}

// Text represents the comment text.
type Text struct {
	R []*CommentRun `xml:"r,omitempty"`
	T string        `xml:"t,omitempty"`
}

// CommentRun represents rich text run in a comment.
type CommentRun struct {
	RPr *CommentRunProps `xml:"rPr,omitempty"`
	T   string           `xml:"t"`
}

// CommentRunProps represents comment run properties.
type CommentRunProps struct {
	Sz     *FloatVal `xml:"sz,omitempty"`
	Color  *CommentColor `xml:"color,omitempty"`
	RFont  *CommentFont  `xml:"rFont,omitempty"`
	Family *IntVal   `xml:"family,omitempty"`
	Scheme *StringVal `xml:"scheme,omitempty"`
}

// CommentColor represents a theme color.
type CommentColor struct {
	Theme *int `xml:"theme,attr,omitempty"`
}

// CommentFont represents a font name.
type CommentFont struct {
	Val string `xml:"val,attr,omitempty"`
}

// FloatVal represents a float value.
type FloatVal struct {
	Val float64 `xml:"val,attr,omitempty"`
}

// IntVal represents an int value.
type IntVal struct {
	Val int `xml:"val,attr,omitempty"`
}

// StringVal represents a string value.
type StringVal struct {
	Val string `xml:"val,attr,omitempty"`
}
