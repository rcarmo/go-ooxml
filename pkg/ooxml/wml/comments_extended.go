package wml

import "encoding/xml"

// CommentsEx represents the commentsExtended part (w15:commentsEx).
type CommentsEx struct {
	XMLName   xml.Name    `xml:"http://schemas.microsoft.com/office/word/2012/wordml commentsEx"`
	CommentEx []*CommentEx `xml:"http://schemas.microsoft.com/office/word/2012/wordml commentEx,omitempty"`
}

// CommentEx represents extended metadata for a comment, including replies.
type CommentEx struct {
	ParaID       string `xml:"http://schemas.microsoft.com/office/word/2012/wordml paraId,attr"`
	ParaIDParent string `xml:"http://schemas.microsoft.com/office/word/2012/wordml paraIdParent,attr,omitempty"`
	Done         *bool  `xml:"http://schemas.microsoft.com/office/word/2012/wordml done,attr,omitempty"`
}
