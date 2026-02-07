package wml

import "encoding/xml"

// Drawing represents a drawing element containing inline or anchored graphics.
type Drawing struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main drawing"`
	Inner   string   `xml:",innerxml"`
}
