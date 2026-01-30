package spreadsheet

import (
	"encoding/xml"
	"strconv"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// SharedStrings manages the shared strings table.
type SharedStrings struct {
	strings []string
	lookup  map[string]int
}

// SST represents the shared strings table XML structure.
type SST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr,omitempty"`
	UniqueCount int      `xml:"uniqueCount,attr,omitempty"`
	SI          []*SI    `xml:"si,omitempty"`
}

// SI represents a shared string item.
type SI struct {
	T string `xml:"t,omitempty"`
	R []*RT  `xml:"r,omitempty"` // Rich text runs
}

// RT represents a rich text run.
type RT struct {
	T string `xml:"t"`
}

func newSharedStrings() *SharedStrings {
	return &SharedStrings{
		strings: make([]string, 0),
		lookup:  make(map[string]int),
	}
}

// Count returns the number of unique strings.
func (ss *SharedStrings) Count() int {
	return len(ss.strings)
}

// Get returns the string at the given index.
func (ss *SharedStrings) Get(index int) string {
	if index < 0 || index >= len(ss.strings) {
		return ""
	}
	return ss.strings[index]
}

// Add adds a string and returns its index.
func (ss *SharedStrings) Add(s string) int {
	// Check if string already exists
	if idx, exists := ss.lookup[s]; exists {
		return idx
	}

	// Add new string
	idx := len(ss.strings)
	ss.strings = append(ss.strings, s)
	ss.lookup[s] = idx
	return idx
}

// Index returns the index of a string, or -1 if not found.
func (ss *SharedStrings) Index(s string) int {
	if idx, exists := ss.lookup[s]; exists {
		return idx
	}
	return -1
}

// parse parses shared strings XML data.
func (ss *SharedStrings) parse(data []byte) error {
	var sst SST
	if err := utils.UnmarshalXML(data, &sst); err != nil {
		return err
	}

	for _, si := range sst.SI {
		var s string
		if si.T != "" {
			s = si.T
		} else if len(si.R) > 0 {
			// Concatenate rich text runs
			for _, r := range si.R {
				s += r.T
			}
		}
		ss.Add(s)
	}

	return nil
}

// marshal returns the shared strings XML data.
func (ss *SharedStrings) marshal() ([]byte, error) {
	sst := &SST{
		Count:       len(ss.strings),
		UniqueCount: len(ss.strings),
		SI:          make([]*SI, len(ss.strings)),
	}

	for i, s := range ss.strings {
		sst.SI[i] = &SI{T: s}
	}

	return utils.MarshalXMLWithHeader(sst)
}

// String returns a string representation for debugging.
func (ss *SharedStrings) String() string {
	return "SharedStrings[" + strconv.Itoa(len(ss.strings)) + " strings]"
}
