package pml

import "encoding/xml"

// Notes represents a notes slide.
type Notes struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main notes"`
	CSld    *CSld    `xml:"cSld"`
	ClrMapOvr *ClrMapOvr `xml:"clrMapOvr,omitempty"`
}

// NotesMaster represents a notes master.
type NotesMaster struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main notesMaster"`
	CSld    *CSld    `xml:"cSld"`
	ClrMap  *ClrMap  `xml:"clrMap,omitempty"`
	NotesStyle *NotesStyle `xml:"notesStyle,omitempty"`
	ExtLst  *NotesExtLst  `xml:"extLst,omitempty"`
}

// ClrMap represents a color map.
type ClrMap struct {
	Bg1      string `xml:"bg1,attr,omitempty"`
	Bg2      string `xml:"bg2,attr,omitempty"`
	Tx1      string `xml:"tx1,attr,omitempty"`
	Tx2      string `xml:"tx2,attr,omitempty"`
	Accent1  string `xml:"accent1,attr,omitempty"`
	Accent2  string `xml:"accent2,attr,omitempty"`
	Accent3  string `xml:"accent3,attr,omitempty"`
	Accent4  string `xml:"accent4,attr,omitempty"`
	Accent5  string `xml:"accent5,attr,omitempty"`
	Accent6  string `xml:"accent6,attr,omitempty"`
	HLink    string `xml:"hlink,attr,omitempty"`
	FolHLink string `xml:"folHlink,attr,omitempty"`
}

// NotesStyle represents the notes text styles in the notes master.
type NotesStyle struct {
	Lvl1pPr *LvlPPr `xml:"lvl1pPr,omitempty"`
	Lvl2pPr *LvlPPr `xml:"lvl2pPr,omitempty"`
	Lvl3pPr *LvlPPr `xml:"lvl3pPr,omitempty"`
	Lvl4pPr *LvlPPr `xml:"lvl4pPr,omitempty"`
	Lvl5pPr *LvlPPr `xml:"lvl5pPr,omitempty"`
	Lvl6pPr *LvlPPr `xml:"lvl6pPr,omitempty"`
	Lvl7pPr *LvlPPr `xml:"lvl7pPr,omitempty"`
	Lvl8pPr *LvlPPr `xml:"lvl8pPr,omitempty"`
	Lvl9pPr *LvlPPr `xml:"lvl9pPr,omitempty"`
}

// ExtLst represents a notes master extension list.
type NotesExtLst struct {
	Ext []*NotesExt `xml:"ext,omitempty"`
}

// Ext represents an extension entry.
type NotesExt struct {
	URI string `xml:"uri,attr,omitempty"`
	Any string `xml:",innerxml"`
}
