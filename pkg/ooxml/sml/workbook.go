// Package sml provides SpreadsheetML types for OOXML spreadsheets.
package sml

import "encoding/xml"

// Namespaces used in SpreadsheetML documents.
const (
	NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// Workbook represents the workbook part.
type Workbook struct {
	XMLName          xml.Name          `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	XMLNS_R          string            `xml:"xmlns:r,attr,omitempty"`
	XMLNS_MC         string            `xml:"xmlns:mc,attr,omitempty"`
	MCIgnorable      string            `xml:"mc:Ignorable,attr,omitempty"`
	XMLNS_X15        string            `xml:"xmlns:x15,attr,omitempty"`
	XMLNS_XR         string            `xml:"xmlns:xr,attr,omitempty"`
	XMLNS_XR6        string            `xml:"xmlns:xr6,attr,omitempty"`
	XMLNS_XR10       string            `xml:"xmlns:xr10,attr,omitempty"`
	XMLNS_XR2        string            `xml:"xmlns:xr2,attr,omitempty"`
	FileVersion      *FileVersion      `xml:"fileVersion,omitempty"`
	WorkbookPr       *WorkbookPr       `xml:"workbookPr,omitempty"`
	AlternateContent *AlternateContent `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent,omitempty"`
	RevisionPtr      *RevisionPtr      `xml:"http://schemas.microsoft.com/office/spreadsheetml/2014/revision revisionPtr,omitempty"`
	BookViews        *BookViews        `xml:"bookViews,omitempty"`
	Sheets           *Sheets           `xml:"sheets"`
	DefinedNames     *DefinedNames     `xml:"definedNames,omitempty"`
	CalcPr           *CalcPr           `xml:"calcPr,omitempty"`
	FileRecoveryPr   *FileRecoveryPr   `xml:"fileRecoveryPr,omitempty"`
	ExtLst           *ExtLst           `xml:"extLst,omitempty"`
}

// FileVersion represents workbook file version info.
type FileVersion struct {
	AppName      string `xml:"appName,attr,omitempty"`
	LastEdited   string `xml:"lastEdited,attr,omitempty"`
	LowestEdited string `xml:"lowestEdited,attr,omitempty"`
	RupBuild     string `xml:"rupBuild,attr,omitempty"`
}

// WorkbookPr represents workbook properties.
type WorkbookPr struct {
	DefaultThemeVersion  string `xml:"defaultThemeVersion,attr,omitempty"`
	Date1904             *bool  `xml:"date1904,attr,omitempty"`
	ChartTrackingRefBase *bool  `xml:"chartTrackingRefBase,attr,omitempty"`
}

// BookViews represents workbook views.
type BookViews struct {
	WorkbookView []*WorkbookView `xml:"workbookView,omitempty"`
}

// WorkbookView represents a workbook view.
type WorkbookView struct {
	XWindow      int `xml:"xWindow,attr,omitempty"`
	YWindow      int `xml:"yWindow,attr,omitempty"`
	WindowWidth  int `xml:"windowWidth,attr,omitempty"`
	WindowHeight int `xml:"windowHeight,attr,omitempty"`
	ActiveTab    int `xml:"activeTab,attr,omitempty"`
	XRUID        string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2 uid,attr,omitempty"`
}

// Sheets is a collection of sheet references.
type Sheets struct {
	Sheet []*Sheet `xml:"sheet,omitempty"`
}

// Sheet represents a sheet reference in the workbook.
type Sheet struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	ID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	State   string `xml:"state,attr,omitempty"` // visible, hidden, veryHidden
}

// DefinedNames is a collection of defined names.
type DefinedNames struct {
	DefinedName []*DefinedName `xml:"definedName,omitempty"`
}

// DefinedName represents a defined name.
type DefinedName struct {
	Name         string `xml:"name,attr"`
	LocalSheetID *int   `xml:"localSheetId,attr,omitempty"`
	Hidden       *bool  `xml:"hidden,attr,omitempty"`
	Value        string `xml:",chardata"`
}

// AlternateContent represents markup-compatibility alternate content.
type AlternateContent struct {
	Choice *AlternateContentChoice `xml:"Choice,omitempty"`
}

// AlternateContentChoice represents a choice within alternate content.
type AlternateContentChoice struct {
	Requires string  `xml:"Requires,attr,omitempty"`
	AbsPath  *AbsPath `xml:"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac absPath,omitempty"`
}

// AbsPath represents the absolute path element.
type AbsPath struct {
	URL string `xml:"url,attr,omitempty"`
}

// RevisionPtr represents revision metadata.
type RevisionPtr struct {
	RevIDLastSave        string `xml:"revIDLastSave,attr,omitempty"`
	DocumentID           string `xml:"documentId,attr,omitempty"`
	CoauthVersionLast    string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6 coauthVersionLast,attr,omitempty"`
	CoauthVersionMax     string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6 coauthVersionMax,attr,omitempty"`
	UIDLastSave          string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10 uidLastSave,attr,omitempty"`
}

// FileRecoveryPr represents file recovery settings.
type FileRecoveryPr struct {
	RepairLoad *bool `xml:"repairLoad,attr,omitempty"`
}

// ExtLst represents extension list.
type ExtLst struct {
	Ext []*Ext `xml:"ext,omitempty"`
}

// Ext represents an extension element.
type Ext struct {
	URI string `xml:"uri,attr,omitempty"`
	Any string `xml:",innerxml"`
}

// CalcPr represents calculation properties.
type CalcPr struct {
	CalcID   int    `xml:"calcId,attr,omitempty"`
	CalcMode string `xml:"calcMode,attr,omitempty"` // auto, manual, autoNoTable
}
