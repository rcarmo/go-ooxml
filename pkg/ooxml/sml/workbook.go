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
	XMLName      xml.Name      `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	FileVersion  *FileVersion  `xml:"fileVersion,omitempty"`
	WorkbookPr   *WorkbookPr   `xml:"workbookPr,omitempty"`
	BookViews    *BookViews    `xml:"bookViews,omitempty"`
	Sheets       *Sheets       `xml:"sheets"`
	DefinedNames *DefinedNames `xml:"definedNames,omitempty"`
	CalcPr       *CalcPr       `xml:"calcPr,omitempty"`
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
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
	Date1904            *bool  `xml:"date1904,attr,omitempty"`
}

// BookViews represents workbook views.
type BookViews struct {
	WorkbookView []*WorkbookView `xml:"workbookView,omitempty"`
}

// WorkbookView represents a workbook view.
type WorkbookView struct {
	XWindow      int    `xml:"xWindow,attr,omitempty"`
	YWindow      int    `xml:"yWindow,attr,omitempty"`
	WindowWidth  int    `xml:"windowWidth,attr,omitempty"`
	WindowHeight int    `xml:"windowHeight,attr,omitempty"`
	ActiveTab    int    `xml:"activeTab,attr,omitempty"`
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

// CalcPr represents calculation properties.
type CalcPr struct {
	CalcID   int    `xml:"calcId,attr,omitempty"`
	CalcMode string `xml:"calcMode,attr,omitempty"` // auto, manual, autoNoTable
}
