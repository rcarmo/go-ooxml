package sml

import "encoding/xml"

// Worksheet represents a worksheet part.
type Worksheet struct {
	XMLName        xml.Name        `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetPr        *SheetPr        `xml:"sheetPr,omitempty"`
	Dimension      *Dimension      `xml:"dimension,omitempty"`
	SheetViews     *SheetViews     `xml:"sheetViews,omitempty"`
	SheetFormatPr  *SheetFormatPr  `xml:"sheetFormatPr,omitempty"`
	Cols           *Cols           `xml:"cols,omitempty"`
	SheetData      *SheetData      `xml:"sheetData"`
	SheetProtection *SheetProtection `xml:"sheetProtection,omitempty"`
	MergeCells     *MergeCells     `xml:"mergeCells,omitempty"`
	Hyperlinks     *Hyperlinks     `xml:"hyperlinks,omitempty"`
	PageMargins    *PageMargins    `xml:"pageMargins,omitempty"`
	PageSetup      *PageSetup      `xml:"pageSetup,omitempty"`
	TableParts     *TableParts     `xml:"tableParts,omitempty"`
}

// SheetPr represents sheet properties.
type SheetPr struct {
	TabColor        *TabColor        `xml:"tabColor,omitempty"`
	OutlinePr       *OutlinePr       `xml:"outlinePr,omitempty"`
	PageSetUpPr     *PageSetUpPr     `xml:"pageSetUpPr,omitempty"`
}

// TabColor represents the sheet tab color.
type TabColor struct {
	RGB   string `xml:"rgb,attr,omitempty"`
	Theme int    `xml:"theme,attr,omitempty"`
	Tint  float64 `xml:"tint,attr,omitempty"`
}

// OutlinePr represents outline properties.
type OutlinePr struct {
	SummaryBelow *bool `xml:"summaryBelow,attr,omitempty"`
	SummaryRight *bool `xml:"summaryRight,attr,omitempty"`
}

// PageSetUpPr represents page setup properties.
type PageSetUpPr struct {
	FitToPage *bool `xml:"fitToPage,attr,omitempty"`
}

// Dimension represents the used range.
type Dimension struct {
	Ref string `xml:"ref,attr"`
}

// SheetViews is a collection of sheet views.
type SheetViews struct {
	SheetView []*SheetView `xml:"sheetView,omitempty"`
}

// SheetView represents a sheet view.
type SheetView struct {
	TabSelected      *bool      `xml:"tabSelected,attr,omitempty"`
	WorkbookViewID   int        `xml:"workbookViewId,attr"`
	ShowGridLines    *bool      `xml:"showGridLines,attr,omitempty"`
	ShowRowColHeaders *bool     `xml:"showRowColHeaders,attr,omitempty"`
	ZoomScale        int        `xml:"zoomScale,attr,omitempty"`
	Selection        *Selection `xml:"selection,omitempty"`
	Pane             *Pane      `xml:"pane,omitempty"`
}

// Selection represents the current selection.
type Selection struct {
	ActiveCell   string `xml:"activeCell,attr,omitempty"`
	SqRef        string `xml:"sqref,attr,omitempty"`
	Pane         string `xml:"pane,attr,omitempty"`
}

// Pane represents a split pane.
type Pane struct {
	XSplit      float64 `xml:"xSplit,attr,omitempty"`
	YSplit      float64 `xml:"ySplit,attr,omitempty"`
	TopLeftCell string  `xml:"topLeftCell,attr,omitempty"`
	ActivePane  string  `xml:"activePane,attr,omitempty"`
	State       string  `xml:"state,attr,omitempty"` // frozen, split
}

// SheetFormatPr represents sheet format properties.
type SheetFormatPr struct {
	DefaultColWidth  float64 `xml:"defaultColWidth,attr,omitempty"`
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr"`
	OutlineLevelRow  int     `xml:"outlineLevelRow,attr,omitempty"`
	OutlineLevelCol  int     `xml:"outlineLevelCol,attr,omitempty"`
}

// Cols is a collection of column definitions.
type Cols struct {
	Col []*Col `xml:"col,omitempty"`
}

// Col represents column properties.
type Col struct {
	Min         int     `xml:"min,attr"`
	Max         int     `xml:"max,attr"`
	Width       float64 `xml:"width,attr,omitempty"`
	Style       int     `xml:"style,attr,omitempty"`
	Hidden      *bool   `xml:"hidden,attr,omitempty"`
	BestFit     *bool   `xml:"bestFit,attr,omitempty"`
	CustomWidth *bool   `xml:"customWidth,attr,omitempty"`
	Collapsed   *bool   `xml:"collapsed,attr,omitempty"`
	OutlineLevel int    `xml:"outlineLevel,attr,omitempty"`
}

// SheetData contains all cell data.
type SheetData struct {
	Row []*Row `xml:"row,omitempty"`
}

// Row represents a row.
type Row struct {
	R            int     `xml:"r,attr,omitempty"` // Row number (1-based)
	Spans        string  `xml:"spans,attr,omitempty"`
	S            int     `xml:"s,attr,omitempty"` // Style index
	CustomFormat *bool   `xml:"customFormat,attr,omitempty"`
	Ht           float64 `xml:"ht,attr,omitempty"` // Height
	Hidden       *bool   `xml:"hidden,attr,omitempty"`
	CustomHeight *bool   `xml:"customHeight,attr,omitempty"`
	OutlineLevel int     `xml:"outlineLevel,attr,omitempty"`
	Collapsed    *bool   `xml:"collapsed,attr,omitempty"`
	C            []*Cell `xml:"c,omitempty"` // Cells
}

// SheetProtection represents sheet protection settings.
type SheetProtection struct {
	Sheet        *bool  `xml:"sheet,attr,omitempty"`
	Objects      *bool  `xml:"objects,attr,omitempty"`
	Scenarios    *bool  `xml:"scenarios,attr,omitempty"`
	Password     string `xml:"password,attr,omitempty"`
	AlgorithmName string `xml:"algorithmName,attr,omitempty"`
	HashValue    string `xml:"hashValue,attr,omitempty"`
	SaltValue    string `xml:"saltValue,attr,omitempty"`
	SpinCount    int    `xml:"spinCount,attr,omitempty"`
}

// MergeCells is a collection of merged cell ranges.
type MergeCells struct {
	Count     int          `xml:"count,attr,omitempty"`
	MergeCell []*MergeCell `xml:"mergeCell,omitempty"`
}

// MergeCell represents a merged cell range.
type MergeCell struct {
	Ref string `xml:"ref,attr"`
}

// Hyperlinks is a collection of hyperlinks.
type Hyperlinks struct {
	Hyperlink []*Hyperlink `xml:"hyperlink,omitempty"`
}

// Hyperlink represents a hyperlink.
type Hyperlink struct {
	Ref      string `xml:"ref,attr"`
	ID       string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Location string `xml:"location,attr,omitempty"`
	Display  string `xml:"display,attr,omitempty"`
	Tooltip  string `xml:"tooltip,attr,omitempty"`
}

// PageMargins represents page margins.
type PageMargins struct {
	Left   float64 `xml:"left,attr"`
	Right  float64 `xml:"right,attr"`
	Top    float64 `xml:"top,attr"`
	Bottom float64 `xml:"bottom,attr"`
	Header float64 `xml:"header,attr"`
	Footer float64 `xml:"footer,attr"`
}

// PageSetup represents page setup.
type PageSetup struct {
	PaperSize          int    `xml:"paperSize,attr,omitempty"`
	Scale              int    `xml:"scale,attr,omitempty"`
	Orientation        string `xml:"orientation,attr,omitempty"` // portrait, landscape
	FitToWidth         int    `xml:"fitToWidth,attr,omitempty"`
	FitToHeight        int    `xml:"fitToHeight,attr,omitempty"`
	HorizontalDpi      int    `xml:"horizontalDpi,attr,omitempty"`
	VerticalDpi        int    `xml:"verticalDpi,attr,omitempty"`
}

// TableParts is a collection of table part references.
type TableParts struct {
	Count     int          `xml:"count,attr,omitempty"`
	TablePart []*TablePart `xml:"tablePart,omitempty"`
}

// TablePart references a table part.
type TablePart struct {
	ID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}
