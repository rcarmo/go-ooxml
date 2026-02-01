package dml

import "encoding/xml"

// Tbl represents a DrawingML table.
type Tbl struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tbl"`
	TblPr   *TblPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tblPr,omitempty"`
	TblGrid *TblGrid `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tblGrid,omitempty"`
	Tr      []*Tr    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tr,omitempty"`
}

// TblPr represents table properties.
type TblPr struct {
	BandRow  *bool `xml:"bandRow,attr,omitempty"`
	BandCol  *bool `xml:"bandCol,attr,omitempty"`
	FirstRow *bool `xml:"firstRow,attr,omitempty"`
	LastRow  *bool `xml:"lastRow,attr,omitempty"`
	FirstCol *bool `xml:"firstCol,attr,omitempty"`
	LastCol  *bool `xml:"lastCol,attr,omitempty"`
}

// TblGrid represents table grid columns.
type TblGrid struct {
	GridCol []*GridCol `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gridCol,omitempty"`
}

// GridCol represents a grid column.
type GridCol struct {
	W int64 `xml:"w,attr,omitempty"`
}

// Tr represents a table row.
type Tr struct {
	H  int64  `xml:"h,attr,omitempty"`
	Tc []*Tc  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tc,omitempty"`
}

// Tc represents a table cell.
type Tc struct {
	TcPr *TcPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tcPr,omitempty"`
	TxBody *TxBody `xml:"http://schemas.openxmlformats.org/drawingml/2006/main txBody,omitempty"`
}

// TcPr represents table cell properties.
type TcPr struct {
	GridSpan *int `xml:"gridSpan,attr,omitempty"`
	RowSpan  *int `xml:"rowSpan,attr,omitempty"`
}
