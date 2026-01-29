package sml

import "encoding/xml"

// Table represents an Excel table.
type Table struct {
	XMLName         xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main table"`
	ID              int            `xml:"id,attr"`
	Name            string         `xml:"name,attr,omitempty"`
	DisplayName     string         `xml:"displayName,attr"`
	Ref             string         `xml:"ref,attr"`
	TotalsRowShown  *bool          `xml:"totalsRowShown,attr,omitempty"`
	TotalsRowCount  int            `xml:"totalsRowCount,attr,omitempty"`
	HeaderRowCount  int            `xml:"headerRowCount,attr,omitempty"`
	AutoFilter      *AutoFilter    `xml:"autoFilter,omitempty"`
	TableColumns    *TableColumns  `xml:"tableColumns,omitempty"`
	TableStyleInfo  *TableStyleInfo `xml:"tableStyleInfo,omitempty"`
}

// AutoFilter represents an auto filter.
type AutoFilter struct {
	Ref          string         `xml:"ref,attr"`
	FilterColumn []*FilterColumn `xml:"filterColumn,omitempty"`
}

// FilterColumn represents a filter on a column.
type FilterColumn struct {
	ColID       int           `xml:"colId,attr"`
	HiddenButton *bool        `xml:"hiddenButton,attr,omitempty"`
	Filters     *Filters      `xml:"filters,omitempty"`
	CustomFilters *CustomFilters `xml:"customFilters,omitempty"`
}

// Filters represents filter values.
type Filters struct {
	Blank  *bool    `xml:"blank,attr,omitempty"`
	Filter []*Filter `xml:"filter,omitempty"`
}

// Filter represents a single filter value.
type Filter struct {
	Val string `xml:"val,attr"`
}

// CustomFilters represents custom filter criteria.
type CustomFilters struct {
	And          *bool          `xml:"and,attr,omitempty"`
	CustomFilter []*CustomFilter `xml:"customFilter,omitempty"`
}

// CustomFilter represents a custom filter criterion.
type CustomFilter struct {
	Operator string `xml:"operator,attr,omitempty"` // equal, lessThan, lessThanOrEqual, etc.
	Val      string `xml:"val,attr"`
}

// TableColumns is a collection of table columns.
type TableColumns struct {
	Count       int            `xml:"count,attr,omitempty"`
	TableColumn []*TableColumn `xml:"tableColumn,omitempty"`
}

// TableColumn represents a table column.
type TableColumn struct {
	ID                  int    `xml:"id,attr"`
	Name                string `xml:"name,attr"`
	UniqueName          string `xml:"uniqueName,attr,omitempty"`
	TotalsRowLabel      string `xml:"totalsRowLabel,attr,omitempty"`
	TotalsRowFunction   string `xml:"totalsRowFunction,attr,omitempty"`
	TotalsRowFormula    string `xml:"totalsRowFormula,attr,omitempty"`
	DataDxfID           *int   `xml:"dataDxfId,attr,omitempty"`
	HeaderRowDxfID      *int   `xml:"headerRowDxfId,attr,omitempty"`
	TotalsRowDxfID      *int   `xml:"totalsRowDxfId,attr,omitempty"`
}

// TableStyleInfo represents table style information.
type TableStyleInfo struct {
	Name              string `xml:"name,attr,omitempty"`
	ShowFirstColumn   *bool  `xml:"showFirstColumn,attr,omitempty"`
	ShowLastColumn    *bool  `xml:"showLastColumn,attr,omitempty"`
	ShowRowStripes    *bool  `xml:"showRowStripes,attr,omitempty"`
	ShowColumnStripes *bool  `xml:"showColumnStripes,attr,omitempty"`
}

// Common totals row functions.
const (
	TotalsRowFunctionAverage  = "average"
	TotalsRowFunctionCount    = "count"
	TotalsRowFunctionCountNums = "countNums"
	TotalsRowFunctionMax      = "max"
	TotalsRowFunctionMin      = "min"
	TotalsRowFunctionStdDev   = "stdDev"
	TotalsRowFunctionSum      = "sum"
	TotalsRowFunctionVar      = "var"
	TotalsRowFunctionCustom   = "custom"
	TotalsRowFunctionNone     = "none"
)
