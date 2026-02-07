// Package chart provides DrawingML chart types.
package chart

import "encoding/xml"

// Namespaces used in DrawingML chart parts.
const (
	NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
)

// ChartSpace represents the root chart part.
type ChartSpace struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chartSpace"`
	Chart   *Chart   `xml:"chart,omitempty"`
}

// Chart represents a chart container.
type Chart struct {
	PlotArea *PlotArea `xml:"plotArea,omitempty"`
	Title    *Title    `xml:"title,omitempty"`
	Legend   *Legend   `xml:"legend,omitempty"`
}

// PlotArea represents the plot area.
type PlotArea struct {
	Layout  *Layout       `xml:"layout,omitempty"`
	Content []interface{} `xml:",any"`
}

// Layout represents chart layout.
type Layout struct{}

// Title represents a chart title.
type Title struct{}

// Legend represents chart legend.
type Legend struct{}
