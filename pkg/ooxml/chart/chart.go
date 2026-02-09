// Package chart provides DrawingML chart types.
package chart

import (
	"encoding/xml"
	"fmt"
)

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
	Title    *Title    `xml:"title,omitempty"`
	PlotArea *PlotArea `xml:"plotArea,omitempty"`
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

// RawXML allows embedding raw chart XML elements.
type RawXML struct {
	XMLName xml.Name
	Inner   string `xml:",innerxml"`
}

// DefaultChartSpace returns a minimal chart definition with a single series.
func DefaultChartSpace() *ChartSpace {
	const (
		catAxisID = 48650112
		valAxisID = 48672768
	)

	barChart := &RawXML{
		XMLName: xml.Name{Space: NS, Local: "barChart"},
		Inner: fmt.Sprintf(
			`<barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><order val="0"/><tx><v>Series 1</v></tx><cat><strLit><ptCount val="1"/><pt idx="0"><v>Category 1</v></pt></strLit></cat><val><numLit><ptCount val="1"/><pt idx="0"><v>1</v></pt></numLit></val></ser><axId val="%d"/><axId val="%d"/>`,
			catAxisID,
			valAxisID,
		),
	}
	catAx := &RawXML{
		XMLName: xml.Name{Space: NS, Local: "catAx"},
		Inner: fmt.Sprintf(
			`<axId val="%d"/><scaling><orientation val="minMax"/></scaling><axPos val="l"/><majorTickMark val="out"/><minorTickMark val="none"/><tickLblPos val="nextTo"/><crossAx val="%d"/><crosses val="autoZero"/><auto val="1"/><lblAlgn val="ctr"/><lblOffset val="100"/>`,
			catAxisID,
			valAxisID,
		),
	}
	valAx := &RawXML{
		XMLName: xml.Name{Space: NS, Local: "valAx"},
		Inner: fmt.Sprintf(
			`<axId val="%d"/><scaling><orientation val="minMax"/></scaling><axPos val="b"/><majorTickMark val="out"/><minorTickMark val="none"/><tickLblPos val="nextTo"/><crossAx val="%d"/><crosses val="autoZero"/><crossBetween val="between"/>`,
			valAxisID,
			catAxisID,
		),
	}

	return &ChartSpace{
		Chart: &Chart{
			PlotArea: &PlotArea{
				Layout:  &Layout{},
				Content: []interface{}{barChart, catAx, valAx},
			},
			Title:  &Title{},
			Legend: &Legend{},
		},
	}
}
