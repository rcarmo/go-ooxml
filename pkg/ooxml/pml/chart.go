package pml

import "github.com/rcarmo/go-ooxml/pkg/ooxml/dml"

// HasChart returns true if the graphic frame contains a chart reference.
func (g *GraphicFrame) HasChart() bool {
	if g == nil || g.Graphic == nil || g.Graphic.GraphicData == nil {
		return false
	}
	return g.Graphic.GraphicData.Chart != nil || g.Graphic.GraphicData.URI == dml.GraphicDataURIChart
}
