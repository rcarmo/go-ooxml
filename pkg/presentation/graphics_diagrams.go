package presentation

import (
	"errors"
	"fmt"
	"math"
	"strings"
	"unicode/utf16"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type DiagramNode struct {
	Key  string `json:"key"`
	Text string `json:"text"`
}
type DiagramEdge struct {
	From string `json:"from"`
	To   string `json:"to"`
}
type DiagramOptions struct {
	X          int64  `json:"x"`
	Y          int64  `json:"y"`
	NodeWidth  int64  `json:"nodeWidth"`
	NodeHeight int64  `json:"nodeHeight"`
	Gap        int64  `json:"gap"`
	Direction  string `json:"direction"`
}
type DiagramNodeReceipt struct {
	Key      string          `json:"key"`
	ShapeID  uint32          `json:"shapeId"`
	Geometry PictureGeometry `json:"geometry"`
}
type DiagramEdgeReceipt struct {
	DiagramEdge
	ShapeID uint32            `json:"shapeId"`
	Start   ConnectorEndpoint `json:"start"`
	End     ConnectorEndpoint `json:"end"`
}
type DiagramReceipt struct {
	PartName string               `json:"partName"`
	Nodes    []DiagramNodeReceipt `json:"nodes"`
	Edges    []DiagramEdgeReceipt `json:"edges"`
}

func graphicsDiagramUnsupported(detail string) error {
	return editRefusal("PPTX_DIAGRAM_UNSUPPORTED", detail)
}
func graphicsDiagramError(err error) error {
	var r *packaging.Refusal
	if errors.As(err, &r) && r.Kind != "PPTX_ID_EXHAUSTED" {
		return graphicsDiagramUnsupported(r.Detail)
	}
	return err
}

// AddDiagram plans all editable rectangles/attached connectors before one commit.
func (s *EditSession) AddDiagram(part string, nodes []DiagramNode, edges []DiagramEdge, o DiagramOptions) (DiagramReceipt, error) {
	var none DiagramReceipt
	if len(nodes) < 1 || len(nodes) > 32 || len(edges) > 64 {
		return none, graphicsDiagramUnsupported("diagram list bounds")
	}
	keys := map[string]bool{}
	for _, n := range nodes {
		if strings.TrimSpace(n.Key) == "" || len(utf16.Encode([]rune(n.Key))) > 64 || len(utf16.Encode([]rune(n.Text))) > 4096 || keys[n.Key] {
			return none, graphicsDiagramUnsupported("diagram key/label bounds")
		}
		if _, e := graphicsXMLString(n.Key); e != nil {
			return none, graphicsDiagramUnsupported("invalid key XML")
		}
		if _, e := graphicsXMLString(n.Text); e != nil {
			return none, graphicsDiagramUnsupported("invalid label XML")
		}
		keys[n.Key] = true
	}
	pairs := map[DiagramEdge]bool{}
	for _, edge := range edges {
		if !keys[edge.From] || !keys[edge.To] || edge.From == edge.To || pairs[edge] {
			return none, graphicsDiagramUnsupported("invalid/duplicate diagram edge")
		}
		pairs[edge] = true
	}
	if o.X < 0 || o.X > math.MaxInt32 || o.Y < 0 || o.Y > math.MaxInt32 || o.NodeWidth < 1 || o.NodeWidth > math.MaxInt32 || o.NodeHeight < 1 || o.NodeHeight > math.MaxInt32 || o.Gap < 0 || o.Gap > math.MaxInt32 || o.Direction != "row" && o.Direction != "column" {
		return none, graphicsDiagramUnsupported("diagram layout bounds")
	}
	geometries := make([]PictureGeometry, len(nodes))
	for i := range nodes {
		g := PictureGeometry{o.X, o.Y, o.NodeWidth, o.NodeHeight}
		if o.Direction == "row" {
			g.X += int64(i) * (o.NodeWidth + o.Gap)
		} else {
			g.Y += int64(i) * (o.NodeHeight + o.Gap)
		}
		if g.X+g.Width > math.MaxInt32 || g.Y+g.Height > math.MaxInt32 {
			return none, graphicsDiagramUnsupported("complete diagram rectangle overflow")
		}
		geometries[i] = g
	}
	if !s.slides[part] {
		return none, graphicsDiagramUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, hash, e := s.pkg.Part(part)
	if e != nil {
		return none, e
	}
	next := source
	receipt := DiagramReceipt{PartName: part, Nodes: []DiagramNodeReceipt{}, Edges: []DiagramEdgeReceipt{}}
	nodeIDs := map[string]uint32{}
	for i, n := range nodes {
		id, at, er := graphicsShapeAppend(next)
		if er != nil {
			return none, graphicsDiagramError(er)
		}
		label, _ := graphicsXMLString("Diagram node " + n.Key)
		text := strings.ReplaceAll(strings.ReplaceAll(n.Text, "\r\n", "\n"), "\r", "\n")
		body := ""
		for _, line := range strings.Split(text, "\n") {
			escaped, _ := graphicsXMLString(line)
			body += `<a:p><a:r><a:t xml:space="preserve">` + escaped + `</a:t></a:r></a:p>`
		}
		g := geometries[i]
		shape := fmt.Sprintf(`<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>`, packaging.NSPresentationML, packaging.NSDrawingML, id, label, g.X, g.Y, g.Width, g.Height, body)
		next = append(append(append([]byte{}, next[:at]...), []byte(shape)...), next[at:]...)
		receipt.Nodes = append(receipt.Nodes, DiagramNodeReceipt{n.Key, id, g})
		nodeIDs[n.Key] = id
	}
	colour := "336699"
	width := int64(12700)
	for _, edge := range edges {
		startSite, endSite := 3, 1
		if o.Direction == "column" {
			startSite, endSite = 2, 0
		}
		start, end := ConnectorEndpoint{nodeIDs[edge.From], startSite}, ConnectorEndpoint{nodeIDs[edge.To], endSite}
		label := "Diagram edge " + edge.From + " -> " + edge.To
		var connector ConnectorReceipt
		next, connector, e = graphicsAddConnector(next, part, start, end, ConnectorOptions{Name: &label, Color: &colour, Width: &width})
		if e != nil {
			return none, graphicsDiagramError(e)
		}
		receipt.Edges = append(receipt.Edges, DiagramEdgeReceipt{edge, connector.ShapeID, start, end})
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}); e != nil {
		return none, e
	}
	s.generation++
	return receipt, nil
}
