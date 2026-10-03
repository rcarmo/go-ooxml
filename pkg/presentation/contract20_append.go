package presentation

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// AppendContractTitleSlide appends one bounded title slide to an ordinary
// retained presentation. It changes only the presentation slide list and
// registries, and adds the new slide and its layout relationship registry.
// Existing slide payloads and held title anchors are not consumed.
func (s *EditSession) AppendContractTitleSlide(title, subtitle string) error {
	if s == nil || s.pkg == nil {
		return editRefusal("PPTX_LAYOUT_UNSAFE", "missing session")
	}
	if err := contractCreatorText(title); err != nil {
		return err
	}
	if err := contractCreatorText(subtitle); err != nil {
		return err
	}
	refuse := func(reason string) error { return editRefusal("PPTX_LAYOUT_UNSAFE", reason) }
	g, err := s.pkg.Graph()
	if err != nil {
		return refuse("unreadable support graph")
	}
	partTypes := map[string]string{}
	existing := map[string]bool{}
	for _, part := range g.Parts {
		partTypes[part.Name] = part.ContentType
		existing[strings.ToLower(part.Name)] = true
	}
	presentationPart := packaging.PresentationPath
	if partTypes[presentationPart] != packaging.ContentTypePresentation || len(s.slides) != 1 {
		return refuse("one ordinary source slide required")
	}
	mainBytes, mainHash, err := s.pkg.Part(presentationPart)
	if err != nil {
		return refuse("presentation missing")
	}
	doc, err := losslessxml.Parse(mainBytes)
	if err != nil {
		return refuse("presentation XML invalid")
	}
	p := func(local string) xml.Name { return name(packaging.NSPresentationML, local) }
	if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != p("presentation") {
		return refuse("presentation root")
	}
	var list losslessxml.Element
	listCount := 0
	maxID := uint64(255)
	count := 0
	oldRID := ""
	for _, node := range doc.Elements() {
		if node.Name() == p("sldIdLst") {
			listCount++
			list = node
		}
		if node.Name() != p("sldId") {
			continue
		}
		parent, ok := node.Parent()
		if !ok || parent.Name() != p("sldIdLst") {
			return refuse("slide list topology")
		}
		count++
		var id, rid string
		for _, attr := range node.Attributes() {
			if attr.Name == (xml.Name{Local: "id"}) {
				id = attr.Value
			}
			if attr.Name == name(packaging.NSDocumentRelationships, "id") {
				rid = attr.Value
			}
		}
		number, e := strconv.ParseUint(id, 10, 32)
		if e != nil || number < 256 || rid == "" {
			return refuse("slide identity")
		}
		if number > maxID {
			maxID = number
		}
		oldRID = rid
	}
	if listCount != 1 || count != 1 || maxID == uint64(^uint32(0)) {
		return refuse("one bounded slide list required")
	}
	sourceSlide := ""
	sourceCount := 0
	usedRID := map[string]bool{}
	for _, edge := range g.Edges {
		if edge.Source == presentationPart {
			usedRID[edge.ID] = true
		}
		if edge.Source == presentationPart && edge.Type == packaging.RelTypeSlide {
			if edge.External || edge.ID != oldRID || !s.slides[edge.ResolvedPart] {
				return refuse("source slide identity")
			}
			sourceCount++
			sourceSlide = edge.ResolvedPart
		}
	}
	if sourceCount != 1 || sourceSlide == "" {
		return refuse("one owned source slide required")
	}
	// Resolve title layout through the old slide, not by guessing a numbered path.
	layout := ""
	layoutCount := 0
	for _, edge := range g.Edges {
		if edge.Source == sourceSlide && edge.Type == packaging.RelTypeSlideLayout {
			if edge.External || partTypes[edge.ResolvedPart] != packaging.ContentTypeSlideLayout {
				return refuse("source layout edge")
			}
			layoutCount++
			layout = edge.ResolvedPart
		}
	}
	if layoutCount != 1 {
		return refuse("unique source layout required")
	}
	layoutData, _, err := s.pkg.Part(layout)
	if err != nil {
		return refuse("layout missing")
	}
	layoutDoc, err := losslessxml.Parse(layoutData)
	if err != nil {
		return refuse("layout XML invalid")
	}
	if len(layoutDoc.Elements()) == 0 || layoutDoc.Elements()[0].Name() != p("sldLayout") {
		return refuse("layout root")
	}
	typeTitle := false
	for _, attr := range layoutDoc.Elements()[0].Attributes() {
		if attr.Name == (xml.Name{Local: "type"}) && attr.Value == "title" {
			typeTitle = true
		}
	}
	if !typeTitle {
		return refuse("title layout required")
	}
	titleKind, subtitleIndex, err := contractAppendLayoutPlaceholders(layoutDoc)
	if err != nil {
		return refuse(err.Error())
	}
	master := ""
	masterCount := 0
	for _, edge := range g.Edges {
		if edge.Source == layout && edge.Type == packaging.RelTypeSlideMaster {
			if edge.External || partTypes[edge.ResolvedPart] != packaging.ContentTypeSlideMaster {
				return refuse("layout master edge")
			}
			master = edge.ResolvedPart
			masterCount++
		}
	}
	if masterCount != 1 {
		return refuse("unique layout master required")
	}
	reciprocal := 0
	for _, edge := range g.Edges {
		if edge.Source == master && edge.Type == packaging.RelTypeSlideLayout && edge.ResolvedPart == layout && !edge.External {
			reciprocal++
		}
	}
	if reciprocal != 1 {
		return refuse("nonreciprocal title layout")
	}
	newSlide := ""
	for n := 1; n <= 10000; n++ {
		candidate := fmt.Sprintf("ppt/slides/slide%d.xml", n)
		rels := packaging.RelationshipsPathForPart(candidate)
		if !existing[strings.ToLower(candidate)] && !existing[strings.ToLower(rels)] {
			newSlide = candidate
			break
		}
	}
	if newSlide == "" {
		return refuse("no bounded slide identity")
	}
	rid := ""
	for n := 1; n <= 10000; n++ {
		candidate := fmt.Sprintf("rId%d", n)
		if !usedRID[candidate] {
			rid = candidate
			break
		}
	}
	if rid == "" {
		return refuse("no bounded relationship identity")
	}
	child := losslessxml.NewElement{Name: p("sldId"), Attributes: []xml.Attr{
		{Name: xml.Name{Local: "id"}, Value: strconv.FormatUint(maxID+1, 10)},
		{Name: name(packaging.NSDocumentRelationships, "id"), Value: rid},
	}}
	updated, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: list, Children: []losslessxml.NewElement{child}}})
	if err != nil {
		return refuse("slide list insertion unsafe")
	}
	// Fresh, fixed title-shape model: the source layout supplies formatting and
	// geometry; no source slide XML or relationship is copied or retargeted.
	shape := func(id int, kind, idx, text string) string {
		index := ""
		if idx != "" {
			index = ` idx="` + idx + `"`
		}
		space := ""
		if strings.TrimSpace(text) != text {
			space = ` xml:space="preserve"`
		}
		return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Contract title %d"/><p:cNvSpPr/><p:nvPr><p:ph type="%s"%s/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t%s>%s</a:t></a:r></a:p></p:txBody></p:sp>`, id, id, kind, index, space, xmlEscapeContractText(text))
	}
	payload := []byte(`<p:sld xmlns:p="` + packaging.NSPresentationML + `" xmlns:a="` + packaging.NSDrawingML + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` + shape(2, titleKind, "", title) + shape(3, "subTitle", subtitleIndex, subtitle) + `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`)
	mutation := packaging.GraphMutation{
		Additions:     []packaging.PartAddition{{Name: newSlide, ContentType: packaging.ContentTypeSlide, Data: payload}},
		Replacements:  []packaging.Replacement{{Part: presentationPart, ExpectedSHA256: mainHash, Data: updated}},
		Relationships: []packaging.RelationshipAddition{{Source: presentationPart, ID: rid, Type: packaging.RelTypeSlide, TargetPart: newSlide}, {Source: newSlide, ID: "rId1", Type: packaging.RelTypeSlideLayout, TargetPart: layout}},
	}
	plan, err := s.pkg.PlanGraphMutation(mutation)
	if err != nil {
		return refuse("append graph preflight: " + err.Error())
	}
	if err := s.pkg.ApplyGraphPlan(plan); err != nil {
		return err
	}
	s.slides[newSlide] = true
	return nil
}
