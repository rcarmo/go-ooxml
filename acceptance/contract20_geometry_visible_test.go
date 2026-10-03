package acceptance

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Independently count visible text and direct object owners in the created
// slide. The subtitle is a real but empty placeholder; the sole other visible
// text is the selected table cell. A title count alone misses extra shapes.
func contract20GeometryVisibleSlide(slide []byte, title, firstCell string) error {
	d, err := losslessxml.Parse(slide)
	if err != nil {
		return err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	es := d.Elements()
	if len(es) == 0 || es[0].Name() != p("sld") {
		return fmt.Errorf("created slide root absent")
	}
	var tree losslessxml.Element
	trees := 0
	for _, n := range es {
		if n.Name() != p("spTree") {
			continue
		}
		parent, ok := n.Parent()
		if ok && parent.Name() == p("cSld") {
			tree = n
			trees++
		}
	}
	if trees != 1 {
		return fmt.Errorf("created shape tree count %d", trees)
	}
	shapeCount, tableCount := 0, 0
	ownerIDs := map[uint64]bool{}
	texts := map[string]string{}
	placeholders := map[string]int{}
	contains := func(n, owner losslessxml.Element) bool {
		for parent, ok := n.Parent(); ok; parent, ok = parent.Parent() {
			if parent == owner {
				return true
			}
		}
		return false
	}
	for _, owner := range es {
		parent, ok := owner.Parent()
		if !ok || parent != tree {
			continue
		}
		kind := owner.Name()
		if kind == p("nvGrpSpPr") || kind == p("grpSpPr") {
			continue
		}
		if kind != p("sp") && kind != p("graphicFrame") {
			return fmt.Errorf("unrequested created tree owner %v", kind)
		}
		if kind == p("sp") {
			shapeCount++
		} else {
			tableCount++
		}
		idCount := 0
		kindText := ""
		for _, n := range es {
			if !contains(n, owner) {
				continue
			}
			if n.Name() == p("cNvPr") {
				idCount++
				id := ""
				for _, attr := range n.Attributes() {
					if attr.Name.Space == "" && attr.Name.Local == "id" {
						id = attr.Value
					}
				}
				value, e := strconv.ParseUint(id, 10, 32)
				if e != nil || value == 0 || ownerIDs[value] {
					return fmt.Errorf("invalid/duplicate created object ID %q", id)
				}
				ownerIDs[value] = true
			}
			if n.Name() == p("ph") {
				for _, attr := range n.Attributes() {
					if attr.Name.Space == "" && attr.Name.Local == "type" {
						placeholders[attr.Value]++
					}
				}
			}
			if n.Name() == a("t") {
				text, leaf := n.Text()
				if !leaf {
					return fmt.Errorf("non-leaf visible text")
				}
				kindText += text
			}
		}
		if idCount != 1 {
			return fmt.Errorf("owner object ID count %d", idCount)
		}
		if kind == p("sp") {
			for _, n := range es {
				if n.Name() != p("ph") || !contains(n, owner) {
					continue
				}
				for _, attr := range n.Attributes() {
					if attr.Name.Space == "" && attr.Name.Local == "type" {
						if _, exists := texts[attr.Value]; exists {
							return fmt.Errorf("duplicate title/subtitle owner %s", attr.Value)
						}
						texts[attr.Value] = kindText
					}
				}
			}
		} else {
			texts["table"] = kindText
		}
	}
	for _, n := range es {
		if n.Name() != a("t") {
			continue
		}
		owned := false
		for _, owner := range es {
			parent, ok := owner.Parent()
			if !ok || parent != tree || owner.Name() != p("sp") && owner.Name() != p("graphicFrame") {
				continue
			}
			if contains(n, owner) {
				owned = true
				break
			}
		}
		if !owned {
			return fmt.Errorf("visible text outside requested owners")
		}
	}
	if shapeCount != 2 || tableCount != 1 || len(ownerIDs) != 3 || placeholders["ctrTitle"] != 1 || placeholders["subTitle"] != 1 || len(placeholders) != 2 {
		return fmt.Errorf("created owner/placeholder inventory %d/%d IDs%d roles%v", shapeCount, tableCount, len(ownerIDs), placeholders)
	}
	if len(texts) != 3 || texts["ctrTitle"] != title || texts["subTitle"] != "" || texts["table"] != firstCell {
		return fmt.Errorf("created visible text/role inventory %v", texts)
	}
	return nil
}
