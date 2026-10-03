package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// ContractBlankEdit fills one existing blank cell with a string or finite number.
// Addresses and values are supplied by the caller; no workbook writer builds inputs.
type ContractBlankEdit struct {
	Address string
	Value   any
}

// SetContractBlankValues edits a bounded, single-sheet batch of existing blank
// cells. It retains untouched XML bytes, invalidates every ordinary formula
// cache after a changed input, and refuses attributed/unknown dependencies.
func (s *EditSession) SetContractBlankValues(sheet string, edits []ContractBlankEdit) error {
	if len(edits) == 0 || len(edits) > 64 {
		return editRefusal("unsupported_structure", "bounded nonempty blank batch required")
	}
	if s.sheets[sheet] == "" {
		return editRefusal("missing_target", "sheet absent")
	}
	if err := s.guardMode(true); err != nil {
		return err
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	chain, err := s.calculationChain(g)
	if err != nil {
		return err
	}
	if chain.part != "" {
		return editRefusal("unsupported_structure", "calculation chain requires separate removal profile")
	}
	enrolled := map[string]bool{}
	for _, part := range s.sheets {
		enrolled[part] = true
	}
	for _, part := range g.Parts {
		if part.ContentType == packaging.ContentTypeWorksheet && !enrolled[part.Name] {
			return editRefusal("unsupported_structure", "unlisted worksheet")
		}
	}
	selected := map[string]ContractBlankEdit{}
	for _, edit := range edits {
		ref, e := utils.ParseCellRef(edit.Address)
		if e != nil || !strictCell.MatchString(edit.Address) || ref.Row > 1048576 || ref.Col > 16384 || selected[edit.Address].Address != "" {
			return editRefusal("unsupported_structure", "unique bounded cell address required")
		}
		switch value := edit.Value.(type) {
		case string:
			if !losslessxml.ValidAuthoredText(value) {
				return editRefusal("unsupported_structure", "invalid authored string")
			}
		case float64:
			if math.IsNaN(value) || math.IsInf(value, 0) {
				return editRefusal("unsupported_structure", "finite number required")
			}
		default:
			return editRefusal("unsupported_structure", "only string or numeric blank fills")
		}
		selected[edit.Address] = edit
	}
	type sheetState struct {
		doc      *losslessxml.Document
		source   []byte
		hash     string
		removals []losslessxml.Element
		blank    map[string]losslessxml.Element
	}
	sheets := map[string]*sheetState{}
	formulaCount := 0
	parts := make([]string, 0, len(enrolled))
	for part := range enrolled {
		parts = append(parts, part)
	}
	sort.Strings(parts)
	for _, part := range parts {
		source, hash, e := s.pkg.Part(part)
		if e != nil {
			return e
		}
		doc, e := losslessxml.Parse(source)
		if e != nil {
			return e
		}
		nodes := doc.Elements()
		if len(nodes) == 0 || nodes[0].Name() != expanded("worksheet") {
			return editRefusal("unsupported_structure", "worksheet root required")
		}
		state := &sheetState{doc: doc, source: source, hash: hash, blank: map[string]losslessxml.Element{}}
		sheets[part] = state
		seen := map[string]bool{}
		rows := map[string]bool{}
		dataCount := 0
		for _, node := range nodes {
			if node.Name() == expanded("sheetData") {
				parent, ok := node.Parent()
				if !ok || parent != nodes[0] {
					return editRefusal("unsupported_structure", "nested sheetData")
				}
				dataCount++
			}
			if node.Name() == expanded("row") {
				id := attr(node, "r")
				if id == "" || rows[id] {
					return editRefusal("ambiguous_target", "duplicate or missing row identity")
				}
				rows[id] = true
			}
		}
		if dataCount != 1 {
			return editRefusal("unsupported_structure", "exactly one sheetData required")
		}
		for _, cell := range nodes {
			if cell.Name() != expanded("c") {
				continue
			}
			address := attr(cell, "r")
			if address == "" || seen[address] {
				return editRefusal("ambiguous_target", "duplicate or missing cell address")
			}
			seen[address] = true
			row, ok := cell.Parent()
			if !ok || row.Name() != expanded("row") {
				return editRefusal("unsupported_structure", "cell without row")
			}
			coord, e := utils.ParseCellRef(address)
			if e != nil || !strictCell.MatchString(address) || attr(row, "r") != strconv.Itoa(coord.Row) {
				return editRefusal("unsupported_structure", "cell/row address differs")
			}
			data, ok := row.Parent()
			if !ok || data.Name() != expanded("sheetData") {
				return editRefusal("unsupported_structure", "cell without sheetData")
			}
			root, ok := data.Parent()
			if !ok || root != nodes[0] {
				return editRefusal("unsupported_structure", "cell owner differs")
			}
			children := map[string]losslessxml.Element{}
			for _, node := range nodes {
				if parent, ok := node.Parent(); ok && parent == cell {
					if node.Name().Space != packaging.NSSpreadsheetML {
						return editRefusal("unsupported_structure", "foreign cell child")
					}
					if _, dup := children[node.Name().Local]; dup {
						return editRefusal("unsupported_structure", "duplicate cell child")
					}
					children[node.Name().Local] = node
				}
			}
			if f, ok := children["f"]; ok {
				formulaCount++
				if len(f.Attributes()) != 0 {
					return editRefusal("xlsx-cache-topology-unsupported", "attributed formula result")
				}
				if _, leaf := f.Text(); !leaf {
					return editRefusal("unsupported_structure", "nonleaf formula")
				}
				for _, a := range cell.Attributes() {
					if a.Name.Space != "" || a.Name.Local != "r" && a.Name.Local != "s" {
						return editRefusal("unsupported_structure", "unclassified formula cell attribute")
					}
				}
				for name := range children {
					if name != "f" && name != "v" {
						return editRefusal("unsupported_structure", "unclassified formula child")
					}
				}
				if v, ok := children["v"]; ok {
					if len(v.Attributes()) != 0 {
						return editRefusal("unsupported_structure", "cached value attributes")
					}
					if _, leaf := v.Text(); !leaf {
						return editRefusal("unsupported_structure", "nonleaf formula cache")
					}
					state.removals = append(state.removals, v)
				}
			}
			if part == s.sheets[sheet] {
				if _, ok := selected[address]; ok {
					state.blank[address] = cell
					if !cell.SelfClosing() || len(children) != 0 || attr(cell, "t") != "" {
						return editRefusal("unsupported_structure", "target is not an existing blank")
					}
					for _, a := range cell.Attributes() {
						if a.Name.Space != "" || a.Name.Local != "r" && a.Name.Local != "s" {
							return editRefusal("unsupported_structure", "unclassified blank attribute")
						}
					}
				}
			}
		}
	}
	target := sheets[s.sheets[sheet]]
	if len(target.blank) != len(selected) {
		return editRefusal("missing_target", "selected blank cell absent")
	}
	replacements := []packaging.Replacement{}
	for _, part := range parts {
		state := sheets[part]
		data := state.source
		if len(state.removals) != 0 {
			data, err = state.doc.RemoveElements(state.removals)
			if err != nil {
				return editRefusal("unsupported_structure", "cache removal failed")
			}
		}
		if part == s.sheets[sheet] {
			doc, e := losslessxml.Parse(data)
			if e != nil {
				return e
			}
			type splice struct {
				start, end int
				bytes      []byte
			}
			splices := []splice{}

			for _, cell := range doc.Elements() {
				if cell.Name() != expanded("c") {
					continue
				}
				edit, ok := selected[attr(cell, "r")]
				if !ok {
					continue
				}
				if !cell.SelfClosing() {
					return editRefusal("stale_target", "blank changed after cache removal")
				}
				raw := cell.Raw()
				if len(raw) < 3 || !bytes.HasSuffix(raw, []byte("/>")) {
					return editRefusal("unsupported_structure", "blank is not self closing")
				}
				name := string(raw[1:])
				if i := strings.IndexAny(name, " \t\r\n/>"); i >= 0 {
					name = name[:i]
				}
				if name == "" {
					return editRefusal("unsupported_structure", "blank QName missing")
				}
				ns := cell.Namespaces()
				prefix := ""
				if i := strings.IndexByte(name, ':'); i >= 0 {
					prefix = name[:i] + ":"
					if ns[name[:i]] != packaging.NSSpreadsheetML {
						return editRefusal("unsupported_structure", "foreign cell prefix")
					}
				} else if ns[""] != packaging.NSSpreadsheetML {
					return editRefusal("unsupported_structure", "wrong default cell namespace")
				}
				var content string
				extra := ""
				switch value := edit.Value.(type) {
				case string:
					var escaped bytes.Buffer
					if e := xml.EscapeText(&escaped, []byte(value)); e != nil {
						return e
					}
					extra = ` t="inlineStr"`
					space := ""
					if len(value) > 0 && (value[0] == ' ' || value[len(value)-1] == ' ') {
						space = ` xml:space="preserve"`
					}
					content = "<" + prefix + "is><" + prefix + "t" + space + ">" + escaped.String() + "</" + prefix + "t></" + prefix + "is>"
				case float64:
					content = "<" + prefix + "v>" + strconv.FormatFloat(value, 'g', -1, 64) + "</" + prefix + "v>"
				}
				result := append([]byte(nil), raw[:len(raw)-2]...)
				result = append(result, extra...)
				result = append(result, '>')
				result = append(result, content...)
				result = append(result, []byte("</"+name+">")...)
				start, end := cell.SourceRange()
				splices = append(splices, splice{start, end, result})
			}
			if len(splices) != len(selected) {
				return editRefusal("stale_target", "selected blanks disappeared")
			}
			sort.Slice(splices, func(i, j int) bool { return splices[i].start < splices[j].start })
			var out bytes.Buffer
			at := 0
			for _, sp := range splices {
				if sp.start < at || sp.end > len(data) {
					return editRefusal("unsupported_structure", "overlapping blank edits")
				}
				out.Write(data[at:sp.start])
				out.Write(sp.bytes)
				at = sp.end
			}
			out.Write(data[at:])
			data = out.Bytes()
			if _, e := losslessxml.Parse(data); e != nil {
				return editRefusal("unsupported_structure", "edited worksheet invalid")
			}
		}
		if !bytes.Equal(state.source, data) {
			replacements = append(replacements, packaging.Replacement{Part: part, ExpectedSHA256: state.hash, Data: data})
		}
	}
	if formulaCount > 0 {
		main, hash, e := s.pkg.Part(s.main)
		if e != nil {
			return e
		}
		doc, e := losslessxml.Parse(main)
		if e != nil {
			return e
		}
		nodes := doc.Elements()
		if nodes[0].Name() != expanded("workbook") {
			return editRefusal("unsupported_structure", "workbook root differs")
		}
		var calc losslessxml.Element
		count := 0
		for _, node := range nodes {
			if node.Name() == expanded("calcPr") {
				parent, ok := node.Parent()
				if !ok || parent != nodes[0] {
					return editRefusal("unsupported_structure", "calcPr owner differs")
				}
				calc = node
				count++
			}
		}
		if count > 1 {
			return editRefusal("unsupported_structure", "duplicate calcPr")
		}
		if count == 1 {
			for _, a := range calc.Attributes() {
				if a.Name.Space != "" {
					return editRefusal("unsupported_structure", "unknown calc metadata namespace")
				}
				switch a.Name.Local {
				case "calcMode":
					if a.Value != "auto" {
						return editRefusal("unsupported_structure", "non-auto calculation mode")
					}
				case "iterate":
					if a.Value != "0" && a.Value != "false" {
						return editRefusal("unsupported_structure", "iterative calculation not modelled")
					}
				case "fullPrecision":
					if a.Value != "1" && a.Value != "true" {
						return editRefusal("unsupported_structure", "precision-as-displayed dependencies not modelled")
					}
				case "calcId", "fullCalcOnLoad", "forceFullCalc", "calcOnSave", "concurrentCalc", "concurrentManualCount":
				default:
					return editRefusal("unsupported_structure", "unknown calculation metadata")
				}
			}
		}
		flags := []xml.Attr{{Name: xml.Name{Local: "calcMode"}, Value: "auto"}, {Name: xml.Name{Local: "fullCalcOnLoad"}, Value: "1"}, {Name: xml.Name{Local: "forceFullCalc"}, Value: "1"}}
		var edited []byte
		if count == 0 {
			edited, e = doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: nodes[0], Children: []losslessxml.NewElement{{Name: expanded("calcPr"), Attributes: flags}}}})
		} else {
			attrs := []losslessxml.AttributeEdit{}
			for _, a := range flags {
				attrs = append(attrs, losslessxml.AttributeEdit{Target: calc, Name: a.Name, Value: a.Value})
			}
			edited, e = doc.Edit(nil, attrs)
		}
		if e != nil {
			return editRefusal("unsupported_structure", "calcPr edit failed")
		}
		if !bytes.Equal(main, edited) {
			replacements = append(replacements, packaging.Replacement{Part: s.main, ExpectedSHA256: hash, Data: edited})
		}
	}
	if err = s.pkg.Replace(replacements); err != nil {
		return err
	}
	s.generation++
	return nil
}
