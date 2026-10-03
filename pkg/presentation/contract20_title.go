package presentation

import (
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"
	"unicode/utf16"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ContractTitleFragment reports stored visible text and its direct run/break
// properties. It is an observation, not an evaluated field or rendered value.
type ContractTitleFragment struct {
	Text       string
	Attributes map[string]string
}

// ContractTitleAnchor is issued for a unique, directly owned title paragraph.
// Edits require its original part fingerprint and issuance identity.
type ContractTitleAnchor struct {
	session          *EditSession
	part, hash, text string
	shapeID          uint32
	doc              *losslessxml.Document
	paragraph        losslessxml.Element
	fragments        []ContractTitleFragment
	consumed         bool
}

func (a *ContractTitleAnchor) Text() string {
	if a == nil {
		return ""
	}
	return a.text
}
func (a *ContractTitleAnchor) Fragments() []ContractTitleFragment {
	if a == nil {
		return nil
	}
	out := make([]ContractTitleFragment, len(a.fragments))
	for i, fragment := range a.fragments {
		attrs := make(map[string]string, len(fragment.Attributes))
		for k, v := range fragment.Attributes {
			attrs[k] = v
		}
		out[i] = ContractTitleFragment{Text: fragment.Text, Attributes: attrs}
	}
	return out
}

func contractTitleChildren(d *losslessxml.Document, parent losslessxml.Element, name xml.Name) []losslessxml.Element {
	var out []losslessxml.Element
	for _, e := range d.Elements() {
		p, ok := e.Parent()
		if ok && p == parent && e.Name() == name {
			out = append(out, e)
		}
	}
	return out
}

// FindContractTitle observes a first title paragraph without rewriting it.
// Unsupported break/field topology is readable but edit-time refused.
func (s *EditSession) FindContractTitle(part string) (*ContractTitleAnchor, error) {
	if !s.slides[part] {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "unenrolled slide")
	}
	data, hash, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	p := func(local string) xml.Name { return name(packaging.NSPresentationML, local) }
	a := func(local string) xml.Name { return name(packaging.NSDrawingML, local) }
	if len(d.Elements()) == 0 || d.Elements()[0].Name() != p("sld") {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "slide root")
	}
	var shape losslessxml.Element
	count := 0
	for _, ph := range d.Elements() {
		if ph.Name() != p("ph") {
			continue
		}
		isTitle := false
		for _, attr := range ph.Attributes() {
			if attr.Name.Space == "" && attr.Name.Local == "type" && (attr.Value == "title" || attr.Value == "ctrTitle") {
				isTitle = true
			}
		}
		if !isTitle {
			continue
		}
		for ancestor, ok := ph.Parent(); ok; ancestor, ok = ancestor.Parent() {
			if ancestor.Name() == p("sp") {
				shape = ancestor
				count++
				break
			}
		}
	}
	if count != 1 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "unique title shape")
	}
	owner, ok := shape.Parent()
	if !ok || owner.Name() != p("spTree") {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "grouped title")
	}
	var shapeID uint32
	identity := contractTitleChildren(d, shape, p("nvSpPr"))
	if len(identity) != 1 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title identity")
	}
	properties := contractTitleChildren(d, identity[0], p("cNvPr"))
	if len(properties) != 1 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title identity")
	}
	for _, attr := range properties[0].Attributes() {
		if attr.Name.Space == "" && attr.Name.Local == "id" {
			n, er := strconv.ParseUint(attr.Value, 10, 32)
			if er != nil || n == 0 {
				return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title ID")
			}
			shapeID = uint32(n)
		}
	}
	if shapeID == 0 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title ID missing")
	}
	for _, other := range d.Elements() {
		if other == properties[0] || other.Name() != p("cNvPr") {
			continue
		}
		for _, attr := range other.Attributes() {
			if attr.Name.Space == "" && attr.Name.Local == "id" && attr.Value == strconv.FormatUint(uint64(shapeID), 10) {
				return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "duplicate title ID")
			}
		}
	}
	bodies := contractTitleChildren(d, shape, p("txBody"))
	if len(bodies) != 1 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title body")
	}
	paragraphs := contractTitleChildren(d, bodies[0], a("p"))
	if len(paragraphs) == 0 {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title paragraph")
	}
	paragraph := paragraphs[0]
	var text strings.Builder
	fragments := []ContractTitleFragment{}
	for _, child := range d.Elements() {
		parent, ok := child.Parent()
		if !ok || parent != paragraph {
			continue
		}
		switch child.Name() {
		case a("pPr"), a("endParaRPr"):
			continue
		case a("br"):
			attrs := map[string]string{}
			for _, rp := range contractTitleChildren(d, child, a("rPr")) {
				for _, attr := range rp.Attributes() {
					if attr.Name.Space == "" {
						attrs[attr.Name.Local] = attr.Value
					}
				}
			}
			fragments = append(fragments, ContractTitleFragment{Text: "\n", Attributes: attrs})
			text.WriteByte('\n')
		case a("r"), a("fld"):
			leaves := contractTitleChildren(d, child, a("t"))
			if len(leaves) != 1 {
				return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "visible title fragment")
			}
			v, leaf := leaves[0].Text()
			if !leaf {
				return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "nested title text")
			}
			attrs := map[string]string{}
			for _, rp := range contractTitleChildren(d, child, a("rPr")) {
				for _, attr := range rp.Attributes() {
					if attr.Name.Space == "" {
						attrs[attr.Name.Local] = attr.Value
					}
				}
			}
			fragments = append(fragments, ContractTitleFragment{Text: v, Attributes: attrs})
			text.WriteString(v)
		default:
			return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "foreign title child")
		}
	}
	issued := &ContractTitleAnchor{session: s, part: part, hash: hash, shapeID: shapeID, doc: d, paragraph: paragraph, text: text.String(), fragments: fragments}
	if s.contractTitles == nil {
		s.contractTitles = make(map[*ContractTitleAnchor]struct{})
	}
	s.contractTitles[issued] = struct{}{}
	return issued, nil
}

func contractRunText(source []byte, run, leaf losslessxml.Element, value string) ([]byte, error) {
	start, end := run.SourceRange()
	a, b := leaf.ContentRange()
	if start < 0 || a < start || b > end || run.SelfClosing() || leaf.SelfClosing() {
		return nil, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "nonordinary run")
	}
	return append(append(bytes.Clone(source[start:a]), []byte(xmlEscapeContractText(value))...), source[b:end]...), nil
}

// ReplaceContractTitleSpan replaces a half-open UTF-16 interval in the held
// ordinary paragraph. A replacement inherits the starting run's properties;
// other runs and all non-title package members retain their source bytes.
func (s *EditSession) ReplaceContractTitleSpan(target *ContractTitleAnchor, start, end int, replacement string) (bool, error) {
	if target == nil || target.session != s {
		return false, editRefusal("PPTX_STALE_ANCHOR", "foreign title anchor")
	}
	if _, ok := s.contractTitles[target]; !ok || target.consumed {
		return false, editRefusal("PPTX_STALE_ANCHOR", "unissued or consumed title anchor")
	}
	current, hash, err := s.pkg.Part(target.part)
	if err != nil {
		return false, err
	}
	if hash != target.hash {
		return false, editRefusal("PPTX_STALE_ANCHOR", "title part changed")
	}
	if start < 0 || end <= start || end > len(utf16.Encode([]rune(target.text))) || !losslessxml.ValidAuthoredText(replacement) || strings.ContainsAny(replacement, "\r\n\t") {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "invalid title span or replacement")
	}
	a := func(local string) xml.Name { return name(packaging.NSDrawingML, local) }
	p := func(local string) xml.Name { return name(packaging.NSPresentationML, local) }
	for _, node := range target.doc.Elements() {
		if !contractDescendant(node, target.paragraph) {
			continue
		}
		parent, ok := node.Parent()
		if !ok {
			return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title owner")
		}
		switch parent.Name() {
		case a("p"):
			if node.Name() != a("r") && node.Name() != a("pPr") && node.Name() != a("endParaRPr") {
				return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "break/field/foreign title")
			}
		case a("r"):
			if node.Name() != a("rPr") && node.Name() != a("t") {
				return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "mixed title run")
			}
		case a("rPr"), a("pPr"), a("endParaRPr"):
			if node.Name() == a("t") || node.Name() == a("br") || node.Name() == a("fld") || node.Name() == a("r") || node.Name() == p("txBody") {
				return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "nested visible title")
			}
		default:
			return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "unsupported title descendants")
		}
	}
	type run struct {
		element, leaf losslessxml.Element
		text          string
		from, to      int
	}
	var runs []run
	position := 0
	for _, element := range contractTitleChildren(target.doc, target.paragraph, a("r")) {
		leaves := contractTitleChildren(target.doc, element, a("t"))
		if len(leaves) != 1 {
			return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title run text count")
		}
		v, leaf := leaves[0].Text()
		if !leaf || leaves[0].SelfClosing() {
			return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title text leaf")
		}
		n := len(utf16.Encode([]rune(v)))
		runs = append(runs, run{element, leaves[0], v, position, position + n})
		position += n
	}
	if position != len(utf16.Encode([]rune(target.text))) || len(runs) == 0 {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title text mismatch")
	}
	// The splice rebuilds the run sequence. Any inter-run source bytes,
	// including XML whitespace, comments or processing instructions, would
	// otherwise disappear; refuse before replacing the part.
	for i := 1; i < len(runs); i++ {
		_, previousEnd := runs[i-1].element.SourceRange()
		nextStart, _ := runs[i].element.SourceRange()
		if previousEnd < 0 || nextStart != previousEnd {
			return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "unmodelled inter-run lexical content")
		}
	}
	if start < 0 || end > position {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title interval")
	}
	var first, last int = -1, -1
	var startOffset, endOffset int
	for i, r := range runs {
		if start >= r.from && start < r.to {
			first, startOffset = i, start-r.from
		}
		if end > r.from && end <= r.to {
			last, endOffset = i, end-r.from
		}
	}
	if first < 0 || last < 0 {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title run boundaries")
	}
	split := func(value string, offset int) (string, string, bool) {
		runes := []rune(value)
		pos := 0
		for i, r := range runes {
			if pos == offset {
				return string(runes[:i]), string(runes[i:]), true
			}
			pos += len(utf16.Encode([]rune{r}))
		}
		if pos == offset {
			return value, "", true
		}
		return "", "", false
	}
	prefix, _, ok := split(runs[first].text, startOffset)
	if !ok {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "split UTF-16 code unit")
	}
	_, suffix, ok := split(runs[last].text, endOffset)
	if !ok {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "split UTF-16 code unit")
	}
	if string([]rune(target.text)) == "" {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "empty title")
	}
	// Compute the expected visible value independently of the XML splice.
	visible := strings.Builder{}
	for _, r := range runs[:first] {
		visible.WriteString(r.text)
	}
	visible.WriteString(prefix)
	visible.WriteString(replacement)
	visible.WriteString(suffix)
	for _, r := range runs[last+1:] {
		visible.WriteString(r.text)
	}
	if visible.String() == target.text {
		return false, nil
	}
	var rebuilt bytes.Buffer
	for _, r := range runs[:first] {
		start, end := r.element.SourceRange()
		rebuilt.Write(current[start:end])
	}
	if prefix != "" {
		raw, e := contractRunText(current, runs[first].element, runs[first].leaf, prefix)
		if e != nil {
			return false, e
		}
		rebuilt.Write(raw)
	}
	if replacement != "" {
		raw, e := contractRunText(current, runs[first].element, runs[first].leaf, replacement)
		if e != nil {
			return false, e
		}
		rebuilt.Write(raw)
	}
	if suffix != "" {
		raw, e := contractRunText(current, runs[last].element, runs[last].leaf, suffix)
		if e != nil {
			return false, e
		}
		rebuilt.Write(raw)
	}
	for _, r := range runs[last+1:] {
		start, end := r.element.SourceRange()
		rebuilt.Write(current[start:end])
	}
	paraStart, paraEnd := target.paragraph.SourceRange()
	firstStart, _ := runs[0].element.SourceRange()
	_, lastEnd := runs[len(runs)-1].element.SourceRange()
	if paraStart < 0 || firstStart < paraStart || lastEnd > paraEnd {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "title lexical range")
	}
	paragraph := append(append(bytes.Clone(current[paraStart:firstStart]), rebuilt.Bytes()...), current[lastEnd:paraEnd]...)
	updated := append(append(bytes.Clone(current[:paraStart]), paragraph...), current[paraEnd:]...)
	if _, err := losslessxml.Parse(updated); err != nil {
		return false, editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "invalid rewritten title")
	}
	if err := s.pkg.Replace([]packaging.Replacement{{Part: target.part, ExpectedSHA256: target.hash, Data: updated}}); err != nil {
		return false, err
	}
	target.consumed = true
	return true, nil
}
