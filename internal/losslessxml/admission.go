package losslessxml

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"strings"
)

var errUnboundPrefix = errors.New("unbound XML prefix")

// Decoder failures are classified using only the token being consumed, not
// later markup. A later comment or CDATA containing an entity-like string
// cannot relabel an earlier malformed attribute/PI as an entity failure.
func classifyDecoderFailure(err error, token []byte) error {
	category := "malformed-xml"
	var syntax *xml.SyntaxError
	if errors.As(err, &syntax) {
		if malformedTokenBeforeEntity(token) {
			return &Refusal{Category: category, Cause: err}
		}
		if containsInvalidXMLRune(token) {
			category = "invalid-character"
		} else if hasUnknownEntity(token) {
			category = "entity-forbidden"
		}
	}
	return &Refusal{Category: category, Cause: err}
}

// Validate the start-tag grammar independently of entity expansion when the
// standard decoder stopped inside an attribute. Otherwise an entity after an
// earlier missing separator would incorrectly own the refusal category.
func malformedTokenBeforeEntity(token []byte) bool {
	if len(token) < 2 || token[0] != '<' || token[1] == '!' || token[1] == '?' || token[1] == '/' {
		return false
	}
	at := 1
	for at < len(token) && token[at] != ' ' && token[at] != '\t' && token[at] != '\r' && token[at] != '\n' && token[at] != '/' && token[at] != '>' {
		at++
	}
	space := func(b byte) bool { return b == ' ' || b == '\t' || b == '\r' || b == '\n' }
	for at < len(token) {
		for at < len(token) && space(token[at]) {
			at++
		}
		if at >= len(token) || token[at] == '/' || token[at] == '>' {
			return false
		}
		start := at
		for at < len(token) && !space(token[at]) && token[at] != '=' && token[at] != '>' {
			at++
		}
		if start == at {
			return true
		}
		for at < len(token) && space(token[at]) {
			at++
		}
		if at >= len(token) || token[at] != '=' {
			return true
		}
		at++
		for at < len(token) && space(token[at]) {
			at++
		}
		if at >= len(token) || (token[at] != '\'' && token[at] != '"') {
			return true
		}
		quote := token[at]
		at++
		for at < len(token) && token[at] != quote {
			at++
		}
		if at >= len(token) {
			return false
		}
		at++
		if at < len(token) && !space(token[at]) && token[at] != '/' && token[at] != '>' {
			return true
		}
	}
	return false
}
func hasUnknownEntity(source []byte) bool {
	for len(source) > 0 {
		at := bytes.IndexByte(source, '&')
		if at < 0 {
			return false
		}
		source = source[at+1:]
		end := bytes.IndexByte(source, ';')
		if end < 0 {
			return false
		}
		entity := string(source[:end])
		source = source[end+1:]
		if entity == "amp" || entity == "lt" || entity == "gt" || entity == "quot" || entity == "apos" || strings.HasPrefix(entity, "#") {
			continue
		}
		return true
	}
	return false
}
func containsInvalidXMLRune(source []byte) bool {
	for _, r := range string(source) {
		if !(r == 9 || r == 10 || r == 13 || r >= 0x20 && r <= 0xd7ff || r >= 0xe000 && r <= 0xfffd || r >= 0x10000 && r <= 0x10ffff) {
			return true
		}
	}
	return false
}

// lexicalAttributeSpacing validates the separator at the point following an
// attribute quote. The decoder accepts adjacent attributes, but XML 1.0 does not.
func lexicalAttributeSpacing(tag []byte) bool {
	at := 1
	for at < len(tag) && tag[at] != ' ' && tag[at] != '\n' && tag[at] != '\r' && tag[at] != '\t' && tag[at] != '>' && tag[at] != '/' {
		at++
	}
	for at < len(tag) {
		for at < len(tag) && (tag[at] == ' ' || tag[at] == '\n' || tag[at] == '\r' || tag[at] == '\t') {
			at++
		}
		if at == len(tag) || tag[at] == '>' || tag[at] == '/' {
			return true
		}
		for at < len(tag) && tag[at] != '=' {
			at++
		}
		if at == len(tag) {
			return false
		}
		at++
		for at < len(tag) && (tag[at] == ' ' || tag[at] == '\n' || tag[at] == '\r' || tag[at] == '\t') {
			at++
		}
		if at == len(tag) || (tag[at] != '\'' && tag[at] != '"') {
			return false
		}
		quote := tag[at]
		at++
		for at < len(tag) && tag[at] != quote {
			at++
		}
		if at == len(tag) {
			return false
		}
		at++
		if at < len(tag) && tag[at] != ' ' && tag[at] != '\n' && tag[at] != '\r' && tag[at] != '\t' && tag[at] != '>' && tag[at] != '/' {
			return false
		}
	}
	return true
}

func validXMLDeclaration(inst []byte) bool {
	// Parse only the XML declaration pseudo-attributes, in their required
	// order. XML S may surround '=', but must separate successive fields.
	i := 0
	space := func(b byte) bool { return b == ' ' || b == '\t' || b == '\n' || b == '\r' }
	// encoding/xml's ProcInst.Inst excludes the mandatory separator after
	// the target; inspect pseudo-attribute order and values here.
	i = skipXMLSpace(inst, i)
	order := []string{"version", "encoding", "standalone"}
	previous := -1
	for i < len(inst) {
		start := i
		for i < len(inst) && (inst[i] >= 'a' && inst[i] <= 'z') {
			i++
		}
		key := string(inst[start:i])
		position := -1
		for j, name := range order {
			if key == name {
				position = j
				break
			}
		}
		if position < 0 || position <= previous || previous < 0 && position != 0 {
			return false
		}
		previous = position
		i = skipXMLSpace(inst, i)
		if i >= len(inst) || inst[i] != '=' {
			return false
		}
		i++
		i = skipXMLSpace(inst, i)
		if i >= len(inst) || (inst[i] != '\'' && inst[i] != '"') {
			return false
		}
		quote := inst[i]
		i++
		start = i
		for i < len(inst) && inst[i] != quote {
			i++
		}
		if i >= len(inst) {
			return false
		}
		value := string(inst[start:i])
		i++
		switch key {
		case "version":
			if value != "1.0" {
				return false
			}
		case "encoding":
			if value != "UTF-8" && value != "utf-8" {
				return false
			}
		case "standalone":
			if value != "yes" && value != "no" {
				return false
			}
		}
		if i < len(inst) && !space(inst[i]) {
			return false
		}
		i = skipXMLSpace(inst, i)
	}
	return previous >= 0
}
func skipXMLSpace(source []byte, i int) int {
	for i < len(source) && (source[i] == ' ' || source[i] == '\t' || source[i] == '\n' || source[i] == '\r') {
		i++
	}
	return i
}

func unboundPrefix(prefix string) error { return fmt.Errorf("%w %s", errUnboundPrefix, prefix) }
