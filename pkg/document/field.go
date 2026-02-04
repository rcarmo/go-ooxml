// Package document provides field functionality.
package document

import (
	"fmt"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// Field represents a Word field.
type Field struct {
	Instruction string
	Display     string
}

// AddField inserts a field with instruction and optional display text.
func (p *paragraphImpl) AddField(instruction, display string) (*Field, error) {
	if instruction == "" {
		return nil, fmt.Errorf("field instruction cannot be empty")
	}

	begin := &wml.R{Content: []interface{}{&wml.FldChar{FldCharType: wml.FldCharBegin}}}
	instr := &wml.R{Content: []interface{}{wml.NewInstrText(instruction)}}
	sep := &wml.R{Content: []interface{}{&wml.FldChar{FldCharType: wml.FldCharSeparate}}}
	end := &wml.R{Content: []interface{}{&wml.FldChar{FldCharType: wml.FldCharEnd}}}

	p.p.Content = append(p.p.Content, begin, instr, sep)
	if display != "" {
		p.p.Content = append(p.p.Content, &wml.R{Content: []interface{}{wml.NewT(display)}})
	}
	p.p.Content = append(p.p.Content, end)

	return &Field{Instruction: instruction, Display: display}, nil
}
