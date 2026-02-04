// Package document provides numbering (lists) functionality.
package document

import (
	"strconv"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Numbering represents a numbering definition instance.
type Numbering struct {
	doc *documentImpl
	num *wml.Num
}

// ID returns the numbering definition ID.
func (n *Numbering) ID() int {
	return n.num.NumID
}

// AbstractNumbering represents an abstract numbering definition.
type AbstractNumbering struct {
	doc *documentImpl
	abs *wml.AbstractNum
}

// ID returns the abstract numbering definition ID.
func (a *AbstractNumbering) ID() int {
	return a.abs.AbstractNumID
}

// Level returns the level definition or nil if out of range.
func (a *AbstractNumbering) Level(level int) *wml.Lvl {
	if level < 0 || level >= len(a.abs.Lvl) {
		return nil
	}
	return a.abs.Lvl[level]
}

// Numbering returns the numbering definitions in the document.
func (d *documentImpl) Numbering() []*Numbering {
	if d.numbering == nil {
		return nil
	}
	result := make([]*Numbering, len(d.numbering.Num))
	for i, num := range d.numbering.Num {
		result[i] = &Numbering{doc: d, num: num}
	}
	return result
}

// AbstractNumberings returns all abstract numbering definitions.
func (d *documentImpl) AbstractNumberings() []*AbstractNumbering {
	if d.numbering == nil {
		return nil
	}
	result := make([]*AbstractNumbering, len(d.numbering.AbstractNum))
	for i, abs := range d.numbering.AbstractNum {
		result[i] = &AbstractNumbering{doc: d, abs: abs}
	}
	return result
}

// NumberingByID returns a numbering definition by ID.
func (d *documentImpl) NumberingByID(id int) *Numbering {
	if d.numbering == nil {
		return nil
	}
	for _, num := range d.numbering.Num {
		if num.NumID == id {
			return &Numbering{doc: d, num: num}
		}
	}
	return nil
}

// AbstractNumberingByID returns an abstract numbering definition by ID.
func (d *documentImpl) AbstractNumberingByID(id int) *AbstractNumbering {
	if d.numbering == nil {
		return nil
	}
	for _, abs := range d.numbering.AbstractNum {
		if abs.AbstractNumID == id {
			return &AbstractNumbering{doc: d, abs: abs}
		}
	}
	return nil
}

// AddNumberingDefinition creates a numbering definition with levels 0..levels-1.
func (d *documentImpl) AddNumberingDefinition(levels int) (*Numbering, *AbstractNumbering, error) {
	if levels < 1 || levels > 9 {
		return nil, nil, utils.ErrInvalidIndex
	}
	if d.numbering == nil {
		d.numbering = &wml.Numbering{}
	}

	absID := d.nextAbstractNumID
	d.nextAbstractNumID++

	abs := &wml.AbstractNum{
		AbstractNumID: absID,
		MultiLevelType: &wml.MultiLevelType{Val: "multilevel"},
	}
	for i := 0; i < levels; i++ {
		abs.Lvl = append(abs.Lvl, &wml.Lvl{
			Ilvl:   i,
			Start:  &wml.NumStart{Val: 1},
			NumFmt: &wml.NumFmt{Val: wml.NumFmtDecimal},
			LvlText: &wml.LvlText{Val: "%"+strconv.Itoa(i+1)+"."},
			LvlJc:  &wml.LvlJc{Val: "left"},
		})
	}

	numID := d.nextNumID
	d.nextNumID++
	num := &wml.Num{
		NumID:        numID,
		AbstractNumID: &wml.AbstractNumIDRef{Val: absID},
	}

	d.numbering.AbstractNum = append(d.numbering.AbstractNum, abs)
	d.numbering.Num = append(d.numbering.Num, num)

	return &Numbering{doc: d, num: num}, &AbstractNumbering{doc: d, abs: abs}, nil
}

// AddNumberedListStyle creates a default numbered list and returns its numbering ID.
func (d *documentImpl) AddNumberedListStyle() (int, error) {
	num, _, err := d.AddNumberingDefinition(1)
	if err != nil {
		return 0, err
	}
	return num.ID(), nil
}

// AddBulletedListStyle creates a default bullet list and returns its numbering ID.
func (d *documentImpl) AddBulletedListStyle() (int, error) {
	num, abs, err := d.AddNumberingDefinition(1)
	if err != nil {
		return 0, err
	}
	lvl := abs.Level(0)
	if lvl != nil {
		lvl.NumFmt = &wml.NumFmt{Val: wml.NumFmtBullet}
		lvl.LvlText = &wml.LvlText{Val: "-"}
	}
	return num.ID(), nil
}
