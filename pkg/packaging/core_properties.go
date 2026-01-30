package packaging

import (
	"fmt"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// CoreProperties returns the package core properties (docProps/core.xml).
func (p *Package) CoreProperties() (*common.CoreProperties, error) {
	if p.closed {
		return nil, utils.ErrDocumentClosed
	}

	var partPath string
	rels := p.GetRelationshipsByType("", RelTypeCoreProps)
	if len(rels) > 0 {
		partPath = ResolveRelationshipTarget("", rels[0].Target)
	}

	if partPath == "" && p.PartExists(CorePropertiesPath) {
		partPath = CorePropertiesPath
	}

	if partPath == "" {
		return common.NewCoreProperties(), nil
	}

	part, err := p.GetPart(partPath)
	if err != nil {
		return nil, err
	}
	data, err := part.Content()
	if err != nil {
		return nil, err
	}

	props := &common.CoreProperties{}
	if err := utils.UnmarshalXML(data, props); err != nil {
		return nil, err
	}
	return props, nil
}

// SetCoreProperties writes core properties to docProps/core.xml and ensures a relationship.
func (p *Package) SetCoreProperties(props *common.CoreProperties) error {
	if p.closed {
		return utils.ErrDocumentClosed
	}
	if props == nil {
		return fmt.Errorf("core properties cannot be nil")
	}

	data, err := utils.MarshalXMLWithHeader(props)
	if err != nil {
		return err
	}

	part, err := p.GetPart(CorePropertiesPath)
	if err != nil {
		if _, err := p.AddPart(CorePropertiesPath, ContentTypeCoreProps, data); err != nil {
			return err
		}
	} else if err := part.SetContent(data); err != nil {
		return err
	}

	p.contentTypes.EnsureContentType(CorePropertiesPath, ContentTypeCoreProps)

	rels := p.GetRelationships("")
	for i := range rels.Relationships {
		if rels.Relationships[i].Type == RelTypeCoreProps {
			rels.Relationships[i].Target = CorePropertiesPath
			rels.Relationships[i].TargetMode = TargetModeInternal
			p.modified = true
			return nil
		}
	}

	rels.Add(RelTypeCoreProps, CorePropertiesPath, TargetModeInternal)
	p.modified = true
	return nil
}
