package document

import (
	"encoding/xml"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// initPackage initializes the OPC package structure for a new document.
func (d *Document) initPackage() error {
	// Create document.xml
	docData, err := utils.MarshalXMLWithHeader(d.document)
	if err != nil {
		return err
	}
	_, err = d.pkg.AddPart(packaging.WordDocumentPath, packaging.ContentTypeWordDocument, docData)
	if err != nil {
		return err
	}

	// Add package-level relationships
	d.pkg.AddRelationship("", packaging.WordDocumentPath, packaging.RelTypeOfficeDocument)

	// Create styles.xml
	stylesData, err := utils.MarshalXMLWithHeader(d.styles)
	if err != nil {
		return err
	}
	_, err = d.pkg.AddPart(packaging.WordStylesPath, packaging.ContentTypeStyles, stylesData)
	if err != nil {
		return err
	}
	d.pkg.AddRelationship(packaging.WordDocumentPath, "styles.xml", packaging.RelTypeStyles)

	// Create settings.xml
	settingsData, err := utils.MarshalXMLWithHeader(d.settings)
	if err != nil {
		return err
	}
	_, err = d.pkg.AddPart(packaging.WordSettingsPath, packaging.ContentTypeSettings, settingsData)
	if err != nil {
		return err
	}
	d.pkg.AddRelationship(packaging.WordDocumentPath, "settings.xml", packaging.RelTypeSettings)

	return nil
}

// updatePackage updates the OPC package with current document state.
func (d *Document) updatePackage() error {
	// Update document.xml
	docData, err := utils.MarshalXMLWithHeader(d.document)
	if err != nil {
		return err
	}
	docPart, err := d.pkg.GetPart(packaging.WordDocumentPath)
	if err != nil {
		return err
	}
	if err := docPart.SetContent(docData); err != nil {
		return err
	}

	// Update styles.xml if exists
	if d.styles != nil {
		stylesData, err := utils.MarshalXMLWithHeader(d.styles)
		if err != nil {
			return err
		}
		stylesPart, err := d.pkg.GetPart(packaging.WordStylesPath)
		if err == nil {
			if err := stylesPart.SetContent(stylesData); err != nil {
				return err
			}
		}
	}

	// Update settings.xml if exists
	if d.settings != nil {
		settingsData, err := utils.MarshalXMLWithHeader(d.settings)
		if err != nil {
			return err
		}
		settingsPart, err := d.pkg.GetPart(packaging.WordSettingsPath)
		if err == nil {
			if err := settingsPart.SetContent(settingsData); err != nil {
				return err
			}
		}
	}

	// Update comments.xml if we have comments
	if d.comments != nil && len(d.comments.Comments) > 0 {
		commentsData, err := utils.MarshalXMLWithHeader(d.comments)
		if err != nil {
			return err
		}
		commentsPart, _ := d.pkg.GetPart(packaging.WordCommentsPath)
		if commentsPart == nil {
			_, err = d.pkg.AddPart(packaging.WordCommentsPath, packaging.ContentTypeComments, commentsData)
			if err != nil {
				return err
			}
			d.pkg.AddRelationship(packaging.WordDocumentPath, "comments.xml", packaging.RelTypeComments)
		} else {
			if err := commentsPart.SetContent(commentsData); err != nil {
				return err
			}
		}
	}

	return nil
}

// parseDocument parses the document.xml part.
func (d *Document) parseDocument() error {
	// Find document part via relationship
	rels := d.pkg.GetRelationshipsByType("", packaging.RelTypeOfficeDocument)
	if len(rels) == 0 {
		return utils.ErrPartNotFound
	}

	docPath := packaging.ResolveRelationshipTarget("", rels[0].Target)
	part, err := d.pkg.GetPart(docPath)
	if err != nil {
		return err
	}

	content, err := part.Content()
	if err != nil {
		return err
	}

	d.document = &wml.Document{}
	return xml.Unmarshal(content, d.document)
}

// parseStyles parses the styles.xml part.
func (d *Document) parseStyles() error {
	rels := d.pkg.GetRelationshipsByType(packaging.WordDocumentPath, packaging.RelTypeStyles)
	if len(rels) == 0 {
		return nil // optional
	}

	stylesPath := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, rels[0].Target)
	part, err := d.pkg.GetPart(stylesPath)
	if err != nil {
		return nil
	}

	content, err := part.Content()
	if err != nil {
		return err
	}

	d.styles = &wml.Styles{}
	return xml.Unmarshal(content, d.styles)
}

// parseSettings parses the settings.xml part.
func (d *Document) parseSettings() error {
	rels := d.pkg.GetRelationshipsByType(packaging.WordDocumentPath, packaging.RelTypeSettings)
	if len(rels) == 0 {
		return nil // optional
	}

	settingsPath := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, rels[0].Target)
	part, err := d.pkg.GetPart(settingsPath)
	if err != nil {
		return nil
	}

	content, err := part.Content()
	if err != nil {
		return err
	}

	d.settings = &wml.Settings{}
	if err := xml.Unmarshal(content, d.settings); err != nil {
		return err
	}

	// Check if track changes is enabled
	if d.settings.TrackRevisions != nil && d.settings.TrackRevisions.Enabled() {
		d.trackChanges = true
	}

	return nil
}

// parseComments parses the comments.xml part.
func (d *Document) parseComments() error {
	rels := d.pkg.GetRelationshipsByType(packaging.WordDocumentPath, packaging.RelTypeComments)
	if len(rels) == 0 {
		return nil // optional
	}

	commentsPath := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, rels[0].Target)
	part, err := d.pkg.GetPart(commentsPath)
	if err != nil {
		return nil
	}

	content, err := part.Content()
	if err != nil {
		return err
	}

	d.comments = &wml.Comments{}
	if err := xml.Unmarshal(content, d.comments); err != nil {
		return err
	}

	// Track max comment ID
	for _, c := range d.comments.Comments {
		if c.ID >= d.nextCommentID {
			d.nextCommentID = c.ID + 1
		}
	}

	return nil
}
