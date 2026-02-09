package spreadsheet

import (
	"fmt"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/chart"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/diagram"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

type drawingKind int

const (
	drawingKindChart drawingKind = iota
	drawingKindDiagram
	drawingKindPicture
)

func (ws *worksheetImpl) addGraphic(fromCell, toCell, title string, kind drawingKind, imagePath string) error {
	if ws == nil || ws.workbook == nil || ws.workbook.pkg == nil {
		return utils.ErrDocumentClosed
	}
	start, err := utils.ParseCellRef(fromCell)
	if err != nil {
		return err
	}
	end, err := utils.ParseCellRef(toCell)
	if err != nil {
		return err
	}
	if start.Row > end.Row || start.Col > end.Col {
		return utils.NewValidationError("range", "fromCell must be top-left of toCell", fmt.Sprintf("%s:%s", fromCell, toCell))
	}

	sheetPath := ws.path
	if sheetPath == "" {
		sheetPath = fmt.Sprintf("xl/worksheets/sheet%d.xml", ws.index+1)
		ws.path = sheetPath
	}
	rels := ws.workbook.pkg.GetRelationships(sheetPath)
	drawingRel := rels.FirstByType(packaging.RelTypeDrawing)
	drawingID := ""
	drawingPath := ""
	if drawingRel != nil {
		drawingID = drawingRel.ID
		drawingPath = packaging.ResolveRelationshipTarget(sheetPath, drawingRel.Target)
		if !strings.HasPrefix(drawingPath, "xl/") {
			drawingPath = "xl/" + strings.TrimPrefix(drawingPath, "/")
		}
	}

	drawingXML, err := ws.loadDrawingXML(drawingPath)
	if err != nil {
		return err
	}
	if drawingXML != "" {
		drawingXML = strings.ReplaceAll(drawingXML, utils.XMLHeader, "")
		drawingXML = strings.TrimSpace(drawingXML)
	}

	nextShapeID := ws.nextDrawingShapeID(drawingXML)
	content, err := ws.buildGraphicContent(sheetPath, kind, title, imagePath, nextShapeID)
	if err != nil {
		return err
	}

	anchorXML := buildAnchorXML(start, end, content)
	if drawingXML == "" {
		drawingXML = buildDrawingXML(anchorXML)
	} else {
		drawingXML = strings.Replace(drawingXML, "</xdr:wsDr>", anchorXML+"</xdr:wsDr>", 1)
	}

	if drawingPath == "" {
		drawingPath = fmt.Sprintf("xl/drawings/drawing%d.xml", ws.index+1)
	}
	if _, err := ws.workbook.pkg.AddPart(drawingPath, packaging.ContentTypeDrawing, []byte(utils.XMLHeader+drawingXML)); err != nil {
		return err
	}
	if drawingID == "" {
		drawingID = rels.NextID()
	}
	rels.AddWithID(drawingID, packaging.RelTypeDrawing, relativeTarget(sheetPath, drawingPath), packaging.TargetModeInternal)
	ws.worksheet.Drawing = &sml.Drawing{ID: drawingID}
	if ws.worksheet.LegacyDrawing != nil && ws.worksheet.LegacyDrawing.ID == "" {
		ws.worksheet.LegacyDrawing.ID = drawingID
	}
	return nil
}

func (ws *worksheetImpl) loadDrawingXML(drawingPath string) (string, error) {
	if drawingPath == "" {
		return "", nil
	}
	part, err := ws.workbook.pkg.GetPart(drawingPath)
	if err != nil {
		return "", nil
	}
	data, err := part.Content()
	if err != nil {
		return "", err
	}
	return string(data), nil
}

func (ws *worksheetImpl) buildGraphicContent(sheetPath string, kind drawingKind, title string, imagePath string, shapeID int) (string, error) {
	drawingPath := ws.ensureDrawingPath()
	drawingRels := ws.workbook.pkg.GetRelationships(drawingPath)
	switch kind {
	case drawingKindChart:
		chartID := shapeID
		chartPath := fmt.Sprintf("xl/charts/chart%d.xml", chartID)
		cs := chart.DefaultChartSpace()
		data, err := utils.MarshalXMLWithHeader(cs)
		if err != nil {
			return "", err
		}
		if _, err := ws.workbook.pkg.AddPart(chartPath, packaging.ContentTypeChart, data); err != nil {
			return "", err
		}
		relID := drawingRels.NextID()
		drawingRels.AddWithID(relID, packaging.RelTypeChart, relativeTarget(drawingPath, chartPath), packaging.TargetModeInternal)
		return fmt.Sprintf(`<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="%s"><c:chart r:id="%s"/></a:graphicData></a:graphic></xdr:graphicFrame>`,
			shapeID,
			escapeXML(chartTitle(title, chartID)),
			packaging.NSDrawingMLChart,
			relID,
		), nil
	case drawingKindDiagram:
		id := shapeID
		dataModel := diagram.DefaultDataModel()
		data, err := utils.MarshalXMLWithHeader(dataModel)
		if err != nil {
			return "", err
		}
		layoutData, err := utils.MarshalXMLWithHeader(diagram.DefaultLayoutDef())
		if err != nil {
			return "", err
		}
		styleData, err := utils.MarshalXMLWithHeader(diagram.DefaultStyleDef())
		if err != nil {
			return "", err
		}
		colorsData, err := utils.MarshalXMLWithHeader(diagram.DefaultColorsDef())
		if err != nil {
			return "", err
		}
		dataPath := fmt.Sprintf("xl/diagrams/data%d.xml", id)
		layoutPath := fmt.Sprintf("xl/diagrams/layout%d.xml", id)
		stylePath := fmt.Sprintf("xl/diagrams/style%d.xml", id)
		colorsPath := fmt.Sprintf("xl/diagrams/colors%d.xml", id)
		if _, err := ws.workbook.pkg.AddPart(dataPath, packaging.ContentTypeDiagramData, data); err != nil {
			return "", err
		}
		if _, err := ws.workbook.pkg.AddPart(layoutPath, packaging.ContentTypeDiagramLayout, layoutData); err != nil {
			return "", err
		}
		if _, err := ws.workbook.pkg.AddPart(stylePath, packaging.ContentTypeDiagramStyle, styleData); err != nil {
			return "", err
		}
		if _, err := ws.workbook.pkg.AddPart(colorsPath, packaging.ContentTypeDiagramColors, colorsData); err != nil {
			return "", err
		}
		dataRelID := drawingRels.NextID()
		drawingRels.AddWithID(dataRelID, packaging.RelTypeDiagramData, relativeTarget(drawingPath, dataPath), packaging.TargetModeInternal)
		layoutRelID := drawingRels.NextID()
		drawingRels.AddWithID(layoutRelID, packaging.RelTypeDiagramLayout, relativeTarget(drawingPath, layoutPath), packaging.TargetModeInternal)
		styleRelID := drawingRels.NextID()
		drawingRels.AddWithID(styleRelID, packaging.RelTypeDiagramStyle, relativeTarget(drawingPath, stylePath), packaging.TargetModeInternal)
		colorsRelID := drawingRels.NextID()
		drawingRels.AddWithID(colorsRelID, packaging.RelTypeDiagramColors, relativeTarget(drawingPath, colorsPath), packaging.TargetModeInternal)
		return fmt.Sprintf(`<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="%s"><dgm:relIds r:dm="%s" r:lo="%s" r:cs="%s" r:qs="%s"/></a:graphicData></a:graphic></xdr:graphicFrame>`,
			shapeID,
			escapeXML(diagramTitle(title, id)),
			packaging.NSDrawingMLDiagram,
			dataRelID, layoutRelID, colorsRelID, styleRelID,
		), nil
	case drawingKindPicture:
		if imagePath == "" {
			return "", utils.ErrPathNotSet
		}
		cleanPath := filepath.Clean(imagePath)
		data, err := os.ReadFile(cleanPath)
		if err != nil {
			return "", err
		}
		ext := strings.TrimPrefix(strings.ToLower(path.Ext(cleanPath)), ".")
		contentType := packaging.ContentTypePNG
		switch ext {
		case "jpg", "jpeg":
			contentType = packaging.ContentTypeJPEG
		case "gif":
			contentType = packaging.ContentTypeGIF
		case "bmp":
			contentType = packaging.ContentTypeBMP
		case "tif", "tiff":
			contentType = packaging.ContentTypeTIFF
		}
		imageName := fmt.Sprintf("xl/media/image%d.%s", shapeID, ext)
		if _, err := ws.workbook.pkg.AddPart(imageName, contentType, data); err != nil {
			return "", err
		}
		relID := drawingRels.NextID()
		drawingRels.AddWithID(relID, packaging.RelTypeImage, relativeTarget(drawingPath, imageName), packaging.TargetModeInternal)
		return fmt.Sprintf(`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="%d" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`,
			shapeID,
			relID,
		), nil
	default:
		return "", utils.ErrInvalidValue
	}
}

func (ws *worksheetImpl) ensureDrawingPath() string {
	return fmt.Sprintf("xl/drawings/drawing%d.xml", ws.index+1)
}

func buildDrawingXML(anchor string) string {
	return fmt.Sprintf(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="%s" xmlns:c="%s" xmlns:dgm="%s">%s</xdr:wsDr>`,
		packaging.NSOfficeDocRels,
		packaging.NSDrawingMLChart,
		packaging.NSDrawingMLDiagram,
		anchor,
	)
}

func buildAnchorXML(start, end utils.CellRef, content string) string {
	return fmt.Sprintf(`<xdr:twoCellAnchor><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>%s<xdr:clientData/></xdr:twoCellAnchor>`,
		start.Col-1,
		start.Row-1,
		end.Col-1,
		end.Row-1,
		content,
	)
}

func (ws *worksheetImpl) nextDrawingShapeID(xmlStr string) int {
	maxID := 0
	scan := xmlStr
	for {
		idx := strings.Index(scan, "cNvPr id=\"")
		if idx == -1 {
			break
		}
		scan = scan[idx+len("cNvPr id=\""):]
		end := strings.Index(scan, `"`)
		if end == -1 {
			break
		}
		if id, err := strconv.Atoi(scan[:end]); err == nil && id > maxID {
			maxID = id
		}
		scan = scan[end+1:]
	}
	if maxID == 0 {
		return 1
	}
	return maxID + 1
}

func chartTitle(title string, id int) string {
	if title != "" {
		return title
	}
	return fmt.Sprintf("Chart %d", id)
}

func diagramTitle(title string, id int) string {
	if title != "" {
		return title
	}
	return fmt.Sprintf("Diagram %d", id)
}

func escapeXML(text string) string {
	replacer := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;", "'", "&apos;")
	return replacer.Replace(text)
}
