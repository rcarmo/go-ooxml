package utils

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"
)

// XMLHeader is the standard XML declaration.
const XMLHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

// MarshalXMLWithHeader marshals v to XML with the standard XML declaration.
func MarshalXMLWithHeader(v interface{}) ([]byte, error) {
	data, err := xml.Marshal(v)
	if err != nil {
		return nil, err
	}
	data = normalizeRelationshipPrefixes(data)
	return append([]byte(XMLHeader), data...), nil
}

// MarshalXMLIndentWithHeader marshals v to indented XML with the standard XML declaration.
func MarshalXMLIndentWithHeader(v interface{}, prefix, indent string) ([]byte, error) {
	data, err := xml.MarshalIndent(v, prefix, indent)
	if err != nil {
		return nil, err
	}
	data = normalizeRelationshipPrefixes(data)
	return append([]byte(XMLHeader), data...), nil
}

func normalizeRelationshipPrefixes(data []byte) []byte {
	data = bytes.ReplaceAll(data, []byte("xmlns:relationships="), []byte("xmlns:r="))
	data = bytes.ReplaceAll(data, []byte("relationships:id="), []byte("r:id="))
	data = bytes.ReplaceAll(data, []byte("relationships:embed="), []byte("r:embed="))
	data = bytes.ReplaceAll(data, []byte("relationships:link="), []byte("r:link="))
	isPresentation := bytes.Contains(data, []byte("presentationml/2006/main")) ||
		bytes.Contains(data, []byte("powerpoint/2018/8/main"))
	if !isPresentation {
		return data
	}
	data = bytes.ReplaceAll(data, []byte("<sld xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""))
	data = bytes.ReplaceAll(data, []byte("</sld>"), []byte("</p:sld>"))
	data = bytes.ReplaceAll(data, []byte("<notes xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:notes xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""))
	data = bytes.ReplaceAll(data, []byte("</notes>"), []byte("</p:notes>"))
	data = bytes.ReplaceAll(data, []byte("<notesMaster xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:notesMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""))
	data = bytes.ReplaceAll(data, []byte("</notesMaster>"), []byte("</p:notesMaster>"))
	data = bytes.ReplaceAll(data, []byte("<bg>"), []byte("<p:bg>"))
	data = bytes.ReplaceAll(data, []byte("</bg>"), []byte("</p:bg>"))
	data = bytes.ReplaceAll(data, []byte("<bgRef"), []byte("<p:bgRef"))
	data = bytes.ReplaceAll(data, []byte("</bgRef>"), []byte("</p:bgRef>"))
	data = bytes.ReplaceAll(data, []byte("<a:bgRef"), []byte("<p:bgRef"))
	data = bytes.ReplaceAll(data, []byte("</a:bgRef>"), []byte("</p:bgRef>"))
	data = bytes.ReplaceAll(data, []byte("<bgPr>"), []byte("<p:bgPr>"))
	data = bytes.ReplaceAll(data, []byte("</bgPr>"), []byte("</p:bgPr>"))
	data = bytes.ReplaceAll(data, []byte("<clrMap"), []byte("<p:clrMap"))
	data = bytes.ReplaceAll(data, []byte("</clrMap>"), []byte("</p:clrMap>"))
	notesExtReplacement := bytes.Contains(data, []byte("notesMaster")) || bytes.Contains(data, []byte("notesMasterId"))
	data = bytes.ReplaceAll(data, []byte("<notesStyle>"), []byte("<p:notesStyle>"))
	data = bytes.ReplaceAll(data, []byte("</notesStyle>"), []byte("</p:notesStyle>"))
	if notesExtReplacement {
		data = bytes.ReplaceAll(data, []byte("<extLst>"), []byte("<p:extLst>"))
		data = bytes.ReplaceAll(data, []byte("</extLst>"), []byte("</p:extLst>"))
	}
	if notesExtReplacement {
		data = bytes.ReplaceAll(data, []byte("<p:extLst><ext"), []byte("<p:extLst><p:ext"))
		data = bytes.ReplaceAll(data, []byte("</ext></p:extLst>"), []byte("</p:ext></p:extLst>"))
	} else {
		data = bytes.ReplaceAll(data, []byte("<p:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"), []byte("<a:extLst>"))
		data = bytes.ReplaceAll(data, []byte("<p:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""), []byte("<a:extLst"))
		data = bytes.ReplaceAll(data, []byte("<p:extLst>"), []byte("<a:extLst>"))
		data = bytes.ReplaceAll(data, []byte("</p:extLst>"), []byte("</a:extLst>"))
	}
	data = bytes.ReplaceAll(data, []byte("<p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\">&lt;p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"2696953737\"/&gt;"), []byte("<p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"2696953737\"/>"))
	data = bytes.ReplaceAll(data, []byte("<schemeClr"), []byte("<a:schemeClr"))
	data = bytes.ReplaceAll(data, []byte("</schemeClr>"), []byte("</a:schemeClr>"))
	data = bytes.ReplaceAll(data, []byte("<cSld>"), []byte("<p:cSld>"))
	data = bytes.ReplaceAll(data, []byte("</cSld>"), []byte("</p:cSld>"))
	data = bytes.ReplaceAll(data, []byte("<spTree "), []byte("<p:spTree "))
	data = bytes.ReplaceAll(data, []byte("<spTree>"), []byte("<p:spTree>"))
	data = bytes.ReplaceAll(data, []byte("</spTree>"), []byte("</p:spTree>"))
	data = bytes.ReplaceAll(data, []byte("<p:spTree xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:spTree>"))
	data = bytes.ReplaceAll(data, []byte("<p:spTree xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:spTree"))
	data = bytes.ReplaceAll(data, []byte("<nvGrpSpPr "), []byte("<p:nvGrpSpPr "))
	data = bytes.ReplaceAll(data, []byte("<nvGrpSpPr>"), []byte("<p:nvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("</nvGrpSpPr>"), []byte("</p:nvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<cNvGrpSpPr "), []byte("<p:cNvGrpSpPr "))
	data = bytes.ReplaceAll(data, []byte("<cNvGrpSpPr>"), []byte("<p:cNvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("</cNvGrpSpPr>"), []byte("</p:cNvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:nvGrpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:nvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:nvGrpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:nvGrpSpPr"))
	data = bytes.ReplaceAll(data, []byte("<p:cNvGrpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:cNvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:cNvGrpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:cNvGrpSpPr"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:grpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:grpSpPr"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:grpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:grpSpPr"))
	data = bytes.ReplaceAll(data, []byte("<p:nvSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:nvSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:nvSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:nvSpPr"))
	data = bytes.ReplaceAll(data, []byte("<p:spPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:spPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:spPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:spPr"))
	data = bytes.ReplaceAll(data, []byte("<nvGrpSpPr>"), []byte("<p:nvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("</nvGrpSpPr>"), []byte("</p:nvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<cNvGrpSpPr>"), []byte("<p:cNvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("</cNvGrpSpPr>"), []byte("</p:cNvGrpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<chOff "), []byte("<a:chOff "))
	data = bytes.ReplaceAll(data, []byte("</chOff>"), []byte("</a:chOff>"))
	data = bytes.ReplaceAll(data, []byte("<chExt "), []byte("<a:chExt "))
	data = bytes.ReplaceAll(data, []byte("</chExt>"), []byte("</a:chExt>"))
	data = bytes.ReplaceAll(data, []byte("<sp xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:sp"))
	data = bytes.ReplaceAll(data, []byte("</sp>"), []byte("</p:sp>"))
	data = bytes.ReplaceAll(data, []byte("<nvSpPr>"), []byte("<p:nvSpPr>"))
	data = bytes.ReplaceAll(data, []byte("</nvSpPr>"), []byte("</p:nvSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<cNvPr "), []byte("<p:cNvPr "))
	data = bytes.ReplaceAll(data, []byte("</cNvPr>"), []byte("</p:cNvPr>"))
	data = bytes.ReplaceAll(data, []byte("<cNvSpPr"), []byte("<p:cNvSpPr"))
	data = bytes.ReplaceAll(data, []byte("</cNvSpPr>"), []byte("</p:cNvSpPr>"))
	// Keep spLocks in DrawingML namespace for notes placeholders.
	data = bytes.ReplaceAll(data, []byte("<nvPr>"), []byte("<p:nvPr>"))
	data = bytes.ReplaceAll(data, []byte("</nvPr>"), []byte("</p:nvPr>"))
	data = bytes.ReplaceAll(data, []byte("<ph "), []byte("<p:ph "))
	data = bytes.ReplaceAll(data, []byte("</ph>"), []byte("</p:ph>"))
	data = bytes.ReplaceAll(data, []byte("<spLocks "), []byte("<a:spLocks "))
	data = bytes.ReplaceAll(data, []byte("<spLocks>"), []byte("<a:spLocks>"))
	data = bytes.ReplaceAll(data, []byte("</spLocks>"), []byte("</a:spLocks>"))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "), []byte("<a:spLocks "))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"), []byte("<a:spLocks>"))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""), []byte("<a:spLocks"))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noGrp=\"true\""), []byte("<a:spLocks noGrp=\"1\""))
	data = bytes.ReplaceAll(data, []byte("noGrp=\"true\""), []byte("noGrp=\"1\""))
	data = bytes.ReplaceAll(data, []byte("<spPr>"), []byte("<p:spPr>"))
	data = bytes.ReplaceAll(data, []byte("</spPr>"), []byte("</p:spPr>"))
	data = bytes.ReplaceAll(data, []byte("<grpSpPr"), []byte("<p:grpSpPr"))
	data = bytes.ReplaceAll(data, []byte("</grpSpPr>"), []byte("</p:grpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<p:grpSpPr>"))
	data = bytes.ReplaceAll(data, []byte("<p:grpSpPr xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\""), []byte("<p:grpSpPr"))
	data = bytes.ReplaceAll(data, []byte("<extLst"), []byte("<a:extLst"))
	data = bytes.ReplaceAll(data, []byte("</extLst>"), []byte("</a:extLst>"))
	data = bytes.ReplaceAll(data, []byte("<a:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"), []byte("<a:extLst>"))
	data = bytes.ReplaceAll(data, []byte("<a:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""), []byte("<a:extLst"))
	data = bytes.ReplaceAll(data, []byte("</p:extLst>"), []byte("</a:extLst>"))
	data = bytes.ReplaceAll(data, []byte("<p:extLst "), []byte("<a:extLst "))
	data = bytes.ReplaceAll(data, []byte("<clrMapOvr>"), []byte("<p:clrMapOvr>"))
	data = bytes.ReplaceAll(data, []byte("</clrMapOvr>"), []byte("</p:clrMapOvr>"))
	data = bytes.ReplaceAll(data, []byte("<masterClrMapping"), []byte("<a:masterClrMapping"))
	data = bytes.ReplaceAll(data, []byte("</masterClrMapping>"), []byte("</a:masterClrMapping>"))
	data = bytes.ReplaceAll(data, []byte("xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\""), []byte("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "), []byte("<a:spLocks "))
	data = bytes.ReplaceAll(data, []byte("<a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"), []byte("<a:spLocks>"))
	data = bytes.ReplaceAll(data, []byte("<blipFill>"), []byte("<a:blipFill>"))
	data = bytes.ReplaceAll(data, []byte("</blipFill>"), []byte("</a:blipFill>"))
	data = bytes.ReplaceAll(data, []byte("<blipFill "), []byte("<a:blipFill "))
	data = bytes.ReplaceAll(data, []byte("<blip "), []byte("<a:blip "))
	data = bytes.ReplaceAll(data, []byte("</blip>"), []byte("</a:blip>"))
	data = bytes.ReplaceAll(data, []byte("<stretch>"), []byte("<a:stretch>"))
	data = bytes.ReplaceAll(data, []byte("<fillRect>"), []byte("<a:fillRect>"))
	data = bytes.ReplaceAll(data, []byte("<fillRect "), []byte("<a:fillRect "))
	data = bytes.ReplaceAll(data, []byte("<fillRect xmlns:a="), []byte("<a:fillRect xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</fillRect>"), []byte("</a:fillRect>"))
	data = bytes.ReplaceAll(data, []byte("</stretch>"), []byte("</a:stretch>"))
	data = bytes.ReplaceAll(data, []byte("<xfrm>"), []byte("<a:xfrm>"))
	data = bytes.ReplaceAll(data, []byte("</xfrm>"), []byte("</a:xfrm>"))
	data = bytes.ReplaceAll(data, []byte("<xfrm "), []byte("<a:xfrm "))
	data = bytes.ReplaceAll(data, []byte("</nvGraphicFramePr><a:xfrm"), []byte("</nvGraphicFramePr><p:xfrm"))
	data = bytes.ReplaceAll(data, []byte("</a:xfrm><graphic"), []byte("</p:xfrm><graphic"))
	data = bytes.ReplaceAll(data, []byte("</a:xfrm><a:graphic"), []byte("</p:xfrm><a:graphic"))
	data = bytes.ReplaceAll(data, []byte("<off "), []byte("<a:off "))
	data = bytes.ReplaceAll(data, []byte("</off>"), []byte("</a:off>"))
	data = bytes.ReplaceAll(data, []byte("<ext "), []byte("<a:ext "))
	data = bytes.ReplaceAll(data, []byte("</ext>"), []byte("</a:ext>"))
	data = bytes.ReplaceAll(data, []byte("<prstGeom "), []byte("<a:prstGeom "))
	data = bytes.ReplaceAll(data, []byte("</prstGeom>"), []byte("</a:prstGeom>"))
	data = bytes.ReplaceAll(data, []byte("<picLocks "), []byte("<a:picLocks "))
	data = bytes.ReplaceAll(data, []byte("</picLocks>"), []byte("</a:picLocks>"))
	data = bytes.ReplaceAll(data, []byte("<graphic xmlns:a="), []byte("<a:graphic xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</graphic>"), []byte("</a:graphic>"))
	data = bytes.ReplaceAll(data, []byte("<graphicData xmlns:a="), []byte("<a:graphicData xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</graphicData>"), []byte("</a:graphicData>"))
	data = bytes.ReplaceAll(data, []byte("<tbl xmlns:a="), []byte("<a:tbl xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</tbl>"), []byte("</a:tbl>"))
	data = bytes.ReplaceAll(data, []byte("<tblPr xmlns:a="), []byte("<a:tblPr xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</tblPr>"), []byte("</a:tblPr>"))
	data = bytes.ReplaceAll(data, []byte("<tblGrid xmlns:a="), []byte("<a:tblGrid xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</tblGrid>"), []byte("</a:tblGrid>"))
	data = bytes.ReplaceAll(data, []byte("<gridCol xmlns:a="), []byte("<a:gridCol xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</gridCol>"), []byte("</a:gridCol>"))
	data = bytes.ReplaceAll(data, []byte("<tr xmlns:a="), []byte("<a:tr xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</tr>"), []byte("</a:tr>"))
	data = bytes.ReplaceAll(data, []byte("<tc xmlns:a="), []byte("<a:tc xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</tc>"), []byte("</a:tc>"))
	if !bytes.Contains(data, []byte("http://schemas.microsoft.com/office/powerpoint/2018/8/main")) {
		data = bytes.ReplaceAll(data, []byte("<txBody>"), []byte("<p:txBody>"))
		data = bytes.ReplaceAll(data, []byte("<txBody "), []byte("<p:txBody "))
		data = bytes.ReplaceAll(data, []byte("</txBody>"), []byte("</p:txBody>"))
		data = normalizeTableCellTxBody(data)
	}
	data = bytes.ReplaceAll(data, []byte("<bodyPr xmlns:a="), []byte("<a:bodyPr xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("<normAutofit>"), []byte("<a:normAutofit>"))
	data = bytes.ReplaceAll(data, []byte("<normAutofit "), []byte("<a:normAutofit "))
	data = bytes.ReplaceAll(data, []byte("</normAutofit>"), []byte("</a:normAutofit>"))
	data = bytes.ReplaceAll(data, []byte("<noAutofit>"), []byte("<a:noAutofit>"))
	data = bytes.ReplaceAll(data, []byte("<noAutofit "), []byte("<a:noAutofit "))
	data = bytes.ReplaceAll(data, []byte("</noAutofit>"), []byte("</a:noAutofit>"))
	data = bytes.ReplaceAll(data, []byte("<spAutoFit>"), []byte("<a:spAutoFit>"))
	data = bytes.ReplaceAll(data, []byte("<spAutoFit "), []byte("<a:spAutoFit "))
	data = bytes.ReplaceAll(data, []byte("</spAutoFit>"), []byte("</a:spAutoFit>"))
	data = bytes.ReplaceAll(data, []byte("</bodyPr>"), []byte("</a:bodyPr>"))
	data = bytes.ReplaceAll(data, []byte("<lstStyle xmlns:a="), []byte("<a:lstStyle xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</lstStyle>"), []byte("</a:lstStyle>"))
	data = bytes.ReplaceAll(data, []byte("<pPr>"), []byte("<a:pPr>"))
	data = bytes.ReplaceAll(data, []byte("<pPr "), []byte("<a:pPr "))
	data = bytes.ReplaceAll(data, []byte("</pPr>"), []byte("</a:pPr>"))
	data = bytes.ReplaceAll(data, []byte("<rPr>"), []byte("<a:rPr>"))
	data = bytes.ReplaceAll(data, []byte("<rPr "), []byte("<a:rPr "))
	data = bytes.ReplaceAll(data, []byte("</rPr>"), []byte("</a:rPr>"))
	data = bytes.ReplaceAll(data, []byte("<solidFill>"), []byte("<a:solidFill>"))
	data = bytes.ReplaceAll(data, []byte("<solidFill "), []byte("<a:solidFill "))
	data = bytes.ReplaceAll(data, []byte("</solidFill>"), []byte("</a:solidFill>"))
	data = bytes.ReplaceAll(data, []byte("<srgbClr>"), []byte("<a:srgbClr>"))
	data = bytes.ReplaceAll(data, []byte("<srgbClr "), []byte("<a:srgbClr "))
	data = bytes.ReplaceAll(data, []byte("</srgbClr>"), []byte("</a:srgbClr>"))
	data = bytes.ReplaceAll(data, []byte("<buAutoNum>"), []byte("<a:buAutoNum>"))
	data = bytes.ReplaceAll(data, []byte("<buAutoNum "), []byte("<a:buAutoNum "))
	data = bytes.ReplaceAll(data, []byte("</buAutoNum>"), []byte("</a:buAutoNum>"))
	data = bytes.ReplaceAll(data, []byte("<buChar>"), []byte("<a:buChar>"))
	data = bytes.ReplaceAll(data, []byte("<buChar "), []byte("<a:buChar "))
	data = bytes.ReplaceAll(data, []byte("</buChar>"), []byte("</a:buChar>"))
	data = bytes.ReplaceAll(data, []byte("<buNone>"), []byte("<a:buNone>"))
	data = bytes.ReplaceAll(data, []byte("<buNone "), []byte("<a:buNone "))
	data = bytes.ReplaceAll(data, []byte("</buNone>"), []byte("</a:buNone>"))
	data = bytes.ReplaceAll(data, []byte("<buBlip>"), []byte("<a:buBlip>"))
	data = bytes.ReplaceAll(data, []byte("<buBlip "), []byte("<a:buBlip "))
	data = bytes.ReplaceAll(data, []byte("</buBlip>"), []byte("</a:buBlip>"))
	data = bytes.ReplaceAll(data, []byte("<latin>"), []byte("<a:latin>"))
	data = bytes.ReplaceAll(data, []byte("<latin "), []byte("<a:latin "))
	data = bytes.ReplaceAll(data, []byte("</latin>"), []byte("</a:latin>"))
	data = bytes.ReplaceAll(data, []byte("<ea>"), []byte("<a:ea>"))
	data = bytes.ReplaceAll(data, []byte("<ea "), []byte("<a:ea "))
	data = bytes.ReplaceAll(data, []byte("</ea>"), []byte("</a:ea>"))
	data = bytes.ReplaceAll(data, []byte("<cs>"), []byte("<a:cs>"))
	data = bytes.ReplaceAll(data, []byte("<cs "), []byte("<a:cs "))
	data = bytes.ReplaceAll(data, []byte("</cs>"), []byte("</a:cs>"))
	data = bytes.ReplaceAll(data, []byte("<p xmlns:a="), []byte("<a:p xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</p>"), []byte("</a:p>"))
	data = bytes.ReplaceAll(data, []byte("<p>"), []byte("<a:p>"))
	data = bytes.ReplaceAll(data, []byte("<p "), []byte("<a:p "))
	data = bytes.ReplaceAll(data, []byte("<r xmlns:a="), []byte("<a:r xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</r>"), []byte("</a:r>"))
	data = bytes.ReplaceAll(data, []byte("<r>"), []byte("<a:r>"))
	data = bytes.ReplaceAll(data, []byte("<r "), []byte("<a:r "))
	data = bytes.ReplaceAll(data, []byte("<t xmlns:a="), []byte("<a:t xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</t>"), []byte("</a:t>"))
	data = bytes.ReplaceAll(data, []byte("<defPPr>"), []byte("<a:defPPr>"))
	data = bytes.ReplaceAll(data, []byte("<defPPr "), []byte("<a:defPPr "))
	data = bytes.ReplaceAll(data, []byte("</defPPr>"), []byte("</a:defPPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl1pPr>"), []byte("<a:lvl1pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl1pPr "), []byte("<a:lvl1pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl1pPr>"), []byte("</a:lvl1pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl2pPr>"), []byte("<a:lvl2pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl2pPr "), []byte("<a:lvl2pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl2pPr>"), []byte("</a:lvl2pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl3pPr>"), []byte("<a:lvl3pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl3pPr "), []byte("<a:lvl3pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl3pPr>"), []byte("</a:lvl3pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl4pPr>"), []byte("<a:lvl4pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl4pPr "), []byte("<a:lvl4pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl4pPr>"), []byte("</a:lvl4pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl5pPr>"), []byte("<a:lvl5pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl5pPr "), []byte("<a:lvl5pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl5pPr>"), []byte("</a:lvl5pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl6pPr>"), []byte("<a:lvl6pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl6pPr "), []byte("<a:lvl6pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl6pPr>"), []byte("</a:lvl6pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl7pPr>"), []byte("<a:lvl7pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl7pPr "), []byte("<a:lvl7pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl7pPr>"), []byte("</a:lvl7pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl8pPr>"), []byte("<a:lvl8pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl8pPr "), []byte("<a:lvl8pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl8pPr>"), []byte("</a:lvl8pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl9pPr>"), []byte("<a:lvl9pPr>"))
	data = bytes.ReplaceAll(data, []byte("<lvl9pPr "), []byte("<a:lvl9pPr "))
	data = bytes.ReplaceAll(data, []byte("</lvl9pPr>"), []byte("</a:lvl9pPr>"))
	data = bytes.ReplaceAll(data, []byte("<t>"), []byte("<a:t>"))
	data = bytes.ReplaceAll(data, []byte("<t "), []byte("<a:t "))
	data = bytes.ReplaceAll(data, []byte("<endParaRPr xmlns:a="), []byte("<a:endParaRPr xmlns:a="))
	data = bytes.ReplaceAll(data, []byte("</endParaRPr>"), []byte("</a:endParaRPr>"))
	data = bytes.ReplaceAll(data, []byte("<endParaRPr>"), []byte("<a:endParaRPr>"))
	data = bytes.ReplaceAll(data, []byte("<endParaRPr "), []byte("<a:endParaRPr "))
	data = bytes.ReplaceAll(data, []byte("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""), []byte("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""))
	data = bytes.ReplaceAll(data, []byte("<presentation xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"), []byte("<presentation xmlns=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"))
	data = bytes.ReplaceAll(data, []byte("<notesMasterIdLst>"), []byte("<p:notesMasterIdLst>"))
	data = bytes.ReplaceAll(data, []byte("</notesMasterIdLst>"), []byte("</p:notesMasterIdLst>"))
	data = bytes.ReplaceAll(data, []byte("<notesMasterId "), []byte("<p:notesMasterId "))
	data = bytes.ReplaceAll(data, []byte("</notesMasterId>"), []byte("</p:notesMasterId>"))
	data = bytes.ReplaceAll(data, []byte("<masterClrMapping"), []byte("<a:masterClrMapping"))
	data = bytes.ReplaceAll(data, []byte("</masterClrMapping>"), []byte("</a:masterClrMapping>"))
	if bytes.Contains(data, []byte("a:")) {
		data = ensureRootNamespace(data, "xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")
	}
	if bytes.Contains(data, []byte("p:")) {
		data = ensureRootNamespace(data, "xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main")
	}
	if bytes.Contains(data, []byte("a:ext")) {
		data = ensureRootNamespace(data, "xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")
	}
	return data
}

func ensureRootNamespace(data []byte, attr, value string) []byte {
	start := bytes.IndexByte(data, '<')
	if start == -1 {
		return data
	}
	end := bytes.IndexByte(data[start:], '>')
	if end == -1 {
		return data
	}
	end += start
	if bytes.Contains(data[start:end], []byte(attr+"=\"")) {
		return data
	}
	insert := []byte(" " + attr + "=\"" + value + "\"")
	return append(data[:end], append(insert, data[end:]...)...)
}

func normalizeTableCellTxBody(data []byte) []byte {
	if !bytes.Contains(data, []byte("<a:tc")) || !bytes.Contains(data, []byte("<p:txBody")) {
		return data
	}
	data = bytes.ReplaceAll(data, []byte("<a:tc xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:txBody"), []byte("<a:tc xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:txBody"))
	data = bytes.ReplaceAll(data, []byte("<a:tc><p:txBody"), []byte("<a:tc><a:txBody"))
	data = bytes.ReplaceAll(data, []byte("<p:txBody"), []byte("<a:txBody"))
	data = bytes.ReplaceAll(data, []byte("</p:txBody>"), []byte("</a:txBody>"))
	data = bytes.ReplaceAll(data, []byte("</p:spPr><a:txBody"), []byte("</p:spPr><p:txBody"))
	data = bytes.ReplaceAll(data, []byte("</p:style><a:txBody"), []byte("</p:style><p:txBody"))
	data = bytes.ReplaceAll(data, []byte("</a:txBody></p:"), []byte("</p:txBody></p:"))
	return data
}

// UnmarshalXML unmarshals XML data, stripping BOM if present.
func UnmarshalXML(data []byte, v interface{}) error {
	// Strip UTF-8 BOM if present
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
	return xml.Unmarshal(data, v)
}

// NewXMLDecoder creates an XML decoder that handles common OOXML quirks.
func NewXMLDecoder(r io.Reader) *xml.Decoder {
	d := xml.NewDecoder(r)
	d.CharsetReader = func(charset string, input io.Reader) (io.Reader, error) {
		// OOXML files are always UTF-8
		return input, nil
	}
	return d
}

// EscapeXMLText escapes special characters in XML text content.
func EscapeXMLText(s string) string {
	var buf strings.Builder
	_ = xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

// BoolPtr returns a pointer to a bool value.
func BoolPtr(v bool) *bool {
	return &v
}

// IntPtr returns a pointer to an int value.
func IntPtr(v int) *int {
	return &v
}

// Int64Ptr returns a pointer to an int64 value.
func Int64Ptr(v int64) *int64 {
	return &v
}

// StringPtr returns a pointer to a string value.
func StringPtr(v string) *string {
	return &v
}

// Float64Ptr returns a pointer to a float64 value.
func Float64Ptr(v float64) *float64 {
	return &v
}

// DerefBool returns the value of a bool pointer, or the default if nil.
func DerefBool(p *bool, def bool) bool {
	if p == nil {
		return def
	}
	return *p
}

// DerefString returns the value of a string pointer, or the default if nil.
func DerefString(p *string, def string) string {
	if p == nil {
		return def
	}
	return *p
}

// DerefInt returns the value of an int pointer, or the default if nil.
func DerefInt(p *int, def int) int {
	if p == nil {
		return def
	}
	return *p
}
