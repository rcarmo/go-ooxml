// Package packaging provides OPC (Open Packaging Conventions) support
// for reading and writing OOXML package files.
package packaging

// Relationship types
const (
	RelTypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RelTypeStyles         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	RelTypeSettings       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	RelTypeComments       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RelTypeNumbering      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
	RelTypeHeader         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	RelTypeFooter         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	RelTypeWorksheet      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	RelTypeSharedStrings  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
	RelTypeTable          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
	RelTypeSlide          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	RelTypeSlideLayout    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	RelTypeSlideMaster    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	RelTypeNotesSlide     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	RelTypeImage          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelTypeHyperlink      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RelTypeTheme          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	RelTypeFontTable      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	RelTypeWebSettings    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	RelTypeEndnotes       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	RelTypeFootnotes      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	RelTypeCoreProps      = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	RelTypeExtendedProps  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	RelTypeCustomXML      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
	RelTypeChart          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	RelTypeDrawing        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	RelTypeVML            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
	RelTypeTableStyles    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
	RelTypePPTXComments   = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
	RelTypePPTXAuthors    = "http://schemas.microsoft.com/office/2018/10/relationships/authors"
	RelTypePresProps      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
	RelTypeViewProps      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
)

// Content types
const (
	ContentTypeWordDocument     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	ContentTypeWordTemplate     = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
	ContentTypeWorkbook         = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	ContentTypePresentation     = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ContentTypeStyles           = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	ContentTypeWordStyles       = ContentTypeStyles
	ContentTypeSettings         = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	ContentTypeComments         = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
	ContentTypeExcelComments    = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
	ContentTypePPTXComments     = "application/vnd.ms-powerpoint.comments+xml"
	ContentTypePPTXAuthors      = "application/vnd.ms-powerpoint.authors+xml"
	ContentTypePresentationProps = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
	ContentTypePresentationViewProps = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
	ContentTypeNumbering        = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
	ContentTypeHeader           = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	ContentTypeWordHeader       = ContentTypeHeader
	ContentTypeFooter           = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	ContentTypeWordFooter       = ContentTypeFooter
	ContentTypeWorksheet        = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	ContentTypeSharedStrings    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
	ContentTypeSlide            = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ContentTypeSlideLayout      = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ContentTypeSlideMaster      = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ContentTypeNotesSlide       = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
	ContentTypeRelationships    = "application/vnd.openxmlformats-package.relationships+xml"
	ContentTypeCoreProps        = "application/vnd.openxmlformats-package.core-properties+xml"
	ContentTypeExtendedProps    = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	ContentTypeTheme            = "application/vnd.openxmlformats-officedocument.theme+xml"
	ContentTypeExcelStyles      = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	ContentTypeTable            = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
	ContentTypeChart            = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ContentTypeDrawing          = "application/vnd.openxmlformats-officedocument.drawing+xml"
	ContentTypeVML              = "application/vnd.openxmlformats-officedocument.vmlDrawing"
	ContentTypeFontTable        = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
	ContentTypeWebSettings      = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
	ContentTypeEndnotes         = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
	ContentTypeFootnotes        = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
	ContentTypePNG              = "image/png"
	ContentTypeJPEG             = "image/jpeg"
	ContentTypeGIF              = "image/gif"
	ContentTypeBMP              = "image/bmp"
	ContentTypeTIFF             = "image/tiff"
	ContentTypeXML              = "application/xml"
)

// XML Namespaces
const (
	NSWordprocessingML    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	NSSpreadsheetML       = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	NSPresentationML      = "http://schemas.openxmlformats.org/presentationml/2006/main"
	NSDrawingML           = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NSRelationships       = "http://schemas.openxmlformats.org/package/2006/relationships"
	NSContentTypes        = "http://schemas.openxmlformats.org/package/2006/content-types"
	NSDocumentRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NSDublinCore          = "http://purl.org/dc/elements/1.1/"
	NSDublinCoreTerms     = "http://purl.org/dc/terms/"
	NSCoreProperties      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	NSExtendedProperties  = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	NSDrawingMLChart      = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	NSDrawingMLPicture    = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	NSMarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	NSOfficeDocRels       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NSVML                 = "urn:schemas-microsoft-com:vml"
	NSWord14              = "http://schemas.microsoft.com/office/word/2010/wordml"
)

// File paths in OPC package
const (
	ContentTypesPath       = "[Content_Types].xml"
	PackageRelsPath        = "_rels/.rels"
	CorePropertiesPath     = "docProps/core.xml"
	AppPropertiesPath      = "docProps/app.xml"
	WordDocumentPath       = "word/document.xml"
	WordStylesPath         = "word/styles.xml"
	WordSettingsPath       = "word/settings.xml"
	WordNumberingPath      = "word/numbering.xml"
	WordCommentsPath       = "word/comments.xml"
	ExcelWorkbookPath      = "xl/workbook.xml"
	ExcelStylesPath        = "xl/styles.xml"
	ExcelSharedStringsPath = "xl/sharedStrings.xml"
	PresentationPath       = "ppt/presentation.xml"
	PPTXPresentationPath   = PresentationPath // Alias for backward compatibility
	PresentationPropsPath  = "ppt/presProps.xml"
	PresentationViewPropsPath = "ppt/viewProps.xml"
)

// TargetMode indicates whether the relationship target is internal or external.
type TargetMode int

const (
	TargetModeInternal TargetMode = iota
	TargetModeExternal
)

func (tm TargetMode) String() string {
	if tm == TargetModeExternal {
		return "External"
	}
	return ""
}
