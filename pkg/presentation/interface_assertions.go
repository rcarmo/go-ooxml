package presentation

var (
	_ Presentation = (*presentationImpl)(nil)
	_ Slide        = (*slideImpl)(nil)
	_ Shape        = (*shapeImpl)(nil)
	_ TextFrame    = (*textFrameImpl)(nil)
	_ TextParagraph = (*textParagraphImpl)(nil)
	_ TextRun      = (*textRunImpl)(nil)
	_ Table        = (*tableImpl)(nil)
	_ TableRow     = (*tableRowImpl)(nil)
	_ TableCell    = (*tableCellImpl)(nil)
	_ SlideMaster  = (*slideMasterImpl)(nil)
	_ SlideLayout  = (*slideLayoutImpl)(nil)
	_ Comment      = (*commentImpl)(nil)
)
