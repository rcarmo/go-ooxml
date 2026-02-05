package spreadsheet

var (
	_ Workbook   = (*workbookImpl)(nil)
	_ Worksheet  = (*worksheetImpl)(nil)
	_ Cell       = (*cellImpl)(nil)
	_ Range      = (*rangeImpl)(nil)
	_ Table      = (*tableImpl)(nil)
	_ TableRow   = (*tableRowImpl)(nil)
	_ Row        = (*rowImpl)(nil)
	_ RowIterator = (*rowIterator)(nil)
	_ NamedRange = (*namedRangeImpl)(nil)
	_ Comment    = (*commentImpl)(nil)
	_ Styles     = (*stylesImpl)(nil)
	_ CellStyle  = (*cellStyleImpl)(nil)
)
