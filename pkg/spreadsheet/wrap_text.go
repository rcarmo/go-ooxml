package spreadsheet

import "github.com/rcarmo/go-ooxml/pkg/ooxml/sml"

// SetCellWrapText creates a direct-alignment XF from the cell's current XF.
// It does not change the legacy CellStyle interface or rewrite other cells.
func SetCellWrapText(cell Cell, wrapped bool) error {
	c, ok := cell.(*cellImpl)
	if !ok || c == nil || c.cell == nil || c.worksheet == nil || c.worksheet.workbook == nil {
		return ErrInvalidValue
	}
	styles, ok := c.worksheet.workbook.Styles().(*stylesImpl)
	if !ok || styles == nil || styles.stylesheet == nil || styles.stylesheet.CellXfs == nil || c.cell.S < 0 || c.cell.S >= len(styles.stylesheet.CellXfs.Xf) {
		return ErrInvalidValue
	}
	current := styles.stylesheet.CellXfs.Xf[c.cell.S]
	if current == nil {
		return ErrInvalidValue
	}
	xf := *current
	alignment := sml.Alignment{}
	if xf.Alignment != nil {
		alignment = *xf.Alignment
	}
	alignment.WrapText = boolPtr(wrapped)
	xf.Alignment = &alignment
	xf.ApplyAlignment = boolPtr(true)
	return c.SetStyle(&cellStyleImpl{styles: styles, xf: &xf})
}
