package spreadsheet

import (
	"strconv"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
)

// CellType represents the type of cell value.
type CellType int

const (
	// CellTypeEmpty indicates an empty cell.
	CellTypeEmpty CellType = iota
	// CellTypeString indicates a string cell.
	CellTypeString
	// CellTypeNumber indicates a numeric cell.
	CellTypeNumber
	// CellTypeBoolean indicates a boolean cell.
	CellTypeBoolean
	// CellTypeDate indicates a date cell.
	CellTypeDate
	// CellTypeFormula indicates a formula cell.
	CellTypeFormula
	// CellTypeError indicates an error cell.
	CellTypeError
)

// Cell represents a cell in a worksheet.
type cellImpl struct {
	worksheet *worksheetImpl
	cell      *sml.Cell
	row       int
	col       int
}

// Reference returns the cell reference (e.g., "A1").
func (c *cellImpl) Reference() string {
	return c.cell.R
}

// Row returns the 1-based row number.
func (c *cellImpl) Row() int {
	return c.row
}

// Column returns the 1-based column number.
func (c *cellImpl) Column() int {
	return c.col
}

// Type returns the cell type.
func (c *cellImpl) Type() CellType {
	if c.cell.F != nil {
		return CellTypeFormula
	}

	switch c.cell.T {
	case sml.CellTypeBool:
		return CellTypeBoolean
	case sml.CellTypeError:
		return CellTypeError
	case sml.CellTypeSharedString, sml.CellTypeString, sml.CellTypeInlineString:
		return CellTypeString
	case sml.CellTypeNumber, "":
		if c.cell.V == "" {
			return CellTypeEmpty
		}
		return CellTypeNumber
	default:
		return CellTypeEmpty
	}
}

// =============================================================================
// Value getters
// =============================================================================

// Value returns the cell value as an interface{}.
func (c *cellImpl) Value() interface{} {
	switch c.Type() {
	case CellTypeString:
		return c.String()
	case CellTypeNumber:
		f, _ := c.Float64()
		return f
	case CellTypeBoolean:
		b, _ := c.Bool()
		return b
	case CellTypeFormula:
		return c.String()
	default:
		return nil
	}
}

// String returns the cell value as a string.
func (c *cellImpl) String() string {
	switch c.cell.T {
	case sml.CellTypeSharedString:
		// Look up in shared strings table
		idx, err := strconv.Atoi(c.cell.V)
		if err != nil {
			return c.cell.V
		}
		return c.worksheet.workbook.getSharedString(idx)
	case sml.CellTypeInlineString:
		if c.cell.Is != nil {
			return c.cell.Is.T
		}
		return ""
	default:
		return c.cell.V
	}
}

// Float64 returns the cell value as a float64.
func (c *cellImpl) Float64() (float64, error) {
	return strconv.ParseFloat(c.cell.V, 64)
}

// Int returns the cell value as an int.
func (c *cellImpl) Int() (int, error) {
	f, err := c.Float64()
	if err != nil {
		return 0, err
	}
	return int(f), nil
}

// Bool returns the cell value as a boolean.
func (c *cellImpl) Bool() (bool, error) {
	if c.cell.V == "1" || c.cell.V == "true" || c.cell.V == "TRUE" {
		return true, nil
	}
	if c.cell.V == "0" || c.cell.V == "false" || c.cell.V == "FALSE" {
		return false, nil
	}
	return false, ErrInvalidValue
}

// Time returns the cell value as a time.Time.
// Excel stores dates as numbers (days since 1900-01-01).
func (c *cellImpl) Time() (time.Time, error) {
	f, err := c.Float64()
	if err != nil {
		return time.Time{}, err
	}

	// Excel epoch is December 30, 1899 (accounting for the 1900 leap year bug)
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	days := int(f)
	fraction := f - float64(days)

	t := epoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(fraction * float64(24*time.Hour)))

	return t, nil
}

// =============================================================================
// Value setters
// =============================================================================

// SetValue sets the cell value.
func (c *cellImpl) SetValue(v interface{}) error {
	if v == nil {
		c.cell.V = ""
		c.cell.T = ""
		c.cell.F = nil
		return nil
	}

	switch val := v.(type) {
	case string:
		return c.setString(val)
	case int:
		c.cell.V = strconv.Itoa(val)
		c.cell.T = ""
	case int64:
		c.cell.V = strconv.FormatInt(val, 10)
		c.cell.T = ""
	case float64:
		c.cell.V = strconv.FormatFloat(val, 'f', -1, 64)
		c.cell.T = ""
	case bool:
		if val {
			c.cell.V = "1"
		} else {
			c.cell.V = "0"
		}
		c.cell.T = sml.CellTypeBool
	case time.Time:
		// Convert to Excel serial date
		epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
		days := val.Sub(epoch).Hours() / 24
		c.cell.V = strconv.FormatFloat(days, 'f', -1, 64)
		c.cell.T = ""
	default:
		return c.setString(valueToString(v))
	}

	return nil
}

func (c *cellImpl) setString(s string) error {
	// Use shared strings for efficiency
	idx := c.worksheet.workbook.addSharedString(s)
	c.cell.V = strconv.Itoa(idx)
	c.cell.T = sml.CellTypeSharedString
	return nil
}

// =============================================================================
// Formula
// =============================================================================

// Formula returns the cell formula (without leading =).
func (c *cellImpl) Formula() string {
	if c.cell.F == nil {
		return ""
	}
	return c.cell.F.Content
}

// SetFormula sets the cell formula.
func (c *cellImpl) SetFormula(formula string) error {
	c.cell.F = &sml.Formula{Content: formula}
	c.cell.T = ""
	if c.cell.V == "" {
		c.cell.V = "0"
	}
	return nil
}

// HasFormula returns true if the cell has a formula.
func (c *cellImpl) HasFormula() bool {
	return c.cell.F != nil && c.cell.F.Content != ""
}

// =============================================================================
// Style (placeholder for future implementation)
// =============================================================================

// Style returns the cell style for this cell.
func (c *cellImpl) Style() CellStyle {
	if c == nil || c.worksheet == nil || c.worksheet.workbook == nil {
		return nil
	}
	return c.worksheet.workbook.Styles().Style()
}

// SetStyle applies a cell style to this cell.
func (c *cellImpl) SetStyle(style CellStyle) error {
	if c == nil || style == nil {
		return ErrInvalidValue
	}
	if cs, ok := style.(*cellStyleImpl); ok {
		c.cell.S = cs.finalize()
		return nil
	}
	return ErrInvalidValue
}

// NumberFormat returns the number format for the cell.
func (c *cellImpl) NumberFormat() string {
	if c == nil || c.worksheet == nil || c.worksheet.workbook == nil {
		return ""
	}
	styles, ok := c.worksheet.workbook.Styles().(*stylesImpl)
	if !ok || styles == nil || styles.stylesheet == nil || styles.stylesheet.CellXfs == nil {
		return ""
	}
	if c.cell.S < 0 || c.cell.S >= len(styles.stylesheet.CellXfs.Xf) {
		return ""
	}
	xf := styles.stylesheet.CellXfs.Xf[c.cell.S]
	return styles.formatCode(xf.NumFmtID)
}

// SetNumberFormat sets the number format for the cell.
func (c *cellImpl) SetNumberFormat(format string) error {
	if c == nil || c.worksheet == nil || c.worksheet.workbook == nil {
		return ErrInvalidValue
	}
	style := c.worksheet.workbook.Styles().Style()
	style.SetNumberFormat(format)
	if cs, ok := style.(*cellStyleImpl); ok {
		c.cell.S = cs.finalize()
	}
	return nil
}
