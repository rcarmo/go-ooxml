package spreadsheet

import "github.com/rcarmo/go-ooxml/pkg/ooxml/sml"

// NamedRange represents a defined name in a workbook.
type namedRangeImpl struct {
	workbook    *workbookImpl
	definedName *sml.DefinedName
}

// Name returns the defined name.
func (nr *namedRangeImpl) Name() string {
	if nr == nil || nr.definedName == nil {
		return ""
	}
	return nr.definedName.Name
}

// RefersTo returns the formula or range reference.
func (nr *namedRangeImpl) RefersTo() string {
	if nr == nil || nr.definedName == nil {
		return ""
	}
	return nr.definedName.Value
}

// SetRefersTo updates the formula or range reference.
func (nr *namedRangeImpl) SetRefersTo(refersTo string) {
	if nr == nil || nr.definedName == nil {
		return
	}
	nr.definedName.Value = refersTo
}

// SheetIndex returns the local sheet index if set.
func (nr *namedRangeImpl) SheetIndex() (int, bool) {
	if nr == nil || nr.definedName == nil || nr.definedName.LocalSheetID == nil {
		return 0, false
	}
	return *nr.definedName.LocalSheetID, true
}

// SetSheetIndex sets the local sheet index.
func (nr *namedRangeImpl) SetSheetIndex(index int) {
	if nr == nil || nr.definedName == nil {
		return
	}
	nr.definedName.LocalSheetID = &index
}

// ClearSheetIndex removes the local sheet index.
func (nr *namedRangeImpl) ClearSheetIndex() {
	if nr == nil || nr.definedName == nil {
		return
	}
	nr.definedName.LocalSheetID = nil
}

// Hidden reports whether the defined name is hidden.
func (nr *namedRangeImpl) Hidden() bool {
	if nr == nil || nr.definedName == nil || nr.definedName.Hidden == nil {
		return false
	}
	return *nr.definedName.Hidden
}

// SetHidden sets whether the defined name is hidden.
func (nr *namedRangeImpl) SetHidden(hidden bool) {
	if nr == nil || nr.definedName == nil {
		return
	}
	nr.definedName.Hidden = &hidden
}
