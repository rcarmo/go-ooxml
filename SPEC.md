# Go OOXML Library Specification

**Version:** 1.0  
**Status:** Historical Phases 1-5 milestone; native enhancement and canonical behaviour reconciliation incomplete.

Current verification, grouped schema2 fixtures and candidate/release setup are
documented in [docs/testing.md](docs/testing.md). Totals and phase tables below
are historical design records, not a current test inventory or release claim.
**Created:** January 29, 2026  
**Updated:** January 2026  
**Purpose:** Actionable specification for building a Go OOXML manipulation library

---

## extended enhancement work (2026-09-26)

The historical completion table below covers the original implementation scope.
extended enhancement parity is a separate, unfinished track defined in
[the design and source mapping](docs/design/native-enhancements.md).

Library runtime packages retain standard-library-only dependencies. Gherkin/Godog
live in the separate acceptance module; make test-batch runs both modules with
bounded concurrency. Features are tagged planned, implemented or external.
Implemented native cases require strict execution and exact result reconciliation;
planned cases remain visible in the inventory and do not count as passes.

New safe-edit APIs will be additive adapters/concrete types. Existing exported
interfaces must not gain methods without a versioned compatibility decision.
No preservation, transaction or full extended-parity guarantee exists yet.

## Document Purpose

This specification provides a detailed blueprint for implementing a Go library capable of reading, writing, and manipulating Office Open XML (OOXML) documents. The library must support Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) formats with specific focus on the features required by the MCP Office Server.

**Implementation Status:**
| Package | Status | Tests |
|---------|--------|-------|
| `pkg/utils` | ✅ Complete | 45 |
| `pkg/packaging` | ✅ Complete | 26 |
| `pkg/ooxml/*` | ✅ Complete | 29 |
| `pkg/document` | ✅ Complete | 369 |
| `pkg/presentation` | ✅ Complete | 94 |
| `pkg/spreadsheet` | ✅ Complete | 117 |
| `internal/testutil` | ✅ Complete | 14 |
| **Total** | **5/6 Phases** | **694 tests** |

> [!IMPORTANT]
> This specification is designed to be consumed by an AI agent or development team. Follow the structure precisely and implement features in the order specified. Do not deviate from interface definitions without updating this specification first.

---

## Table of Contents

1. [Architecture Overview](#1-architecture-overview)
2. [Package Structure](#2-package-structure)
3. [Common Packages](#3-common-packages)
4. [Word Package (document)](#4-word-package-document)
5. [Excel Package (spreadsheet)](#5-excel-package-spreadsheet)
6. [PowerPoint Package (presentation)](#6-powerpoint-package-presentation)
7. [Testing Requirements](#7-testing-requirements)
8. [Implementation Phases](#8-implementation-phases)
9. [Code Quality Standards](#9-code-quality-standards)

---

## 1. Architecture Overview

### 1.1 Design Principles

| Principle | Description |
|-----------|-------------|
| **Layered Architecture** | Separate OOXML types, packaging, and high-level APIs |
| **Interface-First** | Define interfaces before implementations |
| **Composition Over Inheritance** | Use embedding and composition patterns |
| **Zero External Dependencies** | Standard library only (archive/zip, encoding/xml) |
| **Immutable Where Possible** | Prefer returning new objects over mutation |

### 1.2 Layer Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                     HIGH-LEVEL API                          │
│  document.Document  spreadsheet.Workbook  presentation.Pres │
├─────────────────────────────────────────────────────────────┤
│                    ELEMENT WRAPPERS                         │
│   Paragraph, Run, Table, Cell, Slide, Shape, etc.          │
├─────────────────────────────────────────────────────────────┤
│                    OOXML TYPES (ooxml/)                     │
│   wml.CT_P, sml.CT_Cell, pml.CT_Slide, dml.CT_Shape        │
├─────────────────────────────────────────────────────────────┤
│                    PACKAGING (packaging/)                   │
│   Package, Part, Relationships, ContentTypes               │
├─────────────────────────────────────────────────────────────┤
│                    GO STANDARD LIBRARY                      │
│   archive/zip, encoding/xml, io, path                      │
└─────────────────────────────────────────────────────────────┘
```

---

## 2. Package Structure

```
github.com/[org]/ooxml-go/
├── go.mod
├── go.sum
├── README.md
├── SPEC.md                          # This file
│
├── pkg/
│   ├── packaging/                   # OPC package handling
│   │   ├── package.go              # Package struct, Open/Save
│   │   ├── part.go                 # Part interface and types
│   │   ├── relationship.go         # Relationship handling
│   │   ├── content_types.go        # [Content_Types].xml
│   │   └── constants.go            # URIs, namespaces, content types
│   │
│   ├── ooxml/                       # Low-level XML types
│   │   ├── common/                  # Shared types across formats
│   │   │   ├── shared_strings.go   # Shared string table
│   │   │   └── core_properties.go  # Dublin Core properties
│   │   │
│   │   ├── wml/                     # WordprocessingML types
│   │   │   ├── document.go         # CT_Document, CT_Body
│   │   │   ├── paragraph.go        # CT_P, CT_PPr
│   │   │   ├── run.go              # CT_R, CT_RPr, CT_Text
│   │   │   ├── table.go            # CT_Tbl, CT_Row, CT_Tc
│   │   │   ├── styles.go           # CT_Styles, CT_Style
│   │   │   ├── settings.go         # CT_Settings
│   │   │   ├── comments.go         # CT_Comments, CT_Comment
│   │   │   ├── tracking.go         # CT_TrackChange, CT_Ins, CT_Del
│   │   │   ├── numbering.go        # CT_Numbering
│   │   │   └── sdt.go              # CT_SdtBlock (Content Controls)
│   │   │
│   │   ├── sml/                     # SpreadsheetML types
│   │   │   ├── workbook.go         # CT_Workbook
│   │   │   ├── worksheet.go        # CT_Worksheet, CT_SheetData
│   │   │   ├── cell.go             # CT_Cell, CT_CellFormula
│   │   │   ├── styles.go           # CT_Stylesheet
│   │   │   ├── table.go            # CT_Table
│   │   │   └── comments.go         # CT_Comments
│   │   │
│   │   ├── pml/                     # PresentationML types
│   │   │   ├── presentation.go     # CT_Presentation
│   │   │   ├── slide.go            # CT_Slide, CT_CommonSlideData
│   │   │   ├── slide_layout.go     # CT_SlideLayout
│   │   │   ├── slide_master.go     # CT_SlideMaster
│   │   │   ├── notes.go            # CT_NotesSlide
│   │   │   └── comments.go         # CT_CommentList
│   │   │
│   │   ├── chart/                   # DrawingML chart types
│   │   │   └── chart.go            # CT_ChartSpace, CT_Chart
│   │   │
│   │   ├── diagram/                 # DrawingML diagram types
│   │   │   └── diagram.go          # CT_DataModel, CT_LayoutDef
│   │   │
│   │   ├── theme/                   # DrawingML theme types
│   │   │   └── theme.go            # CT_Theme
│   │   │
│   │   └── dml/                     # DrawingML types (shared)
│   │       ├── shape.go            # CT_Shape, CT_ShapeProperties
│   │       ├── text.go             # CT_TextBody, CT_TextParagraph
│   │       ├── table.go            # CT_Table, CT_TableRow
│   │       ├── picture.go          # CT_Picture
│   │       └── color.go            # CT_Color, CT_SchemeColor
│   │
│   ├── document/                    # Word high-level API
│   │   ├── document.go             # Document struct
│   │   ├── paragraph.go            # Paragraph wrapper
│   │   ├── run.go                  # Run wrapper
│   │   ├── table.go                # Table, Row, Cell wrappers
│   │   ├── section.go              # Section handling
│   │   ├── header_footer.go        # Headers/Footers
│   │   ├── styles.go               # Style management
│   │   ├── comments.go             # Comment handling
│   │   ├── tracking.go             # Track changes
│   │   └── sdt.go                  # Content control handling
│   │
│   ├── spreadsheet/                 # Excel high-level API
│   │   ├── workbook.go             # Workbook struct
│   │   ├── worksheet.go            # Worksheet struct
│   │   ├── cell.go                 # Cell struct
│   │   ├── row.go                  # Row struct
│   │   ├── range.go                # Range operations
│   │   ├── table.go                # Table struct
│   │   ├── styles.go               # Cell styling
│   │   └── comments.go             # Comment handling
│   │
│   ├── presentation/                # PowerPoint high-level API
│   │   ├── presentation.go         # Presentation struct
│   │   ├── slide.go                # Slide struct
│   │   ├── shape.go                # Shape struct
│   │   ├── text_frame.go           # TextFrame, Paragraph, Run
│   │   ├── table.go                # Table handling
│   │   ├── layout.go               # Layout management
│   │   ├── master.go               # Master slide handling
│   │   ├── notes.go                # Notes slide
│   │   └── comments.go             # Comment handling
│   │
│   └── utils/                       # Shared utilities
│       ├── emu.go                  # EMU conversions
│       ├── color.go                # Color parsing/formatting
│       ├── xml.go                  # XML helpers
│       ├── cell_ref.go             # A1-style cell reference parsing
│       └── errors.go               # Custom error types
│
├── internal/
│   └── xmlutil/                     # Internal XML utilities
│       └── namespace.go            # Namespace handling
│
└── references/fixtures-ooxml/       # Shared manifest, grouped inputs and workflows
    ├── word/
    │   ├── simple.docx
    │   ├── with_tables.docx
    │   ├── with_track_changes.docx
    │   ├── with_comments.docx
    │   └── template.docx
    ├── excel/
    │   ├── simple.xlsx
    │   ├── with_tables.xlsx
    │   ├── with_formulas.xlsx
    │   └── multi_sheet.xlsx
    └── pptx/
        ├── simple.pptx
        ├── with_tables.pptx
        ├── with_notes.pptx
        └── template.pptx
```

---

## 3. Common Packages

### 3.1 Package: `packaging`

This package handles OPC (Open Packaging Conventions) - the ZIP-based container format.

#### 3.1.1 Interfaces

```go
// pkg/packaging/interfaces.go

// Package represents an OPC package (ZIP archive with relationships)
type Package interface {
    // Core operations
    Open(path string) error
    OpenReader(r io.ReaderAt, size int64) error
    Save() error
    SaveAs(path string) error
    Close() error
    
    // Part management
    GetPart(uri string) (Part, error)
    AddPart(uri string, contentType string, content []byte) (Part, error)
    DeletePart(uri string) error
    Parts() []Part
    
    // Relationship management
    GetRelationships(sourceURI string) []Relationship
    AddRelationship(sourceURI, targetURI, relType string) Relationship
    GetRelationshipsByType(sourceURI, relType string) []Relationship
    
    // Content types
    GetContentType(uri string) string
}

// Part represents a part within the package
type Part interface {
    URI() string
    ContentType() string
    Content() ([]byte, error)
    SetContent([]byte) error
    Stream() (io.ReadCloser, error)
}

// Relationship represents an OPC relationship
type Relationship interface {
    ID() string
    Type() string
    Target() string
    TargetMode() TargetMode
}
```

#### 3.1.2 Constants

```go
// pkg/packaging/constants.go

// Relationship types
const (
    RelTypeOfficeDocument      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    RelTypeStyles              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    RelTypeSettings            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    RelTypeComments            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    RelTypeNumbering           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    RelTypeHeader              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    RelTypeFooter              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    RelTypeWorksheet           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    RelTypeSharedStrings       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    RelTypeSlide               = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    RelTypeSlideLayout         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    RelTypeSlideMaster         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    RelTypeNotesSlide          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    RelTypeImage               = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

// Content types
const (
    ContentTypeWordDocument    = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    ContentTypeWorkbook        = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    ContentTypePresentation    = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    ContentTypeStyles          = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    ContentTypeComments        = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)

// XML Namespaces
const (
    NSWordprocessingML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NSSpreadsheetML    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NSPresentationML   = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NSDrawingML        = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NSRelationships    = "http://schemas.openxmlformats.org/package/2006/relationships"
    NSContentTypes     = "http://schemas.openxmlformats.org/package/2006/content-types"
)
```

### 3.2 Package: `utils`

#### 3.2.1 EMU Conversions

```go
// pkg/utils/emu.go

// EMU (English Metric Units) constants
const (
    EMUsPerInch       = 914400
    EMUsPerPoint      = 12700
    EMUsPerCentimeter = 360000
    EMUsPerPixel      = 9525 // at 96 DPI
)

// Conversion functions
func InchesToEMU(inches float64) int64
func EMUToInches(emu int64) float64
func PointsToEMU(points float64) int64
func EMUToPoints(emu int64) float64
func CentimetersToEMU(cm float64) int64
func EMUToCentimeters(emu int64) float64
```

#### 3.2.2 Cell Reference Parsing

```go
// pkg/utils/cell_ref.go

// CellRef represents a cell reference (e.g., "A1", "Sheet1!B5")
type CellRef struct {
    Sheet  string // Optional sheet name
    Col    int    // 1-based column number
    Row    int    // 1-based row number
    ColAbs bool   // Is column absolute ($A)
    RowAbs bool   // Is row absolute ($1)
}

// ParseCellRef parses "A1", "$A$1", "Sheet1!A1" formats
func ParseCellRef(ref string) (CellRef, error)

// String returns the A1-style string representation
func (c CellRef) String() string

// ColumnToLetter converts 1-based column number to letter(s)
func ColumnToLetter(col int) string

// LetterToColumn converts letter(s) to 1-based column number
func LetterToColumn(letter string) int

// RangeRef represents a range reference (e.g., "A1:B5")
type RangeRef struct {
    Start CellRef
    End   CellRef
}

// ParseRangeRef parses "A1:B5" format
func ParseRangeRef(ref string) (RangeRef, error)
```

#### 3.2.3 Error Types

```go
// pkg/utils/errors.go

var (
    ErrDocumentClosed     = errors.New("document is closed")
    ErrPartNotFound       = errors.New("part not found")
    ErrInvalidCellRef     = errors.New("invalid cell reference")
    ErrInvalidRange       = errors.New("invalid range")
    ErrTableNotFound      = errors.New("table not found")
    ErrSectionNotFound    = errors.New("section not found")
    ErrSlideNotFound      = errors.New("slide not found")
    ErrShapeNotFound      = errors.New("shape not found")
    ErrInvalidIndex       = errors.New("invalid index")
    ErrReadOnly           = errors.New("document is read-only")
)

// ValidationError provides detailed validation failure info
type ValidationError struct {
    Field   string
    Message string
    Value   interface{}
}

func (e *ValidationError) Error() string
```

---

## 4. Word Package (document)

### 4.1 Core Interfaces

```go
// pkg/document/interfaces.go

// Document represents a Word document
type Document interface {
    // Lifecycle
    Save() error
    SaveAs(path string) error
    Close() error
    
    // Content access
    Body() Body
    Paragraphs() []Paragraph
    Tables() []Table
    Sections() []Section
    
    // Content creation
    AddParagraph() Paragraph
    AddTable(rows, cols int) Table
    
    // Styles
    Styles() Styles
    
    // Track changes
    TrackChanges() TrackChanges
    
    // Comments
    Comments() Comments
    
    // Headers/Footers
    Headers() []Header
    Footers() []Footer
    
    // Properties
    Properties() DocumentProperties
}

// Body represents the document body
type Body interface {
    Elements() []BodyElement
    Paragraphs() []Paragraph
    Tables() []Table
    AddParagraph() Paragraph
    AddTable(rows, cols int) Table
    InsertParagraphBefore(target BodyElement) Paragraph
    InsertParagraphAfter(target BodyElement) Paragraph
}

// Paragraph represents a paragraph
type Paragraph interface {
    BodyElement
    
    // Content
    Text() string
    SetText(text string)
    Runs() []Run
    AddRun() Run
    
    // Formatting
    Style() string
    SetStyle(styleID string)
    Properties() ParagraphProperties
    
    // Structure
    IsHeading() bool
    HeadingLevel() int
}

// Run represents a text run
type Run interface {
    Text() string
    SetText(text string)
    
    // Formatting
    Bold() bool
    SetBold(v bool)
    Italic() bool
    SetItalic(v bool)
    Underline() bool
    SetUnderline(v bool)
    Strike() bool
    SetStrike(v bool)
    FontSize() float64
    SetFontSize(points float64)
    FontName() string
    SetFontName(name string)
    Color() string
    SetColor(hex string)
    Highlight() string
    SetHighlight(color string)
    
    Properties() RunProperties
}

// Table represents a table
type Table interface {
    BodyElement
    
    Rows() []Row
    AddRow() Row
    InsertRow(index int) Row
    DeleteRow(index int) error
    
    Cell(row, col int) Cell
    ColumnCount() int
    RowCount() int
    
    // For identification
    FirstRowText() []string
    Purpose() string // Inferred from headers
}

// Row represents a table row
type Row interface {
    Cells() []Cell
    Cell(index int) Cell
    AddCell() Cell
}

// Cell represents a table cell
type Cell interface {
    Text() string
    SetText(text string)
    Paragraphs() []Paragraph
    AddParagraph() Paragraph
    
    // Cell properties
    VerticalMerge() VerticalMerge
    SetVerticalMerge(v VerticalMerge)
    GridSpan() int
    SetGridSpan(span int)
}
```

### 4.2 Track Changes Interface

```go
// pkg/document/tracking.go

// TrackChanges provides track change functionality
type TrackChanges interface {
    // Enable/disable
    Enabled() bool
    Enable()
    Disable()
    
    // Author
    Author() string
    SetAuthor(name string)
    
    // Revisions
    Insertions() []Revision
    Deletions() []Revision
    AllRevisions() []Revision
    
    // Accept/Reject
    AcceptAll()
    RejectAll()
    AcceptRevision(id string) error
    RejectRevision(id string) error
    
    // Create tracked edits
    InsertText(para Paragraph, position int, text string) error
    DeleteText(para Paragraph, start, end int) error
    ReplaceText(para Paragraph, oldText, newText string) error
}

// Revision represents a tracked change
type Revision interface {
    ID() string
    Type() RevisionType
    Author() string
    Date() time.Time
    Text() string
    Location() RevisionLocation
}

type RevisionType int
const (
    RevisionInsert RevisionType = iota
    RevisionDelete
    RevisionFormat
)
```

### 4.3 Comments Interface

```go
// pkg/document/comments.go

// Comments provides comment functionality
type Comments interface {
    All() []Comment
    ByID(id string) (Comment, error)
    
    Add(text, author string, anchorText string) (Comment, error)
    Delete(id string) error
}

// Comment represents a document comment
type Comment interface {
    ID() string
    Author() string
    Initials() string
    Date() time.Time
    Text() string
    
    // The text this comment is attached to
    AnchoredText() string
    
    // Replies
    Replies() []Comment
    AddReply(text, author string) (Comment, error)
}
```

### 4.4 Constructor Functions

```go
// pkg/document/document.go

// New creates a new empty Word document
func New() (Document, error)

// Open opens an existing Word document
func Open(path string) (Document, error)

// OpenReader opens a document from an io.ReaderAt
func OpenReader(r io.ReaderAt, size int64) (Document, error)
```

---

## 5. Excel Package (spreadsheet)

### 5.1 Core Interfaces

```go
// pkg/spreadsheet/interfaces.go

// Workbook represents an Excel workbook
type Workbook interface {
    // Lifecycle
    Save() error
    SaveAs(path string) error
    Close() error
    
    // Sheets
    Sheets() []Worksheet
    Sheet(nameOrIndex interface{}) (Worksheet, error)
    AddSheet(name string) Worksheet
    DeleteSheet(nameOrIndex interface{}) error
    
    // Named ranges
    NamedRanges() []NamedRange
    AddNamedRange(name, refersTo string) NamedRange
    
    // Tables
    Tables() []Table
    Table(name string) (Table, error)
    
    // Styles
    Styles() Styles
}

// Worksheet represents a worksheet
type Worksheet interface {
    // Identity
    Name() string
    SetName(name string) error
    Index() int
    
    // Visibility
    Visible() bool
    SetVisible(v bool)
    Hidden() bool
    SetHidden(v bool)
    
    // Cells
    Cell(ref string) Cell
    CellByRC(row, col int) Cell
    Range(ref string) Range
    
    // Rows
    Row(index int) Row
    Rows() RowIterator
    UsedRange() Range
    
    // Dimensions
    MaxRow() int
    MaxColumn() int
    
    // Tables
    Tables() []Table
    AddTable(ref string, name string) Table
    
    // Merged cells
    MergedCells() []Range
    MergeCells(ref string) error
    UnmergeCells(ref string) error
    
    // Comments
    Comments() []Comment
}

// Cell represents a cell
type Cell interface {
    // Reference
    Reference() string
    Row() int
    Column() int
    
    // Value
    Value() interface{}
    SetValue(v interface{}) error
    String() string
    Float64() (float64, error)
    Int() (int, error)
    Bool() (bool, error)
    Time() (time.Time, error)
    
    // Formula
    Formula() string
    SetFormula(formula string) error
    HasFormula() bool
    
    // Type
    Type() CellType
    
    // Formatting
    Style() CellStyle
    SetStyle(style CellStyle) error
    NumberFormat() string
    SetNumberFormat(format string) error
    
    // Comments
    Comment() (Comment, bool)
    SetComment(text, author string) error
}

// Range represents a range of cells
type Range interface {
    // Reference
    Reference() string
    StartCell() Cell
    EndCell() Cell
    
    // Iteration
    Cells() [][]Cell
    ForEach(fn func(cell Cell) error) error
    
    // Bulk operations
    SetValue(v interface{}) error
    Clear() error
    
    // Properties
    RowCount() int
    ColumnCount() int
}

// Table represents an Excel table
type Table interface {
    Name() string
    DisplayName() string
    Reference() string
    Worksheet() Worksheet
    
    // Structure
    Headers() []string
    DataRange() Range
    HasTotalsRow() bool
    
    // Data operations
    Rows() []TableRow
    AddRow(values map[string]interface{}) error
    UpdateRow(index int, values map[string]interface{}) error
    DeleteRow(index int) error
    
    // Column access
    Column(name string) []Cell
}

// TableRow represents a row in a table (1-based index after header)
type TableRow interface {
    Index() int // 1-based
    Values() map[string]interface{}
    Cell(columnName string) Cell
    SetValue(columnName string, value interface{}) error
}
```

### 5.2 Cell Types and Styles

```go
// pkg/spreadsheet/cell.go

type CellType int
const (
    CellTypeEmpty CellType = iota
    CellTypeString
    CellTypeNumber
    CellTypeBoolean
    CellTypeDate
    CellTypeFormula
    CellTypeError
)

// CellStyle represents cell formatting
type CellStyle interface {
    // Font
    FontName() string
    SetFontName(name string) CellStyle
    FontSize() float64
    SetFontSize(size float64) CellStyle
    Bold() bool
    SetBold(v bool) CellStyle
    Italic() bool
    SetItalic(v bool) CellStyle
    
    // Fill
    FillColor() string
    SetFillColor(hex string) CellStyle
    
    // Border
    Border() Border
    SetBorder(border Border) CellStyle
    
    // Alignment
    HorizontalAlignment() Alignment
    SetHorizontalAlignment(a Alignment) CellStyle
    VerticalAlignment() Alignment
    SetVerticalAlignment(a Alignment) CellStyle
    
    // Number format
    NumberFormat() string
    SetNumberFormat(format string) CellStyle
}
```

### 5.3 Constructor Functions

```go
// pkg/spreadsheet/workbook.go

// New creates a new empty workbook
func New() (Workbook, error)

// Open opens an existing workbook
func Open(path string) (Workbook, error)

// OpenReader opens a workbook from an io.ReaderAt
func OpenReader(r io.ReaderAt, size int64) (Workbook, error)
```

---

## 6. PowerPoint Package (presentation)

### 6.1 Core Interfaces

```go
// pkg/presentation/interfaces.go

// Presentation represents a PowerPoint presentation
type Presentation interface {
    // Lifecycle
    Save() error
    SaveAs(path string) error
    Close() error
    
    // Slides
    Slides() []Slide
    Slide(index int) (Slide, error)
    AddSlide(layoutIndex int) Slide
    InsertSlide(index, layoutIndex int) Slide
    DeleteSlide(index int) error
    DuplicateSlide(index int) Slide
    ReorderSlides(newOrder []int) error
    
    // Masters and Layouts
    Masters() []SlideMaster
    Layouts() []SlideLayout
    
    // Properties
    Properties() PresentationProperties
    SlideSize() (width, height int64) // In EMUs
    SetSlideSize(width, height int64) error
}

// Slide represents a slide
type Slide interface {
    // Identity
    Index() int // 1-based
    ID() string
    
    // Visibility
    Hidden() bool
    SetHidden(v bool)
    
    // Layout
    Layout() SlideLayout
    
    // Shapes
    Shapes() []Shape
    Shape(identifier string) (Shape, error) // By name or index
    AddShape(shapeType ShapeType) Shape
    AddTextBox(left, top, width, height int64) Shape
    AddTable(rows, cols int, left, top, width, height int64) Table
    AddPicture(imagePath string, left, top, width, height int64) (Shape, error)
    DeleteShape(identifier string) error
    
    // Placeholders
    Placeholders() []Shape
    TitlePlaceholder() Shape
    BodyPlaceholder() Shape
    
    // Notes
    Notes() string
    SetNotes(text string) error
    AppendNotes(text string) error
    HasNotes() bool
    
    // Comments
    Comments() []Comment
    AddComment(text, author string, x, y float64) (Comment, error)
}

// Shape represents a shape on a slide
type Shape interface {
    // Identity
    ID() int
    Name() string
    SetName(name string)
    
    // Type
    Type() ShapeType
    IsPlaceholder() bool
    PlaceholderType() PlaceholderType
    
    // Position and size (in EMUs)
    Left() int64
    Top() int64
    Width() int64
    Height() int64
    SetPosition(left, top int64)
    SetSize(width, height int64)
    
    // Text content
    HasTextFrame() bool
    TextFrame() TextFrame
    Text() string // Convenience method
    SetText(text string) error // Convenience method
    
    // Table (if shape contains a table)
    HasTable() bool
    Table() Table
}

// TextFrame represents the text content of a shape
type TextFrame interface {
    Text() string
    SetText(text string)
    
    Paragraphs() []TextParagraph
    AddParagraph() TextParagraph
    ClearParagraphs()
    
    // Autofit
    AutofitType() AutofitType
    SetAutofitType(t AutofitType)
}

// TextParagraph represents a paragraph in a text frame
type TextParagraph interface {
    Text() string
    SetText(text string)
    
    Runs() []TextRun
    AddRun() TextRun
    
    // Bullet
    BulletType() BulletType
    SetBulletType(t BulletType)
    Level() int
    SetLevel(level int)
    
    // Alignment
    Alignment() Alignment
    SetAlignment(a Alignment)
}

// TextRun represents a text run in a paragraph
type TextRun interface {
    Text() string
    SetText(text string)
    
    // Formatting
    Bold() bool
    SetBold(v bool)
    Italic() bool
    SetItalic(v bool)
    FontSize() float64
    SetFontSize(points float64)
    FontName() string
    SetFontName(name string)
    Color() string
    SetColor(hex string)
}

// Table represents a table in a slide
type Table interface {
    Rows() []TableRow
    Row(index int) TableRow
    Cell(row, col int) TableCell
    
    AddRow() TableRow
    InsertRow(index int) TableRow
    DeleteRow(index int) error
    
    RowCount() int
    ColumnCount() int
}

// TableRow represents a table row
type TableRow interface {
    Cells() []TableCell
    Cell(index int) TableCell
    Height() int64
    SetHeight(height int64)
}

// TableCell represents a table cell
type TableCell interface {
    TextFrame() TextFrame
    Text() string
    SetText(text string)
    
    // Spanning
    RowSpan() int
    ColSpan() int
    SetRowSpan(span int)
    SetColSpan(span int)
}
```

### 6.2 Types and Constants

```go
// pkg/presentation/interfaces.go

type ShapeType int
const (
    ShapeTypeRectangle ShapeType = iota
    ShapeTypeEllipse
    ShapeTypeTextBox
    ShapeTypePicture
    ShapeTypeTable
    ShapeTypeChart
    ShapeTypeGroup
    ShapeTypeLine
    ShapeTypeConnector
)

type PlaceholderType int
const (
    PlaceholderTitle PlaceholderType = iota
    PlaceholderBody
    PlaceholderCenteredTitle
    PlaceholderSubtitle
    PlaceholderDate
    PlaceholderFooter
    PlaceholderSlideNumber
    PlaceholderContent
    PlaceholderPicture
    PlaceholderTable
    PlaceholderChart
)

type AutofitType int
const (
    AutofitNone AutofitType = iota
    AutofitNormal   // Shrink text to fit
    AutofitShape    // Resize shape to fit text
)

type BulletType int
const (
    BulletNone BulletType = iota
    BulletAutoNumber
    BulletCharacter
    BulletPicture
)
```

### 6.3 Constructor Functions

```go
// pkg/presentation/presentation.go

// New creates a new empty presentation
func New() (Presentation, error)

// NewWithSize creates a presentation with specified dimensions
func NewWithSize(width, height int64) (Presentation, error)

// NewWidescreen creates a 16:9 widescreen presentation
func NewWidescreen() (Presentation, error)

// Open opens an existing presentation
func Open(path string) (Presentation, error)

// OpenReader opens a presentation from an io.ReaderAt
func OpenReader(r io.ReaderAt, size int64) (Presentation, error)
```

---

## 7. Testing Requirements

> [!CRITICAL]
> Testing is not optional. Every interface method MUST have corresponding tests. The test suite must be comprehensive enough that refactoring can be done with confidence.
> As advanced features are implemented, tests MUST be expanded in lockstep (unit + round-trip + fixture + fuzz + benchmarks where applicable).

### 7.1 Test Categories

| Category | Description | Coverage Target | Status |
|----------|-------------|-----------------|--------|
| **Unit Tests** | Individual function/method tests | 90%+ | ✅ 694 tests |
| **Integration Tests** | Cross-package interactions | 80%+ | ✅ Implemented |
| **Round-Trip Tests** | Open → Modify → Save → Re-open | 100% of features | ✅ All packages |
| **Fixture Tests** | Real-world document handling | All fixtures pass | ⚠️ Programmatic only |
| **Fuzz Tests** | Random input handling | Critical parsers | ✅ Implemented |

> **PowerPoint Repair Prompts:** when debugging Office repair warnings, add a minimal fixture that reproduces the prompt and validate generated files with the OpenXML SDK validator before declaring them fixed.

### 7.1.1 Advanced Feature Test Policy

For every new advanced feature (Phase 7), the following are required before marking it complete:

- **Unit tests** covering all new API surface.
- **Round-trip tests** proving Open → Modify → Save → Re-open behavior.
- **Fixture tests** using Office-authored documents that exercise the feature.
- **Fuzz tests** for any new OOXML types or parsers.
- **Benchmarks** when the feature introduces new heavy parts (charts, pivots, media).

### 7.2 Parameterized Test Pattern ✅ IMPLEMENTED

> [!IMPORTANT]
> Use table-driven tests for all functionality. This ensures consistent coverage and makes adding test cases trivial.

**Current implementation:**
- `internal/testutil/testutil.go` - Shared test infrastructure
- 67 `t.Run` subtests across packages
- 17 table-driven test case arrays
- Common test data: `CommonStringCases`, `CommonNumericCases`, `CommonFormatCombinations`, `CommonCellRefCases`, `CommonRangeCases`

```go
// Example: pkg/document/paragraph_test.go

func TestParagraph_SetStyle(t *testing.T) {
    tests := []struct {
        name      string
        styleID   string
        wantLevel int
        wantErr   bool
    }{
        {"Heading1", "Heading1", 1, false},
        {"Heading2", "Heading2", 2, false},
        {"Normal", "Normal", 0, false},
        {"Empty style", "", 0, false},
        {"Custom style", "CustomHeading", 0, false},
    }
    
    for _, tt := range tests {
        t.Run(tt.name, func(t *testing.T) {
            doc, _ := document.New()
            para := doc.AddParagraph()
            
            para.SetStyle(tt.styleID)
            
            if got := para.Style(); got != tt.styleID {
                t.Errorf("Style() = %v, want %v", got, tt.styleID)
            }
            if got := para.HeadingLevel(); got != tt.wantLevel {
                t.Errorf("HeadingLevel() = %v, want %v", got, tt.wantLevel)
            }
        })
    }
}
```

### 7.3 Fixture Requirements

All reusable documents/media resolve by fixture ID through the shared schema2
manifest. Physical inputs use one file per SHA under
`fixtures/<format>/<scenarioGroup>/`. Native labels map to IDs in
`internal/testutil/fixture_ids.go`; no source-origin trees or compatibility
symlinks are part of the consumer contract. Required provenance and licences
remain in shared metadata. Some inputs are owned generated archives with qualified
provenance; do not infer a producer version where it was not recorded.

The shared root and pin must match. Release checks require the annotated tag and
commit; candidate checks require explicit root/pin overrides and grant no release
status. Full tracked-byte and clean-checkout verification includes facts/workflows
outside the asset manifest. See [docs/testing.md](docs/testing.md).

### 7.4 Round-Trip Test Pattern

Resolve an input ID, open it, record the semantic expectations, save a private edit
to `t.TempDir()`, reopen and assert those expectations and the permitted part
budget. Use ordinary Go test assertions in runtime-module tests. Keep retained
source-custody assertions separate from mutable authoring round trips. Generated
E2E archives belong in local `artifacts/generated`, never shared references.

`docs/ROUNDTRIP-TESTS.md` lists the existing native fixture assertion families.
They are not evidence that every combination or external producer has been tested.

### 7.5 End-to-End Tests

> [!IMPORTANT]
> E2E tests validate complete workflows, not just individual operations. These are critical for MCP Server compatibility.

```go
// Example: e2e/word_workflow_test.go

func TestWordWorkflow_CreateTechnicalReport(t *testing.T) {
    // This test replicates the MCP Server Technical Report generation workflow
    
    // 1. Create new document
    doc, err := document.New()
    require.NoError(t, err)
    defer doc.Close()
    
    // 2. Enable track changes
    doc.TrackChanges().Enable()
    doc.TrackChanges().SetAuthor("Test Author")
    
    // 3. Add heading
    h1 := doc.AddParagraph()
    h1.SetStyle("Heading1")
    h1.AddRun().SetText("Technical Report")
    
    // 4. Add table
    table := doc.AddTable(3, 2)
    table.Cell(0, 0).SetText("Customer")
    table.Cell(0, 1).SetText("[CUSTOMER_NAME]")
    table.Cell(1, 0).SetText("Project")
    table.Cell(1, 1).SetText("[PROJECT_NAME]")
    
    // 5. Replace placeholders (with track changes)
    doc.TrackChanges().ReplaceText(
        table.Cell(0, 1).Paragraphs()[0],
        "[CUSTOMER_NAME]",
        "Acme Corp",
    )
    
    // 6. Add comment
    doc.Comments().Add(
        "Verify customer name with legal",
        "Reviewer",
        "Acme Corp",
    )
    
    // 7. Save
    tmpFile := t.TempDir() + "/technical_report.docx"
    err = doc.SaveAs(tmpFile)
    require.NoError(t, err)
    
    // 8. Verify by reopening
    doc2, err := document.Open(tmpFile)
    require.NoError(t, err)
    defer doc2.Close()
    
    // Verify track changes recorded
    assert.True(t, doc2.TrackChanges().Enabled())
    assert.GreaterOrEqual(t, len(doc2.TrackChanges().AllRevisions()), 1)
    
    // Verify comment exists
    assert.Len(t, doc2.Comments().All(), 1)
    
    // Verify table data
    tables := doc2.Tables()
    require.Len(t, tables, 1)
    assert.Equal(t, "Acme Corp", tables[0].Cell(0, 1).Text())
}
```

---

## 8. Implementation Phases

### Phase 1: Foundation ✅ COMPLETE

| Task | Priority | Status |
|------|----------|--------|
| `packaging` package complete | P0 | ✅ Done |
| `utils` package complete | P0 | ✅ Done |
| `ooxml/wml` basic types | P0 | ✅ Done |
| `ooxml/sml` basic types | P0 | ✅ Done |
| `ooxml/pml` basic types | P0 | ✅ Done |
| `ooxml/dml` basic types | P0 | ✅ Done |
| Test fixtures created | P0 | ⚠️ Partial (programmatic, need real Office docs) |

**Deliverables:**
- [x] Can open and close DOCX/XLSX/PPTX without error
- [x] Can read relationship files
- [x] Can access raw XML parts
- [x] 100+ unit tests passing (694 tests total)

### Phase 2: Word Core ✅ COMPLETE

| Task | Priority | Status |
|------|----------|--------|
| `document.Document` implementation | P0 | ✅ Done |
| `document.Paragraph` implementation | P0 | ✅ Done |
| `document.Run` implementation | P0 | ✅ Done |
| `document.Table` implementation | P0 | ✅ Done |
| Round-trip tests for all fixtures | P0 | ✅ Done |

**Deliverables:**
- [x] Can read all paragraph text from DOCX
- [x] Can add/modify/delete paragraphs
- [x] Can read/write tables
- [x] Can preserve formatting on round-trip

### Phase 3: Word Advanced ✅ COMPLETE

| Task | Priority | Status |
|------|----------|--------|
| `document.TrackChanges` implementation | P0 | ✅ Done |
| `document.Comments` implementation | P0 | ✅ Done |
| `document.Styles` implementation | P1 | ✅ Done |
| Headers/Footers support | P1 | ✅ Done |
| SDT (Content Controls) support | P1 | ✅ Implemented |

**Deliverables:**
- [x] Can enable/disable track changes
- [x] Can create tracked insertions/deletions
- [x] Can add/read comments
- [x] E2E Technical Report workflow test passing (needs SDT support)

### Phase 4: PowerPoint ✅ COMPLETE

| Task | Priority | Status |
|------|----------|--------|
| `presentation.Presentation` implementation | P0 | ✅ Done |
| `presentation.Slide` implementation | P0 | ✅ Done |
| `presentation.Shape` implementation | P0 | ✅ Done |
| `presentation.TextFrame` implementation | P0 | ✅ Done |
| `presentation.Table` implementation | P1 | ✅ Implemented |
| Notes support | P1 | ✅ Done |

**Deliverables:**
- [x] Can add/delete/reorder slides
- [x] Can modify shape text
- [x] Can add bullet points
- [x] Can read/write tables
- [x] Can manipulate notes

### Phase 5: Excel ✅ COMPLETE

| Task | Priority | Status |
|------|----------|--------|
| `spreadsheet.Workbook` implementation | P0 | ✅ Done |
| `spreadsheet.Worksheet` implementation | P0 | ✅ Done |
| `spreadsheet.Cell` implementation | P0 | ✅ Done |
| `spreadsheet.Table` implementation | P1 | ✅ Implemented |
| `spreadsheet.Range` implementation | P1 | ✅ Done |

**Deliverables:**
- [x] Can read/write cell values
- [x] Can work with Excel tables
- [x] Can handle merged cells
- [x] Round-trip all Excel fixtures

### Phase 6: Integration & Polish 🔄 IN PROGRESS

| Task | Priority | Status |
|------|----------|--------|
| Full E2E test suite | P0 | ✅ Done |
| API documentation | P0 | ✅ Done |
| Performance benchmarks | P1 | ✅ Done |
| Memory profiling | P1 | ✅ Done |
| Error message improvements | P2 | ✅ Done |
| Advanced fuzzing stage | P2 | ✅ Done |

**Deliverables:**
- [x] All MCP Server workflows have equivalent Go tests
- [x] GoDoc documentation complete
- [x] Benchmark baselines established
- [x] No memory leaks in stress tests
- [x] Advanced fuzzing stage completed

### Phase 7: Advanced Feature Parity (Python Superset) 🔄 PLANNED

Goal: reach feature parity with advanced surfaces in native-format, native-format, and native-format while maintaining a consistent Go API.

#### Phase 7A: OOXML Foundations (shared)
| Task | Priority | Status |
|------|----------|--------|
| Add OOXML types for charts/diagrams/themes/media | P0 | ✅ Done |
| Extend packaging constants (content types/relationships) | P0 | ✅ Done |
| Add parsers + round-trip support for new parts | P0 | ✅ Done (package-level part preservation + round-trip tests) |

#### Phase 7B: Spreadsheet Advanced (native-format parity)
| Task | Priority | Status |
|------|----------|--------|
| Charts (line/bar/pie/scatter) + series/axes | P0 | ⏳ Planned |
| Pivot tables + pivot cache | P0 | ⏳ Planned |
| Data validation rules | P1 | ⏳ Planned |
| Conditional formatting authoring | P1 | ⏳ Planned |
| Auto-filters + table sort state | P1 | ⏳ Planned |
| Sparklines | P2 | ⏳ Planned |
| Rich text in cells | P2 | ⏳ Planned |
| Sheet protection + workbook protection | P2 | ⏳ Planned |
| External links + formulas metadata | P2 | ⏳ Planned |
| Page setup/print options | P2 | ⏳ Planned |
| Macro stubs (preserve only) | P3 | ⏳ Planned |

#### Phase 7C: Word Advanced (native-format parity)
| Task | Priority | Status |
|------|----------|--------|
| Paragraph keep-lines/page-break-before/widow control | P1 | ✅ Done |
| Run effects (caps/smallCaps/emboss/outline/shadow) | P1 | ✅ Done |
| Fields (read/write) + instruction parsing | P1 | ⏳ Planned |
| Numbering styles (create/edit) | P1 | ⏳ Planned |
| Table cell borders/widths/vertical align/text direction | P1 | ⏳ Planned |
| Section/page layout controls (columns, breaks) | P2 | ⏳ Planned |
| Revision/move tracking extras | P2 | ⏳ Planned |

#### Phase 7D: PowerPoint Advanced (native-format parity)
| Task | Priority | Status |
|------|----------|--------|
| Themes + master/layout editing | P0 | ⏳ Planned |
| Charts + embedded data | P1 | ⏳ Planned |
| Transitions + animations | P2 | ⏳ Planned |
| SmartArt/diagrams | P2 | ⏳ Planned |
| Audio/video media parts | P2 | ⏳ Planned |
| Advanced shape effects (glow/shadow/3D) | P2 | ⏳ Planned |

#### Phase 7E: Testing & Fixtures
| Task | Priority | Status |
|------|----------|--------|
| Office-authored fixtures for advanced features | P0 | ⏳ Planned |
| Round-trip tests for each new part | P0 | ⏳ Planned |
| Performance benchmarks for charts/pivots/media | P1 | ⏳ Planned |
| Fuzz coverage for new OOXML types | P1 | ⏳ Planned |

### Implementation Summary

| Phase | Status | Tests |
|-------|--------|-------|
| Phase 1: Foundation | ✅ Complete | 71 |
| Phase 2: Word Core | ✅ Complete | 369 |
| Phase 3: Word Advanced | ✅ Complete | (included above) |
| Phase 4: PowerPoint | ✅ Complete | 94 |
| Phase 5: Excel | ✅ Complete | 117 |
| Phase 6: Polish | 🔄 In Progress | 14 (testutil) |
| Phase 7: Advanced Parity | ⏳ Planned | 0 |
| **Total** | **6/7 Phases** | **694 tests** |

---

## 9. Code Quality Standards

### 9.1 Code Consolidation

> [!CAUTION]
> Duplication is a maintenance nightmare. Aggressively consolidate common code.

**Rules:**

1. **XML Marshaling:** All XML marshal/unmarshal code goes through `xmlutil` helpers
2. **EMU Conversions:** All unit conversions go through `utils/emu.go`—no inline math
3. **Error Creation:** All errors use `utils/errors.go` types—no `errors.New()` inline
4. **Part Access:** All part reading goes through `packaging.Part`—no direct ZIP access
5. **Style Application:** Common text styling code shared between document/presentation

**Anti-Patterns to Avoid:**

```go
// BAD: Duplicated in multiple places
func (p *Paragraph) HeadingLevel() int {
    style := p.Style()
    if strings.HasPrefix(style, "Heading") {
        level, _ := strconv.Atoi(style[7:])
        return level
    }
    return 0
}

// GOOD: Centralized in utils
func utils.ParseHeadingLevel(styleID string) int { ... }
```

### 9.2 Interface Compliance

> [!IMPORTANT]
> Compile-time interface verification is required for all implementations.

```go
// At the top of each implementation file
var _ document.Document = (*documentImpl)(nil)
var _ document.Paragraph = (*paragraphImpl)(nil)
var _ document.Run = (*runImpl)(nil)
```

### 9.3 Test Organization

```
pkg/document/
├── document.go
├── document_test.go           # Unit tests
├── document_integration_test.go  # Integration tests (build tag)
├── paragraph.go
├── paragraph_test.go
├── internal/testutil/         # Package-specific test helpers
│   └── helpers.go
```

### 9.4 Error Handling

```go
// Always wrap errors with context
func (d *documentImpl) Open(path string) error {
    pkg, err := packaging.Open(path)
    if err != nil {
        return fmt.Errorf("opening package %s: %w", path, err)
    }
    
    mainPart, err := pkg.GetPart("/word/document.xml")
    if err != nil {
        return fmt.Errorf("getting main document part: %w", err)
    }
    
    // ...
}
```

### 9.5 Documentation

Every exported function/type MUST have GoDoc comments:

```go
// Document represents a Word document and provides methods for reading
// and modifying document content.
//
// Documents should be closed after use to release resources:
//
//     doc, err := document.Open("file.docx")
//     if err != nil {
//         return err
//     }
//     defer doc.Close()
//
// Modifications are not persisted until Save() or SaveAs() is called.
type Document interface {
    // ...
}
```

---

## Appendix A: XML Namespace Quick Reference

| Prefix | Namespace URI | Used In |
|--------|--------------|---------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | Word |
| `x` | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | Excel |
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PowerPoint |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML (shared) |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationships |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | Core props |

---

## Appendix B: Priority Feature Matrix

Features required for MCP Server parity:

| Feature | Word | Excel | PPTX | Priority |
|---------|------|-------|------|----------|
| Read document | ✓ | ✓ | ✓ | P0 |
| Write document | ✓ | ✓ | ✓ | P0 |
| Paragraphs/text | ✓ | - | ✓ | P0 |
| Tables | ✓ | ✓ | ✓ | P0 |
| Track changes | ✓ | - | - | P0 |
| Comments | ✓ | ✓ | ✓ | P0 |
| Styles | ✓ | ✓ | - | P1 |
| Images | ✓ | ✓ | ✓ | P1 |
| Headers/Footers | ✓ | - | - | P1 |
| Slide manipulation | - | - | ✓ | P0 |
| Notes | - | - | ✓ | P1 |
| Named ranges | - | ✓ | - | P2 |
| Merged cells | - | ✓ | - | P1 |
| Formulas (read) | - | ✓ | - | P2 |
| Charts | - | - | - | P0 (Phase 7) |

---

## Appendix C: Acceptance Criteria Checklist

Before declaring the library complete, ALL items must pass:

- [x] All interfaces implemented per this spec
- [ ] All test fixtures round-trip without data loss
- [ ] Unit test coverage > 90%
- [x] All E2E workflow tests pass
- [ ] No data races under `go test -race`
- [ ] No memory leaks (verified via `go test -memprofile`)
- [ ] GoDoc documentation for all exported types
- [ ] README with quick start examples
- [ ] Benchmarks establish baseline performance
- [ ] Can be imported and used in MCP Server codebase

## Retained-source package adapter (initial subset)

The additive `packaging.Preserved` concrete type owns immutable source/member
bytes independently of the legacy `Package` models. `OpenPreserved([]byte, Limits)`
validates intake; `Part(name)` returns a cloned payload and SHA-256; `Replace` takes
an atomic batch of existing-member `Replacement{Part, ExpectedSHA256, Data}` values;
`WriteTo` copies exact source bytes on no-op and raw-copies untouched ZIP members
on edits. XML replacements must have one well-formed root. Registry replacements,
duplicate/stale/missing targets and signed-package edits refuse. This is a low-level
payload API; format semantics, reference rewriting, graph edits and native Office
validation are separate contracts. No existing exported interface is extended.

`Preserved.Receipt()` returns schema-1 byte-level changed-member hashes against
the immutable session input. It describes staged changes, not proof of delivery.
`Preserved.SaveAs(path)` stages output, reopens it and checks all expected member
payloads before replacing a regular destination; only then does it return a receipt.
Failed delivery retains staged changes. Symlink/nonregular targets refuse rather
than silently choosing follow-link versus replace-link semantics. Filesystem crash
durability after rename remains outside this initial contract.

`Preserved.Graph()` inspects the understood OPC registry subset and returns sorted
parts/content types/inbound ownership counts and source/ID/type/target edges.
External targets are never fetched. Missing local parts, ambiguous registries and
unknown extensions refuse with relationship_policy. Read-only graph validity is
separate from graph surgery; allocation/copy/delete/import are not implemented.

## Initial Word safe-edit subset

`document.OpenEditing(source, packaging.Limits)` returns an additive concrete
`EditSession` after OPC graph and main-part checks. `FindOne` targets one exact
complete `w:t` leaf. `Replace` checks session identity/generation/fingerprint and
consumes a target only after a changing edit succeeds. It currently permits only
direct body paragraph/run/text ownership; fields, review/range markers, controls,
hyperlinks, protection and unsupported whitespace refuse. No-op and refusal keep
targets reusable; any successful text change invalidates other old targets
conservatively. `SaveAs` returns the preserved package receipt. Cross-run search,
all-story editing, tracked edits and full extended revision/protection semantics are
not implemented by this subset. Existing Document interfaces are unchanged.

## Initial presentation safe-edit subset

`presentation.OpenEditing` validates the retained graph and ordinary transitional
presentation/slide inventory. `EditSession.FindText(slidePart, shapeID, text)`
selects one exact complete text leaf in a plain ungrouped shape; duplicate IDs or
text refuse. `Replace` retains geometry/direct formatting and unrelated XML bytes,
refusing fields, breaks, mixed content, shape locks and unsupported characters.
Targets bind the session, generation and full part fingerprint. No-op/refusal keep
targets reusable; successful change consumes and conservatively invalidates prior
targets. This is not full deck search, inherited formatting or field/group/table
editing. Existing presentation interfaces remain unchanged.

## Initial spreadsheet safe-edit subset

`spreadsheet.OpenEditing` validates the retained OPC graph and exact workbook/sheet
identity. `FindNumber(sheet, cell)` requires one existing numeric value leaf at an
exact uppercase bounded A1 address. `SetNumber` accepts finite values only and
refuses formulas anywhere in the workbook, names, charts/pivots/tables/external
dependencies, validation/conditional formatting, protection, merges and extensions
that it cannot prove independent. It preserves the cell's style and all unrelated
bytes. Targets use session/generation/part fingerprints; no-op/refusal remain
reusable. `SaveAs` delivers through the retained-source verifier. This conservative
formula-free subset performs no dependency rewriting or cache recalculation and is
not the general extended spreadsheet editing contract. Existing interfaces unchanged.

`document.EditSession.Outline(CurrentView|OriginalView|AllView)` returns versioned,
read-only paragraph evidence across the main part and transitively related header,
footer, footnote, endnote and comment stories. It projects basic insertion/deletion/
move wrappers, preserves literal tabs/breaks and assigns nested textbox text to its
nearest paragraph only. AlternateContent branches are skipped with warnings; field
instructions and property revisions are not evaluated. StoryBlock is not a live
mutation target. Full extended outline/report/revision scope remains incomplete.

Word `FindText` now inventories all exact non-overlapping matches within main-story
current-view paragraphs across run fragmentation. `FindOne` uses the same spans
and refuses zero/multiple matches. Internal offsets are Unicode rune offsets;
original text is returned by `TextTarget.Text()`. Paragraph boundaries are not
implicitly joined. In this increment replacement still requires one whole leaf;
substring/multi-run targets refuse until the planner is implemented.

Word ordinary Replace now plans substring/multi-run edits using the unique maximal
exact prefix/suffix split. Ambiguous repeated affixes refuse. Unchanged fragments
remain in original runs; the residual is inserted into its starting run while
later changed fragments become empty text leaves. Each affected leaf retains its
run properties/markup. Pure insertions at interior run boundaries still refuse
until complete formatting/owner equality is proved. Wrappers/fields/protection
retain conservative gates; no tracked edit or revision preservation is implied.

`EditSession.ReplaceBatch([]TextReplacement)` is an explicitly selected atomic
batch: any invalid/stale/overlapping/unsupported target aborts every selected edit.
It merges disjoint rune deltas for targets sharing a text leaf, commits once, and
consumes changing targets only after success. All-no-op batches keep targets live.
This API is deliberately distinct from extended replace_all's per-match refusal
reporting/continue policy; that convenience API remains pending.

`EditSession.Search(text, SearchOptions)` adds opt-in Unicode-15 full default
casefolding, smart punctuation folding, whitespace collapse and soft-hyphen removal
with exact source-rune maps. Matches cannot select only part of an expanded folded
character. Near ranks the complete candidate set by main-story source-character
distance; stable ties and missing contexts retain order. Nth is one-based and
mutually exclusive with Near. Original raw text remains target evidence. Search
is still paragraph-local/current-main-story; cross-paragraph/all-story policy
coverage remains pending. Unicode data licence is retained under tools/casefold.

Word ordinary single/batch replacements now set `xml:space="preserve"` on changed
text leaves when leading/trailing/repeated spaces require it. The attribute is
inserted/updated via the lossless start-tag primitive, alongside the text in one
transaction. Unknown xml:space policies still refuse; tabs/newlines still require
structural Word elements. No-op does not rewrite whitespace attributes. This
supersedes the initial missing-preserve refusal documented in the first subset.

`EditSession.ReplaceAll(needle, replacement, normalized)` searches private targets,
skips exact no-ops, records per-match safe refusals and commits all independently
supported changes together via ReplaceBatch. Unexpected/stale/batch conflicts abort
without commit. Schema1 reports matched/changed/skipped counts and refusal evidence;
it does not yet reproduce extended schema3 formatting/revision result fields. No
private targets escape. Tracked/revision-preserving options remain pending.

Word Search now accepts an exact related story part and current/original/all view.
It joins visible paragraph streams with one literal newline for inspection; a
separator-only match has no target. Cross-paragraph and historical-view matches
remain read-only, and related-story mutation is still unsupported. Source maps
preserve paragraph ownership; ordinary edits still require current body targets.
FindText/FindOne remain their documented main-story paragraph-local convenience
subset until their complete cross-story contracts are unified.

Pure insertion at an interior Word run boundary is now allowed when both sides
prove the same paragraph owner, equal expanded run/text attributes and byte-equal
complete rPr markup. Both structural guards must pass. The left run is selected
deterministically only after equivalence proof. Semantically equal but differently
serialized properties still conservatively refuse. Raw XML remains copied evidence.

`spreadsheet.EditSession.ValidateStyles()` checks cell, row and column style
indices across worksheet parts before guarded numeric writes. An omitted cell `s`
selects zero; when a styles relationship exists, zero must resolve to an actual
cellXfs entry. Empty/duplicate tables, mismatched counts, malformed/overflowing and
out-of-range indices refuse without mutation. With no style part, only default
zero is accepted. This is index integrity, not font/fill dependency closure or
style-authoring parity. Shared-contract pack V2 motivated the regression.

`spreadsheet.EditSession.SetNumberWithInvalidation(target, value)` is an explicit
static-formula subset separate from the formula-free SetNumber contract. It fully
preflights the workbook, traces same/cross-sheet A1/range references transitively,
clears affected cached value text and requests calcMode=auto/fullCalcOnLoad/
forceFullCalc in the same retained multi-part transaction. It never calculates.
CalculationEffect reports applied changes/invalidation, not delivery. Unrelated
caches/calculation metadata remain exact; numeric no-ops do not consume targets.
Dynamic/UDF/shared/array/structured/unknown references, unlisted sheets, manual or
iterative/precision-as-displayed modes, existing calcChain and previously blocked
chart/pivot/protected structures refuse. This is not full X06/X07/extended parity.

## Unresolved extended-comment content-type discrepancy

Pinned Go emits `application/vnd.ms-word.commentsExtended+xml`; pinned extended/Python
sources emit `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`.
No authoring constant is changed on source agreement alone. Retained-source unrelated
edits preserve either input registry spelling and related bytes. This is preservation
characterisation, not schema/native Office/comment-thread certification. Decision
record: `spec/source-discrepancies.json`; details and required checks:
`docs/design/comments-extended-content-type.md`.

## Planned package-graph additions and retargets

`Preserved.PlanGraphMutation(GraphMutation)` creates a private session-bound plan
for explicit `PartAddition` payloads/content types and existing internal
`RelationshipRetarget` edges. Planning patches only required registry elements,
validates the complete candidate OPC graph, and leaves all state unchanged.
`ApplyGraphPlan` commits once; intervening payload changes, foreign sessions and
consumed plans refuse. A semantic no-op preserves the original archive and plan.

New parts receive explicit content-type overrides. Retargets retain absolute or
relative path form with URI escaping, preserve relationship IDs/types and keep old payloads, including
shared or newly unreferenced parts. External/fragment edges, signatures, reserved
registry additions, case collisions, invalid XML and missing targets refuse.
This API does not prove format-specific occurrence ownership, cache policy or
import closure; those checks belong to the format adapter. Explicit edge removals,
detached-leaf deletions and fingerprint-guarded payload replacements can share one
plan. Deleting a part with its own relationship registry refuses; every remaining
inbound edge must resolve after the transaction. The caller proves that removed
edges have no surviving format-level references. Relationship creation and general
dependency import are unsupported.

Saves raw-copy untouched original ZIP members and append new members in sorted
order with fixed metadata. Reopen validation checks payloads and the graph.
Receipts without additions/deletions retain schema1. Otherwise schema2 records
`add`, `replace` or `delete`. Added parts have an empty `before_sha256`; deleted
parts have an empty `after_sha256`. Empty hashes mean absence, not empty payloads.

`spreadsheet.EditSession.FindImage(sheet, shapeID)` selects a loaded directly
anchored picture by cNvPr ID. `ReplaceImage(target, data)` accepts fully decoded
PNG/JPEG only (32 MiB compressed bytes, 16 million pixels maximum), creates fresh
case-collision-free media and retargets its existing relationship. Drawing XML,
geometry, crop, old media, other relationships and all other payloads remain exact.
The drawing must have one inbound edge and the selected relationship exactly one
XML use. Shared drawing/relationship identities, linked/external images, groups,
charts, extensions and protection refuse. Identical image bytes are a no-op;
changed images consume the target and stale other session handles. Existing legacy
drawing APIs are unchanged. This bounded adapter does not complete X10's broader
format/source-policy coverage or establish native Office/rendering compatibility.

Retained intake now verifies local/central ZIP64 sizes and offsets, including tiny
forced declarations and signed/unsigned 64-bit descriptors. ZIP64 end-record extent
must end at its locator; the central directory must end at its declared boundary.
Required extra fields are read only within their TLV length; duplicate/truncated
ZIP64 extras and classic/ZIP64 disagreement refuse. Empty deflated members are
verified through archive/zip payload CRC/size checks. Caller resource limits still
apply; tests use bounded synthetic archives, not multi-gigabyte payloads.

Explicit numeric invalidation can remove one ordinary calculation-chain part when
an input change affects formulas. Chain ownership, MIME, root and direct entry
structure are preflighted; shared/external/unowned/extended/outward-linked chains
refuse. The chain relationship, override and payload are removed in the same
GraphPlan as input/cache/calcPr changes. The effect reports its removed part name.
Unrelated and same-value edits preserve the chain exactly. The legacy rebuild
writer still has its historical synthetic-chain behaviour; this safe API does not
route through it or infer dependencies from chain ordering.

`presentation.EditSession.FindNotes(slidePart)` returns a session-bound NotesTarget
for one existing, uniquely owned notes body containing ordinary paragraphs/runs.
`NotesTarget.Text()` joins paragraph text with LF. `ReplaceNotes` uses the first
paragraph/run formatting as templates for LF-separated lines; empty lines become
empty paragraphs. Nonempty single-leaf corrections still use exact text splices.
Other placeholders/body metadata/registries stay byte-identical. No notes graph is
created. Fields, shared/ambiguous owners, text locks/protection, unproved formatting
and malformed text refuse atomically; grouping-only noGrp locks are allowed.
Tabs/CR are unsupported. No-op targets remain reusable; changes stale session
handles. Broader source/format/rendering coverage is unfinished.

`spreadsheet.EditSession.AllowedValues(sheet,cell)` inspects list validation
without creating a cell. Nil means no covering list; malformed/overlapping or
unproved vocabulary sources return a typed refusal and no partial values.
ValidationValue tags distinguish string, number, boolean and blank. Text retains
literal/string contents or scalar source spelling; numeric formatting/date serials
are not interpreted. Literal lists and finite static1D ranges are supported,
including reversed/quoted-sheet references and blank slots, up to100000 positions.
Formulas are never evaluated and cached formula/error values refuse. Worksheet
extensions, merged interiors, dynamic/named/external/2D sources and ambiguous
cell ownership refuse. Whole-row/column sources use the retained stored-cell extent
(minimum1), never the worksheet dimension hint; result size remains capped.

AllowedValues also resolves shared-string indices against the retained si sequence
without deduplication and reads supported rich inline/shared runs. Missing/external/
ambiguous tables, bad indices and mixed/extended text structures refuse. Values are
raw XML-decoded stored text, not Excel display or _xHHHH_ escape interpretation.

AllowedValues caps literal lists at the same 100000 positions as static ranges.
Unknown declared validation types refuse; an omitted type defaults to none.
Stored numbers require finite decimal/exponent spelling; Go hexadecimal and
underscore numeric spellings refuse. Signed decimals retain their stored text.

Native tests read shared assets from `references/fixtures-ooxml`; the explicit
`OOXML_FIXTURES_ROOT` override supports a candidate distribution before tag freeze.
Missing inputs fail; tests do not fall back to local testdata or opt out. Outputs
remain consumer-local, with generated-output guards against shared-root writes.
The root manifest seals the mutation-safety feature and compact contract, with
fixture bytes addressed by canonical asset IDs. Official Gherkin supplies expanded
cases at verification; no wrapper or generated case inventory is read. The
schema2 shared checkout is pinned to annotated v0.150.0; default batches require no
override. Exact commit/tag/root seal and full tracked-tree custody are checked.
See docs/testing.md for candidate policy and current verification.
