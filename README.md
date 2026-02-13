# go-ooxml

![Icon](docs/icon-256.png)

This is another of my "things that should exist" projects: An in-development Go library for reading, writing, and manipulating Office Open XML (OOXML) documents.

Supports Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) formats, and is slowly being developed against the ECMA 376 specs.

## Installation

```bash
go get github.com/rcarmo/go-ooxml
```

## Usage

```go
package main

import (
	"log"

	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

func main() {
	// Word
	doc, err := document.New()
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()
	doc.AddParagraph().SetText("Hello, World")
	if err := doc.SaveAs("hello.docx"); err != nil {
		log.Fatal(err)
	}

	// Excel
	wb, err := spreadsheet.New()
	if err != nil {
		log.Fatal(err)
	}
	defer wb.Close()
	sheet, _ := wb.Sheet(0)
	sheet.Cell("A1").SetValue("Hello")
	if err := wb.SaveAs("hello.xlsx"); err != nil {
		log.Fatal(err)
	}

	// PowerPoint
	pres, err := presentation.New()
	if err != nil {
		log.Fatal(err)
	}
	defer pres.Close()
	slide := pres.AddSlide(0)
	slide.AddTextBox(0, 0, 4000000, 1000000).SetText("Hello")
	if err := pres.SaveAs("hello.pptx"); err != nil {
		log.Fatal(err)
	}
}
```

### Document (Word)

- **Open/Save:** `document.New()`, `document.Open(path)`, `doc.Save()`, `doc.SaveAs(path)`
- **Content:** `doc.AddParagraph()`, `doc.AddTable(rows, cols)`
- **Formatting:** `Run` setters (`SetBold`, `SetItalic`, `SetFontSize`, `SetColor`, etc.)
- **Track changes:** `doc.EnableTrackChanges(author)`, `doc.TrackChanges()`
- **Comments:** `doc.Comments().Add(text, author, anchorText)`
- **Headers/Footers:** `doc.AddHeader(type)`, `doc.AddFooter(type)`
- **Content controls:** `doc.AddBlockContentControl(tag, alias, text)`

### Spreadsheet (Excel)

- **Open/Save:** `spreadsheet.New()`, `spreadsheet.Open(path)`, `wb.Save()`, `wb.SaveAs(path)`
- **Sheets:** `wb.Sheets()`, `wb.AddSheet(name)`
- **Cells/Ranges:** `sheet.Cell("A1")`, `sheet.Range("A1:C3")`
- **Tables:** `sheet.AddTable("A1:C3", "Sales")`, `table.AddRow(values)`
- **Named ranges:** `wb.AddNamedRange(name, refersTo)`
- **Comments:** `sheet.Cell("A1").SetComment(text, author)`

### Presentation (PowerPoint)

- **Open/Save:** `presentation.New()`, `presentation.Open(path)`, `pres.Save()`, `pres.SaveAs(path)`
- **Slides:** `pres.AddSlide(layoutIndex)`, `pres.Slides()`
- **Shapes:** `slide.AddShape(type)`, `slide.AddTextBox(...)`, `shape.SetText(text)`
- **Tables:** `slide.AddTable(rows, cols, left, top, width, height)`
- **Validation:** `dotnet tools/validator/OoxmlValidator/bin/Release/net10.0/OoxmlValidator.dll <file>` (use when debugging repair prompts)

## Development

### Repository Layout

- `pkg/` - Public Go packages (document/spreadsheet/presentation)
- `internal/` - Internal helpers (testutil/xmlutil)
- `e2e/` - End-to-end workflows and fuzz tests
- `testdata/` - Fixture files
- `docs/` - Specification and reference docs
- `tools/` - Validator and tooling

```bash
# Show available targets
make help

# Full build (clean + deps + lint + test + build)
make build-all

# Run tests
make test

# Run tests with coverage
make coverage

# Run benchmarks
make bench

# Run memory profiling tests
make memprofile

# Format code
make format

# Lint code
make lint
```

## License

MIT
