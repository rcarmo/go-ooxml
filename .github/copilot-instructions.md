# Go OOXML Library - Copilot Instructions

## Agent Coordination

Read `AGENTS.md` for the priority communication contract. Scope changes, stop/hold,
pin corrections, safety blockers and unblock decisions require `chat` with
`target_agent_name: "@alias"` and explicit `mode: "steer"`. Include the current
tag/commit, action/owner and superseded notice. Queue only routine progress;
acknowledge latest state once and do not replay stale notices.

## Build and Test Commands

Use the Makefile for all standard operations:

```bash
make help          # Show all available targets
make build-all     # Full build (clean + deps + lint + test + build)
GOMAXPROCS=2 make test-batch  # Profiled root and acceptance modules after reference setup
make test          # Profiled root module only
make coverage      # Profiled root module plus retained coverage
make lint          # Run golangci-lint
make format        # Format code with gofumpt
make check         # Run lint + tests
```

## Makefile-First Workflow

If you need a new workflow step, add a Make target rather than running ad-hoc commands.

## CI/CD Convention

CI should call lint plus `GOMAXPROCS=2 make test-batch`; root `go test ./...`
does not discover the separate acceptance module. Read `docs/testing.md` for
candidate/released reference setup before running either module. CI must set
project-owned cache/temp paths using `scripts/project-tmp.sh` even outside
this host; invalid explicit overrides fail, and no unprofiled test commands
or home-cache fallback are allowed.

Run related package batches with bounded concurrency and per-package profiles;
do not run individual tests:

```bash
make prepare-temp
GOMAXPROCS=2 bash scripts/test-profile.sh ./pkg/document ./pkg/packaging
```

## Architecture

This library implements Office Open XML (OOXML) support for Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) with a layered architecture:

```
High-Level API (document.Document, spreadsheet.Workbook, presentation.Presentation)
    ↓
Element Wrappers (Paragraph, Cell, Slide, Shape, etc.)
    ↓
OOXML Types (pkg/ooxml/wml, sml, pml, dml - low-level XML structs)
    ↓
Packaging (pkg/packaging - OPC/ZIP handling, relationships, content types)
    ↓
Go Standard Library (archive/zip, encoding/xml)
```

### Package Responsibilities

- **`pkg/packaging`** - OPC package handling (ZIP, relationships, `[Content_Types].xml`)
- **`pkg/ooxml/*`** - Low-level XML types matching ECMA-376 spec (CT_* structs)
  - `wml/` - WordprocessingML
  - `sml/` - SpreadsheetML  
  - `pml/` - PresentationML
  - `dml/` - DrawingML (shared across formats)
- **`pkg/document`** - High-level Word API
- **`pkg/spreadsheet`** - High-level Excel API
- **`pkg/presentation`** - High-level PowerPoint API
- **`pkg/utils`** - Shared utilities (EMU conversions, color parsing, cell references)
- **`internal/testutil`** - Shared test infrastructure

## Key Conventions

### Zero External Dependencies
Runtime packages use only the Go standard library (`archive/zip`, `encoding/xml`,
`io`, `path`). Test-only Godog/Gherkin dependencies are isolated in `acceptance/`.

### Consistent Document Lifecycle
All document types follow the same pattern:
```go
// Create new
doc, err := document.New()   // or spreadsheet.New(), presentation.New()

// Open existing  
doc, err := document.Open("file.docx")

// Open from reader
doc, err := document.OpenReader(reader, size)

// Save
doc.Save()           // Save to original path
doc.SaveAs("new.docx")

// Always close
defer doc.Close()
```

### Table-Driven Tests
Use parameterized tests with the shared test infrastructure:
```go
func TestParagraph_SetText(t *testing.T) {
    for _, tc := range CommonTextCases {
        t.Run(tc.Name, func(t *testing.T) {
            h := NewTestHelper(t)
            doc := h.CreateDocument(nil)
            defer doc.Close()
            // test logic
        })
    }
}
```

Common test data is defined in `internal/testutil/testutil.go`:
- `CommonStringCases`, `CommonNumericCases`
- `CommonFormatCombinations`
- `CommonCellRefCases`, `CommonRangeCases`

### XML Struct Naming
Low-level OOXML types use ECMA-376 naming conventions prefixed with `CT_`:
- `CT_P` (paragraph), `CT_R` (run), `CT_Tbl` (table)
- `CT_Cell`, `CT_Slide`, `CT_Shape`

### EMU Units
Measurements use English Metric Units (EMUs). Use `pkg/utils/emu.go` for conversions:
- 914400 EMUs = 1 inch
- 12700 EMUs = 1 point

### Error Handling
Custom error types are defined in `pkg/utils/errors.go` and `pkg/spreadsheet/errors.go`.

## Test Fixtures

Resolve fixture IDs through the shared schema2 manifest. The only physical input
root is the shared checkout's `fixtures/<format>/<scenarioGroup>/`; paths belong
to manifest records, not hard-coded origin trees. Required licence/provenance data
stays in that repository. Do not copy fixtures or add compatibility symlinks.

See `docs/testing.md` for the exact candidate/released pin rules and clean-checkout
verification. Disposable test output belongs under `/workspace/tmp/go-ooxml/runs/`, with
cache/build under the sibling `cache/` and `build/` directories; retained
profiles, logs and OOXML evidence belong in `artifacts/`. Never write to references. Keep native semantic assertions and use round trips where
appropriate. Local catalogue files are staging for the central canonical registry.
