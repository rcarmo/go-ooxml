# go-ooxml

![Icon](docs/icon-256.png)

This is another of my "things that should exist" projects: An in-development Go library for reading, writing, and manipulating Office Open XML (OOXML) documents.

Supports Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) formats, and is slowly being developed against the ECMA 376 specs.

## Installation

```bash
go get github.com/rcarmo/go-ooxml
```

## Development

```bash
# Show available targets
make help

# Full build (clean + deps + lint + test + build)
make build-all

# Run tests
make test

# Run tests with coverage
make coverage

# Format code
make format

# Lint code
make lint
```

## License

MIT
