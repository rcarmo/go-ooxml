.PHONY: help install lint format test coverage check clean clean-all build build-all bump-patch push validate security

GO ?= go
GOFMT ?= gofumpt
GOLINT ?= golangci-lint
GOSEC ?= gosec
TEST_JOBS ?= 2
DOTNET_ROOT ?= /home/linuxbrew/.linuxbrew/opt/dotnet/libexec
VALIDATOR ?= tools/validator/OoxmlValidator

# Binary name
BINARY_NAME = go-ooxml

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-14s\033[0m %s\n", $$1, $$2}'

# =============================================================================
# Full reproducible build
# =============================================================================

build-all: clean deps lint test build ## Full reproducible build (clean + deps + lint + test + build)
	@echo "Build complete!"

# =============================================================================
# Go targets
# =============================================================================

deps: ## Download and tidy dependencies
	$(GO) mod download
	$(GO) mod tidy

install: ## Install the binary
	$(GO) install ./...

lint: ## Run golangci-lint
	@which $(GOLINT) > /dev/null || (echo "Installing golangci-lint..." && brew install golangci-lint)
	$(GOLINT) run ./...

security: ## Run gosec security scan
	@which $(GOSEC) > /dev/null || (echo "Installing gosec..." && $(GO) install github.com/securego/gosec/v2/cmd/gosec@latest)
	$$(go env GOPATH)/bin/$(GOSEC) ./...

format: ## Format code with gofumpt
	@which $(GOFMT) > /dev/null || (echo "Installing gofumpt..." && $(GO) install mvdan.cc/gofumpt@latest)
	$(GOFMT) -w .

test: ## Run native tests as a bounded batch
	$(GO) test -p $(TEST_JOBS) ./...

.PHONY: acceptance test-batch
acceptance: ## Run strict native Gherkin and report reconciliation
	cd acceptance && $(GO) test -p $(TEST_JOBS) ./...

test-batch: test acceptance ## Run library and acceptance modules

# Bounded graphics candidate tests; never advance the released reference pin.
GRAPHICS_ROOT ?= $(abspath ../fixtures-ooxml)
.PHONY: graphics-test
graphics-test: ## Run graphics-related packages against the sealed shared candidate
	OOXML_GRAPHICS_ROOT="$(GRAPHICS_ROOT)" GOMAXPROCS=2 $(GO) test -p $(TEST_JOBS) ./pkg/presentation ./pkg/packaging ./internal/losslessxml

.PHONY: graphics-insertion-quality
graphics-insertion-quality: ## Run optional LibreOffice insertion save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/insertion-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-insertion-uno.py

.PHONY: graphics-replacement-quality
graphics-replacement-quality: ## Run optional LibreOffice replacement save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/replacement-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-replacement-uno.py

.PHONY: graphics-metadata-quality
graphics-metadata-quality: ## Run optional LibreOffice crop/transform save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/metadata-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-metadata-uno.py

.PHONY: graphics-placement-quality
graphics-placement-quality: ## Run optional LibreOffice fitted-placement save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/placement-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-placement-uno.py

.PHONY: graphics-svg-quality
graphics-svg-quality: ## Run optional LibreOffice paired SVG save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/svg-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-svg-uno.py

.PHONY: graphics-delete-quality
graphics-delete-quality: ## Run optional LibreOffice deletion save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/delete-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-delete-uno.py

.PHONY: graphics-group-quality
graphics-group-quality: ## Run optional LibreOffice grouped-child save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/group-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/shape-group-uno.py

.PHONY: graphics-group-transform-quality
graphics-group-transform-quality: ## Run optional LibreOffice group-frame save/reopen oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/group-transform-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/group-transform-uno.py

.PHONY: graphics-connectors-quality
graphics-connectors-quality: ## Run optional LibreOffice connector attachment oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/connectors-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/connectors-uno.py

.PHONY: graphics-diagrams-quality
graphics-diagrams-quality: ## Run optional LibreOffice diagram edit and attachment oracle
	OOXML_GRAPHICS_OUTPUT="$(abspath artifacts/graphics/diagrams-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/diagrams-uno.py

.PHONY: shared-pack-check
shared-pack-check: ## Verify shared distribution and native readbacks in a complete batch
	cd acceptance && $(GO) test -p $(TEST_JOBS) ./...

coverage: ## Run tests with coverage
	$(GO) test -coverprofile=coverage.out ./...
	$(GO) tool cover -func=coverage.out

bench: ## Run benchmarks
	$(GO) test -run '^$$' -bench . ./...

memprofile: ## Run memory profiling tests (requires ENABLE_MEMPROFILE=1)
	ENABLE_MEMPROFILE=1 $(GO) test -run MemProfile -memprofile=mem.out ./...

check: lint test ## Run lint + tests

build: ## Build the library (verify compilation)
	$(GO) build ./...

# =============================================================================
# Clean targets
# =============================================================================

clean: ## Remove build artifacts and cache
	$(GO) clean
	rm -rf coverage.out $(BINARY_NAME)

clean-all: clean ## Remove everything including vendor
	rm -rf vendor

# =============================================================================
# Version management
# =============================================================================

bump-patch: ## Bump patch version and create git tag
	@CURRENT=$$(git describe --tags --abbrev=0 2>/dev/null || echo "v0.0.0"); \
	MAJOR=$$(echo $$CURRENT | sed 's/v//' | cut -d. -f1); \
	MINOR=$$(echo $$CURRENT | sed 's/v//' | cut -d. -f2); \
	PATCH=$$(echo $$CURRENT | sed 's/v//' | cut -d. -f3); \
	NEW="v$$MAJOR.$$MINOR.$$((PATCH + 1))"; \
	git tag "$$NEW"; \
	echo "Created tag: $$NEW"

push: ## Push commits and current tag to origin
	@TAG=$$(git describe --tags --exact-match 2>/dev/null); \
	git push origin main; \
	if [ -n "$$TAG" ]; then \
		echo "Pushing tag $$TAG..."; \
		git push origin "$$TAG"; \
	else \
		echo "No tag on current commit"; \
	fi

# =============================================================================
# OOXML Validation (requires .NET SDK)
# =============================================================================

validate: ## Validate the two historical root fixtures from the pinned shared checkout using the official SDK
	@set -e; if [ ! -f $(VALIDATOR)/bin/Release/net10.0/OoxmlValidator.dll ]; then \
		echo "Building validator..."; \
		export DOTNET_ROOT=$(DOTNET_ROOT) && cd $(VALIDATOR) && dotnet build -c Release -m:1 -p:UseSharedCompilation=false --disable-build-servers -q; \
	fi; \
	export DOTNET_ROOT=$(DOTNET_ROOT); \
	for id in fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5 fixture-151d747bc37d4f4988c1116f4abb45196b1c1644319ce342ee4dd56d111f3132; do \
		f=$$(go run ./tools/fixturepath "$$id"); \
		dotnet $(VALIDATOR)/bin/Release/net10.0/OoxmlValidator.dll "$$f"; \
	done
