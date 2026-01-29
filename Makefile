.PHONY: help install lint format test coverage check clean clean-all build build-all bump-patch push

GO ?= go
GOFMT ?= gofumpt
GOLINT ?= golangci-lint

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

format: ## Format code with gofumpt
	@which $(GOFMT) > /dev/null || (echo "Installing gofumpt..." && $(GO) install mvdan.cc/gofumpt@latest)
	$(GOFMT) -w .

test: ## Run tests
	$(GO) test -v ./...

coverage: ## Run tests with coverage
	$(GO) test -coverprofile=coverage.out ./...
	$(GO) tool cover -func=coverage.out

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
