.PHONY: help install lint format test coverage check clean clean-all build build-all bump-patch push validate security prepare-temp

# Resolve once before replacing TMPDIR; explicit invalid PROJECT_TMP_ROOT fails.
# The repository-owned resolver is also available to CI without the host Makefile.
PROJECT_TMP := $(shell PROJECT=go-ooxml bash scripts/project-tmp.sh paths 2>/dev/null | sed -n 's/^PROJECT_TMP_ROOT=//p')
ifeq ($(strip $(PROJECT_TMP)),)
$(error Cannot resolve safe Go project temporary root; inspect PROJECT_TMP_ROOT / RUNNER_TEMP / original TMPDIR)
endif
export PROJECT_TMP_ROOT := $(PROJECT_TMP)
RUN_ID ?= $(shell date -u +%Y%m%dT%H%M%S%N)
PROJECT_RUN := $(PROJECT_TMP)/runs/make/$(RUN_ID)
export TMPDIR := $(PROJECT_RUN)
export TMP := $(PROJECT_RUN)
export TEMP := $(PROJECT_RUN)
export GOTMPDIR := $(PROJECT_TMP)/build/go
export GOCACHE := $(PROJECT_TMP)/cache/go/build
export GOMODCACHE := $(PROJECT_TMP)/cache/go/mod
export GOPATH := $(PROJECT_TMP)/cache/go/path
export XDG_CACHE_HOME := $(PROJECT_TMP)/cache/xdg
export NUGET_PACKAGES := $(PROJECT_TMP)/cache/nuget
export DOTNET_CLI_HOME := $(PROJECT_TMP)/cache/dotnet/home
export OOXML_VALIDATOR_DLL := $(PROJECT_TMP)/build/dotnet/OoxmlValidator.dll
export OOXML_ORACLE_SCRATCH := $(PROJECT_RUN)/oracles
export PYTHONDONTWRITEBYTECODE := 1
export PYTHONPYCACHEPREFIX := $(PROJECT_TMP)/cache/python/pycache

GO ?= go
GOFMT ?= gofumpt
GOLINT ?= golangci-lint
GOSEC ?= gosec
TEST_JOBS ?= 2
DOTNET_ROOT ?= /home/linuxbrew/.linuxbrew/opt/dotnet/libexec
VALIDATOR ?= tools/validator/OoxmlValidator

prepare-temp:
	@mkdir -p "$(TMPDIR)" "$(GOTMPDIR)" "$(GOCACHE)" "$(GOMODCACHE)" "$(GOPATH)" "$(XDG_CACHE_HOME)" "$(NUGET_PACKAGES)" "$(DOTNET_CLI_HOME)" "$(OOXML_ORACLE_SCRATCH)" "$(PYTHONPYCACHEPREFIX)" "$(PROJECT_TMP)/build/dotnet/obj"

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

deps: prepare-temp ## Download and tidy dependencies
	$(GO) mod download
	$(GO) mod tidy

install: prepare-temp ## Install the binary
	$(GO) install ./...

lint: prepare-temp ## Run golangci-lint
	@which $(GOLINT) > /dev/null || (echo "Installing golangci-lint..." && brew install golangci-lint)
	$(GOLINT) run ./...

security: prepare-temp ## Run gosec security scan
	@which $(GOSEC) > /dev/null || (echo "Installing gosec..." && $(GO) install github.com/securego/gosec/v2/cmd/gosec@latest)
	$$(go env GOPATH)/bin/$(GOSEC) ./...

format: prepare-temp ## Format code with gofumpt
	@which $(GOFMT) > /dev/null || (echo "Installing gofumpt..." && $(GO) install mvdan.cc/gofumpt@latest)
	$(GOFMT) -w .

test: prepare-temp ## Run profiled native tests as a bounded batch
	GOMAXPROCS=2 bash scripts/test-profile.sh ./...

.PHONY: acceptance test-batch
acceptance: prepare-temp ## Run profiled strict native Gherkin and report reconciliation
	cd acceptance && GOMAXPROCS=2 bash ../scripts/test-profile.sh ./...

test-batch: test acceptance ## Run library and acceptance modules

# Bounded graphics candidate tests; never advance the released reference pin.
GRAPHICS_ROOT ?= $(abspath ../fixtures-ooxml)
.PHONY: graphics-test
graphics-test: prepare-temp ## Run profiled graphics packages against the sealed shared candidate
	OOXML_GRAPHICS_ROOT="$(GRAPHICS_ROOT)" GOMAXPROCS=2 bash scripts/test-profile.sh ./pkg/presentation ./pkg/packaging ./internal/losslessxml

GRAPHICS_EVIDENCE ?= $(abspath artifacts/graphics/final)
.PHONY: graphics-reconcile graphics-quality
graphics-reconcile: prepare-temp ## Reconcile every sealed graphics recipe with profiled Go leaves
	mkdir -p "$(GRAPHICS_EVIDENCE)/outputs"
	OOXML_GRAPHICS_ROOT="$(GRAPHICS_ROOT)" PROFILE_OUTPUT_DIR="$(GRAPHICS_EVIDENCE)/outputs" PROFILE_JSON_EVENTS="$(GRAPHICS_EVIDENCE)/go-events.jsonl" GOMAXPROCS=2 bash scripts/test-profile.sh ./pkg/presentation ./pkg/packaging ./internal/losslessxml
	GOMAXPROCS=2 $(GO) run ./tools/graphicsreconcile -root "$(GRAPHICS_ROOT)" -events "$(GRAPHICS_EVIDENCE)/go-events.jsonl" -output "$(GRAPHICS_EVIDENCE)/reconciliation.json"

# Generate once with graphics-reconcile; all optional oracles share those outputs.
graphics-quality: prepare-temp ## Run all LibreOffice quality oracles on reconciled outputs
	@test -f "$(GRAPHICS_EVIDENCE)/reconciliation.json"
	@set -e; for spec in insertion:picture-insertion replacement:picture-replacement metadata:picture placement:picture-placement svg:picture-svg delete:picture-delete group:shape-group group-transform:group-transform connectors:connectors diagrams:diagrams smartart:smartart autoshapes:autoshapes freeform:freeform z-order:z-order gradients:gradients opacity:opacity outlines:outlines; do \
		name=$${spec%%:*}; prefix=$${spec#*:}; \
		mkdir -p "$(GRAPHICS_EVIDENCE)/quality/$$name-outputs"; \
		cp "$(GRAPHICS_EVIDENCE)/outputs/$$prefix-"*.pptx "$(GRAPHICS_EVIDENCE)/quality/$$name-outputs/"; \
	done
	@set -e; for oracle in picture-insertion picture-replacement picture-metadata picture-placement picture-svg picture-delete shape-group group-transform connectors diagrams autoshapes freeform z-order gradients opacity outlines; do \
		echo "LibreOffice $$oracle"; \
		timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/$$oracle-uno.py "$(GRAPHICS_EVIDENCE)/quality" > "$(GRAPHICS_EVIDENCE)/$$oracle-quality.log" 2>&1 || { tail -30 "$(GRAPHICS_EVIDENCE)/$$oracle-quality.log"; exit 1; }; \
	done
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/smartart-uno.py "$(GRAPHICS_EVIDENCE)/quality/smartart-outputs" > "$(GRAPHICS_EVIDENCE)/smartart-quality.log" 2>&1

.PHONY: graphics-insertion-quality
graphics-insertion-quality: ## Run optional LibreOffice insertion save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/insertion-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-insertion-uno.py

.PHONY: graphics-replacement-quality
graphics-replacement-quality: ## Run optional LibreOffice replacement save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/replacement-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-replacement-uno.py

.PHONY: graphics-metadata-quality
graphics-metadata-quality: ## Run optional LibreOffice crop/transform save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/metadata-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-metadata-uno.py

.PHONY: graphics-placement-quality
graphics-placement-quality: ## Run optional LibreOffice fitted-placement save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/placement-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-placement-uno.py

.PHONY: graphics-svg-quality
graphics-svg-quality: ## Run optional LibreOffice paired SVG save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/svg-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-svg-uno.py

.PHONY: graphics-delete-quality
graphics-delete-quality: ## Run optional LibreOffice deletion save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/delete-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/picture-delete-uno.py

.PHONY: graphics-group-quality
graphics-group-quality: ## Run optional LibreOffice grouped-child save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/group-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/shape-group-uno.py

.PHONY: graphics-group-transform-quality
graphics-group-transform-quality: ## Run optional LibreOffice group-frame save/reopen oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/group-transform-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/group-transform-uno.py

.PHONY: graphics-connectors-quality
graphics-connectors-quality: ## Run optional LibreOffice connector attachment oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/connectors-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/connectors-uno.py

.PHONY: graphics-diagrams-quality
graphics-diagrams-quality: ## Run optional LibreOffice diagram edit and attachment oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/diagrams-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/diagrams-uno.py

.PHONY: graphics-smartart-quality
graphics-smartart-quality: ## Run optional LibreOffice sealed-source SmartArt oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/smartart-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/smartart-uno.py

.PHONY: graphics-autoshapes-quality
graphics-autoshapes-quality: ## Run optional LibreOffice AutoShape label-edit oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/autoshapes-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/autoshapes-uno.py

.PHONY: graphics-freeform-quality
graphics-freeform-quality: ## Run optional LibreOffice custom-vector oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/freeform-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/freeform-uno.py

.PHONY: graphics-order-quality
graphics-order-quality: ## Run optional LibreOffice graphical-order oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/z-order-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/z-order-uno.py

.PHONY: graphics-gradients-quality
graphics-gradients-quality: ## Run optional LibreOffice gradient readback oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/gradients-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/gradients-uno.py

.PHONY: graphics-opacity-quality
graphics-opacity-quality: ## Run optional LibreOffice shape/picture alpha oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/opacity-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/opacity-uno.py

.PHONY: graphics-outlines-quality
graphics-outlines-quality: ## Run optional LibreOffice direct outline scalar oracle
	PROFILE_OUTPUT_DIR="$(abspath artifacts/graphics/outlines-outputs)" $(MAKE) graphics-test
	timeout --kill-after=5s 180s /usr/bin/python3 tools/oracles/outlines-uno.py

.PHONY: shared-pack-check
shared-pack-check: acceptance ## Verify shared distribution and profiled native readbacks in a complete batch

coverage: prepare-temp ## Run profiled tests with coverage in retained profile evidence
	PROFILE_COVERAGE=1 GOMAXPROCS=2 bash scripts/test-profile.sh ./...

bench: prepare-temp ## Run profiled benchmarks (review bytes/op and allocs/op too)
	PROFILE_GO_FLAGS='-run ^$$ -bench . -benchmem' GOMAXPROCS=2 bash scripts/test-profile.sh ./...

memprofile: prepare-temp ## Run profiled memory tests
	ENABLE_MEMPROFILE=1 PROFILE_GO_FLAGS='-run MemProfile' GOMAXPROCS=2 bash scripts/test-profile.sh ./...

check: lint test ## Run lint + tests

build: prepare-temp ## Build the library (verify compilation)
	$(GO) build ./...

# =============================================================================
# Clean targets
# =============================================================================

clean: ## Safe default: do not delete shared caches, active run scratch or retained evidence
	@echo 'No automatic deletion: $(PROJECT_TMP)/{cache,build,runs} may contain active jobs. Review and remove confirmed idle disposable paths manually; preserve artifacts/profiles.'

clean-all: clean ## Same safe scope; never remove vendor, evidence or installed toolchains

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

validate: prepare-temp ## Validate the two historical root fixtures from the pinned shared checkout using the official SDK
	@set -e; if [ ! -f "$(OOXML_VALIDATOR_DLL)" ]; then \
		echo "Building validator into project-owned build output..."; \
		DOTNET_ROOT="$(DOTNET_ROOT)" dotnet build "$(VALIDATOR)/OoxmlValidator.csproj" -c Release -m:1 -p:UseSharedCompilation=false -p:BaseIntermediateOutputPath="$(PROJECT_TMP)/build/dotnet/obj/" -p:OutputPath="$(PROJECT_TMP)/build/dotnet/" --disable-build-servers -q; \
	fi; \
	export DOTNET_ROOT="$(DOTNET_ROOT)"; \
	for id in fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5 fixture-151d747bc37d4f4988c1116f4abb45196b1c1644319ce342ee4dd56d111f3132; do \
		f=$$($(GO) run ./tools/fixturepath "$$id"); \
		dotnet "$(OOXML_VALIDATOR_DLL)" "$$f"; \
	done
