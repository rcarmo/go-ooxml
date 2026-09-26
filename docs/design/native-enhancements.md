## Mapping extended Office enhancements onto go-ooxml

Date: 2026-09-26. Status: architecture proposal and source review; no library implementation changes.

`go-ooxml` already provides the Go object models, OPC packaging and Office fixtures needed to start. extended parity requires a preservation-safe editing path, stronger relationship handling and format-specific mutation engines. Reuse the existing authoring APIs where their contracts fit; add guarded editing without silently changing the behaviour of existing callers.

The scope is the complete set of extended enhancement families, including their refusal behaviour. Recreating every inherited Python API is a separate undertaking. Runtime code stays Go-native. Python may generate reference outcomes during development; LibreOffice remains an optional external calculation/rendering tool.

### Sources and baseline

| Source | Reviewed revision |
|---|---|
| rcarmo/go-ooxml | `43eda6edfe45db20da5879ba332ec6a25d598efa` |
| retired-docx | `[retired implementation revision]` |
| retired-pptx | `[retired implementation revision]` |
| retired-xlsx | `[retired implementation revision]` |

The Go revision matched origin HEAD during this review. extended scope comes from the pinned user/API documentation, implementation reads and the 120 `[retired external test identity]` modules (50 DOCX, 32 PPTX, 38 XLSX). `source-inventory.json` records these files with hashes; inventory is not evidence of port coverage. A function/parameter-case ledger must be completed before any claim of full parity.

`go test ./...` ran under Go 1.26.3. Four tests failed because they address `/workspace/testdata/...` instead of repository-relative fixtures: one presentation test and three spreadsheet tests. The remaining reported packages passed. The unchanged log is included. Fix fixture lookup before establishing the acceptance baseline; missing-fixture failures are not valid behavioural red tests.

There are 74 DOCX/XLSX/PPTX files under the Go repository's `testdata`, plus fixtures under the Python MCP server's `tests/_templates`. extended supplies additional synthetic, malformed, interoperability and specialised fixtures. Record file hashes, provenance and redistribution permissions; do not use the Python MCP checkout as a production dependency.

### Shared implementation first

The existing `pkg/packaging` is the right place to extend package handling. Keep WML, PML and SML semantics separate; their text, inheritance, references and cache rules differ.

| Area | Existing implementation | Required addition |
|---|---|---|
| Archive intake | `pkg/packaging/package.go:52` eagerly reads members into parts | Validated member index; duplicate/collision/path checks, CRC/size/overlap checks, compression policy and configurable read limits. Reject external resource fetching. |
| Source preservation | Parts retain payload bytes; some models retain selected raw XML | Immutable original package/member bytes plus a dirty overlay. Preserve unknown namespaces, attributes, elements, relationships and content-type metadata. |
| XML edits | Typed structs marshal supported fields; selected `innerxml` hooks | Namespace-aware, lossless ordered representation or bounded source splices for changed parts. An unknown sibling inside an edited part must survive too. |
| Relationships | OPC relationships, target resolution and ID allocation exist | Graph ownership, inbound-edge accounting, collision-free imports, explicit shared/deep-copy policies and conservative deletion. Never garbage-collect unrelated unknown parts by default. |
| Mutation | Direct in-memory editing | Preflight plans plus operation-local undo journals or copy-on-write transactions. Refusal must restore held handles as well as document state. |
| Save | `SaveAs` creates the destination directly; `WriteTo` rewrites registries and iterates maps | Stage, validate, write, check ZIP/file close errors, fsync and replace. Stream saves explicitly lack filesystem atomicity. Stable ordering for rebuilt members. |
| Verification | Basic round-trip tests | Part hashes, namespace-aware semantic diffs, relationship/content-type validation, changed-part budgets and reopen checks. |
| Results | Format-specific APIs | Versioned reports, typed refusals and stable codes usable with `errors.As`; preserve programmer/I/O errors as distinct categories. |

A new preservation path must not call existing `updatePackage()` methods unconditionally: they can normalise or discard XML before the package overlay sees it. A semantic diff cannot recover an unknown element already lost during parsing. Keep the original bytes authoritative and serialise only proved edits.

Do not promise identical archive bytes after an edit. Define three separate assertions: untouched member payloads are byte-identical; unchanged regions of a changed part meet their stated structural contract; the reconstructed package is readable and coherent. Exact no-op archive copying can be a stronger, explicit contract. XML comparisons must preserve text whitespace, child ordering where meaningful, namespace bindings and QName-valued attributes; generic whitespace stripping or attribute sorting is insufficient.

`Part.Content()` exposes a mutable slice and relationships are mutable pointers. Existing modified flags therefore cannot alone prove which content changed. Safe sessions need immutable snapshots or independent fingerprint checks. Save-time validation cannot retrofit rollback guarantees onto arbitrary legacy setters.

Use a shared typed refusal structure such as `Kind`, `Operation`, `Part`, `Target` and `Options`. Format-specific errors may wrap it. Expected kinds include ambiguous/missing/stale target, unsupported structure, protected operation, relationship policy, invalid package and unavailable/timed-out oracle. Never downgrade a refusal into a destructive legacy retry.

Signed packages need an explicit policy: edits invalidate signatures; preserve-and-refuse or an explicitly authorised signature-removal operation. Macro parts should remain opaque and unexecuted. Protection checks reproduce editing policy; document protection is not encryption or a security boundary.

### DOCX mapping

Existing Word support includes real insertion/deletion revisions, basic accept/reject, comments and replies, lists, fields, content controls, styles, tables and headers/footers. These are useful implementation pieces, but their narrower semantics need characterisation.

| ID | extended enhancement | Go reuse and missing work |
|---|---|---|
| D01 | All-story outline and current/original/all views | Extend body/table/header traversal to footnotes, endnotes, text boxes, SDTs and revision wrappers. Report unreadable regions. Historical views cannot authorise current edits. |
| D02 | Exact/normalised search, contextual ranking and live spans | Build visible-text-to-run maps and wrapper ownership over ordered XML. Explicitly handle repeated matches, cross-paragraph inspection, Unicode offsets and stale/consumed targets. |
| D03 | Run-preserving replace/replace-all, tracked replacement, correction inside existing insertions | Reuse run properties and insertion/deletion writers. Add unique affix localisation, positional-marker guards, all-or-nothing bulk edits and result flags. Generic `ReplaceText` is insufficient. |
| D04 | Effective formatting and provenance | Use `styles.go`; add defaults/style-chain/direct-format resolution, mixed values and unresolved state. Do not confuse explicit formatting with rendered appearance. |
| D05 | Paragraph/block insert/delete/replace | Reuse paragraph/table operations with live endpoint, field-boundary, section-property and structural-owner validation. |
| D06 | Table lookup, cell updates and template-row operations | Extend `table.go` with ambiguity checks, merge/grid validation and atomic editing. |
| D07 | Numbering and lists | Extend `numbering.go` with validated definitions, levels and copy/remap rules. Preserve real list markup instead of typed bullet characters. |
| D08 | Content controls and bookmarks | Extend `content_control.go` and bookmark helpers with unique targeting, placeholder clearing, protection and marker-pair validation. |
| D09 | Fields and TOC | Extend `field.go` with page/date/reference/TOC helpers and field-boundary checks. Report placeholder results; no built-in Word field calculator. |
| D10 | Full revision inventory and resolution | Extend `tracking.go`/`tracking_manager.go`: all stories, property changes, row changes and paired moves; author filters, stable locations and compound atomicity. `RevisionLocation` is currently empty. |
| D11 | Comment threads | Reuse `comments.go` and `commentsExtended.xml`; add exact multi-run anchors, parent/thread lookup, resolve/reopen, metadata integrity and appropriate protection checks. |
| D12 | Document protection | New central per-mode gate for every safe mutation, including the comments-only exception. |
| D13 | Cross-document composition | New range-copy planner: reconcile styles, numbering, media, hyperlinks, bookmarks, sections and relationships; refuse unresolved ownership; return touched-part reports. |
| D14 | Compare/redline generation | New structural/text comparison over D01-D13. Verify `accept(compare(A,B))` matches B and `reject(compare(A,B))` matches A on private copies within the supported contract. Refuse unrepresentable package differences. |
| D15 | Package/text diff, diagnose, pending changes, patch save | Shared package layer plus Word semantic views; changed-part budget and reopen checks. |

`pkg/document/document.go:172` currently exposes one body section wrapper. Section-aware composition therefore needs a fuller representation even though generic section formatting exists. Do not infer Word parity from the presence of a `Section` interface.

### PPTX mapping

Go preserves many chart, theme, media and diagram parts and has slides, shapes, notes and tables. Preservation on a no-op round trip does not establish that clone/import/edit operations correctly own those parts.

| ID | extended enhancement | Go reuse and missing work |
|---|---|---|
| P01 | Text/deck manifests and stable anchors | Extend shape traversal into groups and table cells; versioned geometry/layout inventory, structural identity, full fingerprints and exact refind. |
| P02 | Effective font/paragraph/shape properties and provenance | New inheritance resolver across run, paragraph, list style, placeholder, layout, master, theme and colour maps. Include independently inherited bullet kind/font/size and unresolved values. |
| P03 | Run-safe text replacement | Replace flattening setters with guarded paragraph/run operations; preserve fields, breaks and untouched regions; refuse ambiguous/stale targets. |
| P04 | Batch transactions | Preflight/rollback boundary with final package checks; no save inside an open batch. Legacy mutation inside a batch needs comprehensive snapshot validation. |
| P05 | Clone/delete/move/reorder slides | Extend existing APIs with relationship closure, chart/workbook and notes deep copies, deliberate media sharing, sections/custom shows and unsupported-relationship refusals. |
| P06 | By-name shapes, copy/delete/z-order, image replacement | Unique group-aware lookup; ownership-safe relationship cleanup; shared-image isolation/deduplication while preserving position, size and crop. |
| P07 | Table rows/columns/merge/extend/split | Extend `table.go` with column surgery, cell-wise merge checks, stale-cell checks, copied direct formatting and owner-frame dimensions. |
| P08 | Bullet authoring and autofit normalisation | Real bullet/numbering markup and explicit autofit policy. Read-only picture-bullet inspection does not imply picture-bullet authoring. |
| P09 | Notes and genuine footer/date/slide-number fields | Reuse notes support, target body placeholders, keep absent notes absent, author `a:fld` and defer field refresh to Office. |
| P10 | Safe chart data replacement | Inspect chart family and ownership; support workbook-less charts; stage cache/workbook changes, refuse unsupported structures, clone dependent workbooks independently. |
| P11 | Layout rebind | Exact-first unique placeholder matching, explicit partial maps/orphan policy, before/after effective-format report. |
| P12 | Slide import/deck append | Explicit adopt-theme/keep-appearance/bake modes; hash-deduplicated master/layout/theme chains, staged whole-deck preflight, section name/GUID targeting and source immutability. |
| P13 | Comments, section bookkeeping, unused layout removal, send-safe delivery | Keep Go's existing comment APIs but honour extended's clone/import refusal/drop policies. Add section/custom-show consistency and safe reachability-based removal only when explicitly requested. General threaded-comment or section authoring is not a extended enhancement to invent. |
| P14 | Semantic deck/package diff and patch save | Lineage-aware slide and shape IDs, exact text snapshots, chart/image/table/notes facets, conservative effective-format pairing and package fallback. Independent decks have no inferred shared identity. |

`presentation.go:425` duplicates slide content without copying the full related-part graph. `presentation.go:1191` uses a fixed modern-comment part name; multi-slide comments need a targeted regression before changing it. These should be checked before building the corresponding safe APIs.

### XLSX mapping

The workbook/cell/style/table APIs provide basic authoring. This format needs the largest new subsystem because one edit can affect references and cached results throughout the workbook.

| ID | extended enhancement | Go reuse and missing work |
|---|---|---|
| X01 | Preserve-mode intake/save/validate and receipts | New immutable-source overlay and complete planner; path/stream classification, workbook content-type preservation and explicit legacy mode. Validate and save must share one plan. |
| X02 | Atomic writes/append and protection | Local undo journals; cell, style, hyperlink, comment and registry rollback; warning/strict protection policy and data-only formula-loss guard. |
| X03 | Search, error scan, validation vocabulary and workbook diff | Search values/formulas; token-aware error operands; deterministic literal/static list validation; refusal for ambiguous/dynamic lists; remap-aware diffs. |
| X04 | Format copy | Finite-range copy of the documented six style groups only; preflight merged interiors/protection and preserve values, formulas and other cell metadata. |
| X05 | Row/column insertion/deletion and sheet/range edits | New address-remap planner; characterise inherited rename/copy/move/delete behaviour individually instead of assuming blanket support. |
| X06 | Reference rewriting | Formula lexer/reference representation covering mixed/absolute A1 refs, sheet quoting, names/scopes, ranges and tables. Walk names, print settings, filters, validation, CF, hyperlinks, charts and pivot sources. Refuse unresolved dynamic, external or unsupported structures according to the operation. |
| X07 | Formula freshness and calculation metadata | Track dirty dependencies, including non-value inputs where relevant; invalidate affected caches only, remove/rebuild real calc-chain metadata and request recalculation. Style-only changes must not indiscriminately clear caches. |
| X08 | Validation, conditional formatting and unknown worksheet extensions | Preserve raw blocks; add only the schema and rewriting required for supported edits, including x14 cases. Missing schema fields must not discard data. |
| X09 | Table append | Prove supported geometry, headers, totals/filter conventions, destination availability and declared calculated-column formulas before touching cells. Refuse query/external/extension/sort/array ambiguities. |
| X10 | Loaded image replacement | Resolve drawing identity, allocate fresh media and retarget the selected relationship while preserving original drawing XML and other users. |
| X11 | Loaded chart repoint/cache invalidation | Add loaded-chart identity and reference patching; remove the corresponding stale cache without replacing the entire chart. |
| X12 | Pivot impact and refresh-on-open | New pivot-source/dependency analysis, explicit scoped permission, shared cache handling, volatile/UDF rules and receipts stating that cached results remain stale until Excel refreshes. |
| X13 | Oracle recalc/certify/evaluate | Optional Go subprocess adapter with deadlines, bounded output, isolated profiles and temporary copies. Selectively splice eligible caches into the preserved source structure; never deliver LibreOffice's rewritten archive. |
| X14 | Certification comparisons | Strict status plus coverage/exclusions; absolute/relative/ULP differences and explicit caller tolerance classification. No automatic Excel-equivalence or financial-correctness claim. |

The current save path has prerequisite defects: `workbook.go:852` replaces relationships with a fixed list; `workbook.go:1001` writes a synthetic calc chain; `updatePackage()` forces the ordinary workbook content type. `pkg/ooxml/sml/worksheet.go` lacks data-validation modelling. Correct these with dedicated regression fixtures before routing existing workbooks through the new path.

Reference rewriting and formula evaluation are different projects. Native Go tokenisation and dependency analysis are required for safe edits; a complete Excel calculation engine is outside this enhancement scope. Static analysis needs conservative fallbacks for dynamic dependencies, volatile formulas, shared/array formulas and unresolved extensions.

### Go API compatibility

Keep `pkg/document`, `pkg/presentation`, `pkg/spreadsheet`, `pkg/ooxml/*` and `pkg/packaging`. Introduce safe sessions/adapters through additive exported constructors and concrete result types. Names such as `OpenPreserving`, `Plan`, `Validate`, `Receipt` and `Refusal` are proposals, not current APIs.

Adding methods to exported Go interfaces breaks downstream implementations even if callers still compile. Prefer new capability interfaces or adapters; change existing interfaces only with a versioned compatibility decision. Internal safe sessions may need small package-owned hooks to avoid exposing mutable XML nodes publicly or creating import cycles.

Keep mutable sessions single-owner; concurrent readers can use immutable snapshots. Anchors should carry package/session identity, stable owner identity and a scoped fingerprint. Unrelated index changes need not invalidate every anchor; deletion, structural movement and changed content must follow each format's explicit rules. Unicode search must define runes/UTF-16/byte offsets at boundaries and test supplementary characters and combining sequences.

### Gherkin -> Go tests -> measured outcomes

Use Gherkin as the behaviour contract. Use Godog for Go bindings in an isolated acceptance module, keeping the shipped library's zero-external-runtime-dependency goal. Pin a maintained runner and the official Cucumber parser. A separate module must have an explicit CI target; root `go test ./...` will not discover it automatically.

Proposed layout:

```text
features/{planned,implemented}/{package,document,presentation,spreadsheet}/
acceptance/go.mod                 # test-only Godog dependencies
acceptance/steps/
acceptance/testdata/manifest.json  # fixture source, licence and hash
internal/testutil/                # package/semantic assertions
spec/native-capability-history.json            # pinned source contract -> scenario IDs
reports/acceptance/               # generated execution results and diffs
```

Give each scenario a stable ID. Use exactly one lifecycle tag (`@planned`, `@implemented` or `@external`) and a runner tag such as `@go` or `@office`. Oracle contract tests can run with a fake process under `@implemented @go`; live LibreOffice and Microsoft Office checks are separate `@external` cases with recorded versions. Missing configured external prerequisites fail that run and never become a green pass.

The delivery loop follows minicore with batched execution: define a coherent group of outcomes; write Gherkin and bindings; run the behavioural red-test batch against unchanged code; implement the changes; run the affected acceptance/unit/regression batch; update the result map and commit promptly after it passes. Characterisation of existing functionality may start green. Undefined steps, missing fixtures and compilation failures do not qualify as behavioural red tests.

### Test batches and commits

Explicit user instruction: no individual test runs; commit regularly after a test batch passes. Optimise elapsed time and CPU usage.

* Group related scenarios and unit/regression tests into package or feature-family batches, including initial red checks and failure reruns. Inspect batch logs and make related fixes together; do not enter single-test retry loops.
* Reuse Go build/test caches, bound concurrency, and avoid multiple agents running the same suites. Run full-repository, race/fuzz and expensive Office/oracle checks at integration/release gates or when a relevant change requires them, in batches.
* After every passing implementation batch, inspect the diff, retain the exact commands/results and covered scope, and commit the tested changes before starting the next batch. Do not accumulate several green batches uncommitted or include unrelated work. Scoped green results do not imply that unrun global suites passed.
* Configure both local and global Git identity as `Rui Carmo <rui.carmo@gmail.com>` before committing. Never rebase; use merge. Keep this design and the plan sidebar aligned with actual batch results and commit references.

For each enhancement, cover success, ambiguity/unsupported input, stale target where relevant, protection where relevant, rollback including held handles, allowed package changes and reopen. Cross-format tests assert integrity rather than identical XML serialisation. Add parser/rewriter fuzzing and Go race tests outside Gherkin.

Use exact scenario and Examples-row identities to reconcile the inventory with execution output. Missing, skipped, undefined, flaky or unexecuted `@implemented` cases fail acceptance. Keep planned cases visible in the denominator. Do not infer pass counts from scenario files, test-file counts or generated bindings.

The supplied `features/*.feature` files are illustrative planned contracts. Their Gherkin syntax and stable-ID shape are checked by `gherkin-check`; they have no Go step bindings and are not implementation evidence. The full family matrix above must be expanded into concrete scenarios and parameter cases during implementation.

Each result record should include source revision/test identity, feature/scenario/Examples identity, fixture hash, Go revision and version, outcome, refusal code if expected, part diff, semantic assertions and any Office renderer/calculator version. Binary outputs and report schemas should be versioned or hash-addressed.

### Delivery order

| Stage | Work | Exit condition |
|---|---|---|
| A | Fixture-path repair, source/test inventory and existing-API characterisation | Reproducible baseline; every known extended family accounted for; known gaps remain planned. |
| B | Package preservation, ZIP/XML validation, transactions, typed refusals and atomic save | No-op and one bounded edit per format preserve the required bytes/structure; all refusal and write-failure tests pass. |
| C | DOCX stories/spans; PPTX manifests/formatting; XLSX reference lexer and dependency index | Read-only inventories are complete within declared scope; unknowns and ambiguous identities are explicit. |
| D | Safe local edits: Word blocks/comments/tables, PowerPoint text/tables/images/notes, spreadsheet cells/formats/tables | Success and rollback contracts pass on synthetic and existing fixtures. |
| E | Word composition/full revisions/compare; presentation graph operations/rebind/import; spreadsheet structural edits/charts/pivots | Cross-part invariants, source immutability and derived effects pass independently checked tests. |
| F | Diffs, reports, optional oracle and Office corpus runs; documentation/API parity audit | Exact source-to-scenario reconciliation, all implemented cases green, exclusions and external checks reported separately. |

The three format tracks can proceed in parallel after Stage B, using its stable transaction and preservation contracts. Keep delivery increments small; a single successful cell edit, text replacement or slide clone is a better first target than simultaneous broad API scaffolding. An effort estimate requires the per-case ledger and measurements from Stage B; package/test counts do not supply one.

### Reference links

* [go-ooxml packaging source](https://github.com/rcarmo/go-ooxml/blob/43eda6edfe45db20da5879ba332ec6a25d598efa/pkg/packaging/package.go)
* [go-ooxml spreadsheet save path](https://github.com/rcarmo/go-ooxml/blob/43eda6edfe45db20da5879ba332ec6a25d598efa/pkg/spreadsheet/workbook.go)
* [extended DOCX enhancement contracts]([retired implementation citation])
* [extended PPTX enhancement contracts]([retired implementation citation])
* [extended XLSX enhancement contracts]([retired implementation citation])
* Workflow references: `rcarmo/minicore/features/README.md`, `docs/development/methodology.md`; `rcarmo/gi/scripts/test-tui-gherkin.sh`.

### Increment 002: existing package I/O hardening

The existing package API now rejects unsafe, duplicate and case-colliding ZIP
member names before reading their payloads, along with encrypted/unsupported
compression members. `packaging.Refusal` exposes a typed invalid-package result.
This increment does not yet implement header/overlap checks or configurable
resource budgets, so B01 remains partial.

`Package.SaveAs` stages a sibling temporary file, checks serialization, file sync
and close, then renames it over the destination. Existing regular-file permission
bits are retained; new outputs use mode 0600. A failed pre-replacement operation
leaves destination bytes and package path/modified state unchanged. Stream output
returns archive-finalisation errors and rejects closed packages. Member groups
serialize in sorted order. Directory fsync/crash durability and platform-specific
replacement semantics are not certified by this increment. Source preservation,
relationship closure validation and guarded format transactions are still pending.
