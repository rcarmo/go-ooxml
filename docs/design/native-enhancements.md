## Mapping extended Office enhancements onto go-ooxml

Date: 2026-09-26. Status: implementation underway on `native-enhancements` in `/workspace/worktrees/go-ooxml-native`; reference checkout unchanged.

Checkpoint after batches 009-018: immutable namespace-aware XML/leaf edits (`a7d3477`), read-only OPC graph (`93cb432`), bounded Word/PPTX/numeric XLSX editing (`fb94bb5`, `c79e906`, `145446d`), refusal hardening (`4766fbb`), descriptor/prefix fixes and corpus checks (`bc6074b`), expanded source collection (`d785bb5`), and Word story views (`5abdec0`). Earlier package foundations are recorded in batches001-008.

Latest checkpoint after batches019-031: exact Word run spans (7e865b4), affix-preserving replacement (73229e2), selected batches (5ea5e9a), normalised/context search (246cf9e), lossless attributes (cc5d7ec), xml:space authoring (3a9f67c), replace-all reports (880352f), scoped story/view search (7a6d7e9), separator/casefold fixes (191041d), proved boundary insertion (b59e695), invisible-content/QName guards (903f2fe), partial source-case links (4ae68c4).

Measured state after batches062-065:225 implemented expanded native Gherkin cases pass;20 original planned plus one external remain unrun. Full/race integration with V2 passed065. Notes body edit29ed125, grouping-lock/source-fixture fixfb01945, namespace intake characterisationc7c8bc3. See reports/batches/065.md. Full/race post-reboot integration passed061: graph deletion/mixed transactions(c1d7f1c),owned calcChain cleanup(de006f4),lexical XML root-boundary fix(c965780),lifecycle checks(dc3e2f0). Read reports/batches/061.md. Local ZIP64/explicit descriptors, terminal boundaries and edited ZIP64 delivery and65535-entry writer/intake checks pass; see reports/batches/055.md (3b223a0 verification,5c7b8d8 delivery tests). Full/race integration with V2 passed. Bounded graph additions/retargets and loaded XLSX picture replacement are implemented; full/race integration with V2 passed. See reports/batches/051.md. Commits722358d graph plans,e711e36 image adapter,9e1f1f4 target-form preservation. One exact sharedV2 cache contract has direct-native execution evidence (schema1 report, formal shared schema2/Gherkin binding pending);18 shared workflows unexecuted. Four pinned inputs continue integrity/readback verification. Batch046 full native/race suites and formula/dependency invariant seeds pass.

New commits:a4be6fc static formula parser;432e136 atomic cache invalidation;7632497 exact shared cache contract/inactive markers;c02505e token-kind fix;3e9c389 worksheet structural guards;072d6d5 insertion reference remap primitive. Scope/remaining X05-X07 gaps in reports/batches/046.md. Previous XML normalization/insertion checkpoint is batch039. No formulas calculated or full Excel reference grammar claimed.

Prior shared coordination: e1c0deb default-style zero; c4402bd V2 pin/r:id namespace; ed123ac native input readbacks. Pin spec/shared-contracts-v2.json. Go/Bun/Python worktree ownership remains separate; shared contracts unchanged.

Earlier corpus: all74 Go fixture archives retain no-op bytes. Ledger:3451 collected source cases, thirteen qualified partial links (3438 unlinked), no completed upstream parity. Batch031 invariant seeds passed; no exploratory fuzz or external Office runs. Full port scope and next work: reports/batches/031.md.

Next: finish B01 ZIP64/overlap/URI policy, B03 structural XML operations, B04 graph surgery and B10 adversarial/corpus gate; expand D01 story coverage, D02-D03 multi-run spans, P01-P03 complete text/inheritance and X06-X07 reference/cache semantics. The three editing adapters implement deliberately narrow subsets, not family completion. Source-parity ledger and native Office/LibreOffice checks remain unfinished. Reports: `reports/batches/001.md` through `018.md`.

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

### Increment 003: explicit intake budgets

`packaging.OpenReaderWithLimits` adds source-byte, entry-count, per-part and total
inflated-byte budgets without changing existing OpenReader defaults. Zero disables
a budget. Source bytes are checked before ZIP parsing; count/declared sizes are
checked before decompression, then actual payload length and CRC are validated.
Payload allocation is bounded by the declared size plus one. Invalid arguments,
resource-limit refusals and invalid-package refusals remain distinct. Central
metadata parsing still requires a finite source budget to bound its input.
Header/overlap validation is pending. See `reports/batches/003.md`.

### Increment 005: physical ZIP validation

Local and central member names, flags, compression, ordinary sizes and CRC must
agree. Payload/descriptor ranges cannot overlap another entry or the central
directory. Ordinary data descriptors are checked against directory metadata.
ZIP64 central offsets are recognised; local ZIP64 sentinel sizes without data
descriptors refuse until their extras can be verified. This is intentionally
conservative and does not establish full ZIP64 parity. Independent overlap/ZIP64
adversarial tests and URI-escape policy remain pending. See batch 005.

### Increment 006: retained-source payload adapter

The additive `packaging.Preserved` API owns cloned source bytes, exposes cloned
payload/fingerprint reads and atomically applies guarded existing-part replacement
batches. A no-op writes exact source archive bytes; edited output raw-copies
untouched members. This low-level API checks XML syntax, not OOXML semantics.
Registry/add/delete/import operations and signed edits refuse. Run-safe editing,
reference/cache updates, graph validation and path delivery integration remain
pending. Read `SPEC.md` and `reports/batches/006.md` for the implemented subset.

### Increment 007: verified path delivery and payload receipts

Retained sessions can deliver through `SaveAs`, which reopens the temporary archive
and compares every member payload before atomic replacement. Schema-1 receipts
contain sorted original/current payload hashes; staged receipts do not claim that
anything was delivered. Failure retains edits for retry. Legacy and preserved
paths share failure-safe delivery; symlink/nonregular targets refuse explicitly.
No semantic diff or directory-sync crash guarantee is implied. See batch 007.

### Increment 009: immutable namespace-aware XML snapshots

`internal/losslessxml` retains source offsets, expanded names and parent ownership.
Atomic plain-text leaf replacements preserve unrelated bytes and existing QName
namespace context; foreign/duplicate/mixed/self-closing insertion targets refuse.
Namespace errors, directives and invalid XML text refuse. This UTF-8-only internal
primitive does not yet apply Office-specific revision/protection/field rules.
Read `reports/batches/009.md`; B03 integration remains partial.

### Increment 010: read-only relationship ownership

`Preserved.Graph` verifies understood content-type and relationship registries,
resolves internal targets and reports inbound counts without rewriting XML or
fetching external links. Unsupported registry extensions refuse. No part deletion,
allocation, import or automatic orphan collection is implied. See batch 010.

### Increment 011: first bounded Word editor

`document.OpenEditing` and concrete `EditSession` provide exact complete-leaf body
run corrections through retained source, namespace-aware splices and OPC graph
checks. Protection, fields, revisions, wrappers and unsupported whitespace refuse.
Successful changes consume targets; no-op/refusal do not. Other prior targets
become stale conservatively. This implements one bounded adapter toward B10,
not full Word/extended search or review parity. See SPEC.md and batch 011.

### Increment 012: first bounded presentation editor

`presentation.OpenEditing` supplies exact slide-part/shape-ID complete-leaf
corrections for plain ungrouped shapes. Fields, locks and mixed content refuse;
source splicing retains formatting/geometry and unrelated parts. This is an
initial adapter, not P01/P03 or B10 completion. See SPEC.md and batch 012.

### Increment 013: first bounded numeric cell editor

The spreadsheet adapter changes existing numeric value leaves only after refusing
formula/dependent/protected structures in its conservative initial subset. It
preserves style and unrelated bytes and delivers through retained-source checks.
All three formats now have one bounded adapter; B10 still needs wider integration
and adversarial/corpus coverage. This is not X01/X02/X07 parity. See batch 013.

### Increment 014: extension refusal boundaries

Namespace-invalid replacement XML now refuses in the retained adapter. Numeric
editing rejects unclassified worksheet/workbook elements and attribute namespaces;
Word checks settings namespace policy before protection. These conservative guards
close three proved acceptance gaps without claiming schema-complete validation.
See batch 014; 61 native expanded cases pass.

### Increment 015: reviewed archive defects and corpus intake

Adjusted prefixed archives now refuse standalone OPC intake. Data-descriptor
signature/CRC ambiguity resolves against the whole central tuple. All 74 checked-in
Office fixtures pass bounded intake and byte-exact no-op delivery in one corpus
batch; this does not establish edit/graph/render parity. See batch 015.

### Increment 016: collected parameter identities

The pinned source suites collect 3,451 cases (DOCX 1,441; PPTX 1,169; XLSX 841).
Schema-2 spec/native-capability-history.json records exact node IDs and collection-log hashes;
collection is not execution or parity. API/refusal/fixture/Go/scenario mapping
remains pending. Raw collection uses Python only as development tooling. Batch 016.

### Increment 017: Word story projections

Read-only current/original/all outlines traverse related story parts and nested
paragraphs without double-counting text boxes. Basic revision wrappers project;
alternate-content branches/field evaluation/property revisions expose limitations.
StoryBlock is inert evidence, not an editing anchor. D01 remains partial pending
per-story fixtures and complete source-case mapping. See batch 017.

### Increment 019: exact run-fragmented Word spans

FindText/FindOne share exact current-view main-story paragraph run maps with Unicode
rune offsets and explicit repeated candidates. Span text is immutable evidence;
substring/cross-run writes still refuse until the mutation planner. Batch019.

### Increment 020: ordinary multi-run replacement

Unique maximal affix localisation preserves unchanged run fragments and puts changed
text into its starting run. Repeated-affix ambiguity, wrappers and interior run-boundary
insertions without owner/format proof refuse. Staged leaves commit once; tracked/bulk
and full result-report parity are not implied. Batch020.

### Increment 021: explicit atomic span batches

ReplaceBatch preflights selected targets, merges disjoint changes within shared
leaves and commits once. Any selected refusal aborts the batch; targets survive
rollback/no-op. This is not extended replace_all's per-match continuation API. Batch021.

### Increment 022: explicit normalised search

Unicode-15 full casefold/punctuation/space matching maps back to whole original
characters. Search Near ranks all candidates with stable ties; Nth is explicit and
mutually exclusive. Runtime remains dependency-free; data/generator/licence retained.
Still paragraph-local/current main story. Batch022.

### Increment 023: lossless attribute updates

Internal XML Edit stages text plus non-namespace attribute updates against original
offsets, preserving unrelated start-tag bytes and existing quote style. New namespaced
attributes require existing bindings; namespace mutation refuses. Batch023.

### Increment 024: Word whitespace attributes

Ordinary and selected-batch text changes atomically insert/update xml:space preserve
when required; opaque attribute bytes and no-op markup remain untouched. This
supersedes the initial conservative missing-preserve refusal. Batch024.

### Increment 025: replace-all refusal reporting

ReplaceAll records independent expected refusals, skips no-ops and commits supported
private matches together. Unexpected/stale/batch conflicts abort. Schema1 counts and
refusals are not full extended formatting/revision result parity. Batch025.

### Increment 026: scoped story/view search

Explicit Search scopes a related story and revision view, preserving literal paragraph
newlines and raw text identity. Historical/cross-paragraph/related-story targets are
inspection-only; main-body editors retain conservative ownership gates. Batch026.

### Increment 027: search boundary regressions

Invalid partial casefold candidates no longer hide later source-aligned matches.
Any synthetic paragraph-separator coverage makes a target inspection-only, including
a match that ends at the separator. Both defects have red/green batches. Batch027.

### Increment 028: proved run-boundary insertion

Boundary insertion succeeds only with equal direct properties/attributes and shared
paragraph ownership, then uses the left run deterministically. Different/interrupting
structures refuse atomically; serialization-equivalent formatting remains conservative.
Batch028.

### Increment 029: invisible content and namespace equality

Run-spanning mutation refuses non-text runs in its paragraph; boundary-format proof
also compares in-scope namespaces, preventing equal lexical rPr under different
bindings from authorising insertion. Two false successes now have red/green regressions.
Batch029; interval-only relaxed guards remain pending.

### Increment 030: qualified source-contract links

retired-native-links links seven pinned DOCX cases to implemented Go scenarios,
all explicitly partial with limits. Acceptance validates exact collected/source IDs
and implemented scenario membership; regeneration cannot erase manual links. No
source equivalence or completed parity is inferred. Batch030.

### Increment 032: shared style-zero regression

Numeric edits now validate cell/row/column style indices. Omitted s means zero,
which must resolve when cellXfs exists; empty tables and invalid indices refuse.
This independently reproduces/fixes the Bun-discovered V2 contract regression,
without claiming shared style-authoring or workflow parity. Batch032.

### Increment 033: shared V2 pin and namespace regression

V2 pack verification independently checks byte-pinned source/member hashes, sentinel
ownership, graph/style/no-op integrity and strict typed JSON. It found incorrect r:id
namespace matching in new Go XLSX/PPTX editors and their synthetic fixtures; both are
corrected to officeDocument relationships. Four fixtures verified, zero of19 shared
workflow cases executed. Pin/spec and report: spec/shared-contracts-v2.json, batch033.
Shared binary redistribution remains unapproved; no fixtures copied into the repo.

### Increment 034: shared fixture native readback

Native Go reads the pinned Word body, PPTX title/subtitle, default-style XLSX string
and cross-sheet formula/cache input facts. Input verification is distinct from the
19 planned mutation scenarios and emits no workflow pass outcomes. Batch034.

### Increment 036: XML value normalization with original offsets

Literal attribute whitespace now follows XML1.0 normalization before reference
decoding, while numeric whitespace references retain their characters. Source
bytes/offsets remain immutable; text/CDATA CRLF/CR no-op/edits are characterized.
Sibling Bun regression suggestions reproduced the attribute defect. Batch036.

### Increment 037: structured child insertion

Internal InsertChildren authors expanded-name nodes with existing or fresh scoped
prefixes, resets default namespaces for unqualified children and expands self-closing
parents without rewriting their attributes. Raw XML/namespace mutations and overlapping
targets refuse. This advances B03, not Office schema/graph authoring parity. Batch037.

### Increment 038: QName and root-boundary conformance

Explicit NCName components reject digit/combining-mark-first local names and invalid
prefix declarations; root-external whitespace is XML whitespace only. Fresh insertion
bindings/default resets are checked across sibling boundaries. Batch038.

### Increment040: static formula references

A native lexer/parser recognises bounded A1/range dependencies, quoted sheets and a
closed nonvolatile expression subset. Unknown syntax refuses without partial analysis;
strings never become references. No formula calculation or remapping yet. Batch040.

### Increment041: static dependency cache invalidation

An explicit numeric edit traces static A1/range dependencies across sheets and cycles,
clears affected cached values, and requests recalculation atomically with the input.
No calculation; unrelated caches/metadata and no-op bytes remain untouched. Conservative
unknown/volatile/shared/array/graph/protection refusals remain. Batch041; not X06/X07 completion.

### Increment042: native shared cache contract

The exact V2 cross-sheet fixture now executes through the invalidation API and passes
independent source/formula/cache/flags/preservation/graph/style/commit-count checks.
Only empty inactive protection/name containers were unblocked; active structures refuse.
One shared contract has direct-native evidence, not yet schema2/Gherkin binding parity.
Batch042;18 others unexecuted, external calculation not performed.

### Increment043: formula lexical-kind guard

Operators/delimiters must be punctuation tokens, never identical text inside quoted
literals. This closes false acceptance and false rejection cases in the static parser.
Batch043; full expression support remains conservative.

### Increment044: dependency worksheet structure

The invalidation planner validates all cell value/metadata structures and unique
sheetData/row ownership before analysis, including unrelated input cells. Malformed
structures refuse before commit. Batch044; shared cache input still executes.

### Increment045: parsed insertion reference remapping

Static formula coordinates can be remapped for row/column insertion on one sheet,
with absolute flags, quoted sheet spelling and non-reference tokens preserved.
Grid overflow or unproved syntax refuses without partial text. This is internal
reference machinery, not worksheet structural mutation. Batch045.

### Increment047: unresolved extended-comment MIME policy

Pinned Go and extended/Python sources disagree on commentsExtended content type.
`spec/source-discrepancies.json` records exact source pins/values and validation needs.
Retained-source body edits preserve either registry spelling and related bytes;
no emitted constant changed or schema/native Office certification inferred. Details:
`docs/design/comments-extended-content-type.md`, batch047. Broader revision/cache
consumer limitations remain explicit; sharedV2 contracts unchanged.

### Increment048: planned graph additions and retargets

Retained sessions now preflight explicit part additions and existing internal-edge
retargets with private plans, generation guards and complete candidate graph checks.
Registry patches preserve unrelated bytes; shared old payloads remain. Delivery
verifies new inventory/payloads/graph; schema2 receipts distinguish additions.
Batch048 passes160 native cases. Format ownership, deletion and import are unfinished.

### Increment049: loaded worksheet picture replacement

FindImage/ReplaceImage isolate a directly anchored PNG/JPEG occurrence by retargeting
one existing drawing relationship to fresh media. Original drawing XML/crop/geometry,
old shared image and other parts remain exact. Shared drawing/edge ownership, protection
and unsupported complex drawings refuse. Batch049 passes166 native cases; X10 is partial.

### Increment050: preserve target path form

Graph retargets now retain relative/absolute path form, including escaped Unicode,
space and quote characters. Pinned extended image tests motivated the regression.
Batch050 passes168 native cases plus chained-plan and exact-restore checks.

### Increment052: worksheet owner isolation

FindImage now requires one inbound worksheet edge, closing an alias outside the
workbook sheet list. Duplicate sheet identities already refused at intake.
Batch052 passes169 native cases; prior full/race checkpoint is051.

### Increment053: local ZIP64 and terminal validation

Local/central ZIP64 values and explicit64-bit descriptors now verify, including
empty deflate payloads. Terminal record gaps/extents and malformed extras refuse.
Batch053 passes183 native cases and the existing74-fixture no-op corpus. B01 remains
partial; no multi-gigabyte or exploratory fuzz run.

### Increment054: ZIP64 delivery boundaries

Edited forced-ZIP64 archives and subsequent graph additions deliver/reopen under
bounded tests. The native writer emits ZIP64 at exactly65535 entries; a lower intake
entry cap refuses. Extra fields cannot borrow a following TLV's bytes. Batch054.

### Increment056: independent ZIP overlap

A distinct-name, coherent-header, CRC-valid inner member embedded in an outer
payload independently exercises physical-overlap refusal. Existing runtime checks
pass; batch056 is test-only characterisation with184 native cases.

### Increment057: explicit leaf deletion and mixed transactions

GraphPlan now combines selected edge removal, detached-leaf deletion and guarded
payload replacement. Dangling edges, stale hashes and conflicting selections refuse.
Whole-element XML removal preserves all bytes outside selected non-root subtrees.
Batch057 passes191 native cases; format owners must authorise edge removal.

### Increment058: calculation-chain cleanup

Affected static numeric edits now remove a uniquely workbook-owned ordinary
calculation chain and registrations atomically with caches/calcPr/input changes.
Unrelated/no-op edits retain chain bytes; ambiguous/extended chains refuse.
Batch058 passes198 native cases. Legacy synthetic-chain authoring is unchanged.

### Increment059: lexical XML boundaries after maintenance

Resumed fromde006f4 after the post-reboot goal update. Prolog/epilog whitespace is
checked lexically; references and empty/whitespace CDATA outside the root refuse.
Batch059 passes202 native cases; internal content and literal boundary bytes remain intact.

### Increment060: graph and chain lifecycle coverage

Transient graph additions can be replaced/deleted back to the exact source. Failed
private plans preserve prior edits and valid held plans; deleted originals remain
reserved. Nonstandard calc-chain deletion, repeated edits and signed refusal pass.
Bounded judge findings were checked against public boundaries and rejected with tests.
Batch060 passes202 native cases; full/race checkpoint follows.

### Increment062: existing notes body editing

FindNotes/ReplaceNotes inspect and edit a uniquely owned existing single-run notes
body without creating missing parts. Other placeholders/registries remain exact;
fields/locks/ambiguous owners/multiline content refuse. Batch062 passes211 native
cases. P09 multiline and broad fixture parity are unfinished.

### Increment063: pinned notes fixture

The existing hash-pinned Go notes fixture now passes exact one-part replacement,
save/reopen and handle/refusal checks. A grouping-only noGrp lock no longer blocks
text editing; unknown/text locks still refuse. Batch063 passes213 native cases.

### Increment064: Office link identity

Twelve XLSX/PPTX intake cases verify expanded attribute names under alias/local
namespace bindings, plain-id exclusion and exact relationship Type URIs. Current
readers pass without runtime changes. Batch064 passes225 native cases.

### Increment066: structured subtree replacement

Immutable replacement of disjoint non-root XML subtrees uses the surviving parent's
namespace context and preserves all surrounding source bytes. Raw XML, overlapping
selections and invalid later nodes refuse atomically. Batch066 passes227 native cases.

### Increment067: multiline notes templates

Ordinary notes paragraphs/runs now read and replace as LF-separated text. First
paragraph/run formatting templates are copied within an explicit supported subset;
empty lines become empty paragraphs. Unproved properties/fields and bad text refuse
before package edits. Single-leaf nonempty corrections keep byte splices.234 cases pass.

### Increment068: empty leaf and whitespace authoring

Self-closing empty notes leaves use structural replacement. Leading/trailing spaces
on ordinary leaf updates set xml:space=preserve atomically. Template-value and
blank-line/no-op invariants pass;235 native cases in068. Review delegate timed out.

### Increment069: frozen extended notes inputs

Opt-in native checks read self-generated and LibreOffice-exported extended notes fixtures
from the pinned read-only checkout. Both pass exact no-op, multiline/clear/reopen and
one-part budgets. Empty effects/inherited underline fill are supported; nonempty
effects refuse.237 native cases pass. No fixture redistribution or live Office run.
