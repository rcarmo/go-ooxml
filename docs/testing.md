## Shared-reference testing

Native tests read a single `fixtures-ooxml` checkout and resolve content IDs
through its schema2 manifest. Documents and media live under
`fixtures/<format>/<scenarioGroup>/`, with one physical file per SHA-256.
Origins and required licences are metadata; consumers do not reconstruct origin
folders, copy fixtures or install compatibility symlinks.

`internal/testutil/fixture_ids.go` maps 38 retained native input labels to
content IDs: 37 logical Office inputs and one PNG. Several labels can resolve to the same
physical file. `FixturePath` accepts these trusted labels; `LookupFixture` accepts
`fixture-<full-sha256>`. Paths, size, content hashes, format/group metadata and
regular-file custody are checked before use. Shared workflow input records use
`assetId` and repository-root-relative paths.

## Candidate and released references

The gitlink and pin name annotated release `v0.111.0`, commit
`f767fa76873a3ae84366294f1b68e09262f7e182`. Initialise the recorded submodule and
run the default batch without overrides:

```sh
git submodule update --init --recursive
GOMAXPROCS=2 make test-batch
```

A future candidate run requires both explicit overrides:

```sh
OOXML_FIXTURES_ROOT=/path/to/clean/candidate-checkout \
OOXML_REFERENCE_PIN=/path/to/candidate-pin.json \
GOMAXPROCS=2 make test-batch
```

The released pin uses schema2 with `commit`, `tag`, `tag_object`,
`manifest_sha256`, `assets`, `facts`, `workflows` and `workflow_cases`.
There is no separate pack seal. The root manifest seals
`contracts/mutation-safety.json` and all five declared operation feature paths.
Go selects six underline-style, six font-name, three colour-getter and five
highlight in-memory rows in `workflows/docx/run-formatting.feature`, alongside
the eight-effects getter case. A separate four-step case tests opposite direct
superscript/subscript flags on two runs. One selected case saves and reopens three
authored runs before checking their named direct properties. [Batch 179](../reports/batches/179.md)
and [Batch 181](../reports/batches/181.md) record the earlier rows, negative
controls and gates. [Batch 184](../reports/batches/184.md) records the two-run
vertical-align binding. [Batch 182](../reports/batches/182.md) records the v0.62
pin adoption; [Batch 183](../reports/batches/183.md) records the v0.63 shared
Go evidence adoption for the nine earlier cases. [Batch 185](../reports/batches/185.md)
records the v0.64 shared Go evidence for the two-run case and Bun-only table/cell
getter evidence. [Batch 186](../reports/batches/186.md) binds just the four-step
`@id-docx-go-table-merge-properties` case: it reads GridSpan 3 at cell (0,0)
and VerticalMerge restart/continue at first-column rows one/two in memory.
[Batch 187](../reports/batches/187.md) records the v0.65 shared Go evidence
adoption for that getter-only case. [Batch 188](../reports/batches/188.md) selects
eight exact dimension rows and the cell-access, cell-text and row-count cases
from the same table feature. These assert in-memory getters, five tested
out-of-range cell coordinates, and an error on deletion at index ten. [Batch 189](../reports/batches/189.md)
records v0.66 shared Go and Bun evidence for all four cases. [Batch 190](../reports/batches/190.md)
adds one new-document empty-body case and one 3×3 table-text save/reopen case.
[Batch 191](../reports/batches/191.md) records v0.67 shared Go and Bun evidence
for both cases; Python remains planned. The in-memory table-value checks do not
establish physical merge topology or safe grid authoring. The selected table
readback checks nine text getters, not Office rendering or general OOXML
validity. [Batch 192](../reports/batches/192.md) binds the in-memory core
Title/Creator/Subject getters and first-section TitlePage/background getters.
[Batch 193](../reports/batches/193.md) records v0.68 shared Go and Bun evidence
for those exact cases; Python remains planned. Other supplied core fields are
setup, without saved or broader setter credit.
[Batch 194](../reports/batches/194.md) binds the five exact JSON rows of
`@id-docx-go-paragraph-text-getter`: empty, Hello World, two-space wrapped,
Japanese and XML punctuation. [Batch 195](../reports/batches/195.md) records
v0.69 shared Go and Bun evidence for those in-memory getters; Python is planned.
Other native text rows, sibling paragraph workflows and saved XML have no Go
execution credit from this binding. [Batch 196](../reports/batches/196.md)
binds four alignment rows (justify→both), four twip-spacing pairs, the three
paragraph flags and three appended runs in memory. [Batch 197](../reports/batches/197.md)
records v0.70 shared Go and Bun evidence for those ten cases; Python is planned.
These getter checks do not test saved XML or Office layout.
[Batch 198](../reports/batches/198.md) binds the four-step in-memory
`@id-docx-go-body-insert-order` case, checking initial empty body, three
operation counts and `First,Second,Third` paragraph order. [Batch 199](../reports/batches/199.md)
records v0.71 shared Go and Bun evidence for that case; Python is planned.
Saved XML and Office layout are outside this binding.
[Batch 200](../reports/batches/200.md) records v0.72 shared Bun-only evidence
for seven DOCX text-slice cases across five IDs: saved cross-run formatting and
whitespace, table-cell paragraph text, stale-span byte custody, and three
unsupported-topology refusals. Go and Python remain planned for those cases.
The Go selector and its 349 in-memory/native cases are unchanged.
[Batch 201](../reports/batches/201.md) records v0.73 shared Bun-only evidence
for eight DOCX creation cases across four IDs: saved minimal package/style,
stale-span and opaque-part custody, and five atomic refusals. Go and Python
remain planned for those cases; this pin adds no Go creation execution credit.
[Batch 202](../reports/batches/202.md) records v0.74 shared Bun-only evidence
for seven DOCX table-authoring cases: saved table/cell readback and opaque-part
custody, stale cell, and four atomic refusals. Go and Python remain planned;
this pin adds no Go table-authoring execution credit.
[Batch 203](../reports/batches/203.md) records v0.75 shared Bun-only evidence
for the six-step saved/reopened DOCX direct-font-size case: setting 10.5pt
produces one direct `w:sz` with value 21. Go and Python remain planned for
this half-point case. [Batch 204](../reports/batches/204.md) binds the exact
saved/reopened Go six-step case. [Batch 205](../reports/batches/205.md) records
v0.76 shared Go evidence for that case and Bun-only XML parser/refusal evidence
for four other scenarios. Go and Python XML execution remain planned. The Go
selector runs 350 cases; the half-point case does not test style inheritance
or rendering. [Batch 206](../reports/batches/206.md) records v0.77 Bun-only
evidence for the four-step XML edit safety case: disjoint fragment and escaped
text success, plus overlap, malformed and DTD refusals. Go and Python remain
planned for that case; this pin adds no Go XML editing credit.
[Batch 207](../reports/batches/207.md) records v0.78 Bun-only evidence
for eleven XML value, namespace and escaping cases across ten IDs. Go and
Python remain planned for these cases; the Go selector remains at 350 cases.
[Batch 208](../reports/batches/208.md) records v0.79 Bun-only evidence
for six exact XML QName and NBSP cases across three IDs. Go and Python remain
planned for these cases; this pin adds no Go XML QName execution credit.
[Batch 209](../reports/batches/209.md) records v0.80 Bun-only evidence
for ten exact conservative XML comparison cases across five IDs: Boolean
results and unchanged UTF-8 inputs. Go and Python remain planned for these
cases; this pin adds no Go XML comparison execution credit.
[Batch 210](../reports/batches/210.md) records v0.81 Bun-only evidence
for 13 exact package-admission refusal cases across four ZIP/XML IDs. Go and
Python remain planned; the physical-overlap case is separate. This pin adds
no Go package-admission execution credit.
[Batch 211](../reports/batches/211.md) records v0.82 Bun-only evidence
for one exact seven-step semantic package-diff case: equivalent `a.xml`,
changed `b.bin`, added `c.bin`, no removed part and unchanged caller bytes.
Go and Python remain planned; this pin adds no Go semantic-diff credit.
[Batch 212](../reports/batches/212.md) records v0.83 Bun-only evidence
for four ZIP32 scenarios: CRC vector (1 case), reader refusal (12), writer
refusal (2) and configured bounds (5), totalling 20 cases / 91 steps.
Go and Python remain planned; this pin adds no Go ZIP32 execution credit.
[Batch 213](../reports/batches/213.md) records v0.84 Bun-only evidence
for positive ZIP32 read (seven steps) and deterministic write (six steps).
Broad unsafe/bounds cases and Go/Python execution remain planned; this pin
adds no Go ZIP32 credit.
[Batch 214](../reports/batches/214.md) records v0.85 Bun-only evidence
for one broad ZIP32 unsafe-structure refusal case (nine steps, eleven samples)
with typed-code and fixture-expectation checks. Broad bounds/pre-expansion
and Go/Python execution remain planned; this pin adds no Go ZIP32 credit.
[Batch 215](../reports/batches/215.md) records v0.86 Bun-only evidence
for one three-step ZIP32 configured-bounds case: five typed thresholds precede
invalid DEFLATE on the same declared expanded payload, with unchanged caller
bytes. It does not measure allocator use. Go and Python remain planned.
[Batch 216](../reports/batches/216.md) records v0.87 Bun-only evidence
for one eight-step unsigned ZIP32 descriptor/signature-collision refusal:
12-byte geometry, independent CRC, typed corrupt-payload refusal, no output
and unchanged input. Go and Python remain planned; the source candidate stays
distinct and this pin adds no Go descriptor credit.
[Batch 217](../reports/batches/217.md) records v0.88 Bun-only evidence
for two six-step negative direct-admission budget cases: source bytes and entry
count set to -1. Invalid-argument refusal precedes broken ZIP parsing; no parts
are delivered and caller bytes remain unchanged. Go and Python remain planned;
the broader Go source candidate is not retired.
[Batch 218](../reports/batches/218.md) records v0.89 Bun-only evidence
for three common OPC cases: 36 Go-origin plus 35 Python-origin no-op archives
retain whole-archive bytes after reopen; a synthetic multi-part transaction
rolls back; and an XML edit/reopen preserves unrelated binary bytes. Go and
Python remain planned; no cross-producer equivalence follows from this pin.
[Batch 219](../reports/batches/219.md) records v0.90 Bun-only evidence
for seven OPC custody/save IDs (11 cases / 108 compiled steps): typed open
refusals, detached bytes, UTF-16LE edit, synchronous transaction/thenable,
existing-file and symlink save custody. Go and Python remain planned; Office
schema and rendering were not tested by this pin.
[Batch 220](../reports/batches/220.md) records v0.91 Bun-only evidence
for two PPTX/XLSX relationship expanded-name IDs (four cases / 16 steps):
`link:id` survives edit/reopen; the wrong URI produces distinct typed
structural refusals without changing the input. Go and Python remain planned;
no live Office or rendering check was run by this pin.
[Batch 221](../reports/batches/221.md) records v0.92 Bun-only evidence
for four existing DOCX comment IDs (14 cases / 46 steps): pinned threaded
inspection, resolve/reopen and restore, no-op and eleven typed refusals with
unchanged package bytes. Missing-metadata creation and arbitrary thread edits
are outside this evidence; Go and Python remain planned.
[Batch 222](../reports/batches/222.md) records v0.93 Bun-only evidence
for eight existing DOCX comment-thread IDs (23 cases / 108 steps): pinned,
nested and reordered inspection; resolve/reopen, no-op, nine refusals,
rollback, encoding, unsupported cases and the 10001 limit. Missing-extension
creation and arbitrary new threads are outside this evidence; Go and Python
remain planned. The Go selector and execution credit do not change.
[Batch 223](../reports/batches/223.md) records v0.94 Bun-only evidence
for one exact six-step PPTX title-slide no-edit path/byte open-save case:
Frankenstein text, byte-identical original after serialization/save/reopen,
source custody and temporary-file cleanup. No rendering was tested; Go and
Python remain planned. The Go selector and execution credit do not change.
[Batch 224](../reports/batches/224.md) records v0.95 Bun-only evidence
for four PPTX table IDs (five cases / 15 steps): saved 2×3 geometry/text,
styled-cell preservation, stale-handle refusal and atomic refusals for merged
or malformed topology. Merged-table editing and Office rendering were not
tested; Go and Python remain planned. The Go selector and execution credit
do not change.
[Batch 225](../reports/batches/225.md) records v0.96 Bun-only evidence
for two PPTX positioned text-box IDs (23 cases / 77 steps): eight saved
variants and 15 atomic refusals preserving package/version. Layout inheritance
and Office rendering were not tested; Go and Python remain planned. The Go
selector and execution credit do not change.
[Batch 226](../reports/batches/226.md) records v0.97 Bun-only evidence
for two PPTX slide-permutation IDs (21 cases / 70 steps): seven reordered
or no-op variants with handle and unrelated-part custody, plus 14 atomic
refusals. Unsupported composition and Office rendering were not tested; Go
and Python remain planned. The Go selector and execution credit do not change.
[Batch 227](../reports/batches/227.md) records v0.98 Bun-only evidence
for seven XLSX creation/missing-cell IDs (seven cases / 29 steps): new
workbook values and parts, independent worksheet IDs, prefixed B2, the
maximum coordinate, row order, atomic invalid-parameter refusal and formatted
fixture append/custody. Calculation and Office rendering were not tested; Go
and Python remain planned. The Go selector and execution credit do not change.
[Batch 228](../reports/batches/228.md) records v0.99 Bun-only evidence
for two XLSX direct cell-style IDs (27 cases / 90 steps): nine saved
selections preserving values, formulas, caches and unrelated bytes, plus 18
atomic refusals. Style dependency closure, recalculation and Office rendering
were not tested; Go and Python remain planned. The Go selector and execution
credit do not change.
[Batch 229](../reports/batches/229.md) records v0.100 Bun-only evidence
for one XLSX existing-comment/VML read-only graph case (seven steps):
distinct internal relationships and exact A2/A3 text. Five preservation-row
and limited-number editor cases remain planned; VML editing, save/reopen and
Office rendering were not tested. Go and Python remain planned. The Go selector
and execution credit do not change.
[Batch 230](../reports/batches/230.md) records v0.101 Bun-only evidence
for nine static XLSX formula-reference API IDs (45 cases / 152 steps):
counts, byte spans, flags, direct-range axis and first coordinate, remaps,
refusals and a finite 288-expression matrix. Formula evaluation, workbook
edit/save and Office rendering were not tested; Go and Python remain planned.
The Go selector and execution credit do not change.
[Batch 232](../reports/batches/232.md) records v0.102 ledger credit for
Go's two exact static XLSX direct-range IDs (13 cases / 45 steps, executed
in [Batch 231](../reports/batches/231.md)) and Bun-only seven DOCX tracking
settings IDs (24 cases / 101 steps). Go's other seven formula-reference IDs
and all DOCX tracking IDs remain planned. No formula evaluation, workbook edit,
Office rendering or cross-consumer execution is credited by this pin. The Go
selector remains 363 cases / 1244 steps.
[Batch 243](../reports/batches/243.md) records v0.111 correction
of Go's already selected direct-admission negative-budget (two cases / 12 steps)
and unsigned ZIP32 descriptor-signature collision (one case / eight steps)
execution credit. It adds no binding or selected case; Go remains at
395 cases / 1351 steps. The credit does not cover positive budgets,
broader descriptors or ZIP64.
[Batch 242](../reports/batches/242.md) records v0.110 Python-only credit
for four exact XML Boolean negative IDs (nine cases / 36 steps). Go adds
no XML binding, selected case or execution credit; its batch remains
395 cases / 1351 steps. The positive OPC-like Relationships ID is planned
because the input omits `Type`.
[Batch 241](../reports/batches/241.md) records v0.109 correction
of Go's latent exact seven-step `@id-zip-physical-member-overlap-refusal`
execution, first selected at v0.41 and confirmed in the v0.108 published
recursive-clone report. It adds no binding or selected case; Go remains at
395 cases / 1351 steps. The ZIP32 STORED overlap refusal is a bounded
admission policy, not a ZIP64 or ECMA conformance assertion. ECMA clauses
were not verified and no PDFs were extracted, converted or read.
[Batch 240](../reports/batches/240.md) records v0.108 shared-ledger
credit for Python alone on one bounded ZIP32 STORED physical-member-overlap
refusal (seven steps). Go remains planned for this exact ID. The pin adds no
Go execution credit or broad overlap, ZIP64, ECMA or Office evidence; Go's
selected batch stays at 395 cases / 1351 steps.
[Batch 239](../reports/batches/239.md) records v0.107 shared-ledger
credit for the four Go static formula-reference IDs (18 cases / 57 steps)
already executed in [Batch 238](../reports/batches/238.md) and published
at `c846cb3`. Together with five earlier formula IDs, Go selects nine IDs /
45 cases / 152 steps in its 395-case / 1351-step batch. The new credit is
for bounded formula-string analysis and remapping; no formula calculation,
workbook editing, Office rendering or ECMA parser completeness is tested.
[Batch 237](../reports/batches/237.md) records v0.106 Bun-only
paragraph-style authoring (two DOCX IDs, 24 cases / 86 steps) and Python-only
read-only existing XLSX comment/VML graph inspection (one ID, one case /
seven steps). Go gains no execution credit. Five XLSX comment editor cases
are planned; no VML editing, Office rendering or broad style synthesis was
tested by this pin. Go retains five formula IDs / 27 cases / 95 steps within
the unchanged 377-case / 1294-step batch.
[Batch 236](../reports/batches/236.md) records v0.105 shared-ledger
credit for the three Go static formula-analysis IDs (14 cases / 50 steps)
already executed in Batch 233 and published at `9ff28dc`. Together with
two direct-range IDs (13 cases / 45 steps), this is five formula IDs /
27 cases / 95 steps. Four sibling IDs / 18 cases remain planned; no formula
evaluation, workbook editing, Office or Python execution is credited.
[Batch 235](../reports/batches/235.md) records v0.104 Bun-only evidence
for two final-section DOCX page-layout IDs (22 cases / 74 steps): eight
four-step bounded layouts and 14 three-step typed refusals with unchanged
serialized input. Go and Python remain planned; no rendering is credited.
Go's formula selection stays at five IDs / 27 cases / 95 steps within the
377-case / 1294-step batch, and its three static analysis IDs still await
separate shared-ledger credit.
[Batch 234](../reports/batches/234.md) records v0.103 Bun-only evidence
for two DOCX effective-run formatting IDs (25 cases / 84 steps): nine bounded
plain-body resolutions and 16 typed refusals. Go and Python remain planned;
no broader provenance or rendering is credited. Go's three newly executed
static formula-analysis IDs (14 cases / 50 steps in Batch 233) await separate
shared-ledger credit; five formula IDs / 27 cases remain selected in Go's
377-case / 1294-step batch.
The v0.50 contract uses an explicit feature list (contract schema 2), selecting
eight mutation IDs and 19 cases. The legacy single-feature schema 1 reader remains
for older distributions; neither contract grants Go workflow execution credit.
Candidate tags begin with `candidate-` and have an empty `tag_object`; release
pins require the exact annotated tag object and its peeled commit. Use the
coordinator-provided pin, not hashes recomputed to accept modified inputs.
Release `v0.50.0` contains 342 manifest assets, including 79 physical fixtures,
plus 149 facts and 304 workflows/791 expanded cases.
The canonical registry has 60 feature files. Another 143 files are staged
consumer candidates, including 55 native Go features and 21 Go behaviour
catalogue candidates. The Go files retain their source IDs and wording; moving
them gives no new canonical execution credit. The shared ZIP-overlap outcome
`@id-zip-physical-member-overlap-refusal` replaces native `@ZIP-003` one-for-one.
Its selected seven steps require independent readable-member and physical-range
checks before a typed overlap refusal. The separate
`@id-package-admission-negative-budget` outline selects two six-step cases on
the sealed default DOCX. The direct `OpenReaderWithLimits` API rejects negative
source-byte and entry-count budgets before any `ReaderAt` access, returns nil
and a plain invalid-argument error distinct from typed package/resource refusals,
and leaves caller bytes unchanged. Zero-budget policy is API-specific. A third selected package outcome,
`@id-zip-unsigned-descriptor-signature-collision`, checks an unsigned 12-byte
DEFLATED data descriptor whose CRC equals the optional signature value. Raw
independent inflation finds `payload` with a different CRC; structural geometry
fits exactly before the central directory, and full admission refuses the
payload checksum with nil package and unchanged caller bytes. Positive
resource-budget cases and staged `@candidate-go-archive-001` and `-004` retain
their source predicates; other canonical admission outcomes remain planned for Go.
The format-grouped workflows retain the previous 229 IDs and 562 cases. In four operation
features, 16 IDs and 18 cases keep their inputs and observable outcomes while
reviewed actor wording and profile tags change. Two package workflows preserve
another 18 IDs/38 cases while separating package and ZIP32 actors from exact
error-code and JavaScript transaction API profiles. Nine Word workflows preserve
29 Go IDs and five anchor IDs (75 cases) while naming document-value and
anchor-response actors and profiles by operation; getter conventions, nullable
cells, heading classification, same-run effects, selected readback and tool
hints retain their existing compatibility boundaries. XML editing and static
formula-reference workflows preserve another 19 IDs/60 cases with neutral actor
and profile names; lexical custody, byte spans, grammar and refusal results
retain their original limits. Comment and template workflows preserve 22 IDs/22
cases while distinguishing existing-comment resolution from authoring and
thread policies and leaving weak response/cache outcomes explicit. XML parsing
retains two IDs/three cases under operation profiles; the XLSX creation change
is description-only. A distinct tracking-settings workflow adds seven IDs and
24 cases for saved preference, custody, refusal and rollback; all remain planned
for Go. A distinct physical horizontal table-merging workflow adds seven IDs
and 26 cases for success, refusal, rollback, encoding and stale handles. Seven
vertical-merge IDs and 26 cases extend that workflow while preserving the
horizontal cases. Both merge groups remain planned for Go. A concrete Word
template-inventory rule adds eight IDs and 22 cases to the existing template
analysis workflow while retaining its response/cache outcomes. These new IDs
also remain planned for Go. An existing-thread rule adds eight IDs and 23
cases to the Word comments workflow; single-comment and authoring profiles
retain their prior scope. The new thread cases remain planned for Go. An
opt-in run-property revision rule adds eight IDs and 33 cases to the Word
revisions workflow. The text-only profile and broader multi-story operations
keep their earlier scope. The new cases remain planned for Go with zero
execution credit. An opt-in paired text-move revision rule adds nine IDs and
45 cases to the same workflow, limited to paired source/destination ranges in
one story; these also remain planned with zero Go execution credit. The first three exact shared outcomes add four canonical execution cases. At v0.45,
`@id-xlsx-owned-calculation-chain-invalidation` replaces native `@CHAIN-001` one-for-one.
The native source stays in the inventory as superseded; `@CHAIN-002..004` stay
selected. The saved synthetic workbook readback checks both dependent caches,
formula text, the unchanged unrelated cache, recalculation flags, removed
nonstandard chain part/edge/override, graph resolution and unrelated payloads.
This does not select broader cache-completeness or external calculation outcomes.
The in-memory Go run-effects case `@id-docx-go-run-effects-getters` also executes
its three exact steps. It sets DoubleStrike, Caps, SmallCaps, Outline, Shadow,
Emboss, Imprint and Vanish on one new run and checks all eight direct getters.
Eight per-getter negative controls fail the assertion when one flag is cleared.
Simultaneous conflicting effects are a Go API observation; saved WordprocessingML
validity and rendering are untested. Other shared run-formatting workflows remain
planned for Go. Acceptance has 295 selected cases, 20 planned native cases and
one external native case. Shared
ECMA specifications, extracts and derived notes are indexed under
[`specs/ecma-376/`](../references/fixtures-ooxml/specs/ecma-376/README.md).
The full PDFs are the specification sources; five earlier ECMA documents and
PDFs from the original main history are recorded with their old paths and hashes
in [`spec/legacy-fixture-migration.json`](../spec/legacy-fixture-migration.json),
not duplicated under `docs/`. Follow the pinned specification index for editions
and provenance. Go's package ledger has
four partial mappings and eight unmapped declarations; its XML ledger has three
partial mappings and one unmapped.
Neither ledger grants canonical execution credit. The semantic-diff assertion
`removed: []` exercises no removal case.

The native runner and inventory read `staging/go/features/` in the pinned shared
checkout. The local behaviour mapping JSON and native source inventory remain in
`docs/behaviors/`; their 21 candidate features are read from
`staging/go/behaviors/`. Candidates do not enter the native execution selector.
For a sealed candidate run, `OOXML_FIXTURES_ROOT` changes all these lookup roots
alongside fixture lookup. The local copies of the 76 feature files have been
removed. The 38 original-main binary `testdata/` files are recorded in
[`spec/legacy-fixture-migration.json`](../spec/legacy-fixture-migration.json).
Its rows preserve the original full-archive identities. Release v0.41 removes two
logically duplicate DOCX archives; the historical `word/minimal.docx` and observed
`generated/word/sdt_content_controls.docx` IDs are recorded, with their original
hashes and provenance, in the sealed
[`fixture-content-consolidation.json`](../references/fixtures-ooxml/ledgers/fixture-content-consolidation.json).
Active test labels select the retained `default-d9d6…` and `SDT-368fe…` files.
The retired bytes remain recoverable at immutable shared v0.40.0; the old and
retained full-archive hashes differ. Release v0.44 separately retires 35
observed-generated Office archives and records their 36 historical input paths,
original hashes and provenance in the sealed
[`observed-generated-retirement.json`](../references/fixtures-ooxml/ledgers/observed-generated-retirement.json).
Those bytes remain recoverable at immutable shared v0.43.0. All 37
`generated/*` input labels are removed, including the SDT label that already
selected a retained archive. No generated input is remapped to a different
committed Office fixture. The no-op corpus now tests 37 Office labels; historical
Go output under `artifacts/generated/` remains unrelated. No runtime fixture
alias or fallback exists. `testdata/FIXTURES.md` remains
as a historical index; fixture lookup has no local fallback.

Update the gitlink and `spec/reference-distribution.json` together when adopting
a release. Use the recorded submodule commit; `git submodule update --remote`
would follow a branch tip instead of the pin.

## Integrity and output custody

Root-module and acceptance checks require the exact reference HEAD, the pinned
root seal, clean index/worktree and tracked bytes/modes. Release tags must be annotated.
Changes to facts or workflows fail even when fixture manifest hashes still match;
`assume-unchanged` does not hide altered tracked bytes. Verification is read-only.
An absent Git checkout, mismatched pin or missing input is a failure, not a skip.

Write outputs under `t.TempDir()` or consumer-local `artifacts/generated`.
Persistent output guards reject reference-root descendants, including symlink
redirects. Never regenerate or rebaseline the shared inputs during a test run.

## Mutation workflow contract

The root-only release removes the old wrapper and generated expanded-case input.
Go compiles the official Gherkin and derives stable scenario/Examples keys and typed
step arguments locally. The compact contract points to canonical asset IDs and
readback facts. Exact part membership is required; preserved hashes are the
complement of `allowedChangedPartsForSuccess`. Native semantic assertions and the
direct cache-invalidation outcome are unchanged. Contract inventory validation
grants no workflow execution credit.

Historical v0.2 compatibility is selected only by a schema1 pin, never by checking
which files happen to exist. Current schema2 verification reads no old wrapper or
generated inventory. Migration results and temporary field-mapping failures are
recorded separately in `../reports/batches/095.md`.

## Batched verification

`make test` covers the root Go module only. `make test-batch` also runs the separate
acceptance module. Use `GOMAXPROCS=2` and `-p 2`; run related packages together,
including failure reruns. Individual-test retry loops are not part of the workflow.
Full/race checks belong at integration points; reuse caches and avoid concurrent
duplicate suites. The runtime has no external dependencies; Godog/Gherkin are
isolated in the acceptance module, and checkout verification uses Git only in tests.

The v0.59 Go run-effects binding is recorded in
[`reports/batches/177.md`](../reports/batches/177.md). It adds one exact
three-step in-memory case; it does not certify saved OOXML or other consumers.
The v0.50 PPTX visibility inventory adoption is recorded in
[`reports/batches/165.md`](../reports/batches/165.md). Two new cases remain
planned for Go, with no Office-confirmed positive or new execution credit. The
v0.49 Python owned-chain evidence adoption is recorded in
[`reports/batches/164.md`](../reports/batches/164.md). It changes no Go binding,
selector or execution credit. The v0.48 Go cross-sheet cache evidence adoption is recorded in
[`reports/batches/163.md`](../reports/batches/163.md). It records the already
published 14-step canonical case replacing native `CACHE-001`, with no added
selected case or execution credit. The v0.47 mapping-only adoption is recorded in
[`reports/batches/162.md`](../reports/batches/162.md). It adds a Python XLSX
dependency contract and five-declaration partial source mapping, with no Go
binding or selection change. The v0.46 evidence-only adoption is recorded in
[`reports/batches/161.md`](../reports/batches/161.md). The Go runner keeps
294 selected cases and 1004 passed steps; v0.46 marks its already published
owned-chain binding implemented in the shared ledger without selecting another
case. The previous v0.45 binding is recorded in
[`reports/batches/160.md`](../reports/batches/160.md): 294 selected cases and
1004 passed steps; 20 native planned and one external case were not run. The
v0.44 default batch is recorded in
[`reports/batches/159.md`](../reports/batches/159.md): 294 selected cases and
996 passed steps; 20 native planned and one external case were not run. The previous v0.43
batch is recorded in [`reports/batches/158.md`](../reports/batches/158.md):
294 selected cases and 996 passed steps, with 20 native planned and one native
external case unrun. The v0.42 batch is recorded in
[`reports/batches/157.md`](../reports/batches/157.md): 293 selected cases and
988 passed steps. The v0.41 batch is recorded in
[`reports/batches/156.md`](../reports/batches/156.md): 291 selected cases and
976 passed steps. The v0.35 batch is recorded in
[`reports/batches/155.md`](../reports/batches/155.md): 291 implemented native
Gherkin cases and 972 steps passed; 20 planned and one external case were not
run. The moved Go source cases keep their IDs, lines and test outcomes. Package/unit
and subtest counts are separate metrics. This batch did not run race, fuzz, live
Office rendering or calculation checks. Older batch reports retain the commands
and results from their own revisions.

## Behaviour catalogue

The central registry owns canonical behaviour IDs and expected outcomes. The
`docs/behaviors` mapping JSON and inventory stage native findings for
reconciliation; their feature sources now live in the shared checkout. They do
not form a second canonical suite. The current inventory has 404 native declarations,
including the checkout guards: 197 have reviewed family/parameter mappings in
207 candidates, and 207 remain unreviewed. This is the frozen inventory snapshot;
later test declarations need a deliberate refresh before entering its denominator.
Importing Gherkin earns no execution credit, and sibling results do not become
Go passes. The historic `reports/batches/` logs remain in the repository; each
report records its own revision and measured scope.
