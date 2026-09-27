# Native behaviour reconciliation staging

The mapping JSON and native source inventory in this directory support review
against the shared `fixtures-ooxml` behaviour registry. The 21 candidate feature
files now live under `references/fixtures-ooxml/staging/go/behaviors/`, where
their original IDs and wording are preserved. They are not an independent Go
specification or an execution report. Canonical scenario IDs and expected
outcomes are assigned or reconciled centrally; moving these candidates adds no
Go execution credit.

`native-inventory.json` inventories tracked native Go test files by SHA256,
function identity, literal/dynamic subtest call sites, literal parameter groups
and called operation names. Every declaration starts `unreviewed` with an empty
scenario mapping. No name-derived scenario automatically earns coverage.

Regenerate using `go run ./tools/nativeinventory`. The observation revision names
the checked-out source commit; test files are independently hashed. Static literal
rows include setup data, and dynamic runtime-generated leaf cases cannot be
counted from source alone. Fuzz seeds and benchmark declarations are inventories,
not exploratory-fuzz or performance results. Helpers and acceptance binding-only
files remain visible, even when they declare no executable test entry point.

The current inventory contains 155 native test files and 404 declarations.
Reviewed family mappings cover 197 declarations in 207 candidate outcomes across
formula, XML, archive, mutable package-model, delivery, graph, test-custody,
utility, OOXML-model, WML-model, spreadsheet-cache, spreadsheet-targets and
spreadsheet-fuzz, presentation-notes, presentation-fuzz, presentation-fixtures and
presentation-parameters, presentation-api, presentation-measurement and
spreadsheet-parameters and spreadsheet-measurement families; 207 declarations
remain unreviewed. This denominator is the static
inventory snapshot, not the number of executed leaf cases. Checkout guard tests
are now included. The executable native Gherkin result is separate at 291 cases.
See [testing.md](../testing.md) for released/candidate reference setup and limits.
The OOXML-model family reviews five tests and nine fuzz targets in seven non-WML
packages. Field equality, marker-only output predicates and decode-error early
returns remain distinct; the 18 explicit fuzz seed tuples do not constitute an
exploratory campaign. New declarations added after the frozen inventory snapshot
remain outside that denominator until a deliberate inventory refresh.
The WML-model family adds 20 unit declarations and four fuzz targets with 12 seed
strings. Its two dynamic `tt.name` sites expand four OnOff and five NewT literal
rows. Single-character output predicates remain weaker than element assertions;
revision-author, container-count and helper results retain their stated limits.
Historical 91/404 and 105/404 checkpoints remain in batches094/104; neither the
model review nor later control-test changes resets the frozen denominator.
The final packaging declaration in this snapshot has two mutable-package fuzz
seeds. Its save/reopen assertions do not compare payloads or MIME values, and a
missing lookup returns silently. Reviewing all packaging declarations does not
complete package-behaviour coverage.
The spreadsheet-cache family captures three unit declarations and ten outcomes.
Selected byte checks, effect counts, typed stale-target errors and helper-owned
intake refusal assertions stay separate. Ignored serialization errors and unnamed
array/nonfinite rows are explicit gaps; no calculation or allocation-bound claim.
The spreadsheet-targets family adds two declarations: eight named style rows and
three image lifecycle subtests. Graph part/MIME and receipt/handle predicates
remain separate from payload equality, geometry, reopened archives and rendering.
The spreadsheet-fuzz family reviews five targets: 11 literal nonfixture seed tuples
and 11 fixture-label seed candidates subject to successful file reads. Nil returns,
ignored write errors and fixture-path early returns remain visible; the zero-XOR
fixture seeds do not mutate their input bytes. No new runtime leaf count is measured.
The presentation-notes family reviews two declarations and seven named subtests.
It separates exact fixture/payload preservation and disk readback from private
paragraph counts, four template-fragment counts and untyped refusal assertions.
Three presentation fuzz targets add four literal tuples plus eleven conditional
fixture-label candidates. Slide-count reopen and in-memory table dimensions do
not verify text-box custody; fixture paths return on operation errors without
asserting equality. No exploratory fuzz or new safe-text-box execution is credited.
The complex presentation fixture declaration dispatches eleven named rows. Its
exact minimal-deck geometry/font/notes checks remain distinct from any-shape text
substrings, conditional selector checks and nonempty collections that source
content can already satisfy. One declaration adds eleven outcomes, not eleven
declarations or whole-deck preservation credit.
Eight presentation parameterised declarations review 61 finite table rows, including
helper-owned strings/formats. Duplicate generated labels, skipped no-shape checks,
empty-substring success and unasserted underline getters are explicit. Helper
source hashes and reviewed row counts supplement the frozen inventory; they add
no exhaustive matrix or measured runtime-leaf count.
The presentation API family adds 21 declarations, separating constructor values,
in-memory reverse-order IDs, comment disk readback and notes-master relationship
checks from count-only round trips, entry presence and a no-panic fill smoke test.
These checks add no canonical safe slide-order, layout or text-box binding credit.
The last five presentation declarations in the snapshot are four benchmarks and
an opt-in memory-profile smoke test. Their conditional error checks are reviewed;
no benchmark or enabled memory-profile run is credited. Ordinary package checks
select no benchmarks and skip the smoke body with `ENABLE_MEMPROFILE` unset.
All snapshot presentation declarations now have scoped staging mappings, without
establishing complete feature coverage or a refreshed current-source inventory.
Twelve spreadsheet parameterised declarations add twelve outcomes with 80 selected
table-row uses and two cell-reference error rows excluded before subtest dispatch.
Helper-owned numeric, string, cell-reference and range tables are hashed alongside
the reviewed source. Fixture round trips compare sheet count only; string round
trips compare exact text while ignoring Open errors. Formula getters do not evaluate
expressions, and date checks compare only Year, Month and Day. These source counts
add no measured runtime-leaf, canonical workflow or calculation coverage.
Four spreadsheet benchmarks and one opt-in memory smoke add five reviewed
conditional error-check outcomes. Save setup ignores cell/table mutation errors,
uses the same workbook across iterations and performs no reopened-content check.
The package batch selects no benchmarks and leaves the enabled memory body unrun;
no performance or memory measurement is credited.
The [graph/ZIP64 handoff](graph-archive-handoff.md) compares selected assertions
with the exact v0.4 released IDs, retaining operation, fixture, error-code and
output-encoding gaps. It adds no reviewed declaration or execution credit.

Next review each remaining outcome group into Given/When/Then preconditions, native operation
and independently observable results. Include malformed/refusal/rollback and
parameter distinctions. Reconcile duplicate expectations with central existing
IDs; report contradictions explicitly. Keep unresolved mappings visible until
reviewed. Importing a feature or mapping never marks its steps executed.
