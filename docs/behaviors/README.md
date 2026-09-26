# Native behaviour reconciliation staging

These files are staging input for the shared `fixtures-ooxml` behaviour registry.
They are not an independent Go specification or an execution report. Canonical
scenario IDs and expected outcomes are assigned/reconciled centrally; the Go
consumer retains only adapters and per-case status/evidence mappings after import.

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
Reviewed family mappings cover 140 declarations in 135 candidate outcomes across
formula, XML, archive, mutable package-model, delivery, graph, test-custody,
utility, OOXML-model, WML-model, spreadsheet-cache, spreadsheet-targets and
spreadsheet-fuzz families; 264 declarations remain unreviewed. This denominator is the static
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
The [graph/ZIP64 handoff](graph-archive-handoff.md) compares selected assertions
with the exact v0.4 released IDs, retaining operation, fixture, error-code and
output-encoding gaps. It adds no reviewed declaration or execution credit.

Next review each remaining outcome group into Given/When/Then preconditions, native operation
and independently observable results. Include malformed/refusal/rollback and
parameter distinctions. Reconcile duplicate expectations with central existing
IDs; report contradictions explicitly. Keep unresolved mappings visible until
reviewed. Importing a feature or mapping never marks its steps executed.
