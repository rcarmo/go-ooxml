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

At the reviewed catalogue checkpoint, 59 of 402 declarations map to 54 candidate
outcomes across formula, XML, archive, mutable package-model, delivery and graph
families; 343 declarations remain unreviewed. The denominator is the committed
inventory snapshot, not the number of executed leaf cases. New checkout-integrity
guard tests added afterwards are additional unmapped work until regeneration.
The executable native Gherkin result remains separate at 291 cases. See
[testing.md](../testing.md) for candidate reference setup and current limits.

Next review each remaining outcome group into Given/When/Then preconditions, native operation
and independently observable results. Include malformed/refusal/rollback and
parameter distinctions. Reconcile duplicate expectations with central existing
IDs; report contradictions explicitly. Keep unresolved mappings visible until
reviewed. Importing a feature or mapping never marks its steps executed.
