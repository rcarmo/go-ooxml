## Shared-reference testing

Native tests read a single `fixtures-ooxml` checkout and resolve content IDs
through its schema2 manifest. Documents and media live under
`fixtures/<format>/<scenarioGroup>/`, with one physical file per SHA-256.
Origins and required licences are metadata; consumers do not reconstruct origin
folders, copy fixtures or install compatibility symlinks.

`internal/testutil/fixture_ids.go` maps 75 native input labels to content IDs:
74 logical Office inputs and one PNG. Several labels can resolve to the same
physical file. `FixturePath` accepts these trusted labels; `LookupFixture` accepts
`fixture-<full-sha256>`. Paths, size, content hashes, format/group metadata and
regular-file custody are checked before use. Shared workflow input records use
`assetId` and repository-root-relative paths.

## Candidate and released references

The gitlink and pin name annotated release `v0.8.0`, commit
`b5729fc0c0fd59a98bfdfa5e4906143c3c6ce56a`. Initialise the recorded submodule and
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
There is no separate pack seal. The root manifest seals the explicit
`workflows/mutation-safety.feature` and `contracts/mutation-safety.json`.
Candidate tags begin with `candidate-` and have an empty `tag_object`; release
pins require the exact annotated tag object and its peeled commit. Use the
coordinator-provided pin, not hashes recomputed to accept modified inputs.
Release `v0.8.0` contains 140 manifest assets, including 115 unique fixtures in
32 format/scenario groups, plus 149 facts and 101 workflows/184 expanded cases.
The added run-formatting feature leaves all prior 139 asset records unchanged.
Its 15 cases add no Go bindings or formatting execution credit. Keep this shared
pin until the coordinator approves a replacement; local catalogue work continues
independently. Go's package ledger has four partial mappings and eight unmapped
declarations; its XML ledger has three partial mappings and one unmapped.
Neither ledger grants canonical execution credit. The semantic-diff assertion
`removed: []` exercises no removal case.

Once a release is approved, update the gitlink and
`spec/reference-distribution.json` together. All consumer branches/version tips
must use that same commit. Do not use `git submodule update --remote` or move an
existing release tag to follow changing content.

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

The v0.8 default released-reference batch is recorded in
`../reports/batches/113.md`: 291 implemented native Gherkin cases; 20 planned and
one external case unrun. The prior v0.3 bounded race checkpoint remains
`../reports/batches/096.md`; the pin-only v0.8 update does not rerun races.
Package/unit/subtest counts are separate metrics. No live Office
rendering/calculation or exploratory fuzz campaign ran in these batches.
Historical reports retain the commands and outcomes from their original runs.

## Behaviour catalogue and publication

The central registry owns canonical behaviour IDs and expected outcomes. Local
`docs/behaviors` files stage native findings for reconciliation; they do not form
a second canonical suite. The current inventory has 404 native declarations,
including the checkout guards: 133 have reviewed family/parameter mappings in
125 candidates, and 271 remain unreviewed. This is the frozen inventory snapshot;
later test declarations remain outside its denominator until a deliberate refresh. Importing Gherkin never earns execution
credit, and sibling results do not become Go passes.

The rewritten history and main/version-tip migrations are local preparations.
The configured GitHub credential has no write permission for this repository.
Go publication has not occurred. The shared release is available, but publishing
the rewritten Go refs still requires write access, final ref audits and explicit
leases against unchanged remote refs. See the batch
reports for historical results, not current release certification.
