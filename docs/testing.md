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

The tracked gitlink/pin currently names release `v0.1.1`, which predates schema2
lookup. This intermediate migration cannot run against that older layout.
The tested grouped-layout candidate is commit
`9ab4a029d6dc4207a48592050795d599df21c19a`; it is not yet a released tag.

A candidate run requires both explicit overrides:

```sh
OOXML_FIXTURES_ROOT=/path/to/clean/candidate-checkout \
OOXML_REFERENCE_PIN=/path/to/candidate-pin.json \
GOMAXPROCS=2 make test-batch
```

The pin uses schema1 with `commit`, `tag`, `tag_object`, `manifest_sha256`,
`shared_pack_sha256`, `assets`, `facts`, `workflows` and `workflow_cases`.
Candidate tags begin with `candidate-` and have an empty `tag_object`; release
pins require the exact annotated tag object and its peeled commit. Use the
coordinator-provided pin, not hashes recomputed to accept modified inputs.
The current candidate contains 123 manifest assets, including 115 unique
fixtures, plus 149 facts and 39 workflows/55 expanded cases.

Once a release is approved, update the gitlink and
`spec/reference-distribution.json` together. All consumer branches/version tips
must use that same commit. Do not use `git submodule update --remote` or move an
existing release tag to follow changing content.

## Integrity and output custody

Root-module and acceptance checks require the exact reference HEAD, root and pack
seals, clean index/worktree and tracked bytes/modes. Release tags must be annotated.
Changes to facts or workflows fail even when fixture manifest hashes still match;
`assume-unchanged` does not hide altered tracked bytes. Verification is read-only.
An absent Git checkout, mismatched pin or missing input is a failure, not a skip.

Write outputs under `t.TempDir()` or consumer-local `artifacts/generated`.
Persistent output guards reject reference-root descendants, including symlink
redirects. Never regenerate or rebaseline the shared inputs during a test run.

## Batched verification

`make test` covers the root Go module only. `make test-batch` also runs the separate
acceptance module. Use `GOMAXPROCS=2` and `-p 2`; run related packages together,
including failure reruns. Individual-test retry loops are not part of the workflow.
Full/race checks belong at integration points; reuse caches and avoid concurrent
duplicate suites. The runtime has no external dependencies; Godog/Gherkin are
isolated in the acceptance module, and checkout verification uses Git only in tests.

The latest full candidate integration and helper race results are in
`../reports/batches/089.md`: 291 implemented native Gherkin cases; 20 planned and
one external case unrun. Package/unit/subtest counts are separate metrics.
No live Office rendering/calculation or exploratory fuzz campaign ran in that batch.
Historical reports retain the commands and outcomes from their original runs.

## Behaviour catalogue and publication

The central registry owns canonical behaviour IDs and expected outcomes. Local
`docs/behaviors` files stage native findings for reconciliation; they do not form
a second canonical suite. At the catalogue checkpoint, 59 of 402 discovered
native declarations had reviewed family/parameter mappings in 54 candidates.
The remaining 343 were unreviewed; subsequent guard tests are additional unmapped
work until the inventory is refreshed. Importing Gherkin never earns execution
credit, and sibling results do not become Go passes.

The rewritten history and main/version-tip migrations are local preparations.
The configured GitHub credential has no write permission for this repository.
Publication has not occurred; it requires the common release pin, renewed tests
and audits, and explicit leases against unchanged remote refs. See the batch
reports for historical results, not current release certification.
