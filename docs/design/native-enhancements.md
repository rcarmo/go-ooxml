# Native preservation-safe editing

The Go runtime provides additive retained-source editors for DOCX, PPTX and XLSX.
Existing legacy authoring APIs remain separate. Mutable editing sessions are
single-owner; readers must not assume concurrent mutation is safe.

## Implemented subsets

- OPC: bounded ZIP intake, immutable payload snapshots, fingerprinted changes,
  atomic delivery, relationship inspection, planned additions/retargets/removals
  and detached-leaf deletion. Replacement-only receipts use schema1; graph
  additions/deletions use schema2.
- XML: namespace-aware lossless text/attribute edits and bounded structured
  insertion/removal/replacement. Unknown or ambiguous target structures refuse.
- DOCX: read-only story views, Unicode-aware search, exact spans and ordinary
  run-preserving replacement. Selected batches are atomic; replace-all reports
  per-match refusals. Formatting, revision, composition and field support is
  incomplete.
- PPTX: bounded plain-shape text and existing notes-body edits. Notes support
  ordinary multiline text and proved first-paragraph/run templates while
  retaining other placeholders. Broader shape, table, layout and import work is
  incomplete.
- XLSX: style-checked numeric edits, static dependency/cache invalidation with
  conditional calculation-chain cleanup, isolated PNG/JPEG replacement and
  read-only list validation vocabulary inspection. No formula calculation,
  display formatting or general structural worksheet editing is provided.

`spec/native-capabilities.json` links partial native capabilities to implemented
regression IDs. It contains explicit limits, not a completeness score.

## Shared references

Every consumer uses the annotated `v0.1.1` tag of `rcarmo/fixtures-ooxml` at
`references/fixtures-ooxml`. `spec/reference-distribution.json` pins the commit,
tag object, manifests and inventory counts. Native tests verify all pinned
assets and shared fact/workflow links. Only locally executed assertions earn
execution credit; shared planned cases and sibling outcomes remain separate.

The shared checkout is read-only. Generated archives go to consumer-local
`artifacts/generated` or temporary directories. Fixture readers never fall back
to legacy local copies. `OOXML_FIXTURES_ROOT` selects an explicit candidate root;
a candidate must still match the pinned distribution. Missing inputs fail.

## Verification and limits

Batch076 passed full library/acceptance and bounded formula/XML/package/all-format
and acceptance races after validation changes. Batches077-078 migrated fixture
readers and checked shared facts with291 native Gherkin cases passing. The cutover
batch rechecks both modules with the installed submodule and verifies a recursive
clone. Historical reports retain native measured results; removed implementation
references are marked where applicable.

No live Office calculation/rendering, schema certification or complete enhancement
family coverage is established. The extended-comment MIME discrepancy is recorded
in the shared fact registry; preservation of both spellings does not establish
comment-authoring interoperability. See `comments-extended-content-type.md`.
