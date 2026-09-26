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

The target layout is one shared schema2 manifest and grouped
`fixtures/<format>/<scenarioGroup>/` inputs. Native labels resolve fixture IDs;
physical files are deduplicated by SHA-256. `spec/reference-distribution.json`
records the commit, annotated release tag and seals. All prepared consumers use
released `v0.6.0`, commit `dc8fdccd5a7e14c9154bb71e68e10c7404fe4fa0`; default
checks need no override. Candidate checking is separate, as described in
`../testing.md`.

Tests verify the pinned full tracked tree, checkout identity, root manifest seal and shared
fact/workflow links. Only locally executed assertions earn execution credit;
planned cases and sibling outcomes remain separate.

The shared checkout is read-only. Generated archives go to consumer-local
`artifacts/generated` or temporary directories. Fixture readers never fall back
to legacy local copies. Candidate checks require both `OOXML_FIXTURES_ROOT` and
`OOXML_REFERENCE_PIN`; exact candidate HEAD/seals and clean tracked bytes are
required without claiming an annotated release. Missing inputs fail.

## Verification and limits

The default released-reference library/acceptance batch passes with 291 native
Gherkin cases; 20 planned and one external case remain unrun. See
`../../reports/batches/107.md` for the v0.6 default commands and checkout custody.
The bounded race checkpoint is the earlier v0.3 batch096; v0.6 changes only
reference pins/counts and documentation.
Historical reports retain original measured results, with sanitised implementation
citations marked as historical. No removed inventory row becomes completed work.

No live Office calculation/rendering, schema certification or complete enhancement
family coverage is established. The extended-comment MIME discrepancy is recorded
in the shared fact registry; preservation of both spellings does not establish
comment-authoring interoperability. See `comments-extended-content-type.md`.
