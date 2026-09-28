## Documentation and specification references

Start with [shared-reference testing](testing.md) for fixture-ID lookup, candidate
versus released pins, batched commands and integrity checks. The native editing
subsets and their limits are in [native-enhancements.md](design/native-enhancements.md).
[Behaviour catalogue staging](behaviors/README.md) tracks unfinished reconciliation
with the single central behaviour registry.

## Standards material

The pinned [ECMA-376 specification index](../references/fixtures-ooxml/specs/ecma-376/README.md)
links all four complete Parts, their exact editions, verbatim extracts, derived
notes, provenance and notices. The PDFs are the specification sources. Check the
edition, part and clause in a complete PDF before treating an extract or note as
a requirement. The presence of a document does not certify this implementation.
Deprecated VML remains relevant to existing packages; Transitional conformance
alone does not mark markup deprecated.

| File | Material |
| --- | --- |
| [Part 1 PDF](../references/fixtures-ooxml/specs/ecma-376/part-1/ECMA-376-Part1-Fundamentals.pdf) | Fundamentals and markup-language reference, October 2016 fifth edition |
| [Part 2 PDF](../references/fixtures-ooxml/specs/ecma-376/part-2/ECMA-376-Part2-OpenPackagingConventions.pdf) | Open Packaging Conventions, December 2021 fifth edition |
| [Part 3 PDF](../references/fixtures-ooxml/specs/ecma-376/part-3/ECMA-376-Part3-MarkupCompatibility.pdf) | Markup Compatibility and Extensibility, December 2015 fifth edition |
| [Part 4 PDF](../references/fixtures-ooxml/specs/ecma-376/part-4/ECMA-376-Part4-TransitionalMigration.pdf) | Transitional Migration Features, October 2016 fifth edition |
| [WordprocessingML notes](../references/fixtures-ooxml/specs/ecma-376/notes/ECMA-376-Phase3-Reference.md) | Project-authored explanations; not normative |
| [WordprocessingML extract](../references/fixtures-ooxml/specs/ecma-376/extracts/ECMA-376-WML-Phase3.md) | Verbatim selected text; check clauses in the full PDF |
| [Part 2 extract](../references/fixtures-ooxml/specs/ecma-376/extracts/ECMA-376-Part2-OPC.md) | Verbatim selected OPC text; check clauses in the full PDF |
| [Historical PDF custody](../spec/legacy-fixture-migration.json) | Original-main Part 1 and Part 2 paths and hashes map to the byte-identical pinned PDFs above; no local copies remain |
| [FIT-GAP-ANALYSIS.md](FIT-GAP-ANALYSIS.md) | Historical January 2026 assessment; percentages are not measured current coverage |
| [ROUNDTRIP-TESTS.md](ROUNDTRIP-TESTS.md) | Native fixture round-trip assertion catalogue |

Specification material is © Ecma International and retains its own terms. See the
[shared notices](../references/fixtures-ooxml/NOTICES.md). Vendor reference
material is available in the [Open XML documentation](https://learn.microsoft.com/en-us/office/open-xml/).

[Batch 190](../reports/batches/190.md) records Go new-document empty-body and table-text
save/reopen bindings. [Batch 189](../reports/batches/189.md) records the v0.66 Go/Bun table-value ledger
evidence pin. [Batch 188](../reports/batches/188.md) records the bounded Go table dimensions,
cell access/text and row-count bindings. [Batch 187](../reports/batches/187.md) records the v0.65 shared Go table merge-property
getter evidence pin. [Batch 186](../reports/batches/186.md) records the bounded Go table/cell merge-property
getter binding. [Batch 185](../reports/batches/185.md) records the v0.64 Go vertical and Bun-only
table/cell getter evidence pin. [Batch 184](../reports/batches/184.md) records the two-run direct vertical-align
getter binding. [Batch 183](../reports/batches/183.md) records the v0.63 Go colour/highlight
and selected saved-readback ledger evidence pin. [Batch 182](../reports/batches/182.md)
records the v0.62 Bun-only run-appearance
and readback evidence pin; it adds no Go credit. [Batch 181](../reports/batches/181.md)
records the colour/highlight and selected save-reopen run bindings. [Batch 180](../reports/batches/180.md) records the v0.61 Bun/Go run-getter
ledger adoption. [Batch 179](../reports/batches/179.md) records the direct run underline/font
getter bindings. [Batch 178](../reports/batches/178.md) records the v0.60 Go run-effect getter
ledger adoption. [Batch 177](../reports/batches/177.md) records the Go binding.
[Batch 176](../reports/batches/176.md) records the v0.59 Bun XLSX-style
readback evidence adoption. [Batch 175](../reports/batches/175.md) records the v0.58 Bun ZIP-overlap
execution ledger adoption. [Batch 174](../reports/batches/174.md) records the v0.57 Bun owned-chain
execution ledger adoption. [Batch 173](../reports/batches/173.md) records the v0.56 Bun owned-chain
refusal ledger adoption. [Batch 172](../reports/batches/172.md) records the v0.55 Python owned-chain
wording adoption. [Batch 171](../reports/batches/171.md) records the v0.54 decoder evidence
ledger adoption. [Batch 170](../reports/batches/170.md) records the Go decoder fix.
[Batch 169](../reports/batches/169.md) records the v0.53 visibility evidence
ledger adoption. [Batch 168](../reports/batches/168.md) records the native Go S/P structural
tests. [Batch 167](../reports/batches/167.md) records the v0.52 CI checkout custody
release adoption. [Batch 166](../reports/batches/166.md) records the v0.51 PPTX visibility
semantic correction. [Batch 165](../reports/batches/165.md) records the v0.50 inventory adoption. [Batch 164](../reports/batches/164.md) records the v0.49 Python owned-chain
evidence-ledger adoption. [Batch 163](../reports/batches/163.md) records the v0.48 Go cross-sheet cache
evidence-ledger adoption. [Batch 162](../reports/batches/162.md) records the v0.47 Python XLSX dependency
mapping-only adoption. [Batch 161](../reports/batches/161.md) records the v0.46 shared evidence-ledger
adoption. [Batch 160](../reports/batches/160.md) records the v0.45 owned-chain canonical
binding. [Batch 159](../reports/batches/159.md) records the v0.44 generated-input
retirement. [Batch 158](../reports/batches/158.md) records the v0.43 descriptor-integrity
adapter. [Batch 157](../reports/batches/157.md) records the v0.42 negative-budget admission
adapter. [Batch 156](../reports/batches/156.md) records the v0.41 ZIP-overlap and fixture
custody adoption. [Batch 155](../reports/batches/155.md) records the v0.35 root and acceptance
checks, including the move of Go feature sources into the pinned shared checkout. Other batch reports record their own commands and outcomes; their
results apply to those revisions. Shared fixtures,
provenance/licences, facts and canonical workflows belong to the pinned
`fixtures-ooxml` checkout; this directory does not mirror their source inventories.
