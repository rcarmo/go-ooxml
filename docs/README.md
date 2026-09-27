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

[Batch 158](../reports/batches/158.md) records the v0.43 descriptor-integrity
adapter. [Batch 157](../reports/batches/157.md) records the v0.42 negative-budget admission
adapter. [Batch 156](../reports/batches/156.md) records the v0.41 ZIP-overlap and fixture
custody adoption. [Batch 155](../reports/batches/155.md) records the v0.35 root and acceptance
checks, including the move of Go feature sources into the pinned shared checkout. Other batch reports record their own commands and outcomes; their
results apply to those revisions. Shared fixtures,
provenance/licences, facts and canonical workflows belong to the pinned
`fixtures-ooxml` checkout; this directory does not mirror their source inventories.
