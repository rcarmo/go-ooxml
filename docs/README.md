## Documentation and specification references

Start with [shared-reference testing](testing.md) for fixture-ID lookup, candidate
versus released pins, batched commands and integrity checks. The native editing
subsets and their limits are in [native-enhancements.md](design/native-enhancements.md).
[Behaviour catalogue staging](behaviors/README.md) tracks unfinished reconciliation
with the single central behaviour registry.

## Standards material

The ECMA-376 documents and extracts below are specification references, not a
statement that this implementation conforms to every listed rule. Keep normative
citations and required notices intact when editing project documentation.

| File | Material |
| --- | --- |
| `ECMA-376-Part1-Fundamentals.pdf` | Fundamentals and markup-language reference, fifth edition |
| `ECMA-376-Part2-OpenPackagingConventions.pdf` | OPC, fifth edition |
| [ECMA-376-Phase3-Reference.md](ECMA-376-Phase3-Reference.md) | Extracts covering revisions, comments, styles, headers and footers |
| [ECMA-376-WML-Phase3.md](ECMA-376-WML-Phase3.md) | WordprocessingML extracts |
| [ECMA-376-Part2-OPC.md](ECMA-376-Part2-OPC.md) | OPC relationships, content types and properties extracts |
| [FIT-GAP-ANALYSIS.md](FIT-GAP-ANALYSIS.md) | Historical January 2026 feature assessment; percentages are not current measured coverage |
| [ROUNDTRIP-TESTS.md](ROUNDTRIP-TESTS.md) | Native fixture round-trip assertion catalogue |

The standard also includes Part 3, Markup Compatibility and Extensibility, and
Part 4, Transitional Migration Features. Neither is bundled here. Legacy markup
can still occur in real files; absence from these documents is not permission to
discard it during retained-source edits.

Specification sources: [Ecma International](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
and the [ECMA-376 fifth-edition mirror](https://github.com/QtExcel/ecma-376-5th).
The specification documents are © Ecma International. Vendor reference material
is available in the [Open XML documentation](https://learn.microsoft.com/en-us/office/open-xml/).

Batch reports under `../reports/batches/` record past commands and outcomes.
They are historical evidence, not current release status. Shared fixtures,
provenance/licences, facts and canonical workflows belong to the pinned
`fixtures-ooxml` checkout; this directory does not mirror their source inventories.
