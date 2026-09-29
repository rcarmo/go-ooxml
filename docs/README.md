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

[Batch 276](../reports/batches/276.md) records v0.133 shared-ledger credit
for Go's published implicit `xml` prefix case from
[Batch 275](../reports/batches/275.md): one case / five steps. Bun and Python
are unchanged. The pin changes no Go runtime or acceptance binding.
[Batch 274](../reports/batches/274.md) records v0.132 Bun-only ledger credit
for the exact caller-owned XML byte seed: one case / four steps from published
Bun `c996c7d`. Go and Python credit are unchanged. The Go pin changes no
runtime or acceptance binding.
[Batch 273](../reports/batches/273.md) records v0.131 shared-ledger credit
for Go's published exact pre-root stylesheet-PI case from
[Batch 272](../reports/batches/272.md): one case / three steps, accepting the
literal PI with root `r`. Bun and Python are unchanged. The pin changes no Go
runtime or acceptance binding; it adds no PI enumeration or general preservation.
[Batch 271](../reports/batches/271.md) records v0.130 shared-ledger credit
for Go's already published exact XML entity-values case from
[Batch 270](../reports/batches/270.md): one case / four steps. Python alone
receives pinned existing Word comment inspection credit (one case / four
steps); Bun is unchanged. The pin changes no Go runtime or acceptance binding.
[Batch 269](../reports/batches/269.md) records v0.129 shared-ledger credit
for Go's published element-replacement refusal outline from
[Batch 268](../reports/batches/268.md): three cases / twelve steps. Python
alone receives styled-blank XLSX cell credit (one case / four steps); Bun is
unchanged. The pin changes no Go runtime or acceptance binding.
[Batch 267](../reports/batches/267.md) records v0.128 shared-ledger credit
for Go's published element-replacement custody binding from
[Batch 266](../reports/batches/266.md): one case / three steps. Python alone
receives staged file-publication OPC rollback credit. The pin changes no Go
runtime or acceptance binding.
[Batch 265](../reports/batches/265.md) records v0.127 shared-ledger credit
for Go's published child-namespace matrix from
[Batch 264](../reports/batches/264.md): one six-step case with 100 derived
combinations. Bun and Python remain planned. The pin changes no Go runtime or
acceptance binding.
[Batch 263](../reports/batches/263.md) records v0.126 shared-ledger credit
for Go's published child-insertion refusal case (one case / four steps) from
[Batch 262](../reports/batches/262.md), plus Python-only XML whitespace
serialization/reparse (one case / three steps). It changes no Go runtime or
binding.
[Batch 261](../reports/batches/261.md) records v0.125 shared-ledger credit
for Go's published structured child-insertion custody case (one case / four
steps) from [Batch 260](../reports/batches/260.md). It changes no Go runtime
or binding; Bun and Python remain planned for this case.
[Batch 259](../reports/batches/259.md) records direct v0.124 adoption:
v0.123 credits Go's published root and nested element-removal refusal rows
(two cases / six steps) from [Batch 258](../reports/batches/258.md);
v0.124 adds only Python's implicit XML-prefix case (one case / five steps).
This pin changes no Go runtime or binding.
[Batch 257](../reports/batches/257.md) records direct v0.122 adoption:
v0.121 credits Go's published two-target XML element removal custody case
(one case / four steps) from [Batch 256](../reports/batches/256.md);
v0.122 adds only Python's expanded XML attribute lookup case (one case /
three steps). This pin changes no Go runtime or binding.
[Batch 255](../reports/batches/255.md) records direct v0.120 adoption:
v0.119 credits Go's published duplicate-attribute refusal (one case / three
steps); v0.120 adds only Python stylesheet processing-instruction evidence
(one case / three steps). The Go pin changes no runtime or binding.
[Batch 254](../reports/batches/254.md) records the Go refusal execution.
[Batch 253](../reports/batches/253.md) records direct v0.118 adoption:
v0.117 credits three published Go attribute-splice rows / nine steps, and
v0.118 adds only Python XML entity-values evidence (one case / four steps).
The Go pin changes no runtime or binding. [Batch 252](../reports/batches/252.md)
records the Go execution.
[Batch 251](../reports/batches/251.md) records the v0.116 Python-only
staged-DOCX OPC package-preservation evidence pin (one case / four steps).
Go gains no credit or new selection.
[Batch 250](../reports/batches/250.md) records the v0.115 shared-ledger
credit for Go's published immutable XML leaf seed (one case / four steps). It
changes no Go runtime or binding. [Batch 249](../reports/batches/249.md)
records the Go execution.
[Batch 248](../reports/batches/248.md) records the v0.114 pin and shared-ledger
credit for Go's published Unicode QName case (one case / four steps) and Python's
five package IDs (14 cases / 47 steps). It changes no Go runtime or binding.
[Batch 247](../reports/batches/247.md) records the Go Unicode QName execution.
[Batch 246](../reports/batches/246.md) records the v0.113 Python-only
negative admission and ZIP32 descriptor evidence pin; Go gains no new credit.
[Batch 245](../reports/batches/245.md) records v0.112 shared-ledger credit
for four bounded Go XML Boolean negative IDs already published in Batch 244.
[Batch 243](../reports/batches/243.md) records v0.111 shared-ledger credit
for two exact Go admission IDs already executed since v0.42 and v0.43.
[Batch 242](../reports/batches/242.md) records the v0.110 Python-only XML
Boolean negative evidence pin; it adds no Go XML execution credit.
[Batch 241](../reports/batches/241.md) records the v0.109 correction of
latent Go ZIP32 overlap execution credit; it adds no new Go code or selection.
[Batch 240](../reports/batches/240.md) records the v0.108 Python-only bounded
ZIP32 overlap-refusal evidence pin; it adds no Go execution credit.
[Batch 239](../reports/batches/239.md) records the v0.107 shared-ledger credit
for four Go static formula-reference IDs already published in Batch 238.
[Batch 237](../reports/batches/237.md) records the v0.106 Bun-only DOCX
paragraph-style and Python-only XLSX comment/VML evidence pin; Go gains no credit.
[Batch 236](../reports/batches/236.md) records the v0.105 shared-ledger credit
for three already-executed Go static formula-analysis IDs, not new execution.
[Batch 235](../reports/batches/235.md) records the v0.104 Bun-only DOCX
final-section page-layout ledger pin; it adds no Go execution credit.
[Batch 234](../reports/batches/234.md) records the v0.103 Bun-only DOCX
effective-run formatting ledger pin; it adds no Go execution credit.
[Batch 232](../reports/batches/232.md) records the v0.102 Go direct-range and
Bun-only DOCX tracking ledger pin; it adds no Go execution beyond Batch 231.
[Batch 230](../reports/batches/230.md) records the v0.101 Bun-only
XLSX static formula-reference API pin; it adds no Go execution credit.
[Batch 229](../reports/batches/229.md) records the v0.100 Bun-only
XLSX existing-comment/VML graph pin; it adds no Go execution credit.
[Batch 228](../reports/batches/228.md) records the v0.99 Bun-only
XLSX direct cell-style outcome pin; it adds no Go execution credit.
[Batch 227](../reports/batches/227.md) records the v0.98 Bun-only
XLSX creation and missing-cell outcome pin; it adds no Go execution credit.
[Batch 226](../reports/batches/226.md) records the v0.97 Bun-only
PPTX slide-permutation outcome pin; it adds no Go execution credit.
[Batch 225](../reports/batches/225.md) records the v0.96 Bun-only
PPTX positioned text-box outcome pin; it adds no Go execution credit.
[Batch 224](../reports/batches/224.md) records the v0.95 Bun-only
PPTX table outcome pin; it adds no Go execution credit.
[Batch 223](../reports/batches/223.md) records the v0.94 Bun-only
PPTX title-slide no-edit custody pin; it adds no Go execution credit.
[Batch 222](../reports/batches/222.md) records the v0.93 Bun-only
existing DOCX comment-thread ledger pin; it adds no Go execution credit.
[Batch 221](../reports/batches/221.md) records the v0.92 Bun-only
existing DOCX comment ledger pin; it adds no Go execution credit.
[Batch 220](../reports/batches/220.md) records the v0.91 Bun-only
PPTX/XLSX relationship expanded-name ledger pin; it adds no Go execution credit.
[Batch 219](../reports/batches/219.md) records the v0.90 Bun-only
OPC custody and save ledger pin; it adds no Go execution credit.
[Batch 218](../reports/batches/218.md) records the v0.89 Bun-only
common OPC preservation ledger pin; it adds no Go execution credit.
[Batch 217](../reports/batches/217.md) records the v0.88 Bun-only
negative admission-budget ledger pin; it adds no Go execution credit.
[Batch 216](../reports/batches/216.md) records the v0.87 Bun-only
unsigned ZIP32 descriptor-collision refusal pin; it adds no Go execution credit.
[Batch 215](../reports/batches/215.md) records the v0.86 Bun-only
ZIP32 configured-bounds ledger pin; it adds no Go execution credit.
[Batch 214](../reports/batches/214.md) records the v0.85 Bun-only broad
ZIP32 unsafe-structure refusal pin; it adds no Go execution credit.
[Batch 213](../reports/batches/213.md) records the v0.84 Bun-only positive
ZIP32 read/write ledger pin; it adds no Go execution credit.
[Batch 212](../reports/batches/212.md) records the v0.83 Bun-only ZIP32
ledger pin; it adds no Go execution credit.
[Batch 211](../reports/batches/211.md) records the v0.82 Bun-only
semantic package-diff ledger pin; it adds no Go execution credit.
[Batch 210](../reports/batches/210.md) records the v0.81 Bun-only
package-admission refusal ledger pin; it adds no Go execution credit.
[Batch 209](../reports/batches/209.md) records the v0.80 Bun-only
conservative XML comparison ledger pin; it adds no Go execution credit.
[Batch 208](../reports/batches/208.md) records the v0.79 Bun-only XML QName
and NBSP ledger pin; it adds no Go execution credit.
[Batch 207](../reports/batches/207.md) records the v0.78 Bun-only XML value,
namespace and escaping ledger pin; it adds no Go execution credit.
[Batch 206](../reports/batches/206.md) records the v0.77 Bun-only XML edit safety
ledger pin; it adds no Go execution credit. [Batch 205](../reports/batches/205.md) records the v0.76 Go direct-size and Bun-only
XML parser/refusal evidence pin. [Batch 204](../reports/batches/204.md) records the
Go six-step saved direct-size binding. [Batch 203](../reports/batches/203.md) records the v0.75 Bun-only saved DOCX direct
half-point font-size ledger pin; it adds no Go execution credit. [Batch 202](../reports/batches/202.md) records the v0.74 Bun-only DOCX table-authoring
ledger pin; it adds no Go execution credit. [Batch 201](../reports/batches/201.md) records the v0.73 Bun-only DOCX creation
ledger pin; it adds no Go execution credit. [Batch 200](../reports/batches/200.md) records the v0.72 Bun-only DOCX text-slice
ledger pin; it adds no Go execution credit. [Batch 199](../reports/batches/199.md) records the v0.71 Go/Bun body-insertion
ledger pin. [Batch 198](../reports/batches/198.md) records the exact Go in-memory
body-order binding. [Batch 197](../reports/batches/197.md) records the v0.70 Go/Bun paragraph-value getter
ledger pin. [Batch 196](../reports/batches/196.md) records the ten exact Go in-memory
alignment, spacing, flags and runs cases. [Batch 195](../reports/batches/195.md) records the v0.69 Go/Bun paragraph-text getter
ledger pin. [Batch 194](../reports/batches/194.md) records the five exact Go in-memory
paragraph text rows. [Batch 193](../reports/batches/193.md) records the v0.68 Go/Bun core and section
getter ledger pin. [Batch 192](../reports/batches/192.md) records the bounded Go core-property and
section/background getter bindings. [Batch 191](../reports/batches/191.md) records the v0.67 Go/Bun new-body and table-text
readback ledger pin. [Batch 190](../reports/batches/190.md) records Go new-document empty-body and table-text
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
