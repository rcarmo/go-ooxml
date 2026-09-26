# Graph and ZIP64 reconciliation handoff

Reviewed Go `395b1e99556b3e36e230511a9fc09f15ba30217e` against released shared
`v0.4.0` / `40eb26e684b12073956e4f24915444075a60c212`. This is a bounded
comparison of the existing graph/archive staging assertions. No declaration gains
complete coverage, no candidate becomes a canonical binding and no canonical
execution credit is assigned. Go refs are local and unpublished.

Shared feature seals:

- `workflows/native/opc-graph.feature`: `20f2ca5d050d96372286a8b03920bcc301709504a4578d6e7bde212c2411e3fe`
- `workflows/native/opc-zip64.feature`: `18da944b62e7c88c7558627bca9fbc211a121799efbd7cefeb04051dd573af91`

The authoritative ZIP64 IDs are `parity-zip64`, `zip64-preflight-count`,
`zip64-unsafe-offset`, and `zip64-resource-limit`. Earlier requested names
`zip64-count-write`, `zip64-resource-preflight`, and `zip64-corruption-refusal`
do not occur in the released feature. The coordinator corrected those names;
writer-count and corruption assertions stay separate native scope.

## Released scenarios and native overlap

Source keys below resolve to exact test declarations and file seals in the next
section. Candidate references resolve to the existing `*-mapping.json` files.

| Released ID | Existing assertion overlap | Differences and untested outcomes |
| --- | --- | --- |
| `@id-opc-add-related-part` | G3 private payload snapshots, addition receipt fields and preserved compressed member streams/metadata; G2 accumulates additions, reopens and validates the graph. Candidates graph-002/004/005/006. | Synthetic Go package, not the shared DOCX. Go adds a part then retargets an existing document relationship; `GraphMutation` has no relationship-creation operation. No fresh internal root relationship or independent exact MIME assertion for that canonical action. Compressed-byte custody is a separate assertion from canonical payload custody. |
| `@id-opc-graph-rollback` | G1 failed mixed plan preserves prior edited bytes and a held valid plan; G3 malformed later addition refuses without serialised changes. Candidates graph-003/008. | These failure conditions are stale payload fingerprints and malformed additions, not deletion of the new referenced opaque leaf. G1's nonleaf deletion also refuses, but that part has its own registry and its incoming edge is explicitly removed. No selected declaration asserts canonical code `opc-part-referenced`. Go `Refusal.Kind`/private preflight is not a Bun direct transaction/error-code contract. |
| `@id-opc-remove-related-part` | G1 retargets one document edge, removes the second and deletes the formerly shared media leaf in one plan; checks surviving edge, schema2 receipt, absence after save/reopen and valid graph. Candidate graph-010. | Different source, edge owner and operation sequence; no shared DOCX/root-edge detach. The deleted original media uses a default type, so this subtest does not independently assert removal of an explicit content-type override. Transient add/delete cancellation in G1 returns exact source bytes, but is a separate lifecycle operation. |
| `@id-opc-diff-content-type` | None in the selected graph declarations. | Go receipts assert payload/add/delete changes and registry patches. No type-only part diff with equal payload hashes or independent unrelated-add/remove assertion is tested here. Leave unmapped. |
| `@id-parity-zip64` | A7 edits `data.bin` in forced-ZIP64 input, saves/reopens, checks new bytes and exact unchanged `[Content_Types].xml`/`empty.bin` payloads, then adds a part and saves again. Candidate archive-008. A3 checks exact no-op archive bytes. | Go fixture has three members (XML registry plus two binary members), not two XML members. Go has no tested ZIP64-enabled output option in this subtest and it does not assert that the small edited output retains ZIP64 encoding. Its payload assertions overlap; input/output geometry and runtime operation differ. |
| `@id-zip64-preflight-count` | A3 rejects a classic-count mismatch; A7 rejects a valid 65,535-entry archive under a 65,534-entry budget. | Neither assertion builds an impossible ZIP64 declared count relative to directory size or observes pre-allocation refusal. No `zip-structure-invalid` code assertion. Leave this exact scenario unmapped. |
| `@id-zip64-unsafe-offset` | A3 rejects local-size and ZIP64-record-length overflow cases. | These are not locator offsets above JavaScript's safe integer range. Go uses integer arithmetic; no exact safe-integer boundary or `zip-zip64-unsupported` code is asserted. Leave unmapped. |
| `@id-zip64-resource-limit` | A7 admits a valid 65,535-entry ZIP64 archive through `OpenBytes`, then asserts typed `resource_limit` from `OpenPreserved` with `MaxEntries=65534`. Candidate archive-009. A2 covers source/entry/part/total budgets on ordinary ZIP. | Canonical two-entry/one-entry values are not exercised by the selected ZIP64 assertions. Error taxonomy is Go `Refusal.Kind=resource_limit`, not `zip-too-many-entries`. A2 ordinary-ZIP tests alone cannot supply ZIP64 coverage or pre-allocation evidence. |

## Source declaration identities and seals

All paths are relative to this Go repository. File SHA256 values match both the
committed staging map and Go `395b1e9`; they are not hashes of individual functions.

| Key | Exact native ID | Source SHA256 |
| --- | --- | --- |
| G1 | `pkg/packaging/graph_delete_lifecycle_test.go::TestGraphDeletionLifecycleBatch` | `735b40e7df1fc466ce4388c95d52e7638798e35ea05724ae8a662615e30b6453` |
| G2 | `pkg/packaging/graph_edit_properties_test.go::TestRetargetURIAndChainedPlans` | `e63c6307b3ba9ff78e3d09d27ed786f936dcdeb94e62d5ae95881f73bf06b64b` |
| G3 | `pkg/packaging/graph_edit_test.go::TestGraphPlanContracts` | `3fb4fda1664a6b2e11efc267faab2d8c1478c5438e0d3883cd993e060243ac02` |
| G4 | `pkg/packaging/graph_test.go::TestGraphURIResolutionBatch` | `a7e99fa69a76c222ee4c64e0bc0a4dae2570e6450feddc12fa5e5c9ec63f4526` |
| A1 | `pkg/packaging/fixture_corpus_test.go::TestRetainedNoOpFixtureCorpus` | `4049233ac8b30bf0086947360bc756a98db6457bf79bede4ac2c7045eb99f6dc` |
| A2 | `pkg/packaging/limits_test.go::TestIntakeBudgetAndIntegrityContracts` | `dbe634a45fd262860cc84d2ffe7dd7aa4346979296d94a570b697f8833f58442` |
| A3 | `pkg/packaging/zip64_test.go::TestZIP64IntakeBatch` | `4e7b20044712613deb6d019baa6eb29c35403106092fdd716f2d321a522295a3` |
| A4 | `pkg/packaging/zip64_test.go::TestZIP64TerminalGeometryBatch` | `4e7b20044712613deb6d019baa6eb29c35403106092fdd716f2d321a522295a3` |
| A5 | `pkg/packaging/zip_overlap_test.go::TestIndependentPhysicalOverlap` | `0c63da0c273dcd9f1b850741bc558a00dd80c901d87517b67965366680580614` |
| A6 | `pkg/packaging/zip_structure_test.go::TestZIPStructureContracts` | `0b20478a0823f579a2ca90e7223f03fc7f67f22b032c4442df62432ca239fe15` |
| A7 | `pkg/packaging/zip64_delivery_test.go::TestZIP64DeliveryBatch` | `bf9df4fee8a39b8b46113261b324d2e8b74ea9df10c4975e81f5d112e625daa7` |
| A8 | `pkg/packaging/zip_adversarial_test.go::TestZIPDescriptorAndPrefixBatch` | `1134e2d4fe16c97795b2a604c6332d4b159479830015199c4774a7b1f1a780ad` |

## Parameter and extra native scope

- G2's unlabelled nested loop covers 4 sources × 4 targets × 3 original target
  forms (48 combinations), including relative/absolute forms and Unicode/escaped
  characters. It checks URI resolution round trips; those are not 48 canonical
  graph-edit cases. Plan owner/ABA/consumption, snapshots and repeated no-ops in
  G3 remain additional native assertions.
- A3 has 2 compression choices × 3 descriptor forms (6 valid variants), plus 15
  named malformed-input variants. It requires errors for the malformed rows, not
  exact canonical refusal codes. A4 separately has one valid terminal geometry
  and three invalid ones. These deterministic tables are not exploratory fuzzing.
- A7 has three edited deflate variants (no/signed/unsigned descriptor), one
  65,535-entry writer-boundary case, and direct ZIP64-extra-field bounds checks.
  The count boundary is exercised via Go's standard `archive/zip` writer. It does
  not certify large payloads, all count boundaries or forced output encoding.
- A5 independently reads three overlapping, distinct-name/CRC-valid members with
  `archive/zip`, then asserts Go's physical-overlap refusal and unchanged source.
  A6 checks four local/central inconsistencies. A8 checks prefixed-archive refusal
  and the ambiguous descriptor-signature CRC geometry; full intake still refuses
  that deliberately incorrect payload CRC. These are extra corruption assertions,
  not bindings to any absent corruption scenario ID.
- A1's corpus no-op traversal is dynamic over the committed fixture ID map;
  ordinary payload retention does not establish that every producer fixture is
  ZIP64. No new leaf-case count was measured for this handoff.

Batch098 already covers the released default root/acceptance gates. The handoff
adds no runtime change. Any subsequent scoped validation is recorded separately;
no independent audit completed (the read-only delegate timed out).
