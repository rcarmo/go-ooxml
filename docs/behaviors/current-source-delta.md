# Native test source added after the frozen inventory

At Go `8e6fa288d10c309602406a28482b2fa5fcbae2f5`, the tracked source has 165 `*_test.go` files and 410 runnable declarations. The historical [`native-inventory.json`](native-inventory.json) observes `ec71b6cf4c6ac8bebec49ee195ac64520b104be9`: 155 files and 404 declarations. Static discovery with `go run ./tools/nativeinventory -root . -out <path>` found ten added files and six added declaration IDs, with none removed. The output path is resolved under `-root`; move temporary output out of the repository after generation, without staging it. Eleven existing declaration records also differ between observations, so the six new IDs do not describe every source change.

| Added file | New declaration or role |
| --- | --- |
| `acceptance/canonical_cache_test.go` | Godog bindings; no new Go test declaration |
| `acceptance/descriptor_collision_test.go` | Godog bindings; no new Go test declaration |
| `acceptance/feature_roots_test.go` | Godog bindings; no new Go test declaration |
| `acceptance/mutation_contract_test.go` | `TestMutationContractAndCompilerBatch` |
| `acceptance/negative_budget_test.go` | Godog bindings; no new Go test declaration |
| `acceptance/zip_overlap_test.go` | Godog bindings; no new Go test declaration |
| `internal/testutil/legacy_fixture_migration_test.go` | `TestLegacyFixtureMigrationBatch` |
| `pkg/ooxml/pml/slide_visibility_test.go` | `TestSlideShowAttributeQualification` |
| `pkg/presentation/retained_visibility_test.go` | `TestRetainedSlideVisibilityStructure`; `TestRetainedSlideVisibilityRejectsBrokenStructure` |
| `pkg/presentation/slide_visibility_semantics_test.go` | `TestSlideHiddenAttributeQualification` |

The six new declarations have these *source-level* parameter groups:

- `TestMutationContractAndCompilerBatch`: 11 named invalid-contract rows (`unknown membership`, `duplicate scenario`, `duplicate fixture`, `unknown allowed member`, `duplicate allowed member`, `malformed member hash`, `schema 1 rejects feature list`, `schema 2 requires feature list`, `schema 2 rejects duplicate feature`, `schema 2 rejects escaping feature`, `schema 2 rejects absolute feature`); five named schema-1 feature-rejection rows (`wrong expanded count`, `unknown scenario`, `claim executed lifecycle`, `duplicate destination row`, `invalid typed JSON`); six named schema-2 feature-rejection rows (`missing selected scenario`, `claim executed lifecycle`, `duplicate selected scenario in second feature`, `wrong expanded count`, `duplicate destination row`, `invalid typed JSON`). Four other literal groups have 1, 2, 1 and 1 rows and serve as fixture/data lists, not additional named acceptance cases. There are three dynamic subtest sites; the schema-1/schema-2 branches are alternatives.
- `TestLegacyFixtureMigrationBatch`: two dynamic subtest sites over historical migration entries; no literal parameter group. The original v0.35 identities remain historical.
- `TestSlideShowAttributeQualification`: seven named rows: `default`, `namespaced-only`, `foreign-only`, `unqualified-hidden`, `unqualified-visible`, `both-qualified-first`, `both-unqualified-first`.
- `TestRetainedSlideVisibilityStructure`: two named input rows, `retained-S` and `visible-P`, plus two four-value relationship-ID arrays. It inspects retained package structure and source-byte custody only.
- `TestRetainedSlideVisibilityRejectsBrokenStructure`: 14 named corruption rows (`unqualified-show`, `wrong-show-namespace`, `missing-marker`, `wrong-marker-slide`, `duplicate-slide-relationship`, `reused-slide-part`, `escaping-slide-part`, `external-slide-relationship`, `missing-slide-relationship`, `wrong-slide-relationship`, `duplicate-slide-id-list`, `reordered-slide-id`, `malformed-slide-xml`, `duplicate-root-marker`), plus fixed `missing-slide-part` and `control-unqualified-show` subtests. Two four-value relationship-ID arrays are expected identities, not further cases.
- `TestSlideHiddenAttributeQualification`: five named rows: `retained-namespaced`, `unqualified-hidden`, `unqualified-visible`, `absent-default`, `both-unqualified-hidden`.

All six discovered records start `unreviewed` with empty `scenario_ids`. The generator records source declarations, call sites, dynamic subtest sites and literal groups; it does not enumerate runtime leaves or prove which literal rows are behavioural cases. Godog step registrations and helper-only files do not create new cases. Fuzz seeds and benchmarks do not establish fuzz or performance results. The frozen 197 reviewed / 207 unreviewed declaration split belongs to `ec71b6c` and has not been extended to the 410-declaration source observation.

These tests add **zero canonical execution credit**. Go's separate selected Gherkin gate at v0.57 remains 294 cases / 1013 steps. Neither retained slide case is bound to the shared PPTX visibility workflows, and `p:show="0"` supplies no Office-confirmed hiding. `@CACHE-001` remains deselected in favour of the one-for-one canonical cross-sheet-cache case. No shared, Bun or Python denominator is changed here.

## Verification

Static discovery was run at `8e6fa288d10c309602406a28482b2fa5fcbae2f5` with `go run ./tools/nativeinventory -root . -out workspace/tmp/go-native-inventory-8e6fa28.json`; its output was moved to `/workspace/tmp/go-native-inventory-8e6fa28.json` before comparison. Declaration IDs and file paths were compared with `docs/behaviors/native-inventory.json` using sorted set differences. No frozen JSON was regenerated in place.

With this note and the README link in the working tree, the uncached, no-reference-override batch `env -u OOXML_FIXTURES_ROOT -u OOXML_REFERENCE_PIN GOMAXPROCS=2 GOFLAGS=-count=1 make test-batch TEST_JOBS=2` passed the root and acceptance modules: 294 selected Cucumber cases, 1013 passed steps, zero failures or skips. Log: `/workspace/tmp/go-current-source-delta-local-gate.log`, SHA-256 `7c49dce7f22c1644a8415de753948dd4b98a458654f9259c80d6384a53886ebd`. Race, fuzz and Office gates were not run for this documentation change.
