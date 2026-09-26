## Extended-comment content type disagreement

Pinned Go and Python sources emit different content types for `word/commentsExtended.xml`. The authoring decision is unresolved; no MIME constant changed in this batch.

| Source | Revision | Symbol and file | Content type |
|---|---|---|---|
| go-ooxml | `43eda6edfe45db20da5879ba332ec6a25d598efa` | `ContentTypeCommentsExtended`, `pkg/packaging/constants.go` | `application/vnd.ms-word.commentsExtended+xml` |
| retired-docx | `[retired implementation revision]` | `_CT_COMMENTS_EXTENDED`, `[retired code reference]` | `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` |
| Python Office MCP | `36ac406ad9d4bd3e7538b4bcc7aa2fb0e51cc943` | `CT_COMMENTS_EXTENDED`, `[retired code reference]` | `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` |

All three use relationship type `http://schemas.microsoft.com/office/2011/relationships/commentsExtended`. Relationship agreement does not resolve the MIME disagreement.

The machine-readable record is [`spec/source-discrepancies.json`](../../spec/source-discrepancies.json). The retained-source editor preserves the input content-type registry, relationship registries and extended-comment payload during an unrelated body edit. `@MIME-001` characterises that behaviour for both spellings. The synthetic fixture has an unanchored review note; it tests preservation, not valid comment anchoring or native Office acceptance.

Before changing authoring policy, identify a versioned Microsoft extension specification or Open XML SDK part contract, then run authorised native Office fixtures through open/save with the Office version and any repair messages recorded. Test replies/resolution and paragraph-ID links independently of package preservation. Neither spelling is certified by the current tests.

The sibling Python audit also limits revision acceptance to its tested main-document scope and cache freshness to its tested cell-formula consumers. Go has separate partial implementations: read-only story projections do not establish all-story revision resolution, and static numeric invalidation refuses names, tables, chart/pivot sources, dynamic/shared/array formulas and other unproved consumers. Those boundaries are recorded in `SPEC.md` and `reports/batches/046.md`.
