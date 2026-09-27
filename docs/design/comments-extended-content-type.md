# Extended-comment content type

The runtime constant `ContentTypeCommentsExtended` remains
`application/vnd.ms-word.commentsExtended+xml`. The shared fact registry marks
it disputed. `ContentTypeCommentsExtendedSpecified` records
`application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`
with vendor-metadata evidence. This migration does not change runtime authoring.

Read `references/fixtures-ooxml/facts/content-types.json` and
`facts/evidence.json` for versioned values, specification citations and provenance.
Both spellings use the relationship type
`http://schemas.microsoft.com/office/2011/relationships/commentsExtended`.

`@MIME-001` verifies that an unrelated retained-source body edit preserves either
input content-type registry, extended-comment payload and relationships exactly.
It does not validate comment anchors, thread semantics, schema conformance or
native Office behaviour. New thread authoring needs separate semantic and
interoperability tests.
