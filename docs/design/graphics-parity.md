# Bounded PowerPoint graphics parity

The Go port targets the twenty Bun graphics operations and its admitted SmartArt source encoding, using centrally sealed workflows, recipes and literal expected results. LibreOffice is the application quality gate. Microsoft Office testing is not required for this task.

## Shared inputs and tests

Shared candidate: `fixtures-ooxml` commit `0cf83c156c8e9433347e446e58c415db011b17be`. The released gitlink and reference pin stay at `v0.152.0`, commit `28e492f50979aaec6ab8d8d001cd9c37790e7fc6`.

```sh
GOMAXPROCS=2 make graphics-test
# An isolated clone can use another clean checkout of the exact shared commit:
GOMAXPROCS=2 make graphics-test GRAPHICS_ROOT=/path/to/fixtures-ooxml
# Include the candidate tests with the root and acceptance full gates:
OOXML_GRAPHICS_ROOT=/path/to/fixtures-ooxml GOMAXPROCS=2 make test-batch
```

The new graphics helper verifies the clean committed checkout, root manifest seal, fixture IDs and sealed recipe/contract/feature bytes. It uses an independent ZIP reader to apply the literal recipes to temporary in-memory archives. Shared files remain unchanged. These tests do not change the default acceptance selector or grant shared lifecycle credit.

## Implemented slices

### Picture inspection

`EditSession.InspectPictures(slidePart)` returns detached, source-order `PictureInfo` records with exact shape identities/names/descriptions, embedded and linked assets, declared MIME types, payload lengths, direct transforms/crop and group ancestry. Missing transforms are nil. Coordinates, extents and crop follow the shared read-only contract; no inheritance, composed group placement, pixel decoding or external fetching occurs.

All twelve shared cases across six IDs pass, including five successful input profiles and seven `PPTX_PICTURE_INVALID` refusals. Tests compare complete expected records, mutate returned data, inspect again, check caller-byte/member custody and save/reopen readback. Package/XML admission retains its own error categories. Successful empty input yields an empty non-nil collection.

The existing root/acceptance baseline and the full candidate-enabled `make test-batch` pass. An isolated local clone with the unchanged released gitlink passes the graphics-related presentation, packaging and lossless XML package batch. This inspection slice does not generate changed graphics for a new LibreOffice rendering check.

### PNG/JPEG insertion

`EditSession.AddPicture(slidePart, payload, PictureGeometry, PictureOptions)` appends a signature-checked PNG/JPEG with explicit signed positions and positive extents. It allocates collision-free shape/media/relationship identities, retains existing default content types and relationships, and appends before terminal extensions. A validated graph plan commits the slide, payload and registries together; refusals retain session generation and original members. Native Go field types exclude accessor/non-byte values; strict recipe decoding tests fractional and unknown-field refusals.

All eighteen shared cases pass, including twelve typed refusals and six successful variants. Additional controls cover oversized payloads, negative/out-of-range extents, invalid description text, default metadata and signed-package late-plan refusal. Tests verify complete receipt/geometry/media metadata, defensive copying, exact appended-node removal, existing relationships, package membership and save/reopen custody.

```sh
# Requires LibreOffice and system python3-uno; no external programs in runtime APIs.
GOMAXPROCS=2 GOFLAGS=-count=1 make graphics-insertion-quality
```

LibreOffice 24.2.7.2 loads, saves and reopens all six outputs without interaction requests. Inserted raster counts and names survive, with position/extent drift at most 0.01 mm. All six render as one-page PDFs and pass optional SDK structural validation. The pre-existing vector image in the source is converted into a shape by LibreOffice; this oracle counts imported raster objects separately and does not establish visual fidelity.

### Isolated picture replacement

`EditSession.ReplacePicture(slidePart, shapeID, payload, PictureReplacementOptions)` replaces only the selected expanded-name embed attribute through a lossless splice. It allocates fresh media and a fresh relationship and retains every previous dependency, so shared relationships/media and grouped pictures are isolated. Linked, dual and paired-extension assets refuse. Protection and inspection errors retain the shared semantic categories. A graph plan commits payload, registries and the selected attribute together.

All fourteen shared cases pass with literal XML restoration, old-edge/media custody, defensive bytes and save/reopen metadata. The full candidate batch and isolated related-package batch pass. `make graphics-replacement-quality` checks all six successful variants through LibreOffice load/save/reopen, with retained raster names/counts and geometry drift at most 0.01 mm. Their PDF conversions and optional SDK checks pass. Image fidelity and the converted pre-existing vector picture are outside the measured quality scope.

### Bounded crop and direct orientation

`GetPictureCrop`/`SetPictureCrop` read or set all four bounded source sides. `PatchPictureTransform` uses optional pointer fields for rotation and flips and never synthesises missing placement. Both support grouped/linked metadata without fetching, preserve unchanged lexical scalar spellings, return a changed count, and leave versions unchanged on no-op. Negative retained crop/rotation values remain inspectable but refuse bounded editing.

All seventeen crop and eighteen orientation recipe cases pass, including exact no-ops, unsupported grammar and typed refusals. Tests mask only the selected crop node or transform opening tag, compare remaining slide XML, preserve every unrelated member and compare saved/reopened records. Full root/acceptance and isolated related-package batches pass.

`make graphics-metadata-quality` imports, saves and reopens fourteen embedded successful variants in LibreOffice; linked outputs are excluded to avoid external fetching. Raster object counts/names, rotation, mirror/crop properties survive; position/extent drift is at most 0.01 mm. General visual fidelity is not measured.

### Explicit contain/cover/stretch placement

`CalculatePicturePlacement` is a pure native calculator; `AddFittedPicture` authors its geometry/crop and media in the insertion graph plan. It uses caller-supplied intrinsic integer dimensions, floor-scaled contain extents and floor-centred offsets, with exact cross-products and big-integer cover rounding. No DPI or pixel-size inference occurs.

All sixteen shared cases pass, including odd/fractional rounding, large rational inputs and overflow/empty-region refusals. Saved inspection retains the exact calculated integer geometry/crop. `make graphics-placement-quality` checks eight practical successful cases in LibreOffice: raster names/counts, orientation/mirror/crop properties survive save/reopen with geometry drift at most 0.01 mm. The native large-rational stress case is excluded from the application gate because its near-2^31 extents are not practical slide dimensions. Full root/acceptance and isolated related-package gates pass.

### Passive SVG with caller raster fallback

`AddSvgPicture` admits bounded passive UTF-8 SVG with an element/attribute/paint allowlist and a signature-checked PNG/JPEG supplied by the caller. It never rasterises or fetches references. One graph plan owns two new payloads, two image relationships and the picture's SVG extension/fallback reference. Both payload arrays are defensively copied; ordinary replacement refuses paired extensions.

All eleven shared cases pass with exact paired receipts, payload/graph custody, rejected executable/reference SVG and saved readback. The full root/acceptance and isolated related-package batches pass. `make graphics-svg-quality` checks both PNG/JPEG fallback variants in LibreOffice; picture names/counts persist through save/reopen with at most 0.01 mm geometry drift. This does not compare the visual content of SVG and fallback.

### Deletion and conservative collection

`DeletePicture` removes one exact picture span, retaining dependencies by default. Collection checks every remaining slide attribute and opaque lexical barrier before removing owner edges, then scans the complete package graph for incoming references before removing only PNG/JPEG/SVG media leaves. Connected picture endpoints refuse. One graph plan owns slide, relationship and content-type removals.

All seventeen shared cases pass, including shared relationships/media, other-owner references, grouped/link/SVG pairs and opaque reference/comment/text conservation. Tests assert exact removed-span XML, literal unrelated payloads, receipts, complete graph and save/reopen records. Full root/acceptance and isolated related-package batches pass. `make graphics-delete-quality` loads/saves/reopens ten unlinked success outputs in LibreOffice and checks target absence and retained imported picture names/counts. The remaining source vector image is converted into a shape; this measures deletion acceptance, not complete rendering equivalence.

### Editable shape groups

`GroupShapes` selects 2–100 contiguous direct shapes/pictures in source order, validates placeholder/lock/connector constraints and wraps their exact lexical slice with an identity parent/child rectangle. Wrapper namespace prefixes never shadow original bindings. Only the slide payload changes; no assets or relationships are copied.

All sixteen shared cases pass with exact receipts, identity-frame coordinates, literal child/outer XML reconstruction, typed refusals and saved picture ancestry. Full root/acceptance and isolated related-package batches pass. `make graphics-group-quality` checks four success variants through LibreOffice save/reopen: one group with two children, the child's text and raster identity survive. Raster geometry drift is at most 0.01 mm; general visual fidelity is not measured.

### Group transforms and point mapping

`GetGroupTransform`/`PatchGroupTransform` inspect and splice direct group frames while preserving child XML and scalar spelling. Pure `MapGroupPoint`/`UnmapGroupPoint` and outer-to-inner chain helpers support scale/translation, centre reflections and clockwise rotation; quarter turns use exact constants. Invalid/singular/unbounded frames and intermediates refuse.

All twenty-one shared cases pass. A separate 216-trial polar-reference matrix measures forward median 0, p95 3.64e-12, maximum 1.09e-11 EMUs; inverse median 0, p95 4.27e-14, maximum 1.71e-13 EMUs. Shared tolerances remain 1e-5 EMUs absolute plus relative 1e-12; serialized frames use exact integers.

`make graphics-group-transform-quality` checks six variants in LibreOffice, preserving one group and its picture child after export/reopen. Five have at most 0.01 mm geometry drift; the changed child-frame case has 0.02 mm position drift and 0.01 mm extent drift. Its parent/child coordinate composition accumulates two separately quantised 0.01 mm UNO positions. The application check therefore admits 0.02 mm for composed positions, without changing native mapping/serialization tolerances. The extreme contract probe projects its large source picture beyond practical slide bounds; visual fidelity is outside this oracle. Full root/acceptance and isolated related-package gates pass.

### Attached rectangle connectors

`AddConnector` appends an editable straight connector between distinct exact direct rectangle IDs/sites. It validates ordinary unrotated/unflipped geometry, locks/placeholders and bounded computed sites. Sites top/left/bottom/right use floor midpoints; direction controls explicit flip flags. Only the selected new connector span is appended, with exact line colour/width and fresh shape identity.

All fifteen shared cases pass with exact receipt/endpoints/geometry, attachment XML and full original-slide recovery. Full root/acceptance and isolated related-package gates pass. `make graphics-connectors-quality` imports/saves/reopens six variants in LibreOffice, retaining one attached connector and its endpoint names and glue-point indices. Existing raster geometry drift is at most 0.01 mm; routing/visual quality is not independently compared.

### Editable node-and-edge diagrams

`AddDiagram` validates 1–32 unique ordered keys/labels and up to 64 directed edges, computes bounded row/column placements and plans nodes then attached connectors before a single payload commit. Cycles/reverse edges are allowed; self/duplicate/missing edges refuse. Node XML remains editable through the existing native text and retained-style APIs.

All fourteen shared cases pass with exact receipts, node-label/attachment custody, original-slide reconstruction and save/reopen. Native text and fill edits succeed on each generated positive graph. Full root/acceptance and isolated related-package gates pass. `make graphics-diagrams-quality` checks four contract probes plus a practical-sized sample in LibreOffice, preserving three node labels and connector endpoints/site indices. The practical sample's input label is edited through UNO, then saved/reopened and exported to PDF. The source raster drift is at most 0.01 mm; routing/visual equivalence is outside the oracle.

### SmartArt dependency inspection

`InspectSmartArt` reports detached frame identities, four role roots, bounded transitive parts/edges and persisted drawing dependencies. It handles both owner-local synthetic drawing edges and the admitted source encoding's slide-owned metadata relationship. External targets remain literal data and are never fetched. No graph inventory implies SmartArt editing or layout evaluation.

All twelve synthetic shared cases and the sealed source inspection case pass with literal full records, detached returns, membership/payload custody and save/reopen. Full root/acceptance and isolated related-package gates pass. Existing synthetic graph schema limits remain unchanged. The sealed source is centrally licensed/provenance-addressed; no changed graphics are generated by inspection.

## Remaining slices

SmartArt copying, AutoShapes/freeforms/order and gradient/opacity/outline editing are not yet at Bun parity. Each slice needs native recipe assertions, atomic refusals, save/reopen custody and bounded application checks where it generates visible output. Preserve legacy APIs and unrelated XML/media.
