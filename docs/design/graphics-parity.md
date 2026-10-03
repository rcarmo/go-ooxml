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

## Remaining slices

Picture replacement/placement/crop/transforms, SVG fallback, deletion, groups/mapping, connectors/diagrams, SmartArt, AutoShapes/freeforms/order and gradient/opacity/outline editing are not yet at Bun parity. Each slice needs native recipe assertions, atomic refusals, save/reopen custody and bounded application checks where it generates visible output. Preserve legacy APIs and unrelated XML/media.
