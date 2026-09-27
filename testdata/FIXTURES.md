# Historical test fixture index

The original-main binary files listed below are no longer copied into `testdata/`.
Their bytes and historical paths are preserved in
[`spec/legacy-fixture-migration.json`](../spec/legacy-fixture-migration.json).
Native tests resolve fixture IDs through the pinned shared manifest; labels
such as `word/minimal.docx` are lookup keys, not local filesystem paths.
The shared checkout is read-only. No symlink or binary fallback is installed.

The original files were generated with python-docx, openpyxl and
python-pptx from public-domain *Frankenstein* text, with author `Test Author`.
These producer files are not substitutes for Microsoft Office-created
samples. The shared manifest retains origin and licence metadata.

| Historical path | Description | Pinned shared fixture | Asset ID |
| --- | --- | --- | --- |
| `testdata/default.docx` | Default Word package (historical fixture label) | [`fixtures/docx/creation/default-d9d6a313182a.docx`](../references/fixtures-ooxml/fixtures/docx/creation/default-d9d6a313182a.docx) | `fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5` |
| `testdata/default.pptx` | Default PowerPoint package (historical fixture label) | [`fixtures/pptx/creation/default-151d747bc37d.pptx`](../references/fixtures-ooxml/fixtures/pptx/creation/default-151d747bc37d.pptx) | `fixture-151d747bc37d4f4988c1116f4abb45196b1c1644319ce342ee4dd56d111f3132` |
| `testdata/excel/comments.xlsx` | Excel Cell comments | [`fixtures/xlsx/comments/comments-264be55e012d.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/comments/comments-264be55e012d.xlsx) | `fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f` |
| `testdata/excel/conditional_format.xlsx` | Excel Conditional formatting rules | [`fixtures/xlsx/conditional-formatting/conditional-format-7124469770b5.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/conditional-formatting/conditional-format-7124469770b5.xlsx) | `fixture-7124469770b5368c036f9dcc13d0c09c382797827d1b58934d9b2a0b61efee28` |
| `testdata/excel/data_types.xlsx` | Excel String, number, date, boolean, formula cells | [`fixtures/xlsx/cells/data-types-13bf5f08697e.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/cells/data-types-13bf5f08697e.xlsx) | `fixture-13bf5f08697e7de3da52b47d61e0f5af42081a4f1bdfbc1d51a3926910098b68` |
| `testdata/excel/formatting.xlsx` | Excel Colors, fonts, borders | [`fixtures/xlsx/formatting/formatting-dab1d6dbdfa5.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/formatting/formatting-dab1d6dbdfa5.xlsx) | `fixture-dab1d6dbdfa594680115bc87fc1707fbd06adea87b45ef0a1115d7cf4c4f72b5` |
| `testdata/excel/formulas.xlsx` | Excel Various formulas (SUM, VLOOKUP, etc.) | [`fixtures/xlsx/formulas/formulas-61d8806a5fd3.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/formulas/formulas-61d8806a5fd3.xlsx) | `fixture-61d8806a5fd3ff6c9eb62d716a6d929d0947e35ac6d9ac56626704690ecd3877` |
| `testdata/excel/merged_cells.xlsx` | Excel Merged cell regions | [`fixtures/xlsx/tables/merged-cells-691d1d459bd3.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/tables/merged-cells-691d1d459bd3.xlsx) | `fixture-691d1d459bd3d2ac3704b889bde121186b802381cf618c0a2daf12ca5b76ebec` |
| `testdata/excel/minimal.xlsx` | Excel Empty workbook with one sheet | [`fixtures/xlsx/creation/minimal-5140bb7ea228.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/creation/minimal-5140bb7ea228.xlsx) | `fixture-5140bb7ea22875bc86e3a9790b3dce686a8adc518f7c8597fbddd5d83264b59d` |
| `testdata/excel/multiple_sheets.xlsx` | Excel Three sheets with data | [`fixtures/xlsx/worksheets/multiple-sheets-a88d1e934bea.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/worksheets/multiple-sheets-a88d1e934bea.xlsx) | `fixture-a88d1e934bea9d3483325937f29a5d0880065bd54f95bbeceb612f7d4b5575f6` |
| `testdata/excel/named_ranges.xlsx` | Excel Named ranges defined | [`fixtures/xlsx/defined-names/named-ranges-cc8da5b59690.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/defined-names/named-ranges-cc8da5b59690.xlsx) | `fixture-cc8da5b5969032cbd2bdc6a40b26fc61c2366cf87e3d9455d4fdb8308ceb14fa` |
| `testdata/excel/single_cell.xlsx` | Excel One cell with a value | [`fixtures/xlsx/cells/single-cell-78318622a642.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/cells/single-cell-78318622a642.xlsx) | `fixture-78318622a642a298a0ec9626fc03a86719e0d9d2f170aa7e8d15db99ee29dbd9` |
| `testdata/excel/tables.xlsx` | Excel Excel tables (not just cell ranges) | [`fixtures/xlsx/tables/tables-51eaa2ee2fa0.xlsx`](../references/fixtures-ooxml/fixtures/xlsx/tables/tables-51eaa2ee2fa0.xlsx) | `fixture-51eaa2ee2fa09f384e5c177cebd62801d628f20724c57a221a25112cf374a8fb` |
| `testdata/pptx/bullet_points.pptx` | PowerPoint Slide with bullet points | [`fixtures/pptx/text/bullet-points-5c615f27a23c.pptx`](../references/fixtures-ooxml/fixtures/pptx/text/bullet-points-5c615f27a23c.pptx) | `fixture-5c615f27a23cb0d5fe3d603080ba057c7fc320f2af1df277b5b0dd128f1a8e3d` |
| `testdata/pptx/comments.pptx` | PowerPoint Slide with author metadata | [`fixtures/pptx/comments/comments-ababb5bed8ee.pptx`](../references/fixtures-ooxml/fixtures/pptx/comments/comments-ababb5bed8ee.pptx) | `fixture-ababb5bed8eeed4511f25e23da8e9039642ed3020b1517622d3ed71f5d5ef319` |
| `testdata/pptx/hidden_slides.pptx` | PowerPoint Mix of visible and hidden slides | [`fixtures/pptx/slides/hidden-slides-e01ded1106a2.pptx`](../references/fixtures-ooxml/fixtures/pptx/slides/hidden-slides-e01ded1106a2.pptx) | `fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c` |
| `testdata/pptx/image1.png` | Sample PowerPoint image | [`fixtures/png/media/sample-image.png`](../references/fixtures-ooxml/fixtures/png/media/sample-image.png) | `fixture-21e40e91bf948fd8b179c59915b5b840c0f8438ea7f2f144b4287cfd9b158e7f` |
| `testdata/pptx/images.pptx` | PowerPoint Embedded images (placeholder shapes) | [`fixtures/pptx/media/images-88f5cf211cbe.pptx`](../references/fixtures-ooxml/fixtures/pptx/media/images-88f5cf211cbe.pptx) | `fixture-88f5cf211cbedea688b081b2822c9207670e58a665afb1222506672fda3de2a8` |
| `testdata/pptx/layouts.pptx` | PowerPoint All standard layouts used | [`fixtures/pptx/layouts/layouts-ef5de48dec43.pptx`](../references/fixtures-ooxml/fixtures/pptx/layouts/layouts-ef5de48dec43.pptx) | `fixture-ef5de48dec43f165d513086506ba52c8f1b9ebbbc2fc99a6b4f56731f6115862` |
| `testdata/pptx/minimal.pptx` | PowerPoint Single blank slide | [`fixtures/pptx/creation/minimal-6a28461a0085.pptx`](../references/fixtures-ooxml/fixtures/pptx/creation/minimal-6a28461a0085.pptx) | `fixture-6a28461a00850a9297690a0a53ad8c109aea9e6f279df309639e73be0cc4c762` |
| `testdata/pptx/multiple_masters.pptx` | PowerPoint Multiple slide layouts used | [`fixtures/pptx/layouts/multiple-masters-e6bbd95968d9.pptx`](../references/fixtures-ooxml/fixtures/pptx/layouts/multiple-masters-e6bbd95968d9.pptx) | `fixture-e6bbd95968d9fd04bfc13cecc671a7438ab3c99987bc36b4fe118ef99241296f` |
| `testdata/pptx/notes.pptx` | PowerPoint Slides with speaker notes | [`fixtures/pptx/notes/notes-04faba67841d.pptx`](../references/fixtures-ooxml/fixtures/pptx/notes/notes-04faba67841d.pptx) | `fixture-04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc` |
| `testdata/pptx/shapes.pptx` | PowerPoint Various shape types (rectangles, arrows, etc.) | [`fixtures/pptx/shapes/shapes-2610748ab308.pptx`](../references/fixtures-ooxml/fixtures/pptx/shapes/shapes-2610748ab308.pptx) | `fixture-2610748ab308da10d61ebc6e5677f774427a207bdb474b74ebd3ae234152589e` |
| `testdata/pptx/tables.pptx` | PowerPoint Table on slide | [`fixtures/pptx/tables/tables-3465194945a0.pptx`](../references/fixtures-ooxml/fixtures/pptx/tables/tables-3465194945a0.pptx) | `fixture-3465194945a0f8084084ceb843c219d59a4ad92b25a271d66182587d637db0be` |
| `testdata/pptx/title_slide.pptx` | PowerPoint Title layout slide with content | [`fixtures/pptx/text/title-slide-836b5c7d7917.pptx`](../references/fixtures-ooxml/fixtures/pptx/text/title-slide-836b5c7d7917.pptx) | `fixture-836b5c7d7917a5e053ca994d96af0409d5f6650dd28b02dda9f878f8600c8656` |
| `testdata/word/bullet_list.docx` | Word Bullet list items | [`fixtures/docx/numbering/bullet-list-cb2c2610f99f.docx`](../references/fixtures-ooxml/fixtures/docx/numbering/bullet-list-cb2c2610f99f.docx) | `fixture-cb2c2610f99f786484b1941f25d068f6347cd0cd95213a57db849842d9d4a64b` |
| `testdata/word/comments.docx` | Word Multiple comments with replies | [`fixtures/docx/comments/comments-2029abbda3bd.docx`](../references/fixtures-ooxml/fixtures/docx/comments/comments-2029abbda3bd.docx) | `fixture-2029abbda3bdacb270137f6210791ba2ee3f8be62affb8cc3181f3ac72a7c35a` |
| `testdata/word/complex_table.docx` | Word Merged cells, nested tables | [`fixtures/docx/tables/complex-table-28da984e3fc4.docx`](../references/fixtures-ooxml/fixtures/docx/tables/complex-table-28da984e3fc4.docx) | `fixture-28da984e3fc4c6079579c0e8b3b9a901a46ed3f9081f3a074e0fb7333b984695` |
| `testdata/word/formatted_text.docx` | Word Bold, italic, underline, colors, font sizes | [`fixtures/docx/formatting/formatted-text-9a92eba3dc84.docx`](../references/fixtures-ooxml/fixtures/docx/formatting/formatted-text-9a92eba3dc84.docx) | `fixture-9a92eba3dc84f293a82de9496e571c780fd258cc3bf276b74464faaef5835dbc` |
| `testdata/word/headers_footers.docx` | Word Different first page, odd/even headers/footers | [`fixtures/docx/headers-footers/headers-footers-3bafa1552422.docx`](../references/fixtures-ooxml/fixtures/docx/headers-footers/headers-footers-3bafa1552422.docx) | `fixture-3bafa155242222dbd3529b56af6b8f1939cbd2e954de5b80afeadbbd53a432aa` |
| `testdata/word/headings.docx` | Word All heading levels 1-9 | [`fixtures/docx/styles/headings-8513f0537071.docx`](../references/fixtures-ooxml/fixtures/docx/styles/headings-8513f0537071.docx) | `fixture-8513f05370714f5e288ec1ca2fb76fe21b74b458636f325666b20b100c9b021a` |
| `testdata/word/minimal.docx` | Word Empty document with just body element | [`fixtures/docx/creation/minimal-9726b477472d.docx`](../references/fixtures-ooxml/fixtures/docx/creation/minimal-9726b477472d.docx) | `fixture-9726b477472ddb7595875c9f30493df2577e587416b18418d0dc7221046690fe` |
| `testdata/word/numbered_list.docx` | Word Numbered list items | [`fixtures/docx/numbering/numbered-list-37d3c408403d.docx`](../references/fixtures-ooxml/fixtures/docx/numbering/numbered-list-37d3c408403d.docx) | `fixture-37d3c408403dbecf4310f0a0b1313dc0c3756302e1778cc32d975823988ea3b2` |
| `testdata/word/sdt_content_controls.docx` | Word Content controls/placeholders | [`fixtures/docx/content-controls/sdt-content-controls-368fe96cb3ae.docx`](../references/fixtures-ooxml/fixtures/docx/content-controls/sdt-content-controls-368fe96cb3ae.docx) | `fixture-368fe96cb3ae55a0cc5fecbb599eda1d1058596d4914992f596083f291071ae4` |
| `testdata/word/simple_table.docx` | Word 3x3 table, no merged cells | [`fixtures/docx/tables/simple-table-87e3c67cb73b.docx`](../references/fixtures-ooxml/fixtures/docx/tables/simple-table-87e3c67cb73b.docx) | `fixture-87e3c67cb73bdbf5c8389791fd2bb459ba0eaac862149af6fd7c5b641dafd56b` |
| `testdata/word/single_paragraph.docx` | Word One paragraph, no formatting | [`fixtures/docx/text/single-paragraph-395b75b992c6.docx`](../references/fixtures-ooxml/fixtures/docx/text/single-paragraph-395b75b992c6.docx) | `fixture-395b75b992c6662e2bf44c72401020efccc212b41a345e082cb09abee450d97f` |
| `testdata/word/styles.docx` | Word Custom styles applied | [`fixtures/docx/styles/styles-d35e32ab35d4.docx`](../references/fixtures-ooxml/fixtures/docx/styles/styles-d35e32ab35d4.docx) | `fixture-d35e32ab35d4c95e81c15c5baac18b842696adcea960664e329a5a9a699ab7a1` |
| `testdata/word/track_changes.docx` | Word Document with insertions and deletions tracked | [`fixtures/docx/revisions/track-changes-2e1022f7358f.docx`](../references/fixtures-ooxml/fixtures/docx/revisions/track-changes-2e1022f7358f.docx) | `fixture-2e1022f7358f0deeaeb542a6dd10005f455c433a66f1b9b7e8964a48c7f0fb2a` |

The historical checklist recorded 13 Word, 11 Excel and 11 PowerPoint
scenario files, plus two root `default` fixtures and one PNG: **38 binary
files total**. Six proposed PowerPoint repair/relationship cases were unchecked
and have no retired local binary or asserted shared mapping:

- `frankenstein_notes_chart_rels.pptx` — notes, comments and chart relationship order
- `frankenstein_repair_prompt_repro.pptx` — PowerPoint repair-prompt reproducer
- `frankenstein_diagram_rels.pptx` — diagram relationships
- `frankenstein_notes_master_rids.pptx` — notes master relationship IDs
- `frankenstein_slide_rel_conflicts.pptx` — cross-slide relationship ID collisions
- `frankenstein_repaired_reference.pptx` — Office-repaired golden reference

The old checklist's "40 fixtures created" and PowerPoint "11/16" totals
were inconsistent with its entries: 13 Word + 11 Excel + 11 checked
PowerPoint = 35 checked scenario files, and six unchecked proposals would
make 41 entries, not 40. The 38 retired binaries also include two root
defaults and the PNG. Use the ledger and manifest rather than the old totals
as the physical inventory.
