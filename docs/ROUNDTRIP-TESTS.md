# Round-Trip Fixture Tests

This document summarizes the complex round-trip fixture tests for Word, Excel, and PowerPoint. Each test opens a fixture, mutates it, saves, re-opens, and verifies expected behavior.

## Summary Table

| Area | Fixture | Mutations | Verifications |
| --- | --- | --- | --- |
| Document | minimal.docx | track changes + paragraph + run bold + tracked text + block content control + header/footer + table + comment + field | header/footer text, content control, comment, revisions, paragraph presence, table cell |
| Document | single_paragraph.docx | paragraph + hyperlink + bookmark + comment | paragraph text, hyperlink URL, comment |
| Document | formatted_text.docx | formatted run + paragraph toggles (keep lines/page break/widow) | run effects + paragraph toggles |
| Document | headings.docx | heading paragraph style | paragraph style |
| Document | simple_table.docx | add row to fixture table | table cell text |
| Document | complex_table.docx | new table + shaded cell | table cell text |
| Document | track_changes.docx | enable track changes + tracked insert | revisions present |
| Document | comments.docx | add comment | comment present |
| Document | styles.docx | add paragraph style + apply | style exists + paragraph style |
| Document | headers_footers.docx | set header/footer text | header/footer text |
| Document | sdt_content_controls.docx | add content control + date config | content control text |
| Document | numbered_list.docx | add numbered list style + list item | list level preserved |
| Document | bullet_list.docx | add list item with level | list level preserved |
| Spreadsheet | minimal.xlsx | values + formula + merge + comment + style + named range | value, formula, merged range, comment, named range, style present |
| Spreadsheet | single_cell.xlsx | values + formula + number formats (currency/percent/negative) | value, formula, number formats |
| Spreadsheet | data_types.xlsx | string/int/bool/date values | value accessors |
| Spreadsheet | formatting.xlsx | styled cell (bold/fill/border) | style present |
| Spreadsheet | multiple_sheets.xlsx | add hidden sheet + named range | sheet hidden + named range |
| Spreadsheet | tables.xlsx | add table + header + rows + delete/update + number format | table headers, rows, values, number format |
| Spreadsheet | merged_cells.xlsx | merge + value | merged range |
| Spreadsheet | named_ranges.xlsx | add named range | named range present |
| Spreadsheet | comments.xlsx | add comment | comment present |
| Spreadsheet | formulas.xlsx | add formula | formula preserved |
| Spreadsheet | conditional_format.xlsx | value edit | conditional formatting present |
| Presentation | minimal.pptx | add textbox + notes | notes text + textbox text |
| Presentation | title_slide.pptx | add textbox title | title text |
| Presentation | bullet_points.pptx | add bullet paragraph | bullet text |
| Presentation | shapes.pptx | add rectangle + fill/line | shapes present |
| Presentation | tables.pptx | add table + cell text | table present |
| Presentation | notes.pptx | set notes | notes text |
| Presentation | comments.pptx | add comment | comments present |
| Presentation | hidden_slides.pptx | add hidden slide | hidden slide present |
| Presentation | multiple_masters.pptx | add slide | masters + layouts present |
| Presentation | layouts.pptx | add slide | layouts present |
| Presentation | images.pptx | add textbox placeholder | text present |
