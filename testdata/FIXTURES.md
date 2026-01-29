# Test Fixtures Checklist

These fixtures must be created in **actual Microsoft Office applications** (not programmatically generated) to ensure realistic testing. OOXML files created by Office have subtle differences from programmatically generated ones.

## Word Documents (`testdata/word/`)

- [ ] `minimal.docx` - Empty document with just body element
- [ ] `single_paragraph.docx` - One paragraph, no formatting
- [ ] `formatted_text.docx` - Bold, italic, underline, colors, font sizes
- [ ] `headings.docx` - All heading levels 1-9
- [ ] `simple_table.docx` - 3x3 table, no merged cells
- [ ] `complex_table.docx` - Merged cells, nested tables
- [ ] `track_changes.docx` - Document with insertions and deletions tracked
- [ ] `comments.docx` - Multiple comments with replies
- [ ] `styles.docx` - Custom styles applied
- [ ] `headers_footers.docx` - Different first page, odd/even headers/footers
- [ ] `sdt_content_controls.docx` - Content controls/placeholders
- [ ] `numbered_list.docx` - Numbered list items
- [ ] `bullet_list.docx` - Bullet list items

## Excel Workbooks (`testdata/excel/`)

- [ ] `minimal.xlsx` - Empty workbook with one sheet
- [ ] `single_cell.xlsx` - One cell with a value
- [ ] `data_types.xlsx` - String, number, date, boolean, formula cells
- [ ] `formatting.xlsx` - Colors, fonts, borders
- [ ] `multiple_sheets.xlsx` - Three sheets with data
- [ ] `tables.xlsx` - Excel tables (not just cell ranges)
- [ ] `merged_cells.xlsx` - Merged cell regions
- [ ] `named_ranges.xlsx` - Named ranges defined
- [ ] `comments.xlsx` - Cell comments
- [ ] `formulas.xlsx` - Various formulas (SUM, VLOOKUP, etc.)
- [ ] `conditional_format.xlsx` - Conditional formatting rules

## PowerPoint Presentations (`testdata/pptx/`)

- [ ] `minimal.pptx` - Single blank slide
- [ ] `title_slide.pptx` - Title layout slide with content
- [ ] `bullet_points.pptx` - Slide with bullet points
- [ ] `shapes.pptx` - Various shape types (rectangles, arrows, etc.)
- [ ] `tables.pptx` - Table on slide
- [ ] `images.pptx` - Embedded images
- [ ] `notes.pptx` - Slides with speaker notes
- [ ] `comments.pptx` - Slide comments
- [ ] `hidden_slides.pptx` - Mix of visible and hidden slides
- [ ] `multiple_masters.pptx` - Multiple slide masters
- [ ] `layouts.pptx` - All standard layouts used

---

## Fixture Creation Guidelines

1. **Use Microsoft Office** (Word, Excel, PowerPoint) - LibreOffice/Google Docs produce different XML
2. **Save as .docx/.xlsx/.pptx** (not .doc/.xls/.ppt)
3. **Keep files minimal** - Include only what's needed to test the specific feature
4. **Use Office 2016 or later** for best compatibility
5. **Don't password protect** the files

## Total: 35 fixtures needed

| Format | Count | Status |
|--------|-------|--------|
| Word   | 13    | 0/13   |
| Excel  | 11    | 0/11   |
| PowerPoint | 11 | 0/11  |
