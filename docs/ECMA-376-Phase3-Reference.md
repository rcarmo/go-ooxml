# ECMA-376 WordprocessingML - Phase 3 Quick Reference

This document provides essential ECMA-376 specifications for Phase 3 features:
- Track Changes (Revisions)
- Comments
- Styles
- Headers and Footers

## 1. Track Changes (Revisions) - §17.13.5

### 1.1 Enabling Track Changes

Track changes must be enabled in document settings (`word/settings.xml`):

```xml
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:trackRevisions w:val="true"/>
</w:settings>
```

### 1.2 Inserted Content - `<w:ins>` (§17.13.5.20)

Marks inline-level content that has been **inserted** while tracking revisions.

**XML Structure:**
```xml
<w:ins w:id="0" w:author="Joe Smith" w:date="2006-03-31T12:50:00Z">
    <w:r>
        <w:t>inserted text</w:t>
    </w:r>
</w:ins>
```

**Attributes:**
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `w:id` | ST_DecimalNumber | Yes | Unique identifier for the revision |
| `w:author` | ST_String | No | Author who made the change |
| `w:date` | ST_DateTime | No | Date/time of the change (ISO 8601) |

**Parent Elements:** `<w:p>`, `<w:body>`, `<w:tc>`, `<w:hdr>`, `<w:ftr>`, etc.

**Child Elements:** `<w:r>`, `<w:del>`, `<w:ins>` (nested), bookmarks, comments, etc.

### 1.3 Deleted Content - `<w:del>` (§17.13.5.12)

Marks inline-level content that has been **deleted** while tracking revisions.

**XML Structure:**
```xml
<w:del w:id="1" w:author="Jane Doe" w:date="2006-04-01T09:30:00Z">
    <w:r>
        <w:delText>deleted text</w:delText>
    </w:r>
</w:del>
```

**Key Difference:** Deleted text uses `<w:delText>` instead of `<w:t>`.

**Attributes:** Same as `<w:ins>`

### 1.4 Property Changes

| Element | Description |
|---------|-------------|
| `<w:rPrChange>` | Run (character) properties change |
| `<w:pPrChange>` | Paragraph properties change |
| `<w:sectPrChange>` | Section properties change |
| `<w:tblPrChange>` | Table properties change |
| `<w:trPrChange>` | Table row properties change |
| `<w:tcPrChange>` | Table cell properties change |

**Example (formatting change):**
```xml
<w:rPr>
    <w:b/>
    <w:rPrChange w:id="2" w:author="Editor" w:date="2006-04-02T14:00:00Z">
        <w:rPr>
            <!-- Previous properties (before bold was added) -->
        </w:rPr>
    </w:rPrChange>
</w:rPr>
```

### 1.5 Move Operations

| Element | Description |
|---------|-------------|
| `<w:moveFrom>` | Content moved FROM this location |
| `<w:moveTo>` | Content moved TO this location |
| `<w:moveFromRangeStart>` | Start of move source range |
| `<w:moveFromRangeEnd>` | End of move source range |
| `<w:moveToRangeStart>` | Start of move destination range |
| `<w:moveToRangeEnd>` | End of move destination range |

---

## 2. Comments - §17.13.4

### 2.1 Comments Part Structure

Comments are stored in a separate part: `/word/comments.xml`

**Root Element:**
```xml
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:comment w:id="0" w:author="John Doe" w:date="2006-04-06T13:50:00Z" w:initials="JD">
        <w:p>
            <w:r>
                <w:t>This is a comment.</w:t>
            </w:r>
        </w:p>
    </w:comment>
</w:comments>
```

### 2.2 Comment Element - `<w:comment>` (§17.13.4.2)

**Attributes:**
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `w:id` | ST_DecimalNumber | Yes | Unique identifier |
| `w:author` | ST_String | No | Comment author's name |
| `w:date` | ST_DateTime | No | Date/time created (ISO 8601) |
| `w:initials` | ST_String | No | Author's initials (for display) |

**Child Elements:** Block-level content - `<w:p>`, `<w:tbl>`, etc.

### 2.3 Comment Anchoring in Document

Comments are anchored to document text using three elements:

```xml
<w:p>
    <w:r><w:t>Some </w:t></w:r>
    <w:commentRangeStart w:id="0"/>
    <w:r><w:t>commented text</w:t></w:r>
    <w:commentRangeEnd w:id="0"/>
    <w:r>
        <w:commentReference w:id="0"/>
    </w:r>
</w:p>
```

| Element | Description |
|---------|-------------|
| `<w:commentRangeStart>` | Marks start of commented text |
| `<w:commentRangeEnd>` | Marks end of commented text |
| `<w:commentReference>` | Reference marker (where comment bubble appears) |

**Note:** The `w:id` must match between the anchor elements and the `<w:comment>` in comments.xml.

### 2.4 Relationship

The comments part must be declared in `/word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId5" 
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="comments.xml"/>
```

---

## 3. Styles - §17.7

### 3.1 Styles Part Structure

Styles are stored in: `/word/styles.xml`

```xml
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>...</w:docDefaults>
    <w:latentStyles>...</w:latentStyles>
    <w:style w:type="paragraph" w:styleId="Normal">...</w:style>
    <w:style w:type="paragraph" w:styleId="Heading1">...</w:style>
    <!-- more styles -->
</w:styles>
```

### 3.2 Style Types - ST_StyleType (§17.18.83)

| Type | Description |
|------|-------------|
| `paragraph` | Paragraph formatting (includes runs) |
| `character` | Character/run formatting only |
| `table` | Table formatting |
| `numbering` | List/numbering formatting |

### 3.3 Style Element - `<w:style>` (§17.7.4.17)

**Example:**
```xml
<w:style w:type="paragraph" w:styleId="Heading1" w:default="0">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
        <w:keepNext/>
        <w:spacing w:before="240" w:after="60"/>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        <w:b/>
        <w:sz w:val="32"/>
    </w:rPr>
</w:style>
```

**Attributes:**
| Attribute | Description |
|-----------|-------------|
| `w:type` | Style type (paragraph, character, table, numbering) |
| `w:styleId` | Unique identifier for referencing |
| `w:default` | Whether this is the default style for its type |
| `w:customStyle` | Whether this is a user-defined custom style |

**Key Child Elements:**
| Element | Description |
|---------|-------------|
| `<w:name>` | Display name of the style |
| `<w:basedOn>` | Parent style ID (for inheritance) |
| `<w:next>` | Style to apply after pressing Enter |
| `<w:link>` | Linked character style (for paragraph styles) |
| `<w:uiPriority>` | Sort order in UI |
| `<w:qFormat>` | Show in Quick Style gallery |
| `<w:pPr>` | Paragraph properties |
| `<w:rPr>` | Run (character) properties |
| `<w:tblPr>` | Table properties (for table styles) |

### 3.4 Style Inheritance

Styles inherit from their `<w:basedOn>` parent:
- Properties not specified are inherited from parent
- Specified properties override parent values
- Chain: Style → basedOn → ... → docDefaults

### 3.5 Applying Styles

**Paragraph style:**
```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading1"/>
    </w:pPr>
    ...
</w:p>
```

**Character style:**
```xml
<w:r>
    <w:rPr>
        <w:rStyle w:val="Strong"/>
    </w:rPr>
    ...
</w:r>
```

**Table style:**
```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
    </w:tblPr>
    ...
</w:tbl>
```

---

## 4. Headers and Footers - §17.10

### 4.1 Header/Footer Parts

Headers and footers are stored as separate parts:
- `/word/header1.xml`, `/word/header2.xml`, etc.
- `/word/footer1.xml`, `/word/footer2.xml`, etc.

### 4.2 Header Element - `<w:hdr>` (§17.10.4)

```xml
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:r>
            <w:t>Header text</w:t>
        </w:r>
    </w:p>
</w:hdr>
```

### 4.3 Footer Element - `<w:ftr>` (§17.10.3)

```xml
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:r>
            <w:t>Footer text</w:t>
        </w:r>
    </w:p>
</w:ftr>
```

### 4.4 Header/Footer References in Section Properties

Headers and footers are linked via `<w:sectPr>`:

```xml
<w:sectPr>
    <w:headerReference w:type="default" r:id="rId6"/>
    <w:headerReference w:type="first" r:id="rId7"/>
    <w:headerReference w:type="even" r:id="rId8"/>
    <w:footerReference w:type="default" r:id="rId9"/>
    <w:footerReference w:type="first" r:id="rId10"/>
    <w:footerReference w:type="even" r:id="rId11"/>
    ...
</w:sectPr>
```

### 4.5 Header/Footer Types - ST_HdrFtr (§17.18.36)

| Type | Description |
|------|-------------|
| `default` | Default header/footer (odd pages if different) |
| `first` | First page only |
| `even` | Even pages only |

### 4.6 Relationships

Each header/footer must be declared in `/word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
<Relationship Id="rId9"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Target="footer1.xml"/>
```

### 4.7 Title Page Settings

To enable different first page:
```xml
<w:sectPr>
    <w:titlePg/>
    ...
</w:sectPr>
```

---

## 5. Content Types

### 5.1 Required Content Types

```xml
<!-- [Content_Types].xml -->
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <!-- Comments -->
    <Override PartName="/word/comments.xml" 
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
    
    <!-- Styles -->
    <Override PartName="/word/styles.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    
    <!-- Headers -->
    <Override PartName="/word/header1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
    
    <!-- Footers -->
    <Override PartName="/word/footer1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>
```

---

## 6. Relationship Types

| Feature | Relationship Type |
|---------|-------------------|
| Comments | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| Styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| Header | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/header` |
| Footer | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer` |
| Settings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings` |

---

## 7. Validation Rules

### 7.1 Track Changes
- `w:id` must be unique within the document
- `w:date` must be valid ISO 8601 format
- Nested `<w:ins>` inside `<w:del>` is allowed (and vice versa)

### 7.2 Comments
- `w:id` must match between anchor elements and comment
- Orphaned comments (no matching reference) may be ignored
- Duplicate IDs: only one comment loaded

### 7.3 Styles
- `w:styleId` must be unique within styles.xml
- `w:basedOn` must reference an existing style
- Circular inheritance is not allowed

### 7.4 Headers/Footers
- Each type can only have one reference per section
- Referenced parts must exist in package
- Relationships must be correctly declared
