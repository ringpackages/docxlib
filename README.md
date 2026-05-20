# DOCXLib

**Word Document Generation and Reading Library for the Ring Programming Language**

---

## Table of Contents

**Writer**

1. [Introduction](#1-introduction)
2. [Document Setup](#2-document-setup)
3. [Headers, Footers, and Page Numbers](#3-headers-footers-and-page-numbers)
4. [Text and Paragraphs](#4-text-and-paragraphs)
5. [Lists](#5-lists)
6. [Tables](#6-tables)
7. [Images](#7-images)
8. [Text Watermarks](#8-text-watermarks)
9. [Page Borders and Background](#9-page-borders-and-background)
10. [Footnotes and Endnotes](#10-footnotes-and-endnotes)
11. [Hyperlinks, Bookmarks, and Cross-References](#11-hyperlinks-bookmarks-and-cross-references)
12. [Table of Contents](#12-table-of-contents)
13. [Captions and Lists of Figures / Tables](#13-captions-and-lists-of-figures--tables)
14. [Tab Stops and Line Numbers](#14-tab-stops-and-line-numbers)
15. [Comments](#15-comments)
16. [Mail Merge Fields](#16-mail-merge-fields)
17. [Content Controls](#17-content-controls)
18. [Drawing Shapes](#18-drawing-shapes)
19. [Text Boxes](#19-text-boxes)
20. [RTL Support](#20-rtl-support)
21. [Saving the Document](#21-saving-the-document)
22. [Native OOXML Charts](#22-native-ooxml-charts)
23. [Complete Example](#23-complete-example)
24. [Mail Merge Template Engine](#24-mail-merge-template-engine)
25. [Chart Data Tables](#25-chart-data-tables)
26. [Version History](#26-version-history)

**Reader**

27. [Introduction to WordReader](#27-introduction-to-wordreader)
28. [Loading and Saving](#28-loading-and-saving)
29. [Document Properties](#29-document-properties)
30. [Content Query Methods](#30-content-query-methods)
31. [Table Query Methods](#31-table-query-methods)
32. [Image Query Methods](#32-image-query-methods)
33. [Section and Layout Queries](#33-section-and-layout-queries)
34. [Round-Trip: toWriter()](#34-round-trip-towriter)
35. [Reader Quick Reference](#35-reader-quick-reference)

**Appendix**

36. [Quick Reference Index](#36-quick-reference-index)

---

## 1. Introduction

DOCXLib is a **pure-Ring** library for generating and reading Microsoft Word (`.docx`)
documents programmatically — no external tools, no COM automation, no Office installation
required. 

### 1.1 Key Features

| Category | Capability |
|---|---|
| Zero dependencies | Pure Ring — no Python, Java, COM, or Office needed |
| Fluent API | Every method returns `self`, enabling method chaining |
| OOXML compliant | Files open without repair warnings in Word 2016–365 |
| Rich tables | Cell merging, per-side borders, row heights, text direction, images, lists inside cells |
| Full typography | Bold/italic/colour/font/size per run, language tags, character styles, custom paragraph styles |
| Paragraph control | Widow/orphan control, hyphenation suppression, keep-together, contextual spacing |
| Navigation | TOC, Table of Figures, Table of Tables, bookmarks, cross-refs |
| Forms | Checkboxes, dropdowns, and text input content controls |
| Shapes | Rectangle, ellipse, diamond, triangle, line (DrawingML) |
| Themes | 8 built-in colour palettes applied to all heading levels |
| Charts | 8 native OOXML chart types with axis formatting, data tables, and style presets |
| Mail Merge | Template engine: `setMergeTemplate()`, `mergeRecord()`, `mergeAll()` with `{{FIELD}}` tokens |
| Image cropping | Non-destructive `a:srcRect` crop by percentage on inline and floating images |
| Rich notes | Footnote and endnote bodies with full per-run formatting (bold, italic, colour, size) |
| WordReader | Parse any `.docx` — extract all content and styles; full round-trip reconstruct and save |

### 1.2 Requirements

Ring programming language (any recent version). No additional packages required.

### 1.3 Installation

	ringpm install docxlib from ringpackages

### 1.4 Quick Start

```ring
load "docxlib.ring"

doc = new WordWriter()
doc.setTitle("My First Document")
doc.setAuthor("Your Name")
doc.setPageSize("a4")

doc.addHeading("Hello, DOCXLib!", 1)
doc.addParagraph("This document was created entirely from Ring code.", NULL)
doc.addBulletList(["Feature A", "Feature B", "Feature C"])

if doc.save("hello.docx")
    ? "Saved successfully"
ok
```

---

## 2. Document Setup

Setup methods configure global document properties and should be called **before** adding
content.

### 2.1 Document Metadata

| Method | Parameters | Description |
|---|---|---|
| `setTitle()` | `title : String` | Sets the document title (written to core.xml) |
| `setAuthor()` | `author : String` | Sets the document author (written to core.xml) |
| `setDefaultFont()` | `fontName, fontSize` | Body font and size in pt. Defaults: Calibri, 11pt |

### 2.2 Page Size and Orientation

| Method | Parameters | Description |
|---|---|---|
| `setPageSize()` | `size : String` | `"letter"`, `"a4"`, `"legal"`, `"a3"`, `"b5"`. Default: `letter` |
| `setCustomPageSize()` | `widthCm, heightCm` | Exact page dimensions in centimetres |
| `setOrientation()` | `orientation` | `"portrait"` or `"landscape"` |
| `setMargins()` | `top, bottom, left, right` | Page margins in centimetres |
| `setMarginsInches()` | `top, bottom, left, right` | Page margins in inches |
| `setNarrowMargins()` | — | 0.5" margins on all four sides |
| `setWideMargins()` | — | 1.5" margins on all four sides |

```ring
doc.setPageSize("a4")
doc.setMargins(2.54, 2.54, 2.54, 2.54)   # 1 inch all sides
```

### 2.3 Multi-Column Layouts

| Method | Parameters | Description |
|---|---|---|
| `setColumns()` | `numColumns, spaceCm` | Set column count and gap in cm |
| `setTwoColumns()` | — | Two equal columns, 0.5 cm gap |
| `setThreeColumns()` | — | Three equal columns, 0.5 cm gap |
| `addColumnBreak()` | — | Insert a column break at the current position |
| `addSectionBreak()` | `breakType, options` | Insert a section break with optional layout |
| `addLandscapeStart()` | — | Begin a landscape section |
| `addLandscapeEnd()` | — | End the landscape section |

`addSectionBreak()` break types: `"continuous"`, `"nextPage"`, `"evenPage"`, `"oddPage"`.  
The second argument is an **options list** — pass `[]` when no options are needed:

| Option | Type | Description |
|---|---|---|
| `:numColumns` | Number | Column count for the new section (`2` or more enables multi-column) |
| `:columnSpaceCm` | Number | Gap between columns in cm |

```ring
# Two-column newsletter — begin on same page
doc.addSectionBreak("continuous", [:numColumns=2, :columnSpaceCm=0.6])
doc.addParagraph("Left column content...", [])
doc.addColumnBreak()
doc.addParagraph("Right column content...", [])

# Return to single-column on the next page
doc.addSectionBreak("nextPage", [])

# Landscape page within a portrait document
doc.addLandscapeStart()
doc.addTable(wideData, options)
doc.addLandscapeEnd()
```

### 2.4 Document Themes

`setTheme()` applies a colour palette to all heading levels and writes `word/theme/theme1.xml`.
Call it before `save()`.

```ring
doc.setTheme("Blue")
```

| Theme | Accent 1 (H1/H2) | Accent 2 (H3/H4) |
|---|---|---|
| `Office` | #4472C4 Blue | #ED7D31 Orange |
| `Blue` | #1E6BB5 Deep Blue | #2E75B6 Mid Blue |
| `Dark` | #88C0D0 Ice Blue | #81A1C1 Steel Blue |
| `Green` | #375623 Forest | #70AD47 Lime |
| `Red` | #C00000 Dark Red | #FF0000 Red |
| `Purple` | #7030A0 Purple | #9B59B6 Violet |
| `Teal` | #00796B Teal | #009688 Medium Teal |
| `Warm` | #E65100 Deep Orange | #F57C00 Orange |

### 2.5 Custom Paragraph Styles

`defineStyle()` registers a named paragraph style. Use the style name via `:style` in
`addParagraph()`.

```ring
doc.defineStyle("CalloutBox", [
    :basedOn     = "Normal",
    :bold        = true,
    :size        = 11,
    :color       = "1E6BB5",
    :bgColor     = "DEEAF1",
    :align       = "center",
    :spaceBefore = 120,
    :spaceAfter  = 120
])

doc.addParagraph("This uses the custom style.", [:style = "CalloutBox"])
```

| `defineStyle` option | Type | Description |
|---|---|---|
| `:basedOn` | String | Parent style ID, e.g. `"Normal"`, `"Heading1"` |
| `:bold` | true/false | Bold text |
| `:italic` | true/false | Italic text |
| `:underline` | true/false | Underlined text |
| `:font` | String | Font family name |
| `:size` | Number | Font size in points |
| `:color` | String | Text colour hex without `#` |
| `:align` | String | `"left"`, `"center"`, `"right"`, `"both"` |
| `:bgColor` | String | Paragraph background fill colour hex |
| `:spaceBefore` | Number | Space before in twips (20 twips = 1 pt) |
| `:spaceAfter` | Number | Space after in twips |
| `:lineSpacing` | Number | Line spacing multiplier: 1.5, 2.0, etc. |
| `:indent` | Number | Left indent in twips |
| `:keepNext` | true/false | Keep with the next paragraph |
| `:keepLines` | true/false | Prevent splitting across pages |

---

## 3. Headers, Footers, and Page Numbers

| Method | Parameters | Description |
|---|---|---|
| `setHeader()` | `text` | Default header text (centred) |
| `setFooter()` | `text` | Default footer text |
| `setHeaderFooter()` | `headerText, footerText` | Set both in one call |
| `showPageNumbers()` | `align` | Add page numbers: `"left"`, `"center"`, `"right"` |
| `setFirstPageDifferent()` | `bEnable` | Enable a different header/footer on page 1 |
| `setFirstPageHeader()` | `text` | Header text for the first page only |
| `setFirstPageFooter()` | `text` | Footer text for the first page only |
| `setEvenAndOddHeaders()` | `bEnable` | Enable separate headers for even/odd pages |
| `setEvenPageHeader()` | `text` | Header text for even-numbered (left) pages |
| `setEvenPageFooter()` | `text` | Footer text for even-numbered pages |

```ring
# Standard header with page numbers
doc.setHeader("Annual Report 2025")
doc.showPageNumbers("right")

# Cover page — no header, subsequent pages with header
doc.setFirstPageDifferent(true)
doc.setFirstPageHeader("")          # blank on cover
doc.setHeader("Chapter 1 — Introduction")

# Book-style even/odd headers
doc.setEvenAndOddHeaders(true)
doc.setHeader("Ring Language Study")           # odd pages (right side)
doc.setEvenPageHeader("Mahmoud Fayed, 2025")   # even pages (left side)
```

---

## 4. Text and Paragraphs

### 4.1 `addParagraph()`

The primary content method. Pass `[]` or `NULL` as the second argument to use all defaults.

```ring
doc.addParagraph("Plain text paragraph.", [])
doc.addParagraph("Bold red heading", [:bold=true, :color="C00000", :size=14])
doc.addParagraph("Centred note",     [:align="center", :italic=true])
doc.addParagraph("Custom style",     [:style="CalloutBox"])
```

| Option | Type | Description |
|---|---|---|
| `:bold` | true/false | Bold text |
| `:italic` | true/false | Italic text |
| `:underline` | true/false | Underlined text |
| `:size` | Number | Font size in points |
| `:font` | String | Font family name |
| `:color` | String | Text colour hex, e.g. `"C00000"` |
| `:charStyle` | String | Named character style, e.g. `"Strong"`, `"Emphasis"` |
| `:lang` | String | BCP 47 language tag, e.g. `"ar-SA"`, `"fr-FR"`, `"de-DE"` |
| `:align` | String | `"left"`, `"center"`, `"right"`, `"both"` |
| `:style` | String | Named paragraph style |
| `:bgColor` | String | Paragraph background colour hex |
| `:spaceBefore` | Number | Space before in twips (20 twips = 1 pt) |
| `:spaceAfter` | Number | Space after in twips |
| `:lineSpacing` | Number | Line spacing multiplier, e.g. `1.5` |
| `:indent` | Number | Left indent in twips |
| `:keepNext` | true/false | Keep this paragraph on the same page as the next |
| `:keepLines` | true/false | Prevent the paragraph from splitting across pages |
| `:suppressLineNumbers` | true/false | Exclude from line numbering |
| `:contextualSpacing` | true/false | Suppress spacing between same-style paragraphs |
| `:widowControl` | false | Set `false` to disable widow/orphan protection (Word default is on) |
| `:noHyphenate` | true | Suppress automatic hyphenation on this paragraph |

### 4.2 Headings and Titles

```ring
doc.addHeading("Chapter 1 — Introduction", 1)   # Heading 1
doc.addHeading("1.1 Background", 2)              # Heading 2
doc.addHeading("Sub-section detail", 3)          # Heading 3  (levels 1–6)
doc.addTitle("Document Title")                   # Title style
doc.addSubtitle("A subtitle line")               # Subtitle style
```

### 4.3 `addRichParagraph()` — Mixed Formatting

Build a paragraph from an explicit list of `[text, options]` run pairs. Each run can
carry its own bold, italic, colour, size, font, language, and character style.

```ring
doc.addRichParagraph([
    ["Status: ",  []],
    ["Passed",    [:bold=true, :color="375623"]],
    ["  —  Score: ", []],
    ["98/100",    [:bold=true, :color="1E6BB5", :size=14]]
], [:align="center"])

# Language-tagged runs (Arabic + English in one paragraph)
doc.addRichParagraph([
    ["مرحبا",    [:font="Arial", :lang="ar-SA"]],
    [" — Hello", [:lang="en-US"]]
], [])

# Character styles
doc.addRichParagraph([
    ["See ",         []],
    ["ring-lang.net", [:charStyle="Hyperlink"]]
], [])
```

Run-level option keys: `:text`, `:bold`, `:italic`, `:underline`, `:color`, `:size`,
`:font`, `:bgColor`, `:charStyle`, `:lang`.

### 4.4 Shorthand Methods

| Method | Description |
|---|---|
| `bold(text)` | Bold paragraph |
| `italic(text)` | Italic paragraph |
| `underline(text)` | Underlined paragraph |
| `colored(text, color)` | Coloured text (hex colour) |
| `styled(text, font, size, color)` | Single-call full style |
| `centered(text)` | Centre-aligned paragraph |
| `rightAligned(text)` | Right-aligned paragraph |
| `justified(text)` | Justified paragraph |
| `addBlockQuote(text)` | Indented block-quote paragraph |
| `addCaption(text)` | Small centred italic caption |
| `addAbstract(text)` | Indented abstract block with bold "Abstract" heading |
| `addKeywords(text)` | Keywords line with bold "Keywords:" prefix |
| `addEmptyParagraph()` | One blank line |
| `addHorizontalLine()` | Full-width horizontal rule |
| `addPageBreak()` | Hard page break |
| `addLineBreak()` | Soft line break within a paragraph |

### 4.5 Shaded and Bordered Paragraphs

```ring
# Solid background colour
doc.addShadedParagraph("Important note", "FFF3CD", [:bold=true])

# Full border box with optional background
doc.addBorderedParagraph("Critical warning", [
    :borderStyle = "single",
    :borderColor = "C00000",
    :bgColor     = "FDECEA",
    :spaceBefore = 80,
    :spaceAfter  = 80,
    :indent      = 240
])
```

### 4.6 Widow/Orphan Control and Auto-Hyphenation

Word protects against widows (last line stranded on the next page) and orphans (first
line stranded at the bottom) by default. You can also suppress automatic hyphenation
per paragraph — useful for technical terms, URLs, and headings.

```ring
# Disable widow/orphan protection for a data-dense paragraph
doc.addParagraph(denseText, [:align="both", :widowControl=false])

# Suppress auto-hyphenation (useful for URLs, product codes, headings)
doc.addParagraph(longTechnicalTerm, [:noHyphenate=true])

# Both together — dense justified paragraph that must not be hyphenated
doc.addParagraph(body, [:align="both", :widowControl=false, :noHyphenate=true])
```

---

## 5. Lists

DOCXLib generates proper Word OOXML list numbering — not manually inserted bullet
characters.

```ring
# Simple bullet list
doc.addBulletList(["Ring language", "Python", "JavaScript"])

# Numbered list
doc.addNumberedList(["Install Ring", "Load docxlib.ring", "Call save()"])

# Nested bullet list — inner list items are sub-lists
doc.addNestedBulletList([
    "Programming Languages",
    ["Ring", "Python", "Java"],
    "Databases",
    ["SQLite", "MySQL"]
])

# Nested numbered list
doc.addNestedNumberedList([
    "Phase 1 — Design",
    ["Requirements", "Architecture"],
    "Phase 2 — Development"
])
```

---

## 6. Tables

### 6.1 Table Creation Methods

| Method | Parameters | Description |
|---|---|---|
| `addTable()` | `data, options` | Full-featured table with all formatting options |
| `addSimpleTable()` | `data` | Plain table with no header styling |
| `addStyledTable()` | `data, headerBg, evenBg` | Quick styled table with two colour parameters |

### 6.2 Table Options

| Option | Type | Description |
|---|---|---|
| `:headerRow` | true/false | First row is a header: bold, coloured background |
| `:headerBgColor` | String | Header row background colour hex |
| `:headerTextColor` | String | Header text colour. Default: `FFFFFF` |
| `:evenRowBgColor` | String | Alternating even-row background colour |
| `:borderStyle` | String | `"single"`, `"double"`, `"dashed"`, `"dotted"`, `"none"` |
| `:borderColor` | String | Border colour hex. Default: `000000` |
| `:colWidths` | List | Column widths in cm, e.g. `[2, 6, 3, 5.5]` |
| `:headerVAlign` | String | Header vertical alignment: `"top"`, `"center"`, `"bottom"` |
| `:cellVAlign` | String | Body cell vertical alignment |
| `:repeatHeader` | true/false | Repeat header row on each page when the table spans pages |
| `:rowHeights` | List | Per-row heights in cm — `0` means auto-size, e.g. `[1.5, 0, 2.0]` |
| `:rowHRule` | String | Height rule for `:rowHeights`: `"atLeast"` (default) or `"exact"` |
| `:fontSize` | Number | Font size for all cells in points |
| `:conditionalRules` | List | Conditional cell formatting rules (see §6.5) |

```ring
doc.addTable([
    ["ID", "Description",                  "Status",      "Owner" ],
    ["1",  "Design new dashboard",          "In Progress", "Ahmed" ],
    ["2",  "Fix login timeout on mobile",   "Done",        "Sara"  ],
    ["3",  "Write API documentation v2",    "Pending",     "Omar"  ]
], [
    :headerRow      = true,
    :borderStyle    = "single",
    :headerBgColor  = "1565C0",
    :evenRowBgColor = "E3F2FD",
    :headerVAlign   = "center",
    :cellVAlign     = "center",
    :colWidths      = [2, 6, 3, 5.5]
])

# Per-row heights — header 1.5 cm tall, rows 2 & 4 auto-height, row 3 is 2.5 cm
doc.addTable(data, [
    :headerRow  = true,
    :colWidths  = [5, 10],
    :rowHeights = [1.5, 0, 2.5, 0],
    :rowHRule   = "atLeast"
])
```

### 6.3 Rich Cell Content — `WordCell`

For cells needing mixed formatting, images, lists, or cell spanning:

```ring
# Styled cell spanning 3 columns
titleCell = new WordCell()
titleCell.setText("Status Report", [:bold=true, :color="1E6BB5"])
titleCell.setBgColor("DEEAF1")
titleCell.setColSpan(3)
titleCell.setVerticalAlign("center")

# Row-spanning category cell
cat = new WordCell()
cat.addRun("Electronics", [:bold=true, :color="1565C0"])
cat.setRowSpan(3)
cat.setVerticalAlign("center")
cat.setAlign("center")
cat.setBgColor("E3F2FD")

# Convenience function: wordCell(text, options)
hdr = wordCell("Category", [:bold=true, :color="white", :bgColor="37474F"])
```

| `WordCell` method | Parameters | Description |
|---|---|---|
| `setText()` | `text, options` | Set main text with formatting options |
| `addRun()` | `text, options` | Append an additional text run |
| `addText()` | `text, options` | Alias for `addRun()` |
| `addCellBulletList()` | `items, listId` | Add a bullet list inside the cell |
| `addCellNumberedList()` | `items, listId` | Add a numbered list inside the cell |
| `addCellImage()` | `path, widthCm, heightCm` | Embed an image in the cell |
| `addCellHyperlink()` | `text, relId` | Add a hyperlink inside the cell |
| `setBgColor()` | `color` | Cell background colour hex |
| `setColSpan()` | `n` | Merge across n columns |
| `setRowSpan()` | `n` | Merge down n rows |
| `setVerticalAlign()` | `align` | `"top"`, `"center"`, `"bottom"` |
| `setAlign()` | `align` | Text alignment within the cell |
| `setBorder()` | `style, color` | All-sides border (legacy — prefer `setBorderSide`) |
| `setBorderSide()` | `side, style, color, size` | Per-side border — see §6.4 |
| `setTopBorder()` | `style, color, size` | Top edge only |
| `setBottomBorder()` | `style, color, size` | Bottom edge only |
| `setLeftBorder()` | `style, color, size` | Left edge only |
| `setRightBorder()` | `style, color, size` | Right edge only |
| `setNoBorder()` | — | Remove all borders from this cell |
| `setTextDir()` | `dir` | Rotate cell text — see §6.4 |
| `setPadding()` | `top, bot, left, right` | Cell padding in cm |
| `setWidth()` | `widthCm` | Override cell width in cm |

> **Tip:** Use `wordCell(text, options)` as a shorthand to create a `WordCell` with a
> single text run in one line.

### 6.4 Per-Side Cell Borders and Text Direction

**Per-side borders** give each cell edge an independent style, colour, and weight.
This makes it easy to build accent-left sidebars, coloured header underlines, and
borderless inner cells with only outer rules showing.

```ring
# Left accent sidebar — only a thick teal left edge, everything else removed
sidebar = wordCell("Note", [])
sidebar.setLeftBorder("thick",  "1B998B", 16)
sidebar.setTopBorder("none",    NULL, NULL)
sidebar.setBottomBorder("none", NULL, NULL)
sidebar.setRightBorder("none",  NULL, NULL)

# Double outer box
boxed = wordCell("Total", [:bold=true])
boxed.setTopBorder("double",    "000000", 8)
boxed.setBottomBorder("double", "000000", 8)
boxed.setLeftBorder("double",   "000000", 8)
boxed.setRightBorder("double",  "000000", 8)

# Header cell with a thick coloured bottom accent only
hdr = wordCell("January", [:bold=true, :align="center"])
hdr.setBottomBorder("thick", "E94F37", 12)
```

`setBorderSide(side, style, color, size)` parameter values:

| Parameter | Values |
|---|---|
| `side` | `"top"`, `"bottom"`, `"left"`, `"right"`, `"insideH"`, `"insideV"` |
| `style` | `"single"`, `"double"`, `"dashed"`, `"dotted"`, `"thick"`, `"none"` |
| `color` | Hex string, or `NULL` for auto |
| `size` | Eighths-of-a-point: `4`=½pt, `8`=1pt, `16`=2pt, `24`=3pt — or `NULL` for default |

**Text direction** rotates the cell's text. Most commonly used for compact column headers
in data tables where many narrow columns need rotated labels.

```ring
# Text reads upward (BTL) — saves horizontal space for many columns
colHdr = wordCell("January", [:bold=true, :align="center"])
colHdr.setTextDir("btLr")

# Text reads downward (TBR)
colHdr2 = wordCell("Q1 Target", [:bold=true])
colHdr2.setTextDir("tbRl")
```

| `setTextDir` value | Description |
|---|---|
| `"btLr"` | Bottom-to-top: text rotated 90° CCW (reads upward) |
| `"tbRl"` | Top-to-bottom: text rotated 90° CW (reads downward) |
| `"lrTb"` | Normal horizontal (default — also clears a previous direction) |

### 6.5 Conditional Table Formatting

---

## 7. Images

### 7.1 Inline Images

| Method | Parameters | Description |
|---|---|---|
| `addImage()` | `imagePath, widthCm, heightCm, options` | Inline image, left-aligned |
| `addImageCentered()` | `imagePath, widthCm, heightCm, options` | Inline image, centred on page |

Supported formats: PNG, JPG/JPEG, GIF, BMP. Pass `[]` for `options` when no options are needed.

**Image options:**

| Option | Type | Description |
|---|---|---|
| `:altText` | String | Accessibility description |
| `:cropLeft` | Number | Trim left side by this percentage (0–100) |
| `:cropRight` | Number | Trim right side by this percentage |
| `:cropTop` | Number | Trim top by this percentage |
| `:cropBottom` | Number | Trim bottom by this percentage |

```ring
# Basic inline image
doc.addImage("logo.png", 5, 3, [])
doc.addImageCentered("chart.png", 12, 8, [:altText="Benchmark chart"])

# Crop: remove 15% from each side, 10% from top (non-destructive)
doc.addImageCentered("photo.jpg", 12, 8, [
    :cropLeft = 15, :cropRight = 15, :cropTop = 10
])

# Centre punch-out — shows only the inner 60×60% of the image
doc.addImage("landscape.png", 10, 6, [
    :cropLeft=20, :cropRight=20, :cropTop=20, :cropBottom=20
])
```

> **Note:** Cropping is non-destructive. The full image is embedded; only the displayed
> viewport changes. Uses the OOXML `a:srcRect` element.

### 7.2 Floating Images

`addFloatingImage()` positions an image independently of the text flow.

```ring
doc.addFloatingImage("logo.png", 4, 3, [
    :wrapType = "wrapSquare",
    :wrapSide = "right",
    :posX     = 100,
    :posY     = 200,
    :distCm   = 0.3,
    :altText  = "Company logo",
    :cropTop  = 5,
    :cropBottom = 5
])
```

### 7.3 Image Watermarks

| Method | Parameters | Description |
|---|---|---|
| `setImageWatermark()` | `imagePath, options` | Background image watermark on every page |
| `removeImageWatermark()` | — | Remove the image watermark |

| Option | Type | Description |
|---|---|---|
| `:widthCm` | Number | Watermark width in cm. Default: 15 |
| `:heightCm` | Number | Watermark height in cm. Default: 15 |
| `:opacity` | Number | Opacity 0–100. Default: 50 |

```ring
doc.setImageWatermark("logo_grey.png", [:widthCm=10, :opacity=25])
```

---

## 8. Text Watermarks

| Method | Parameters | Description |
|---|---|---|
| `setWatermark()` | `text, options` | Diagonal text watermark on every page |
| `setWatermarkText()` | `text` | Change watermark text, keep other settings |
| `removeWatermark()` | — | Remove the text watermark |

| Option | Type | Description |
|---|---|---|
| `:color` | String | Hex colour. Default: `C0C0C0` |
| `:opacity` | Number | Opacity 0–100. Default: 50 |
| `:rotation` | Number | Rotation in degrees (negative = tilt left). Default: –45 |
| `:font` | String | Font name. Default: Arial |
| `:size` | Number | Font size in pt. Default: 72 |

```ring
# Simple DRAFT watermark
doc.setWatermarkText("DRAFT")

# Custom watermark
doc.setWatermark("CONFIDENTIAL", [
    :color    = "FF0000",
    :opacity  = 30,
    :rotation = -45,
    :size     = 96
])
```

---

## 9. Page Borders and Background

| Method | Parameters | Description |
|---|---|---|
| `setPageBorder()` | `options` | Add borders around every page |
| `setSimplePageBorder()` | — | Single-line 3pt black border on all four sides |
| `removePageBorder()` | — | Remove the page border |
| `setPageBackground()` | `color` | Solid background colour for all pages hex |

| `setPageBorder()` option | Type | Description |
|---|---|---|
| `:style` | String | `"single"`, `"double"`, `"thick"`, `"dashed"`, `"dotted"`, `"shadow"` |
| `:color` | String | Border colour hex. Default: `000000` |
| `:size` | Number | Width in eighths of a point. 24 = 3pt |
| `:space` | Number | Gap from page edge in points. Default: 24 |
| `:offsetFrom` | String | `"page"` or `"text"` |
| `:sides` | List | `["top","bottom","left","right"]`. Default: all four |

```ring
# Simple border
doc.setSimplePageBorder()

# Custom — top and bottom only, thick blue
doc.setPageBorder([
    :style = "single",
    :color = "1E6BB5",
    :size  = 48,
    :sides = ["top", "bottom"]
])

# Light blue background
doc.setPageBackground("EBF3FB")
```

---

## 10. Footnotes and Endnotes

| Method | Parameters | Description |
|---|---|---|
| `addFootnote()` | `text, footnoteContent, options` | Inline text with a superscript footnote reference |
| `addEndnote()` | `text, endnoteContent, options` | Inline text with an endnote reference |
| `registerFootnote()` | `footnoteContent` | Register a note, return its ID for use in `addRichParagraph` |
| `registerEndnote()` | `endnoteContent` | Register an endnote, return its ID |

The `footnoteContent` / `endnoteContent` argument accepts **either**:

- A **plain string** — simple unformatted note body
- A **list of `[text, options]` run pairs** — rich formatted body with per-run bold, italic, colour, size, and font (same format as `addRichParagraph`)

```ring
# Plain-text footnote (backward-compatible)
doc.addFootnote(
    "Ring is a multi-paradigm language.",
    "Fayed, M. (2013). Ring Programming Language. ring-lang.net",
    []
)

# Rich footnote — bold author name, coloured URL, italic date
doc.addFootnote(
    "See the official documentation.",
    [
        [" See: ",            []],
        ["Mahmoud Fayed",     [:bold=true, :color="1B5E20"]],
        [", Ring Language, ", []],
        ["2013–present",      [:italic=true, :color="666666"]],
        [". ring-lang.net",   []]
    ],
    []
)

# Rich endnote — italic title with underline, bold source
doc.addEndnote(
    "See the OOXML standard.",
    [
        ["Source: ",              []],
        ["OOXML Standard ECMA-376", [:italic=true, :underline=true]],
        [", 5th edition — ",      []],
        ["ISO/IEC 29500",         [:bold=true, :color="1A237E"]]
    ],
    []
)
```

> **Tip:** Use `registerFootnote(content)` to attach a note to a specific run inside
> `addRichParagraph()` via `:footnoteId` on the run options.

---

## 11. Hyperlinks, Bookmarks, and Cross-References

### 11.1 Hyperlinks

```ring
doc.addHyperlink("Ring Language Website", "http://ring-lang.net")
doc.addHyperlink("Send Email", "mailto:info@ring-lang.net")
```

### 11.2 Bookmarks

| Method | Parameters | Description |
|---|---|---|
| `addBookmarkedParagraph()` | `name, text, options` | Paragraph with an invisible named bookmark anchor |
| `addBookmark()` | `name` | Zero-width anchor inserted at the current position |

```ring
# Bookmark an entire heading
doc.addBookmarkedParagraph("bm_intro", "1. Introduction", [:style = "Heading1"])

# Bookmark a specific point (no text)
doc.addParagraph("See the results ", NULL)
doc.addBookmark("results_anchor")
```

### 11.3 Cross-References

`addCrossRef()` inserts a Word REF field that resolves to the page number of the named
bookmark when the user presses F9.

```ring
doc.addCrossRef("bm_intro", "see Introduction on page ")
# → renders as: 'see Introduction on page 3' after F9
```

---

## 12. Table of Contents

`addTableOfContents()` inserts a Word TOC field that collects all Heading 1–3 paragraphs.

```ring
doc.addTableOfContents("Table of Contents")
doc.addPageBreak()

doc.addHeading("Chapter 1", 1)
doc.addHeading("1.1 Background", 2)
```

> **Note:** Press F9 in Word to populate page numbers. The field uses
> `TOC \o "1-3" \h \z \u` and collects Headings 1 through 3.

---

## 13. Captions and Lists of Figures / Tables

### 13.1 SEQ Captions

```ring
# After an image:
doc.addImageCentered("arch.png", 12, 8)
doc.addFigureCaption("Ring VM architecture overview")
# → renders as:  Figure 1 — Ring VM architecture overview

# Above or below a table:
doc.addTableCaption("Feature comparison across Ring versions")
doc.addTable(data, options)
# → renders as:  Table 1 — Feature comparison across Ring versions
```

### 13.2 Table of Figures and Table of Tables

Caption entries are **pre-populated** from captions already added — lists are visible
immediately on open without pressing F9.

```ring
# Place these early in the document (before the figures/tables)
doc.addTableOfFigures("List of Figures")
doc.addTableOfTables("List of Tables")
doc.addPageBreak()

# For any custom SEQ label (Equation, Chart, Listing, etc.):
doc.addTableOfSeq("Equation", "List of Equations")
```

> **Tip:** Press F9 in Word to refresh page numbers. Caption text is always visible
> immediately; only the page numbers need updating.

---

## 14. Tab Stops and Line Numbers

### 14.1 Tabbed Paragraphs

```ring
# TOC-style dot leader:
tocTabs = [:tabStops = [[:pos=8500, :align="right", :leader="dot"]]]
doc.addTabbedParagraph(["Introduction",            "1"], tocTabs)
doc.addTabbedParagraph(["Chapter 1 — Background",  "3"], tocTabs)
doc.addTabbedParagraph(["Chapter 2 — Results",     "7"], tocTabs)

# Price list:
priceTabs = [:tabStops = [[:pos=8500, :align="right", :leader="dot"]]]
doc.addTabbedParagraph(["Ring Language",  "$0.00 (Open Source)"], priceTabs)
doc.addTabbedParagraph(["DOCXLib",    "$0.00 (Open Source)"], priceTabs)

# Multi-column layout:
multiTabs = [:tabStops = [
    [:pos=3000, :align="center", :leader="none"],
    [:pos=6000, :align="left",   :leader="none"],
    [:pos=9000, :align="right",  :leader="none"]
]]
doc.addTabbedParagraph(["Name", "Role", "Score", "Grade"], multiTabs)
```

| Tab stop option | Values | Description |
|---|---|---|
| `:pos` | Number (twips) | Distance from left margin. 1440 twips = 1 inch |
| `:align` | `"left"` / `"center"` / `"right"` / `"decimal"` | Alignment at the stop |
| `:leader` | `"none"` / `"dot"` / `"hyphen"` / `"underscore"` | Fill character between stops |

### 14.2 Line Numbers

```ring
doc.enableLineNumbers([
    :start   = 1,          # first line number
    :step    = 5,          # show number every 5 lines
    :restart = "newPage"   # restart on each page
])

doc.addParagraph("This paragraph shows line numbers.", NULL)
doc.addParagraph("This one is excluded.", [:suppressLineNumbers=true])

doc.disableLineNumbers()   # turn off for remaining paragraphs
```

| `enableLineNumbers` option | Type | Description |
|---|---|---|
| `:start` | Number | First line number. Default: 1 |
| `:step` | Number | Show number every N lines. Default: 1 |
| `:distance` | Number | Gap from text margin in twips. Default: 360 |
| `:restart` | String | `"newPage"`, `"newSection"`, `"continuous"` |

---

## 15. Comments

| Method | Parameters | Description |
|---|---|---|
| `addCommentedParagraph()` | `text, commentText, options` | Paragraph with a comment balloon attached |
| `addComment()` | `text, commentText, author` | Shortcut with explicit author name |

```ring
doc.addCommentedParagraph(
    "The Ring VM is stack-based and written in C.",
    "Please verify against the source",
    [:commentAuthor = "Dr. Mahmoud",
     :commentDate   = "2025-03-15T09:00:00Z"]
)

# Shortcut form:
doc.addComment("See Table 3 for benchmark data.",
               "Check these numbers against the latest run.",
               "Reviewer")
```

---

## 16. Mail Merge Fields

Mail merge fields create `.docx` templates that Word merges with data sources. Each
`MERGEFIELD` placeholder is replaced with a real value during the merge.

```ring
# Formal letter template
doc.addMergeFieldParagraph([[:field="RecipientName"]], [:bold=true, :size=13])
doc.addMergeFieldParagraph([[:field="CompanyName"]], NULL)
doc.addMergeFieldParagraph([[:field="Address"]], NULL)
doc.addMergeFieldParagraph([[:field="City"], ", ", [:field="Country"]], NULL)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph(["Dear ", [:field="Salutation"], " ", [:field="LastName"], ","], NULL)

# Batch template for invoice fields:
doc.addMergeTemplate([
    ["Invoice #: ",   [:field="InvoiceNo"]],
    ["Date: ",        [:field="InvoiceDate"]],
    ["Client: ",      [:field="ClientName"]],
    ["Amount Due: $", [:field="TotalAmount"]]
])
```

> **Note:** In `addMergeFieldParagraph()`, pass a plain string for literal text and
> `[:field="Name"]` for a `MERGEFIELD` placeholder. They can be freely mixed.

---

## 17. Content Controls

Content controls embed interactive form elements in the DOCX. Users can fill in the form
without enabling macros.

### 17.1 Checkbox

```ring
doc.addCheckbox("I accept the Terms and Conditions",      true)   # pre-checked
doc.addCheckbox("Subscribe to the Ring newsletter",       false)  # unchecked
doc.addCheckbox("Enable two-factor authentication",       true)
doc.addCheckbox("Allow anonymous usage statistics",       false)
```

### 17.2 Dropdown List

```ring
doc.addDropdown("Country",
    ["Saudi Arabia", "Egypt", "UAE", "Kuwait", "Jordan", "Morocco", "Other"],
    "Saudi Arabia"    # default selection
)

doc.addDropdown("Experience Level",
    ["Beginner", "Intermediate", "Advanced", "Expert"],
    "Intermediate"
)
```

### 17.3 Text Input

```ring
doc.addTextInput("Full Name",    "",                    "Enter your full name")
doc.addTextInput("Email",        "",                    "name@example.com")
doc.addTextInput("Organisation", "",                    "Company or university")
doc.addTextInput("Comments",     "Pre-filled default",  NULL)
```

---

## 18. Drawing Shapes

Inline DrawingML shapes embedded in the document flow. All shapes support fill colour,
border, and an optional text label.

### 18.1 Shape Methods

| Method | Shape type | Notes |
|---|---|---|
| `addRect(options)` | Rectangle | Use `:rounded=true` for rounded corners |
| `addEllipse(options)` | Ellipse / circle | Equal width and height for a circle |
| `addDiamond(options)` | Diamond / rhombus | Useful for flowchart decision nodes |
| `addShape("triangle", options)` | Right triangle | |
| `addLine(options)` | Horizontal rule | Thin coloured rect; use `:height` for thickness |
| `addShape(type, options)` | Any of the above | Generic form |

### 18.2 Shape Options

| Option | Type | Description |
|---|---|---|
| `:width` | Number | Width in cm. Default: 5 |
| `:height` | Number | Height in cm. Default: 3 (line: 0.1) |
| `:fillColor` | String | Fill colour hex. Default: `4472C4` |
| `:lineColor` | String | Border colour hex. Default: `2E4E7E` |
| `:lineWidth` | Number | Border width in pt. Default: 1 |
| `:noFill` | true/false | Transparent shape (no fill) |
| `:noBorder` | true/false | No border |
| `:rounded` | true/false | Rounded corners (rect only) |
| `:text` | String | Label text inside the shape |
| `:textColor` | String | Text colour hex. Default: `FFFFFF` |
| `:textBold` | true/false | Bold text label. Default: true |
| `:textSize` | Number | Text size in pt. Default: 11 |
| `:align` | String | Outer paragraph alignment: `"left"`, `"center"`, `"right"` |

```ring
# Solid rectangle with white text
doc.addRect([:width=8, :height=1.8, :fillColor="1E6BB5",
             :text="Section Header", :textColor="FFFFFF"])

# Rounded border-only box
doc.addRect([:width=6, :height=1.4, :noFill=true,
             :lineColor="C00000", :lineWidth=2,
             :text="Border Only", :textColor="C00000", :rounded=true])

# Thin horizontal rule
doc.addLine([:width=14, :height=0.08, :fillColor="4472C4", :noBorder=true])

# Circle with text
doc.addEllipse([:width=3, :height=3, :fillColor="7030A0",
               :text="75%", :textSize=18])
```

---

## 19. Text Boxes

`addTextBox()` inserts a floating VML text box positioned absolutely relative to the page.

```ring
doc.addTextBox("This content floats beside the main body text.", [
    :width       = 6,       # cm
    :height      = 4,       # cm
    :top         = 2,       # cm from top of page
    :left        = 10.5,    # cm from left of page
    :bgColor     = "F0F4FF",
    :borderColor = "1E6BB5",
    :bold        = true,
    :fontSize    = 11,
    :align       = "left"
])
```

---

## 20. RTL Support

| Method | Parameters | Description |
|---|---|---|
| `setDocumentRTL()` | `bEnable` | Enable document-wide RTL — flips paragraph alignment and list indents |
| `addRTLParagraph()` | `text, options` | Single RTL paragraph in an otherwise LTR document |

```ring
# Full RTL document (Arabic)
doc.setDocumentRTL(true)
doc.addHeading("مرحباً بكم في DOCXLib", 1)
doc.addParagraph("هذا مستند منشأ باستخدام لغة Ring.", NULL)

# Single RTL paragraph inside an LTR document
doc.addRTLParagraph("النص العربي هنا", NULL)
```

---

## 21. Saving the Document

```ring
if doc.save("output.docx")
    ? "Document saved successfully."
else
    ? "Error: save() failed."
ok
```

`save()` returns `true` on success. It automatically:

1. Creates a temporary working directory
2. Writes all XML parts: `document.xml`, `styles.xml`, `settings.xml`, `numbering.xml`,
   `fontTable.xml`, `footnotes.xml`, `endnotes.xml`, `comments.xml`, header/footer files,
   and `theme/theme1.xml` (when a theme is set)
3. Copies all referenced image files into `word/media/`
4. Assembles `[Content_Types].xml` and all `.rels` relationship files
5. Creates the ZIP archive (DOCX format) using Ring's built-in zip functions
6. Cleans up the temporary directory

---

## 22. Native OOXML Charts

Chart types are stored as pure XML inside the DOCX package with
inline data caches, so charts are completely self-contained with no linked spreadsheet.

### 22.1 Chart Methods

#### Category-based charts (use string category labels)

| Method | Chart Type | Best Used For |
|---|---|---|
| `addColumnChart(title, cats, series, opts)` | Vertical bars | Category comparisons across groups |
| `addBarChart(title, cats, series, opts)` | Horizontal bars | Long category labels |
| `addLineChart(title, cats, series, opts)` | Lines / trends | Time-series, continuous data |
| `addAreaChart(title, cats, series, opts)` | Filled areas | Cumulative or stacked trends |
| `addPieChart(title, cats, series, opts)` | Pie slices | Proportions — uses first series only |
| `addDoughnutChart(title, cats, series, opts)` | Hollow-centre pie | Proportions with inner space |
| `addChart(type, title, cats, series, opts)` | Any type above | Generic entry point |

#### Numeric-axis charts (no category labels — pass `NULL` for `cats`)

| Method | Chart Type | Best Used For |
|---|---|---|
| `addScatterChart(title, series, opts)` | XY scatter | Correlation, scientific data, irregular X spacing |
| `addBubbleChart(title, series, opts)` | Bubble | Three-variable data — X, Y, and a size dimension |

### 22.2 Series Format

#### Category-based series (column, bar, line, area, pie, doughnut)

| Key | Type | Required | Description |
|---|---|---|---|
| `:name` | String | Yes | Series label shown in the legend |
| `:values` | List | Yes | Numeric values, one per category |
| `:color` | String | No | Hex colour override, e.g. `"4472C4"` |

#### Scatter series

| Key | Type | Required | Description |
|---|---|---|---|
| `:name` | String | Yes | Series label shown in the legend |
| `:xValues` | List | Yes | X-axis numeric values |
| `:yValues` | List | Yes | Y-axis numeric values (same length as `:xValues`) |
| `:color` | String | No | Hex colour override |

#### Bubble series

| Key | Type | Required | Description |
|---|---|---|---|
| `:name` | String | Yes | Series label shown in the legend |
| `:xValues` | List | Yes | X-axis numeric values |
| `:yValues` | List | Yes | Y-axis numeric values |
| `:sizes` | List | Yes | Bubble size values — positive numbers, relative scale |
| `:color` | String | No | Hex colour override |

If `:color` is omitted for any series type the library cycles through an 8-colour
Office palette automatically.

### 22.3 Options Reference

#### Layout and legend options (all chart types)

| Option | Type | Default | Description |
|---|---|---|---|
| `:widthCm` | Number | `14` | Display width in centimetres |
| `:heightCm` | Number | `9` | Display height in centimetres |
| `:centered` | Bool | `true` | Centre chart horizontally on the page |
| `:showLegend` | Bool | `true` | Show or hide the legend |
| `:legendPos` | String | `"r"` | Legend position: `"r"` `"b"` `"t"` `"l"` |
| `:showDataLabels` | Bool | `false` | Show value labels on each data point |

#### Axis formatting options (all non-pie/doughnut chart types)

| Option | Type | Default | Description |
|---|---|---|---|
| `:yAxisMin` | Number | auto | Force value axis lower bound (e.g. `0` to start at zero) |
| `:yAxisMax` | Number | auto | Force value axis upper bound |
| `:yAxisFormat` | String | `"General"` | Number format for value axis labels — see table below |
| `:yAxisTitle` | String | `""` | Rotated label alongside the value axis |
| `:xAxisTitle` | String | `""` | Label below the category axis (or X axis for scatter/bubble) |
| `:showGridlines` | Bool | `true` | Show major gridlines on the value axis |
| `:showDataTable` | Bool | `false` | Show source data in a table below the chart (column/bar/line/area only) |
| `:dataTableShowKeys` | Bool | `true` | Include colour legend keys in the data table |

#### Axis format string reference

| Format string | Example output | Use case |
|---|---|---|
| `"General"` | `1234567` | Default — no formatting |
| `"#,##0"` | `1,234,567` | Large integers with thousands separator |
| `"$#,##0"` | `$1,234,567` | Currency (whole dollars) |
| `"$#,##0.00"` | `$1,234,567.89` | Currency with cents |
| `"0%"` | `32%` | Percentage (value `0.32` → `32%`) |
| `"0.0%"` | `32.5%` | Percentage with one decimal |
| `"0.00"` | `3.14` | Two decimal places |
| `"0.0"` | `3.1` | One decimal place |
| `"0.00E+00"` | `3.14E+06` | Scientific notation |

> **Note:** These are standard OOXML/Excel number format strings. Any Excel format
> string is accepted — the library passes it through verbatim to the XML.

#### Category-chart options

| Option | Type | Default | Description |
|---|---|---|---|
| `:grouping` | String | `"clustered"` | Bar/column/area: `"clustered"` `"stacked"` `"percentstacked"` |
| `:smooth` | Bool | `false` | Line chart only — bezier curves |
| `:colors` | List | *(palette)* | Override full palette with a list of hex strings |

#### Scatter options

| Option | Type | Default | Description |
|---|---|---|---|
| `:markerStyle` | String | `"circle"` | `"circle"` `"square"` `"diamond"` `"triangle"` `"x"` `"star"` `"dot"` `"dash"` `"none"` |
| `:markerSize` | Number | `7` | Marker size in half-points |
| `:lines` | Bool | `false` | Connect points with lines |
| `:smooth` | Bool | `false` | Bezier curves on lines (requires `:lines = true`) |
| `:xAxisMin` | Number | auto | Force X axis lower bound |
| `:xAxisMax` | Number | auto | Force X axis upper bound |
| `:xAxisFormat` | String | `"General"` | Number format for X axis labels |
| `:xAxisTitle` | String | `""` | X axis label |
| `:yAxisTitle` | String | `""` | Y axis label |

#### Bubble options

| Option | Type | Default | Description |
|---|---|---|---|
| `:bubble3D` | Bool | `false` | Render bubbles with 3-D specular shading |
| `:xAxisMin` | Number | auto | Force X axis lower bound |
| `:xAxisMax` | Number | auto | Force X axis upper bound |
| `:xAxisFormat` | String | `"General"` | Number format for X axis labels |
| `:xAxisTitle` | String | `""` | X axis label |
| `:yAxisTitle` | String | `""` | Y axis label |

### 22.4 Code Examples

**Column chart — three series, data labels:**

```ring
doc.addColumnChart(
    "Quarterly Revenue vs Costs ($K)",
    ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"],
    [
        [:name = "Revenue",  :values = [320, 410, 380, 490]],
        [:name = "Costs",    :values = [210, 255, 230, 290]],
        [:name = "Profit",   :values = [110, 155, 150, 200]]
    ],
    [:widthCm = 14, :heightCm = 9, :showDataLabels = true, :legendPos = "b"])
```

**Line chart — smooth curves:**

```ring
doc.addLineChart(
    "Monthly Website Traffic",
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    [
        [:name = "2023", :values = [12000, 14500, 13200, 16800, 18000, 17400]],
        [:name = "2024", :values = [15000, 18200, 17500, 21000, 23500, 22800]]
    ],
    [:smooth = true, :legendPos = "r"])
```

**Pie chart — percent labels:**

```ring
doc.addPieChart(
    "Product Revenue Share",
    ["Ring", "PWCT", "Supernova", "PWCT2"],
    [[:name = "Revenue ($K)", :values = [450, 280, 160, 110]]],
    [:showDataLabels = true])
```

**Stacked column — `:grouping` option:**

```ring
doc.addColumnChart(
    "Revenue by Product",
    ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"],
    [
        [:name = "Ring",      :values = [120, 155, 140, 180]],
        [:name = "PWCT",      :values = [80,  95,  90,  110]],
        [:name = "Supernova", :values = [60,  75,  70,   90]]
    ],
    [:grouping = "stacked"])
```

**Scatter chart — markers only (note: `NULL` for categories, `:xValues`/`:yValues` in series):**

```ring
doc.addScatterChart(
    "Study Hours vs Exam Score",
    [
        [:name = "Cohort A",
         :xValues = [2, 3, 4, 5, 6, 7, 8, 9, 10],
         :yValues = [45, 52, 60, 67, 73, 79, 84, 88, 93]],
        [:name = "Cohort B",
         :xValues = [1, 2, 4, 5, 7, 8, 9, 10, 11],
         :yValues = [38, 48, 55, 63, 70, 75, 81, 86, 90]]
    ],
    [:markerStyle = "circle", :xAxisTitle = "Study Hours", :yAxisTitle = "Score (%)"])
```

**Scatter chart — smooth lines + markers:**

```ring
doc.addScatterChart(
    "Algorithm Complexity",
    [
        [:name = "O(n log n)", :xValues = [100, 500, 1000, 5000, 10000],
         :yValues = [0.8, 5.2, 11.5, 68.3, 145.2], :color = "4472C4"],
        [:name = "O(n^2)",     :xValues = [100, 500, 1000, 5000, 10000],
         :yValues = [1.2, 28.5, 112.4, 2802.5, 11210.0], :color = "C00000"]
    ],
    [:lines = true, :smooth = true, :markerStyle = "diamond",
     :xAxisTitle = "Input Size (n)", :yAxisTitle = "Time (ms)"])
```

**Scatter chart — lines only, no markers:**

```ring
doc.addScatterChart(
    "Sine vs Cosine Wave",
    [
        [:name = "sin(x)", :xValues = aX, :yValues = aSin, :color = "2E75B6"],
        [:name = "cos(x)", :xValues = aX, :yValues = aCos, :color = "70AD47"]
    ],
    [:lines = true, :smooth = true, :markerStyle = "none"])
```

**Bubble chart — flat shading (each product is its own series for distinct colours):**

```ring
doc.addBubbleChart(
    "Product Portfolio — Share / Growth / Revenue",
    [
        [:name = "Ring",      :xValues = [35], :yValues = [22], :sizes = [450], :color = "4472C4"],
        [:name = "PWCT",      :xValues = [28], :yValues = [15], :sizes = [280], :color = "ED7D31"],
        [:name = "Supernova", :xValues = [18], :yValues = [31], :sizes = [160], :color = "70AD47"]
    ],
    [:xAxisTitle = "Market Share (%)", :yAxisTitle = "Growth (%)"])
```

**Bubble chart — 3-D shading, bubble-size labels, multiple points per series:**

```ring
doc.addBubbleChart(
    "Research Projects — Team / Duration / Impact",
    [
        [:name = "Project Alpha",
         :xValues = [3, 5, 8, 12], :yValues = [6, 9, 14, 18],
         :sizes = [45, 120, 280, 500], :color = "1F3864"],
        [:name = "Project Beta",
         :xValues = [2, 4, 7, 10], :yValues = [4, 7, 11, 15],
         :sizes = [30, 90, 200, 380], :color = "C00000"]
    ],
    [:bubble3D = true, :showDataLabels = true,
     :xAxisTitle = "Team Size", :yAxisTitle = "Months"])
```

**Column chart — fixed axis bounds and currency format:**

```ring
doc.addColumnChart(
    "Regional Revenue vs Target",
    ["North", "South", "East", "West", "Central"],
    [
        [:name = "Actual",  :values = [1250000, 980000, 1420000, 1100000, 870000]],
        [:name = "Target",  :values = [1200000, 1050000, 1350000, 1150000, 950000]]
    ],
    [:yAxisMin = 0, :yAxisMax = 1800000,
     :yAxisFormat = "$#,##0",
     :yAxisTitle = "Revenue (USD)", :xAxisTitle = "Region",
     :showGridlines = false, :legendPos = "b"])
```

**Line chart — percentage format on value axis:**

```ring
doc.addLineChart(
    "Gross Margin Trend",
    ["Q1", "Q2", "Q3", "Q4"],
    [
        [:name = "2023", :values = [0.38, 0.41, 0.39, 0.43]],
        [:name = "2024", :values = [0.41, 0.44, 0.42, 0.46]]
    ],
    [:yAxisMin = 0, :yAxisMax = 0.6,
     :yAxisFormat = "0.0%",
     :yAxisTitle = "Gross Margin", :xAxisTitle = "Quarter"])
```

**Scatter chart — fixed bounds with formatted numeric axes:**

```ring
doc.addScatterChart(
    "Ad Spend vs Revenue",
    [
        [:name = "2024",
         :xValues = [50000, 100000, 200000, 350000],
         :yValues = [210000, 430000, 870000, 1500000]]
    ],
    [:xAxisMin = 0, :xAxisMax = 400000,
     :yAxisMin = 0, :yAxisMax = 2000000,
     :xAxisFormat = "$#,##0", :yAxisFormat = "$#,##0",
     :xAxisTitle = "Ad Spend (USD)", :yAxisTitle = "Revenue (USD)",
     :markerStyle = "circle", :lines = true, :smooth = true])
```


### 22.4.1 Pie and Doughnut: Single Series Only

Pie and doughnut charts render **only the first series**. If more than one series is
passed, DOCXLib prints a warning to the console and silently discards the
extra series. Use a column or bar chart when you need
to compare multiple series.

```ring
# WRONG — triggers warning; only "2023" series will appear in the chart
doc.addPieChart("Sales", cats, [
    [:name = "2023", :values = [380, 240, 140, 90]],
    [:name = "2024", :values = [450, 280, 160, 110]]   # ignored
], [])

# CORRECT — grouped column for side-by-side comparison
doc.addColumnChart("Sales 2023 vs 2024", cats, [
    [:name = "2023", :values = [380, 240, 140, 90]],
    [:name = "2024", :values = [450, 280, 160, 110]]
], [:legendPos = "b"])
```

### 22.5 OOXML Internals

Each chart adds three components to the DOCX package automatically:

| File | Purpose |
|---|---|
| `word/charts/chartN.xml` | Full chart definition — data stored inline as XML caches; no spreadsheet needed |
| `word/charts/_rels/chartN.xml.rels` | Chart relationship manifest (no external dependencies) |
| Drawing reference in `document.xml` | `<c:chart r:id="rIdN"/>` anchors the chart in the document body |

`[Content_Types].xml` receives one `Override` entry per chart automatically.

Category-based charts use `<c:cat>` + `<c:val>` for series data. Scatter and bubble
charts use `<c:xVal>` + `<c:yVal>` (and `<c:bubbleSize>` for bubble) instead —
both axes are `<c:valAx>` rather than `<c:catAx>` + `<c:valAx>`.

> **Compatibility:** Charts render in Word 2016+, Microsoft 365, and LibreOffice 7+.
> The DrawingML chart schema (`http://schemas.openxmlformats.org/drawingml/2006/chart`)
> is part of the ECMA-376 standard.

---

## 23. Complete Example

```ring
load "docxlib.ring"

doc = new WordWriter()
doc.setTitle("Ring Language Performance Study")
doc.setAuthor("Mahmoud Fayed")
doc.setPageSize("a4")
doc.setTheme("Blue")

# Header / footer
doc.setFirstPageDifferent(true)
doc.setFirstPageHeader("")          # blank on cover
doc.setHeader("Ring Language Performance Study")
doc.showPageNumbers("right")

# Cover page
doc.addTitle("Ring Language Performance Study")
doc.addSubtitle("A Comparative Analysis — v1.11.0 Edition")
doc.addPageBreak()

# Front matter
doc.addTableOfContents("Table of Contents")
doc.addTableOfFigures("List of Figures")
doc.addTableOfTables("List of Tables")
doc.addPageBreak()

# Chapter 1
doc.addHeading("1. Introduction", 1)
doc.addParagraph("Ring is a multi-paradigm language designed for application development.", NULL)
doc.addFootnote("See ring-lang.net", "Ring official website.", NULL)
doc.addParagraph("", NULL)

doc.addHeading("1.1 Methodology", 2)
doc.addBulletList(["Benchmark suite", "Reference VM", "Metrics collected"])
doc.addParagraph("", NULL)

# Image + caption
doc.addImageCentered("chart.png", 12, 7)
doc.addFigureCaption("Benchmark results across five test configurations")
doc.addParagraph("", NULL)

# Table + caption
doc.addTableCaption("Raw timing data in milliseconds")
doc.addTable([
    ["Test Case",  "Ring", "Python", "Node.js"],
    ["Fibonacci",  "42",   "180",    "95"     ],
    ["Sort 10k",   "18",   "72",     "38"     ],
    ["File I/O",   "55",   "140",    "60"     ]
], [
    :headerRow      = true,
    :borderStyle    = "single",
    :headerBgColor  = "1E6BB5",
    :evenRowBgColor = "DEEAF1"
])

doc.save("report.docx")
```

---

## 24. Mail Merge Template Engine

A full template-based mail merge engine. Define a template once with `setMergeTemplate()`, then call `mergeRecord()` or `mergeAll()` to produce personalised documents for any number of records.

### 24.1 API

| Method | Description |
|---|---|
| `setMergeTemplate(template)` | Register a template definition (list of items) |
| `clearMergeTemplate()` | Remove the registered template |
| `mergeRecord(data)` | Render the template once, substituting `data` into all `{{FIELD}}` tokens |
| `mergeAll(dataList, sep)` | Render for every record in `dataList`; insert separator between records |

**`mergeAll` separator values:** `"pagebreak"` (default) · `"emptyline"` · `"line"` · `"none"`

### 24.2 Template Item Types

A template is a Ring list. Each item is either a plain string or an associative list:

| Item type | Description |
|---|---|
| `"Dear {{FirstName}},"` | Plain string — rendered as a paragraph; `{{FIELD}}` tokens filled |
| `[:type="heading", :level=N, :text="{{Field}} Report"]` | Heading at level 1–6 |
| `[:type="paragraph", :text="...", :options=[...]]` | Paragraph with full formatting options |
| `[:type="table", :data=[...], :options=[...]]` | Table — `{{FIELD}}` tokens in cells are filled |
| `[:type="pagebreak"]` | Page break |
| `[:type="emptyline"]` | Empty paragraph |

`{{FIELD}}` tokens work in all text fields including table cells. Unknown fields are left as-is.

### 24.3 Data Record Format

A data record is a Ring associative list:

```ring
data = [
    :FirstName = "Alice",
    :LastName  = "Smith",
    :Score     = "94"
]
```

`mergeAll()` takes a list of such records:

```ring
dataList = [
    [:Name = "Alice", :Score = "94"],
    [:Name = "Bob",   :Score = "78"]
]
```

### 24.4 Code Examples

**Simple letter — three recipients, page break between each:**

```ring
doc.setMergeTemplate([
    [:type = "paragraph", :text = "{{Date}}", :options = [:align = "right"]],
    [:type = "emptyline"],
    "Dear {{Salutation}} {{LastName}},",
    [:type = "emptyline"],
    [:type = "paragraph",
     :text = "Your application to {{ProgramName}} has been accepted. " +
             "Your enrolment number is {{EnrolmentNumber}}."],
    [:type = "emptyline"],
    [:type = "paragraph", :text = "Yours sincerely,"],
    [:type = "paragraph", :text = "Professor Mahmoud Fayed", :options = [:bold = true]]
])

recipients = [
    [:Date = "1 March 2025", :Salutation = "Ms", :LastName = "Benali",
     :ProgramName = "Ring MSc", :EnrolmentNumber = "MSC-2025-001"],
    [:Date = "1 March 2025", :Salutation = "Mr", :LastName = "Mendes",
     :ProgramName = "Ring MSc", :EnrolmentNumber = "MSC-2025-002"]
]

doc.mergeAll(recipients, "pagebreak")
doc.clearMergeTemplate()
```

**Invoice template with an embedded table:**

```ring
doc.setMergeTemplate([
    [:type = "heading", :level = 1, :text = "INVOICE #{{InvoiceNum}}"],
    [:type = "table",
     :data = [
         ["Bill To",  "{{ClientName}}",  "Date",    "{{InvoiceDate}}"],
         ["Address",  "{{Address}}",     "Due",     "{{DueDate}}"]
     ],
     :options = [:borderStyle = "single", :colWidths = [3.5, 5.0, 2.5, 5.0]]],
    [:type = "emptyline"],
    [:type = "paragraph",
     :text = "Total Due: {{Total}}", :options = [:bold = true, :align = "right"]]
])

doc.mergeAll(invoices, "pagebreak")
```

---

## 25. Chart Data Tables

The `:showDataTable = true` option renders the chart's source data in a compact bordered table directly beneath the plot area — useful for printed reports where readers need exact values.

### 25.1 Options

| Option | Type | Default | Description |
|---|---|---|---|
| `:showDataTable` | Bool | `false` | Render source data table beneath the chart |
| `:dataTableShowKeys` | Bool | `true` | Include colour legend keys in the leftmost column |

**Supported chart types:** column, bar, line, area
**Not applicable to:** pie, doughnut, scatter, bubble

### 25.2 Code Examples

**Column chart with data table and legend keys:**

```ring
doc.addColumnChart(
    "Quarterly Revenue",
    ["Q1", "Q2", "Q3", "Q4"],
    [
        [:name = "Ring", :values = [320000, 410000, 380000, 490000]],
        [:name = "PWCT", :values = [210000, 265000, 245000, 305000]]
    ],
    [:yAxisFormat = "$#,##0", :legendPos = "b",
     :showDataTable = true, :dataTableShowKeys = true])
```

**Line chart — data table without colour keys:**

```ring
doc.addLineChart(
    "Monthly Active Users",
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    [[:name = "Web",    :values = [142, 158, 171, 185, 201, 218]],
     [:name = "Mobile", :values = [89,  104, 118, 135, 152, 171]]],
    [:smooth = true, :showDataTable = true, :dataTableShowKeys = false])
```

**Combined with chart style presets:**

```ring
doc.setChartDefaults([
    :widthCm       = 14,
    :legendPos     = "b",
    :showDataTable = true,
    :dataTableShowKeys = true,
    :yAxisFormat   = "#,##0"
])

# Every subsequent chart automatically gets a data table
doc.addColumnChart("Revenue", quarters, series, [])
doc.addLineChart("Trend", months, trendSeries, [:smooth = true])
```

---

## 26. Version History

| Version | Key Features Added |
|---|---|
| 1.0–1.5 | Paragraphs, tables, images, lists, styles, hyperlinks, footnotes, RTL |
| 1.6 | Text boxes, custom styles, watermarks, multi-column, paragraph shading |
| 1.7 | Odd/even headers, page background, per-side page borders |
| 1.8 | Bookmarks, cross-references, tab stops with leaders, line numbers |
| 1.9 | Comments/annotations, SEQ captions, mail merge MERGEFIELD |
| 1.10 | Content controls (checkbox, dropdown, text), DrawingML shapes, themes |
| 1.11 | Table of Figures, Table of Tables, `addTableOfSeq()` |
| 1.12 | 6 native OOXML chart types: column, bar, line, area, pie, doughnut |
| 1.12.1 | `addScatterChart()`, `addBubbleChart()` — numeric dual-axis charts |
| 1.12.2 | Axis formatting: min/max, number formats, axis titles, gridlines control |
| 1.12.3 | `setChartDefaults()`, `clearChartDefaults()`; conditional table formatting (8 conditions) |
| 1.12.4 | Mail merge template engine: `setMergeTemplate()`, `mergeRecord()`, `mergeAll()`, `clearMergeTemplate()` |
| 1.12.5 | Rich `WordCell` in merge templates: `mergeCell()`, `mergeFillWordCell()` |
| Reader 3l | `addSectionBreak()` options: `:numColumns`, `:columnSpaceCm`; run `:lang` and `:charStyle` |
| Reader 3m | Per-side cell borders: `setBorderSide()`, `setTopBorder()`, `setBottomBorder()`, `setLeftBorder()`, `setRightBorder()`, `setNoBorder()`; per-row heights: `:rowHeights`, `:rowHRule`; cell text direction: `setTextDir()` |
| Reader 3n | Rich footnote/endnote bodies (run lists with per-run formatting); image cropping: `:cropLeft/Right/Top/Bottom`; widow/orphan control: `:widowControl=false`; suppress hyphenation: `:noHyphenate=true` |

---

## 27. Introduction to WordReader

`WordReader` (`docxlib.ring`) parses any `.docx` file and provides:

1. **Structured access** to all document content as Ring lists
2. **Semantic query methods** — `getHeadings()`, `getTables()`, `getSectionLayouts()`, etc.
3. **Full round-trip** — `toWriter()` converts the parsed document back into a `WordWriter`
   that you can modify and re-save

```ring
load "docxlib.ring"

reader = new WordReader("report.docx")
reader.loadDocx()

? reader.summary()

headings = reader.getHeadings()
for h in headings
    ? string(h[:level]) + ": " + h[:text]
next

reader.save("report_copy.docx")
reader.cleanup()
```

---

## 28. Loading and Saving

| Method | Parameters | Description |
|---|---|---|
| `new WordReader(filePath)` | `filePath : String` | Constructor — sets the source `.docx` path |
| `loadDocx()` | — | Unzip and parse the document. Call once after construction |
| `save(outputPath)` | `outputPath : String` | Round-trip: parse → reconstruct → write |
| `cleanup()` | — | Delete the temporary extraction folder |
| `toWriter()` | — | Return a populated `WordWriter` from the parsed content |

> `loadDocx()` resets all internal state each call, so one `WordReader` instance can be
> reused across multiple files.

---

## 29. Document Properties

| Method | Returns | Description |
|---|---|---|
| `getFilePath()` | String | Source file path |
| `getTitle()` | String | Document title from core.xml |
| `getAuthor()` | String | Document author |
| `getDefaultFont()` | String | Body font name |
| `getDefaultFontSize()` | Number | Body font size in points |
| `getPageWidthCm()` | Number | Page width in cm |
| `getPageHeightCm()` | Number | Page height in cm |
| `getMarginTopCm()` | Number | Top margin in cm |
| `getMarginBottomCm()` | Number | Bottom margin in cm |
| `getMarginLeftCm()` | Number | Left margin in cm |
| `getMarginRightCm()` | Number | Right margin in cm |
| `getOrientation()` | String | `"portrait"` or `"landscape"` |
| `getHeaderText()` | String | Default header text |
| `getFooterText()` | String | Default footer text |
| `getEvenPageHeaderText()` | String | Even-page header |
| `getEvenPageFooterText()` | String | Even-page footer |
| `getFirstPageHeaderText()` | String | First-page header |
| `getFirstPageFooterText()` | String | First-page footer |
| `hasEvenAndOddHeaders()` | Bool | Even/odd header mode enabled |
| `hasFirstPageDifferent()` | Bool | First-page-different mode enabled |
| `hasPageBorder()` | Bool | Page border present |
| `getPageBorderStyle()` | String | Page border style |
| `getPageBorderColor()` | String | Page border colour hex |
| `getPageBorderSizePt()` | Number | Page border width in points |
| `getPageBgColor()` | String | Page background colour hex |
| `hasPageBackground()` | Bool | Solid page background present |
| `getColumnCount()` | Number | Number of columns in the document |
| `getColumnSpaceCm()` | Number | Column gap in cm |
| `getTOCTitle()` | String | Table of contents title |
| `hasTOC()` | Bool | TOC field present |
| `hasWatermark()` | Bool | Text watermark present |
| `getWatermarkText()` | String | Watermark text |
| `getWatermarkOptions()` | List | Watermark properties |
| `isDocumentRTL()` | Bool | Document-wide RTL enabled |
| `hasPageNumbers()` | Bool | Page number field present |
| `getPageNumberAlign()` | String | Page number alignment |
| `getCustomProperties()` | List | All custom document properties |
| `getCustomProperty(name)` | String | A single custom property by name |

---

## 30. Content Query Methods

### 30.1 Summary and Block Listing

```ring
? reader.summary()        # human-readable content overview
reader.listBlocks()       # numbered block list with type + preview
blocks = reader.getBlocks()  # raw list of all parsed blocks
```

### 30.2 Paragraphs and Runs

| Method | Returns | Description |
|---|---|---|
| `getHeadings()` | List of `[:level, :text, :style]` | All heading blocks (levels 1–6) |
| `getParagraphs()` | List of paragraph blocks | All body paragraphs with formatting |
| `getRTLParagraphs()` | List | Right-to-left paragraphs |
| `getCommentedParagraphs()` | List of `[:text, :commentText, :commentAuthor]` | Paragraphs with attached comments |
| `getKeepNextParagraphs()` | List | Paragraphs with keep-with-next set |
| `getPageBreakBeforeParagraphs()` | List | Paragraphs preceded by a page break |
| `getBorderedParagraphs()` | List | Paragraphs with a border box |
| `getNamedStyleParagraphs(name)` | List | Paragraphs using a specific named paragraph style |
| `getOutlinedParagraphs()` | List | Paragraphs with an explicit outline level |
| `getNumberedHeadings()` | List | Headings that are part of a numbered list |
| `getTabbedParagraphs()` | List | Paragraphs containing tab characters |
| `getHiddenRuns()` | List of `[:text]` | Runs with `vanish` (hidden text) |
| `getCapsRuns()` | List of `[:text, :type]` | Runs with allCaps or smallCaps |
| `getDStrikeRuns()` | List of `[:text]` | Runs with double-strikethrough |
| `getRunBorderRuns()` | List | Runs with a character-level border |
| `getRunLanguages()` | List of `[:lang, :text]` | Runs with explicit language tags |
| `getCharStyleRuns()` | List of `[:styleName, :text]` | Runs referencing a named character style |
| `getCharStyles()` | List | All character style names found in the document |
| `getTextBoxes()` | List of `[:text]` | Floating text box content |

### 30.3 Lists, Hyperlinks, Captions, Bookmarks

| Method | Returns | Description |
|---|---|---|
| `getListItems()` | List of `[:text, :level, :isBullet]` | All list items with nesting level |
| `getListRestartPoints()` | List | List items where numbering restarts |
| `getHyperlinks()` | List of `[:text, :url]` | All external hyperlinks |
| `getCaptions()` | List of `[:text, :label, :seqNum]` | SEQ figure/table captions |
| `getBookmarks()` | List of `[:name, :text]` | Named bookmark anchors |

### 30.4 Footnotes and Endnotes

| Method | Returns | Description |
|---|---|---|
| `getFootnotes()` | List of note blocks | All footnotes |
| `getEndnotes()` | List of note blocks | All endnotes |

Each note block contains:

```
block[:id]       — note reference number
block[:refText]  — anchor text in the document body
block[:noteText] — plain concatenated note body text
block[:runs]     — list of formatted runs, each: [:text, :bold, :italic,
                    :underline, :color, :font, :size]
```

```ring
footnotes = reader.getFootnotes()
for fn in footnotes
    ? "Note " + fn[:id] + ": " + fn[:noteText]
    for r in fn[:runs]
        if r[:bold] = true  ? "  BOLD: " + r[:text]  ok
    next
next
```

### 30.5 Comments, Fields, Charts

| Method | Returns | Description |
|---|---|---|
| `getComments()` | List of `[:id, :text, :author]` | All comment entries |
| `getMergeFields()` | List of `[:name]` | All MERGEFIELD entries |
| `getMergeFieldNames()` | List of Strings | MERGEFIELD names as a flat list |
| `getFields()` | List of `[:fieldType, :fieldResult]` | All Word fields (DATE, PAGE, TOC, SEQ, …) |
| `getCharts()` | List of `[:type, :title, :widthCm, :heightCm]` | Embedded chart metadata |
| `getChartData()` | List with full series data | Charts with categories and series values |

### 30.6 Form Fields

| Method | Returns | Description |
|---|---|---|
| `getFormFields()` | List | All form fields (any type) |
| `getCheckboxes()` | List of `[:label, :checked]` | Checkbox controls |
| `getDropdowns()` | List of `[:label, :choices, :default]` | Dropdown controls |
| `getTextInputs()` | List of `[:label, :value]` | Text input controls |

---

## 31. Table Query Methods

| Method | Returns | Description |
|---|---|---|
| `getTables()` | List of table blocks | All tables with full row/cell data |
| `getTableLayouts()` | List of layout summaries | Dimension and style summary per table |
| `getTableStyle(blockIndex)` | List | Style info for a specific table |
| `getTableRowProperties()` | List | Row heights and header-repeat flags |
| `getNestedTables()` | List | Tables nested inside cell content |
| `getCellsWithBorders()` | List | Cells with explicit border overrides |
| `getCellsWithPadding()` | List | Cells with custom padding |

**Table block structure:**

```
block[:rows]        — list of rows; each row is a list of cell dicts
block[:colWidths]   — column widths in twips
block[:rowHeights]  — per-row heights in twips (0 = auto)
block[:rowHRules]   — per-row height rule strings ("atLeast" / "exact")
block[:rowIsHeader] — per-row boolean (tblHeader repeat flag)
block[:borderStyle] — table-level border style
block[:headerBg]    — header background colour hex
block[:evenRowBg]   — even-row stripe colour
```

**Cell dict structure:**

```
cell[:text]         — concatenated plain text
cell[:runs]         — list of runs: [:text, :bold, :italic, :color, :font,
                       :size, :lang, :styleName]
cell[:bgColor]      — background colour hex
cell[:align]        — text alignment
cell[:vAlign]       — vertical alignment
cell[:colspan]      — column span count
cell[:rowspan]      — row span count
cell[:borderStyle]  — legacy all-sides border style
cell[:borderSides]  — list of per-side entries: [:side, :style, :color, :size]
cell[:textDir]      — "btLr", "tbRl", or "" (normal)
cell[:padTop/Bottom/Left/Right] — padding in twips
```

---

## 32. Image Query Methods

| Method | Returns | Description |
|---|---|---|
| `getImages()` | List of image blocks | All inline and floating images |
| `getImagesWithAlt()` | List of `[:path, :altText, :widthCm, :heightCm]` | Images with alt text |
| `getFloatingImages()` | List | Floating images with position and wrap data |
| `getFloatingImageWrapTypes()` | List of `[:wrapType, :wrapSide]` | Wrap type summary |
| `getLandscapeSections()` | List | Sections using landscape orientation |

Image blocks include crop data when present:

```
block[:widthCm]   — display width in cm
block[:heightCm]  — display height in cm
block[:floating]  — true for floating images
block[:altText]   — accessibility description
block[:cropL]     — left crop in percent
block[:cropR]     — right crop in percent
block[:cropT]     — top crop in percent
block[:cropB]     — bottom crop in percent
```

---

## 33. Section and Layout Queries

| Method | Returns | Description |
|---|---|---|
| `getSectionLayouts()` | List | Per-section column, header/footer info |
| `getSectionBreaks()` | List | Section break points with type and column settings |
| `getSectionHeaders()` | List | Section-level header/footer text |

`getSectionLayouts()` entry structure:

```
item[:breakType]     — "nextPage", "continuous", "evenPage", "oddPage"
item[:numColumns]    — column count in this section
item[:columnSpaceCm] — column gap in cm
item[:sectHeader]    — section-specific header text
item[:sectFooter]    — section-specific footer text
```

`getSectionBreaks()` entry structure:

```
item[:breakType]     — break type string
item[:numColumns]    — column count after the break
item[:columnSpaceCm] — column spacing in cm
```

---

## 34. Round-Trip: `toWriter()`

`toWriter()` rebuilds the entire parsed document as a `WordWriter` instance. You can
then add new content, modify page settings, or just call `save()` to produce a clean
copy.

```ring
reader = new WordReader("template.docx")
reader.loadDocx()

writer = reader.toWriter()

# Append new content
writer.addPageBreak()
writer.addHeading("Appendix", 1)
writer.addParagraph("Content appended after round-trip reconstruction.", [])

writer.save("extended_output.docx")
reader.cleanup()
```

**Round-trip fidelity table:**

| Element | Preserved |
|---|---|
| Paragraphs | Text, bold/italic/underline/strike, colour, font, size, alignment, indent, spacing |
| Run properties | Language (`:lang`), character style (`:charStyle`), superscript/subscript |
| Headings 1–6 | Level and text |
| Lists | Bullet and numbered, nesting levels, restart points |
| Tables | All cells, merges, bg colours, per-side borders, row heights, text direction |
| Images | Inline and floating, dimensions, wrap type, alt text, crop percentages |
| Footnotes | Full run-level formatting (bold, italic, colour, size per run) |
| Endnotes | Full run-level formatting |
| Hyperlinks | URL and display text |
| Bookmarks | Name and position |
| Section breaks | Break type, column count, column spacing |
| Page properties | Size, margins, orientation, background, page border |
| Headers/Footers | Default, first-page, even-page |
| Paragraph control | keepNext, keepLines, `widowControl=false`, `noHyphenate=true` |
| Comments | Text and author |
| Fields | DATE, PAGE, TOC, SEQ, and other auto-updating fields |
| Watermarks | Text and image |
| RTL | RTL paragraphs and document RTL mode |

---

## 35. Reader Quick Reference

```ring
load "docxlib.ring"

reader = new WordReader("doc.docx")
reader.loadDocx()

# Metadata
? reader.getTitle()  ? reader.getAuthor()  ? reader.getPageWidthCm()

# Overview
? reader.summary()
reader.listBlocks()

# Content
headings   = reader.getHeadings()
tables     = reader.getTables()
images     = reader.getImages()
footnotes  = reader.getFootnotes()
charts     = reader.getCharts()
hyperlinks = reader.getHyperlinks()
listItems  = reader.getListItems()

# Run intelligence
langs    = reader.getRunLanguages()
styles   = reader.getCharStyleRuns()
hiddenTx = reader.getHiddenRuns()

# Table detail
layouts  = reader.getTableLayouts()
borders  = reader.getCellsWithBorders()
rowProps = reader.getTableRowProperties()

# Layout
sections = reader.getSectionLayouts()
breaks   = reader.getSectionBreaks()

# Form fields
boxes    = reader.getCheckboxes()
drops    = reader.getDropdowns()
inputs   = reader.getTextInputs()

# Round-trip
reader.save("output.docx")
reader.cleanup()
```

---

## 36. Quick Reference Index

### Writer — Document Setup

| Methods | Section |
|---|---|
| `setTitle()`, `setAuthor()`, `setDefaultFont()` | §2.1 |
| `setPageSize()`, `setCustomPageSize()`, `setOrientation()` | §2.2 |
| `setMargins()`, `setMarginsInches()`, `setNarrowMargins()`, `setWideMargins()` | §2.2 |
| `setColumns()`, `setTwoColumns()`, `setThreeColumns()` | §2.3 |
| `addColumnBreak()`, `addSectionBreak(type, options)` | §2.3 |
| `addLandscapeStart()`, `addLandscapeEnd()` | §2.3 |
| `setTheme()` | §2.4 |
| `defineStyle()` | §2.5 |

### Writer — Headers, Watermarks, Borders

| Methods | Section |
|---|---|
| `setHeader()`, `setFooter()`, `setHeaderFooter()`, `showPageNumbers()` | §3 |
| `setFirstPageDifferent()`, `setFirstPageHeader()`, `setFirstPageFooter()` | §3 |
| `setEvenAndOddHeaders()`, `setEvenPageHeader()`, `setEvenPageFooter()` | §3 |
| `setWatermark()`, `setWatermarkText()`, `removeWatermark()` | §8 |
| `setImageWatermark()`, `removeImageWatermark()` | §7.3 |
| `setPageBorder()`, `setSimplePageBorder()`, `removePageBorder()` | §9 |
| `setPageBackground()` | §9 |

### Writer — Text, Lists, Tables, Images

| Methods | Section |
|---|---|
| `addParagraph()`, `addRichParagraph()` | §4.1, §4.3 |
| `addHeading()`, `addTitle()`, `addSubtitle()` | §4.2 |
| `bold()`, `italic()`, `underline()`, `colored()`, `styled()` | §4.4 |
| `centered()`, `rightAligned()`, `justified()` | §4.4 |
| `addBlockQuote()`, `addCaption()`, `addAbstract()`, `addKeywords()` | §4.4 |
| `addShadedParagraph()`, `addBorderedParagraph()` | §4.5 |
| `addEmptyParagraph()`, `addHorizontalLine()`, `addPageBreak()`, `addLineBreak()` | §4.4 |
| `:widowControl=false`, `:noHyphenate=true` | §4.6 |
| `addBulletList()`, `addNumberedList()` | §5 |
| `addNestedBulletList()`, `addNestedNumberedList()` | §5 |
| `addTable()`, `addSimpleTable()`, `addStyledTable()` | §6.1 |
| `:rowHeights`, `:rowHRule`, `:repeatHeader`, `:conditionalRules` | §6.2 |
| `wordCell()`, `mergeCell()`, `wordMergeCell()` | §6.3 |
| `.setBorderSide()`, `.setTopBorder()`, `.setBottomBorder()` | §6.4 |
| `.setLeftBorder()`, `.setRightBorder()`, `.setNoBorder()` | §6.4 |
| `.setTextDir()` | §6.4 |
| `addImage()`, `addImageCentered()` — `:cropLeft/Right/Top/Bottom` | §7.1 |
| `addFloatingImage()` | §7.2 |
| `addFootnote()`, `addEndnote()` — plain or rich run-list content | §10 |
| `registerFootnote()`, `registerEndnote()` | §10 |

### Writer — Navigation, Automation, Charts

| Methods | Section |
|---|---|
| `addHyperlink()` | §11.1 |
| `addBookmarkedParagraph()`, `addBookmark()`, `addCrossRef()` | §11.2–3 |
| `addTableOfContents()` | §12 |
| `addFigureCaption()`, `addTableCaption()` | §13.1 |
| `addTableOfFigures()`, `addTableOfTables()`, `addTableOfSeq()` | §13.2 |
| `addTabbedParagraph()` | §14.1 |
| `enableLineNumbers()`, `disableLineNumbers()` | §14.2 |
| `addCommentedParagraph()`, `addComment()` | §15 |
| `addMergeFieldParagraph()` | §16 |
| `setMergeTemplate()`, `mergeRecord()`, `mergeAll()`, `clearMergeTemplate()` | §24 |
| `addCheckbox()`, `addDropdown()`, `addTextInput()` | §17 |
| `addShape()`, `addRect()`, `addEllipse()`, `addLine()`, `addDiamond()` | §18 |
| `addTextBox()` | §19 |
| `setDocumentRTL()`, `addRTLParagraph()` | §20 |
| `save()` | §21 |
| `addColumnChart()`, `addBarChart()`, `addLineChart()`, `addAreaChart()` | §22 |
| `addPieChart()`, `addDoughnutChart()`, `addScatterChart()`, `addBubbleChart()` | §22 |
| `setChartDefaults()`, `clearChartDefaults()` | §22.3 |
| `:showDataTable`, `:dataTableShowKeys` | §25 |

### Reader — All Methods

| Methods | Section |
|---|---|
| `new WordReader(path)`, `loadDocx()` | §28 |
| `save(path)`, `cleanup()`, `toWriter()` | §28 |
| `getTitle()`, `getAuthor()`, `getPageWidthCm()`, … | §29 |
| `hasPageBorder()`, `isDocumentRTL()`, `hasTOC()`, `hasWatermark()` | §29 |
| `summary()`, `listBlocks()`, `getBlocks()` | §30.1 |
| `getHeadings()`, `getParagraphs()`, `getRTLParagraphs()` | §30.2 |
| `getRunLanguages()`, `getCharStyleRuns()`, `getCharStyles()` | §30.2 |
| `getHiddenRuns()`, `getCapsRuns()`, `getDStrikeRuns()` | §30.2 |
| `getListItems()`, `getListRestartPoints()` | §30.3 |
| `getHyperlinks()`, `getCaptions()`, `getBookmarks()` | §30.3 |
| `getFootnotes()`, `getEndnotes()` | §30.4 |
| `getComments()`, `getMergeFields()`, `getFields()` | §30.5 |
| `getCharts()`, `getChartData()` | §30.5 |
| `getCheckboxes()`, `getDropdowns()`, `getTextInputs()` | §30.6 |
| `getTables()`, `getTableLayouts()`, `getTableStyle()` | §31 |
| `getTableRowProperties()`, `getNestedTables()` | §31 |
| `getCellsWithBorders()`, `getCellsWithPadding()` | §31 |
| `getImages()`, `getImagesWithAlt()`, `getFloatingImages()` | §32 |
| `getFloatingImageWrapTypes()`, `getLandscapeSections()` | §32 |
| `getSectionLayouts()`, `getSectionBreaks()`, `getSectionHeaders()` | §33 |
