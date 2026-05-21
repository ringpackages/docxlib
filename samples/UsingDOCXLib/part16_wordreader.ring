/*
	DOCXLib - Demo (WordReader)
*/

load "docxlib.ring"

/*
    Demonstrates the complete read -> inspect -> modify -> save workflow.

    The demo is fully self-contained:
      Step 1  Generate a rich source document (product catalogue + report)
      Step 2  Open it with WordReader and print a structural summary
      Step 3  Inspect headings, paragraphs, tables programmatically
      Step 4  Perform modifications:
                a) Global find/replace (year update)
                b) Table cell edits   (price/status updates)
                c) Heading rename
                d) Paragraph rewrite
                e) Append a new summary section
      Step 5  Export the modified document via toWriter() + save()
      Step 6  Open the saved file, verify the changes round-tripped
*/

? "======================================================="
? "  DOCXLib -- WordReader Demo"
? "======================================================="
? ""

# =============================================================================
# STEP 1 -- Generate the source document
# =============================================================================
? "STEP 1: Generating source document (demo_wordreader_source.docx) ..."

src = new WordWriter()
src.setTitle("Ring Technology -- Product Catalogue 2024")
src.setAuthor("Ring Documentation Team")
src.setDefaultFont("Calibri", 11)
src.setTheme("Blue")
src.setHeader("Ring Technology -- Confidential -- 2024")
src.setFooter("Page | Ring Technology Ltd")

# Cover section
src.addHeading("Ring Technology", 1)
src.addHeading("Product Catalogue 2024", 2)
src.addParagraph(
    "This catalogue provides an overview of all Ring Technology products " +
    "available in the 2024 financial year. Pricing reflects the standard " +
    "enterprise tier. Volume discounts are available on request.", NULL)
src.addEmptyParagraph()

# Products section
src.addHeading("1. Core Products", 1)

src.addHeading("1.1 Ring Language Platform", 2)
src.addParagraph(
    "The Ring Language Platform is a multi-paradigm programming environment " +
    "featuring natural language syntax, built-in visual programming support, " +
    "and a comprehensive standard library. Version 1.26 is the current release.", NULL)

src.addTable([
    ["Product",                 "Version", "Licence",       "Unit Price"],
    ["Ring Language IDE",       "1.26",    "Annual",        "$1,200"],
    ["Ring Enterprise Licence", "1.26",    "Perpetual",     "$8,500"],
    ["Ring Developer Starter",  "1.26",    "Annual",        "$450"],
    ["Ring Education Bundle",   "1.26",    "Per-Seat/Year", "$150"]
], [:headerRow = true, :borderStyle = "single",
    :colWidths = [5.5, 2.0, 2.5, 2.5]])

src.addEmptyParagraph()

src.addHeading("1.2 PWCT Visual Programming Suite", 2)
src.addParagraph(
    "PWCT (Programming Without Coding Technology) is a visual programming " +
    "environment that generates Ring source code from graphical flowcharts. " +
    "Suitable for education, rapid prototyping, and non-developer audiences.", NULL)

src.addTable([
    ["Product",             "Version", "Licence",   "Unit Price"],
    ["PWCT Desktop",        "2.0",     "Perpetual", "$950"],
    ["PWCT Team Edition",   "2.0",     "Annual/5",  "$3,200"],
    ["PWCT2 Visual Studio", "2.0",     "Annual",    "$750"],
    ["PWCT Mobile SDK",     "1.5",     "Annual",    "$600"]
], [:headerRow = true, :borderStyle = "single",
    :colWidths = [5.5, 2.0, 2.5, 2.5]])

src.addEmptyParagraph()

src.addHeading("1.3 RingStar Cloud Platform", 2)
src.addParagraph(
    "RingStar is Ring Technology's cloud-native application platform. " +
    "It provides managed infrastructure, auto-scaling, and a REST API gateway " +
    "optimised for Ring applications.", NULL)

src.addTable([
    ["Plan",           "Instances", "Storage", "Support",  "Monthly"],
    ["Starter",        "2",         "50 GB",   "Community","$99"],
    ["Professional",   "10",        "500 GB",  "Email",    "$499"],
    ["Business",       "50",        "2 TB",    "Phone",    "$1,999"],
    ["Enterprise",     "Unlimited", "10 TB",   "Dedicated","$7,500"]
], [:headerRow = true, :borderStyle = "single",
    :colWidths = [3.0, 2.5, 2.0, 2.5, 2.5]])

src.addEmptyParagraph()

# Financial section
src.addHeading("2. Financial Summary", 1)
src.addParagraph(
    "The following table summarises annual revenue by product line for the " +
    "2024 fiscal year. All figures in USD thousands.", NULL)

src.addTable([
    ["Product Line",  "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024", "Total 2024"],
    ["Ring Platform", "320",     "415",     "389",     "502",     "1,626"],
    ["PWCT Suite",    "210",     "268",     "241",     "315",     "1,034"],
    ["RingStar",     "185",     "240",     "312",     "418",     "1,155"],
    ["Services",      "125",     "138",     "149",     "171",     "583"],
    ["TOTAL",         "840",     "1,061",   "1,091",   "1,406",   "4,398"]
], [:headerRow = true, :borderStyle = "single",
    :colWidths = [3.5, 2.0, 2.0, 2.0, 2.0, 2.5]])

src.addEmptyParagraph()

# Status section
src.addHeading("3. Product Status", 1)
src.addParagraph(
    "Current development and release status for all major products as of " +
    "the 2024 catalogue publication date.", NULL)

src.addTable([
    ["Product",             "Status",        "Next Release", "Notes"],
    ["Ring Language",       "RELEASED",      "Q2 2025",      "v1.27 in development"],
    ["PWCT Visual Suite",   "RELEASED",      "Q3 2025",      "Mobile beta Q1 2025"],
    ["RingStar Cloud",     "BETA",          "Q1 2025",      "GA expected March 2025"],
    ["Ring AI Toolkit",     "IN DEVELOPMENT","Q4 2025",      "LLM integration"],
    ["Ring Enterprise IDE", "PLANNING",      "Q1 2026",      "Roadmap confirmed"]
], [:headerRow = true, :borderStyle = "single",
    :colWidths = [4.0, 3.0, 2.5, 5.0]])

src.addEmptyParagraph()

# Notes
src.addHeading("4. Important Notes", 1)
src.addParagraph(
    "All pricing shown is exclusive of applicable taxes. Enterprise customers " +
    "should contact their account manager for volume pricing. Ring Technology " +
    "reserves the right to revise pricing at any time with 30 days notice.", NULL)
src.addParagraph(
    "This document is confidential and intended for internal distribution only. " +
    "Do not distribute outside Ring Technology without written authorisation.", NULL)

src.save("demo_wordreader_source.docx")
? "  Source document saved."
? ""

# =============================================================================
# STEP 2 -- Open with WordReader and print summary
# =============================================================================
? "STEP 2: Opening with WordReader ..."
wr = new WordReader("demo_wordreader_source.docx")
? wr.summary()

# =============================================================================
# STEP 3 -- Inspect structure programmatically
# =============================================================================
? "STEP 3: Inspecting document structure ..."
? ""
? "--- Block index ---"
wr.listBlocks()
? ""

? "--- All headings ---"
blocks = wr.getBlocks()
bLen = len(blocks)
for i = 1 to bLen
    b = blocks[i]
    if b[:type] = "heading"
        ? "  H" + string(b[:level]) + ": " + b[:text]
    ok
next
? ""

? "--- Table 4 contents (Product Status) ---"
tblIdx4 = wr.getTableBlockIndex(4)
if tblIdx4 > 0
    tbl = wr.getBlock(tblIdx4)
    rows = tbl[:rows]
    rowLen = len(rows)
    for ri = 1 to rowLen
        row = rows[ri]
        line = "  Row " + string(ri) + ":  "
        cLen = len(row)
        for ci = 1 to cLen
            cell = row[ci]
            ct = ""
            if isList(cell)  ct = cell[:text]
            else             ct = "" + cell
            ok
            line += ct + "  |  "
        next
        ? line
    next
ok
? ""

? "--- Find blocks containing BETA ---"
found = wr.findBlocks("BETA")
if len(found) > 0
    for idx in found
        b = wr.getBlock(idx)
        ? "  Block " + string(idx) + " (" + b[:type] + ")"
    next
else
    ? "  Not found."
ok
? ""

? "--- Specific cells ---"
? "  Table 1, Row 3, Col 1: " + wr.getCell(wr.getTableBlockIndex(1), 3, 1)
? "  Table 1, Row 3, Col 4: " + wr.getCell(wr.getTableBlockIndex(1), 3, 4)
? "  Table 4, Row 3, Col 2: " + wr.getCell(tblIdx4, 3, 2)
? ""

# =============================================================================
# STEP 4 -- Modify the document
# =============================================================================
? "STEP 4: Applying modifications ..."
? ""

wr.cTitle = substr(wr.cTitle,"2024","2025")
wr.cHeaderText = substr(wr.cHeaderText,"2024","2025")

# a) Global find/replace: update year throughout
? "  a) Global find/replace: 2024 -> 2025 ..."
wr.replaceText("2024", "2025")
? "     Verifying title heading: " + wr.getBlock(1)[:text]
? ""

# b) Table cell edits -- update prices in the Ring Platform table (table 1)
? "  b) Updating Ring Platform prices (table 1) ..."
t1 = wr.getTableBlockIndex(1)
wr.setCell(t1, 2, 4, "$1,400")   # Ring Language IDE
wr.setCell(t1, 3, 4, "$9,800")   # Ring Enterprise Licence
wr.setCell(t1, 4, 4, "$520")     # Ring Developer Starter
wr.setCell(t1, 5, 4, "$175")     # Ring Education Bundle
? "     New price for Ring Enterprise: " + wr.getCell(t1, 3, 4)
? ""

# c) Update status for RingStar (table 4, row 3: BETA -> RELEASED)
? "  c) Updating RingStar status: BETA -> RELEASED ..."
wr.setTableCell(4, 3, 2, "RELEASED")
wr.setTableCell(4, 3, 3, "Q1 2025")
wr.setTableCell(4, 3, 4, "GA released January 2025")
? "     New status: " + wr.getCell(wr.getTableBlockIndex(4), 3, 2)
? ""

# d) Rename a heading
? "  d) Updating catalogue title heading ..."
wr.setHeadingText(2, "Product Catalogue 2025 (Revised Edition)")
? "     New heading: " + wr.getBlock(2)[:text]
? ""

# e) Rewrite the intro paragraph
? "  e) Rewriting introduction paragraph ..."
introIdx = 0
for i = 1 to bLen
    if blocks[i][:type] = "paragraph" and len(blocks[i][:text]) > 50
        introIdx = i
        break
    ok
next
if introIdx > 0
    wr.setBlockText(introIdx,
        "This catalogue provides an updated overview of all Ring Technology " +
        "products available in the 2025 financial year. Pricing has been " +
        "revised to reflect current market conditions. Volume discounts " +
        "remain available on request and have been expanded for education.")
    ? "     Paragraph rewritten."
ok
? ""

# f) Append a new summary section
? "  f) Appending 2025 update summary section ..."
wr.appendPageBreak()
wr.appendHeading("5. 2025 Edition -- Change Summary", 1)
wr.appendParagraph(
    "This section documents all changes made to the catalogue in the 2025 " +
    "revised edition compared to the original 2024 publication.", NULL)

wr.appendTable([
    ["Change",                    "Section",       "Detail"],
    ["Year references updated",   "Throughout",    "All 2024 references changed to 2025"],
    ["Ring Platform pricing",     "Section 1.1",   "Prices increased by ~15%"],
    ["RingStar status",          "Section 3",     "BETA -> RELEASED; GA confirmed Jan 2025"],
    ["RingStar notes",           "Section 3",     "Updated to reflect GA release"],
    ["Introduction rewritten",    "Cover",         "Updated to reflect 2025 pricing policy"],
    ["Title updated",             "Cover",         "Revised Edition designation added"]
], [5.5, 3.0, 6.5])

wr.appendEmptyParagraph()
wr.appendParagraph(
    "Document revision prepared by the Ring Documentation Team. " +
    "For queries contact: docs@ringtech.example.com", NULL)
? "     Section appended."
? ""

# =============================================================================
# STEP 5 -- Save the modified document
# =============================================================================
? "STEP 5: Saving modified document (demo_wordreader_modified.docx) ..."
wr.save("demo_wordreader_modified.docx")
wr.cleanup()
? "  Saved and temp files cleaned up."
? ""

# =============================================================================
# STEP 6 -- Re-read the saved file to verify round-trip
# =============================================================================
? "STEP 6: Re-reading modified document to verify round-trip ..."
wr2 = new WordReader("demo_wordreader_modified.docx")
? wr2.summary()

? "--- Verification checks ---"
checkTitle = wr2.getTitle()
? "  Title metadata  : " + checkTitle

# Find the title heading
blocks2 = wr2.getBlocks()
b2Len = len(blocks2)
titleHeadingText = ""
for i = 1 to b2Len
    if blocks2[i][:type] = "heading" and blocks2[i][:level] = 2
        titleHeadingText = blocks2[i][:text]
        break
    ok
next
? "  Title heading   : " + titleHeadingText

# Verify RingStar status changed
t4b2 = wr2.getTableBlockIndex(4)
? "  RingStar status: " + wr2.getCell(t4b2, 3, 2)

# Verify price change
t1b2 = wr2.getTableBlockIndex(1)
? "  Enterprise price: " + wr2.getCell(t1b2, 3, 4)

# Check 2025 section heading exists
foundH5 = false
for i = 1 to b2Len
    b = blocks2[i]
    if b[:type] = "heading" and substr(b[:text], "2025 Edition") > 0
        foundH5 = true
    ok
next
? "  Change summary section present: " + foundH5

# Check no stray 2024 references remain in headings
stray2024 = false
for i = 1 to b2Len
    b = blocks2[i]
    bTxt = b[:text]
    if bTxt = NULL  bTxt = ""  ok
    if substr(bTxt, "2024") > 0 and b[:type] = "heading"
        ? "  WARNING: 2024 still in heading " + string(i) + ": " + bTxt
        stray2024 = true
    ok
next
if !stray2024
    ? "  No 2024 references in headings: true"
ok

wr2.cleanup()
? ""
? "======================================================="
? "  Demo complete."
? "  Files: demo_wordreader_source.docx"
? "         demo_wordreader_modified.docx"
? "======================================================="
