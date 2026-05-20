/*
	DOCXLib - Demo (Table of Contents)
*/

load "docxlib.ring"

/*
    Demonstrates:
      * addTableOfFigures(title) - collects all Figure N captions
      * addTableOfTables(title)  - collects all Table N captions
      * addTableOfSeq(label, title) - generic version for any SEQ label
      * Both lists live-update in Word with F9 / Update Fields
*/

doc = new WordWriter()
doc.setTitle("DOCXLib - Table of Figures & Tables Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setTheme("Blue")
doc.setHeader("DOCXLib - Table of Figures & Tables")
doc.showPageNumbers("right")

# -----------------------------------------------------------
# PAGE 1 - Table of Figures + Table of Tables (front matter)
# -----------------------------------------------------------

doc.addHeading("Document Overview", 1)
doc.addParagraph("This document demonstrates the  Table of Figures and Table of Tables feature. The lists below are live Word fields - press F9 or right-click and choose Update Field to populate page numbers.", NULL)
doc.addParagraph("", NULL)

# --- Table of Figures ---
doc.addTableOfFigures("List of Figures")
doc.addParagraph("", NULL)

# --- Table of Tables ---
doc.addTableOfTables("List of Tables")
doc.addParagraph("", NULL)

doc.addPageBreak()

# -----------------------------------------------------------
# PAGE 2 - Figures with captions
# -----------------------------------------------------------

doc.addHeading("Chapter 1 - Ring Language Architecture", 1)
doc.addParagraph("This chapter provides an overview of the Ring language runtime and its major components.", NULL)
doc.addParagraph("", NULL)

doc.addHeading("1.1 Virtual Machine", 2)
doc.addParagraph("The Ring VM is a stack-based bytecode interpreter written in C. It processes compiled Ring bytecode and manages the execution stack, variable scopes, and garbage collection.", NULL)
doc.addParagraph("", NULL)

# Figure 1 - placeholder box (grey rect)
doc.addRect([:width=12, :height=4, :fillColor="E8E8E8", :lineColor="AAAAAA",
             :text="[ Figure Placeholder: Ring VM Architecture Diagram ]",
             :textColor="555555", :textBold=false])
doc.addFigureCaption("Ring Virtual Machine architecture overview")
doc.addParagraph("", NULL)

doc.addHeading("1.2 Compiler Pipeline", 2)
doc.addParagraph("Ring source code passes through a multi-stage compiler pipeline before execution. Each stage transforms the representation closer to executable bytecode.", NULL)
doc.addParagraph("", NULL)

doc.addRect([:width=12, :height=3.5, :fillColor="E8F4E8", :lineColor="88AA88",
             :text="[ Figure Placeholder: Compiler Pipeline Stages ]",
             :textColor="336633", :textBold=false])
doc.addFigureCaption("Multi-stage Ring compiler pipeline: Lexer → Parser → Code Generation")
doc.addParagraph("", NULL)

doc.addPageBreak()

# -----------------------------------------------------------
# PAGE 3 - More figures + tables
# -----------------------------------------------------------

doc.addHeading("Chapter 2 - DOCXLib Feature Coverage", 1)
doc.addParagraph("DOCXLib provides comprehensive Word document generation from Ring code. The following tables and figures summarise its capabilities.", NULL)
doc.addParagraph("", NULL)

# Table 1
doc.addTableCaption("DOCXLib version history and major milestones")
doc.addTable(
    [
        ["Version", "Release Focus",         "Key Features Added"],
        ["0.1–0.5", "Core document elements","Paragraphs, tables, images, lists, styles"],
        ["0.6",     "Layout & formatting",   "Text boxes, watermarks, multi-column, RTL"],
        ["0.7",     "Headers & structure",   "Odd/even headers, page background, page borders"],
        ["0.8",     "Reference features",    "Bookmarks, cross-refs, tab stops, line numbers"],
        ["0.9",     "Review & automation",   "Comments, SEQ captions, mail merge fields"],
        ["0.95",    "Rich content",          "Content controls, drawing shapes, themes"],
        ["1.0",     "Navigation lists",      "Table of Figures, Table of Tables"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "1E6BB5", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "DEE9F7", :colWidths = [2, 4, 9.5]
    ]
)
doc.addParagraph("", NULL)

doc.addHeading("2.1 Feature Distribution", 2)
doc.addParagraph("The chart below illustrates how library features are distributed across functional categories.", NULL)
doc.addParagraph("", NULL)

doc.addRect([:width=12, :height=4.5, :fillColor="F0F4FF", :lineColor="4472C4",
             :text="[ Figure Placeholder: Feature Distribution Chart ]",
             :textColor="1E6BB5", :textBold=false])
doc.addFigureCaption("DOCXLib feature distribution by category")
doc.addParagraph("", NULL)

# Table 2
doc.addTableCaption("Feature count by category in DOCXLib")
doc.addTable(
    [
        ["Category",            "Feature Count", "Example Methods"],
        ["Text & Paragraphs",   "12",  "addParagraph, addHeading, addHyperlink"],
        ["Tables",              "8",   "addTable, addTableCaption"],
        ["Images & Shapes",     "10",  "addImage, addShape, addRect, addEllipse"],
        ["Headers & Footers",   "8",   "setHeader, setFooter, showPageNumbers"],
        ["Lists",               "6",   "addBulletList, addNumberedList"],
        ["References",          "8",   "addBookmark, addCrossRef, addTOC, addTableOfFigures"],
        ["Form Controls",       "3",   "addCheckbox, addDropdown, addTextInput"],
        ["Document Settings",   "10",  "setTheme, setPageSize, setMargins"],
        ["Total",               "65+", "-"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "375623", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "E2EFDA", :colWidths = [4.5, 3, 8]
    ]
)
doc.addParagraph("", NULL)

doc.addPageBreak()

# -----------------------------------------------------------
# PAGE 4 - More figures on a new page
# -----------------------------------------------------------

doc.addHeading("Chapter 3 - Usage Examples", 1)
doc.addParagraph("", NULL)

doc.addHeading("3.1 Document Generation Workflow", 2)
doc.addParagraph("The typical workflow for generating a Word document with DOCXLib:", NULL)
doc.addParagraph("", NULL)

doc.addRect([:width=12, :height=3, :fillColor="FFF3E0", :lineColor="E65100",
             :text="[ Figure Placeholder: Document Generation Workflow Diagram ]",
             :textColor="6D3B00", :textBold=false])
doc.addFigureCaption("Typical DOCXLib document generation workflow")
doc.addParagraph("", NULL)

doc.addHeading("3.2 Table of Figures API", 2)
doc.addParagraph("The three methods:", NULL)
doc.addParagraph("", NULL)

# Table 3
doc.addTableCaption("Methods for auto-generated caption lists")
doc.addTable(
    [
        ["Method",                           "Purpose"],
        ["addTableOfFigures(title)",          "Collects all Figure N captions"],
        ["addTableOfTables(title)",           "Collects all Table N captions"],
        ["addTableOfSeq(seqLabel, title)",    "Generic - any SEQ label (Equation, Chart, etc.)"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "7030A0", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "EDE7F6", :colWidths = [7, 8.5]
    ]
)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "How it works: Word uses a TOC field with the \c switch: " +
    '  TOC \c "Figure" \h' + NL +
    "This instructs Word to scan the document for all SEQ Figure fields, " +
    "collect their caption text and page numbers, and render them as a " +
    "dot-leader list - exactly what Insert > Table of Figures produces. " +
    "Press F9 to update after document generation.",
    [:borderStyle = "single", :borderColor = "7030A0", :bgColor = "EDE7F6",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)

# -----------------------------------------------------------
# SAVE
# -----------------------------------------------------------
filename = "demo_tableofcontents.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
