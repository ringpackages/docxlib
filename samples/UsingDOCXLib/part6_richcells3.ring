/*
	DOCXLib - Demo (RichCells3)
*/

load "docxlib.ring"

/*
    Column Widths & Cell Quotes/Captions Demo
    Demonstrates custom column widths and block quotes/captions inside table cells.
*/

? "=============================================="
? "   DOCXLib - Column Widths & Quotes Demo"
? "=============================================="

# =====================================================================
# Demo 1: Custom Column Widths
# =====================================================================
? "Demo 1: Custom column widths..."

doc = new WordWriter()
doc.setTitle("Custom Column Widths")

doc.addHeading("Custom Column Widths", 1)

# --- Table 1: Simple custom widths ---
doc.addHeading("Narrow ID + Wide Description", 2)
doc.addParagraph("First column 2cm, second 6cm, third 3cm, fourth 5.5cm:", NULL)

doc.addTable([
    ["ID", "Description", "Status", "Assigned To"],
    ["1", "Design new user dashboard with responsive layout", "In Progress", "Ahmed"],
    ["2", "Fix login timeout bug on mobile devices", "Done", "Sara"],
    ["3", "Write API documentation for v2.0 endpoints", "Pending", "Omar"],
    ["4", "Optimize database queries for reports module", "In Progress", "Fatima"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "1565C0",
    :evenRowBgColor = "E3F2FD",
    :headerHeight = 600,
    :headerVAlign = "center",
    :cellVAlign = "center",
    :colWidths = [2, 6, 3, 5.5]
])

doc.addEmptyParagraph()

# --- Table 2: Narrow label + wide value (form-like layout) ---
doc.addHeading("Form-Style Layout (Label : Value)", 2)

doc.addTable([
    ["Field", "Value"],
    ["Full Name", "Mahmoud Fayed"],
    ["Email", "msfclipper@yahoo.com"],
    ["Position", "Programming Language Designer & Researcher"],
    ["Languages Created", "Ring, PWCT, PWCT2, Supernova"],
    ["Research Areas", "Compiler Design, Visual Programming, Machine Learning"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "37474F",
    :headerHeight = 600,
    :headerVAlign = "center",
    :colWidths = [4.5, 12]
])

doc.addEmptyParagraph()

# --- Table 3: Mixed with WordCell and custom widths ---
doc.addHeading("Product Catalog with Custom Widths", 2)

h1 = wordCell("SKU", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Product Name", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Description", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])
h4 = wordCell("Price", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])

row1a = wordCell("RL-001", [:align = "center"])
row1a.setVerticalAlign("center")
row1b = new WordCell()
row1b.addRun("Ring Language Book", [:bold = true])
row1b.setVerticalAlign("center")
row1c = new WordCell()
row1c.addRun("Comprehensive guide to the Ring programming language covering all paradigms, libraries, and advanced features with practical examples.", [])
row1d = wordCell("$0.00", [:bold = true, :color = "green", :align = "center"])
row1d.setVerticalAlign("center")

row2a = wordCell("PW-002", [:align = "center"])
row2a.setVerticalAlign("center")
row2b = new WordCell()
row2b.addRun("PWCT Course", [:bold = true])
row2b.setVerticalAlign("center")
row2c = new WordCell()
row2c.addRun("Online video course for Programming Without Coding Technology with hands-on exercises.", [])
row2d = wordCell("$0.00", [:bold = true, :color = "green", :align = "center"])
row2d.setVerticalAlign("center")

doc.addTable([
    [h1, h2, h3, h4],
    [row1a, row1b, row1c, row1d],
    [row2a, row2b, row2c, row2d]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "center",
    :colWidths = [2, 4, 8, 2.5]
])

if doc.save("demo_richcells_widths.docx")
    ? "  Created: demo_richcells_widths.docx"
else
    ? "  FAILED: demo_richcells_widths.docx"
ok

# =====================================================================
# Demo 2: Block Quotes in Table Cells
# =====================================================================
? "Demo 2: Block quotes in cells..."

doc = new WordWriter()
doc.setTitle("Cell Block Quotes")

doc.addHeading("Block Quotes Inside Table Cells", 1)
doc.addParagraph("Cells can contain styled block quotes with left border accent.", NULL)

h1 = wordCell("Topic", [:bold = true, :color = "white", :bgColor = "1B5E20", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Content", [:bold = true, :color = "white", :bgColor = "1B5E20", :align = "center", :verticalAlign = "center"])

topic1 = new WordCell()
topic1.addRun("Programming", [:bold = true, :size = 12, :color = "1B5E20"])
topic1.setVerticalAlign("center")
topic1.setAlign("center")

content1 = new WordCell()
content1.addRun("Why Ring Language?", [:bold = true, :size = 11])
content1.addCellBlockQuote("Ring is designed to be a practical general-purpose language that is simple, natural, and encourages organized programming.")
content1.addRun("Ring supports multiple programming paradigms:", [:bold = true])
bulletId1 = doc.registerCellList(WORD_LIST_BULLET)
content1.addCellBulletList(["Procedural programming", "Object-Oriented programming", "Functional programming", "Declarative programming"], bulletId1)
content1.addCellBlockQuote("The language is designed to produce small executables and can be embedded in C/C++ projects with just a few files.")

topic2 = new WordCell()
topic2.addRun("Research", [:bold = true, :size = 12, :color = "1B5E20"])
topic2.setVerticalAlign("center")
topic2.setAlign("center")

content2 = new WordCell()
content2.addRun("Visual Programming Research:", [:bold = true, :size = 11])
content2.addCellBlockQuote("Visual programming reduces the barrier to entry for new programmers by representing code as visual elements rather than text.")
content2.addRun("Key findings from published papers demonstrate significant improvements in learning outcomes when visual tools are combined with traditional coding.", [])

doc.addTable([
    [h1, h2],
    [topic1, content1],
    [topic2, content2]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :colWidths = [3, 13.5],
    :cellVAlign = "top"
])

if doc.save("demo_richcells_quotes.docx")
    ? "  Created: demo_richcells_quotes.docx"
else
    ? "  FAILED: demo_richcells_quotes.docx"
ok

# =====================================================================
# Demo 3: Captions in Table Cells
# =====================================================================
? "Demo 3: Captions in cells..."

doc = new WordWriter()
doc.setTitle("Cell Captions")

doc.addHeading("Captions Inside Table Cells", 1)
doc.addParagraph("Captions provide centered, italic descriptive text, useful for labeling images or data within cells.", NULL)

h1 = wordCell("Chart", [:bold = true, :color = "white", :bgColor = "0D47A1", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Summary", [:bold = true, :color = "white", :bgColor = "0D47A1", :align = "center", :verticalAlign = "center"])

# Row 1 - Sales chart placeholder
chart1 = new WordCell()
chart1.addRun("Q1 2024", [:bold = true, :size = 18, :color = "1565C0"])
chart1.addLineBreak()
chart1.addRun("$1.2M Revenue", [:size = 14])
chart1.setAlign("center")
chart1.setBgColor("E3F2FD")
chart1.addCellCaption("Figure 1: Q1 2024 Revenue Overview")

summary1 = new WordCell()
summary1.addRun("Key Metrics:", [:bold = true])
numId1 = doc.registerCellList(WORD_LIST_NUMBER)
summary1.addCellNumberedList(["Revenue up 23% YoY", "New customers: 450", "Retention rate: 94%"], numId1)
summary1.addCellCaption("Source: Internal Sales Dashboard, March 2024")
summary1.setVerticalAlign("center")

# Row 2
chart2 = new WordCell()
chart2.addRun("Q2 2024", [:bold = true, :size = 18, :color = "2E7D32"])
chart2.addLineBreak()
chart2.addRun("$1.5M Revenue", [:size = 14])
chart2.setAlign("center")
chart2.setBgColor("E8F5E9")
chart2.addCellCaption("Figure 2: Q2 2024 Revenue Overview")

summary2 = new WordCell()
summary2.addRun("Key Metrics:", [:bold = true])
numId2 = doc.registerCellList(WORD_LIST_NUMBER)
summary2.addCellNumberedList(["Revenue up 25% QoQ", "New customers: 520", "Retention rate: 96%"], numId2)
summary2.addCellCaption("Source: Internal Sales Dashboard, June 2024")
summary2.setVerticalAlign("center")

doc.addTable([
    [h1, h2],
    [chart1, summary1],
    [chart2, summary2]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :colWidths = [7, 9.5],
    :cellVAlign = "top"
])

if doc.save("demo_richcells_captions.docx")
    ? "  Created: demo_richcells_captions.docx"
else
    ? "  FAILED: demo_richcells_captions.docx"
ok

# =====================================================================
# Demo 4: Combined - Project Report Table
# =====================================================================
? "Demo 4: Combined project report table..."

doc = new WordWriter()
doc.setTitle("Project Status Report")

doc.addHeading("Project Status Report", 1)
doc.addParagraph("A realistic example combining column widths, quotes, captions, lists, and rich formatting.", NULL)

# Register resources
bulletId = doc.registerCellList(WORD_LIST_BULLET)
numId = doc.registerCellList(WORD_LIST_NUMBER)
linkRing = doc.registerHyperlink("https://ring-lang.github.io")

# Header
h1 = wordCell("#", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Project", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Details & Progress", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h4 = wordCell("Status", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])

# Row 1 - Ring 1.22
num1 = wordCell("1", [:bold = true, :align = "center"])
num1.setVerticalAlign("center")

proj1 = new WordCell()
proj1.addRun("Ring 1.22", [:bold = true, :size = 12, :color = "1565C0"])
proj1.addLineBreak()
proj1.addRun("Programming Language", [:italic = true, :size = 9, :color = "gray"])
proj1.setVerticalAlign("center")

details1 = new WordCell()
details1.addRun("Release Highlights:", [:bold = true])
details1.addCellBulletList(["Closures support", "RingSlint UI toolkit", "Rust language bindings", "Archive and Proc packages"], bulletId)
details1.addCellBlockQuote("Ring 1.22 represents a major step forward with modern language features while maintaining backward compatibility.")
details1.addRun("Links:", [:bold = true])
details1.addCellHyperlink("Ring Official Website", linkRing)
details1.addCellCaption("Released: January 2025")

status1 = new WordCell()
status1.addRun("RELEASED", [:bold = true, :color = "green", :size = 11])
status1.setAlign("center")
status1.setVerticalAlign("center")
status1.setBgColor("E8F5E9")

# Row 2 - Research paper
num2 = wordCell("2", [:bold = true, :align = "center"])
num2.setVerticalAlign("center")

proj2 = new WordCell()
proj2.addRun("Research Paper", [:bold = true, :size = 12, :color = "E65100"])
proj2.addLineBreak()
proj2.addRun("MDPI Electronics", [:italic = true, :size = 9, :color = "gray"])
proj2.setVerticalAlign("center")

details2 = new WordCell()
details2.addRun("Submission Progress:", [:bold = true])
numId2 = doc.registerCellList(WORD_LIST_NUMBER)
details2.addCellNumberedList(["Initial submission", "Reviewer feedback received", "Revisions completed", "Resubmission done"], numId2)
details2.addCellBlockQuote("The paper presents a novel approach to prompt-driven development using Ring language for TUI framework creation.")
details2.addCellCaption("Journal: MDPI Electronics, Impact Factor: 2.6")

status2 = new WordCell()
status2.addRun("REVIEW", [:bold = true, :color = "orange", :size = 11])
status2.setAlign("center")
status2.setVerticalAlign("center")
status2.setBgColor("FFF8E1")

# Row 3 - Website
num3 = wordCell("3", [:bold = true, :align = "center"])
num3.setVerticalAlign("center")

proj3 = new WordCell()
proj3.addRun("Ring Website", [:bold = true, :size = 12, :color = "2E7D32"])
proj3.addLineBreak()
proj3.addRun("Bootstrap 5 Migration", [:italic = true, :size = 9, :color = "gray"])
proj3.setVerticalAlign("center")

details3 = new WordCell()
details3.addRun("Migration Tasks:", [:bold = true])
bulletId3 = doc.registerCellList(WORD_LIST_BULLET)
details3.addCellBulletList(["Updated to Bootstrap 5", "Responsive design improvements", "New documentation pages", "Performance optimization"], bulletId3)
details3.addCellCaption("Completion: 95%")

status3 = new WordCell()
status3.addRun("ACTIVE", [:bold = true, :color = "blue", :size = 11])
status3.setAlign("center")
status3.setVerticalAlign("center")
status3.setBgColor("E3F2FD")

doc.addTable([
    [h1, h2, h3, h4],
    [num1, proj1, details1, status1],
    [num2, proj2, details2, status2],
    [num3, proj3, details3, status3]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top",
    :colWidths = [1.2, 3.5, 9.5, 2.3]
])

if doc.save("demo_richcells_report.docx")
    ? "  Created: demo_richcells_report.docx"
else
    ? "  FAILED: demo_richcells_report.docx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
