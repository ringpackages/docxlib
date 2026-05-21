/*
	DOCXLib - Demo (RichCells2)
*/

load "docxlib.ring"

/*
    DOCXLib - Cell Content Demos
    Demonstrates lists, hyperlinks, and nested tables inside table cells.
*/

? "=============================================="
? "   DOCXLib - Cell Content Demos"
? "=============================================="

# =====================================================================
# Demo 1: Bullet Lists Inside Table Cells
# =====================================================================
? "Demo 1: Bullet lists in cells..."

doc = new WordWriter()
doc.setTitle("Cell Content - Bullet Lists")

doc.addHeading("Bullet Lists Inside Table Cells", 1)
doc.addParagraph("Each cell can contain a bullet list alongside other content.", NULL)

# Register list numbering definitions
bulletId1 = doc.registerCellList(WORD_LIST_BULLET)
bulletId2 = doc.registerCellList(WORD_LIST_BULLET)
bulletId3 = doc.registerCellList(WORD_LIST_BULLET)

# Header
h1 = wordCell("Department", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Responsibilities", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Tools", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])

# Row 1
dept1 = new WordCell()
dept1.addRun("Engineering", [:bold = true, :size = 12, :color = "1565C0"])
dept1.setVerticalAlign("center")

resp1 = new WordCell()
resp1.addRun("Key Tasks:", [:bold = true])
resp1.addCellBulletList(["Software development", "Code review", "Architecture design", "Testing & QA"], bulletId1)

tools1 = new WordCell()
tools1.addCellBulletList(["Ring Language", "Git", "VS Code", "Docker"], bulletId2)

# Row 2
dept2 = new WordCell()
dept2.addRun("Marketing", [:bold = true, :size = 12, :color = "2E7D32"])
dept2.setVerticalAlign("center")

resp2 = new WordCell()
resp2.addRun("Key Tasks:", [:bold = true])
resp2.addCellBulletList(["Brand strategy", "Content creation", "Social media", "Analytics"], bulletId3)

tools2 = wordCell("Various marketing tools", [])

tableData = [
    [h1, h2, h3],
    [dept1, resp1, tools1],
    [dept2, resp2, tools2]
]

doc.addTable(tableData, [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top"
])

if doc.save("demo_richcells2_lists.docx")
    ? "  Created: demo_richcells2_lists.docx"
else
    ? "  FAILED: demo_richcells2_lists.docx"
ok

# =====================================================================
# Demo 2: Numbered Lists Inside Table Cells
# =====================================================================
? "Demo 2: Numbered lists in cells..."

doc = new WordWriter()
doc.setTitle("Cell Content - Numbered Lists")

doc.addHeading("Numbered Lists Inside Table Cells", 1)
doc.addParagraph("Cells can contain numbered step-by-step instructions.", NULL)

numId1 = doc.registerCellList(WORD_LIST_NUMBER)
numId2 = doc.registerCellList(WORD_LIST_NUMBER)

h1 = wordCell("Recipe", [:bold = true, :color = "white", :bgColor = "E65100", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Steps", [:bold = true, :color = "white", :bgColor = "E65100", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Notes", [:bold = true, :color = "white", :bgColor = "E65100", :align = "center", :verticalAlign = "center"])

# Row 1
recipe1 = new WordCell()
recipe1.addRun("Pasta Carbonara", [:bold = true, :size = 12])
recipe1.addLineBreak()
recipe1.addRun("Serves 4", [:italic = true, :color = "gray"])
recipe1.setVerticalAlign("center")

steps1 = new WordCell()
steps1.addCellNumberedList(["Boil pasta in salted water", "Fry pancetta until crispy", "Mix eggs with pecorino", "Combine hot pasta with egg mix", "Add pancetta and serve"], numId1)

notes1 = new WordCell()
notes1.addRun("Important: ", [:bold = true, :color = "red"])
notes1.addRun("Do not scramble the eggs! Remove pan from heat before mixing.", [])
notes1.setVerticalAlign("center")

# Row 2
recipe2 = new WordCell()
recipe2.addRun("Quick Salad", [:bold = true, :size = 12])
recipe2.addLineBreak()
recipe2.addRun("Serves 2", [:italic = true, :color = "gray"])
recipe2.setVerticalAlign("center")

steps2 = new WordCell()
steps2.addCellNumberedList(["Wash and chop lettuce", "Slice tomatoes and cucumber", "Add olive oil and lemon"], numId2)

notes2 = wordCell("Ready in 5 minutes!", [:italic = true, :color = "green"])
notes2.setVerticalAlign("center")

doc.addTable([
    [h1, h2, h3],
    [recipe1, steps1, notes1],
    [recipe2, steps2, notes2]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top"
])

if doc.save("demo_richcells2_numbered.docx")
    ? "  Created: demo_richcells2__numbered.docx"
else
    ? "  FAILED: demo_richcells2_numbered.docx"
ok

# =====================================================================
# Demo 3: Hyperlinks Inside Table Cells
# =====================================================================
? "Demo 3: Hyperlinks in cells..."

doc = new WordWriter()
doc.setTitle("Cell Content - Hyperlinks")

doc.addHeading("Hyperlinks Inside Table Cells", 1)
doc.addParagraph("Cells can contain clickable hyperlinks to external URLs.", NULL)

# Register hyperlinks
link1 = doc.registerHyperlink("https://ring-lang.github.io")
link2 = doc.registerHyperlink("https://doublesvsoop.sf.net")
link3 = doc.registerHyperlink("https://supernova.sf.net")
link4 = doc.registerHyperlink("https://ring-lang.github.io/doc1.26/index.html")
link5 = doc.registerHyperlink("https://github.com/PWCT/PWCT2")

h1 = wordCell("Project", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Description", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Links", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])

# Row 1
proj1 = new WordCell()
proj1.addRun("Ring Language", [:bold = true, :size = 12, :color = "1565C0"])
proj1.setVerticalAlign("center")

desc1 = new WordCell()
desc1.addRun("Practical general-purpose multi-paradigm programming language.", [])
desc1.setVerticalAlign("center")

links1 = new WordCell()
links1.addRun("Resources:", [:bold = true])
links1.addCellHyperlink("Ring Official Website", link1)
links1.addCellHyperlink("Ring Documentation", link4)

# Row 2
proj2 = new WordCell()
proj2.addRun("PWCT2", [:bold = true, :size = 12, :color = "2E7D32"])
proj2.setVerticalAlign("center")

desc2 = new WordCell()
desc2.addRun("Programming Without Coding Technology - visual programming environment.", [])
desc2.setVerticalAlign("center")

links2 = new WordCell()
links2.addRun("Resources:", [:bold = true])
links2.addCellHyperlink("PWCT2 on GitHub", link5)

doc.addTable([
    [h1, h2, h3],
    [proj1, desc1, links1],
    [proj2, desc2, links2]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top"
])

if doc.save("demo_richcells2_hyperlinks.docx")
    ? "  Created: demo_richcells2_hyperlinks.docx"
else
    ? "  FAILED: demo_richcells2_hyperlinks.docx"
ok

# =====================================================================
# Demo 4: Nested Tables Inside Table Cells
# =====================================================================
? "Demo 4: Nested tables in cells..."

doc = new WordWriter()
doc.setTitle("Cell Content - Nested Tables")

doc.addHeading("Nested Tables Inside Table Cells", 1)
doc.addParagraph("A table cell can contain another table for complex layouts.", NULL)

# Generate the nested table XML
nestedXml1 = doc.generateNestedTable([
    ["Q1", "Q2", "Q3", "Q4"],
    ["$12K", "$15K", "$18K", "$22K"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "4472C4"
])

nestedXml2 = doc.generateNestedTable([
    ["Q1", "Q2", "Q3", "Q4"],
    ["$8K", "$9K", "$11K", "$14K"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "4472C4"
])

# Outer table
h1 = wordCell("Product Line", [:bold = true, :color = "white", :bgColor = "1B5E20", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Quarterly Revenue", [:bold = true, :color = "white", :bgColor = "1B5E20", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Status", [:bold = true, :color = "white", :bgColor = "1B5E20", :align = "center", :verticalAlign = "center"])

prod1 = new WordCell()
prod1.addRun("Enterprise Suite", [:bold = true, :size = 11])
prod1.addLineBreak()
prod1.addRun("B2B software platform", [:italic = true, :size = 9, :color = "gray"])
prod1.setVerticalAlign("center")

revenue1 = new WordCell()
revenue1.addRun("2025 Revenue Breakdown:", [:bold = true, :size = 9])
revenue1.addCellTable(nestedXml1)

status1 = new WordCell()
status1.addRun("Growing", [:bold = true, :color = "green", :size = 14])
status1.setAlign("center")
status1.setVerticalAlign("center")
status1.setBgColor("E8F5E9")

prod2 = new WordCell()
prod2.addRun("Consumer App", [:bold = true, :size = 11])
prod2.addLineBreak()
prod2.addRun("Mobile application", [:italic = true, :size = 9, :color = "gray"])
prod2.setVerticalAlign("center")

revenue2 = new WordCell()
revenue2.addRun("2025 Revenue Breakdown:", [:bold = true, :size = 9])
revenue2.addCellTable(nestedXml2)

status2 = new WordCell()
status2.addRun("Stable", [:bold = true, :color = "orange", :size = 14])
status2.setAlign("center")
status2.setVerticalAlign("center")
status2.setBgColor("FFF8E1")

doc.addTable([
    [h1, h2, h3],
    [prod1, revenue1, status1],
    [prod2, revenue2, status2]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top"
])

if doc.save("demo_richcells2_nested_tables.docx")
    ? "  Created: demo_richcells2_nested_tables.docx"
else
    ? "  FAILED: demo_richcells2_nested_tables.docx"
ok

# =====================================================================
# Demo 5: Mixed Content - All Features Combined
# =====================================================================
? "Demo 5: Mixed content (lists + hyperlinks + tables combined)..."

doc = new WordWriter()
doc.setTitle("Cell Content - Mixed Content")

doc.addHeading("All Cell Content Types Combined", 1)
doc.addParagraph("A single cell can contain text, lists, hyperlinks, and nested tables.", NULL)

# Register resources
bulletId = doc.registerCellList(WORD_LIST_BULLET)
numId = doc.registerCellList(WORD_LIST_NUMBER)
linkRing = doc.registerHyperlink("https://ring-lang.github.io")
linkDocs = doc.registerHyperlink("https://ring-lang.github.io/doc1.26/index.html")

nestedXml = doc.generateNestedTable([
    ["Feature", "Status"],
    ["Closures", "New"],
    ["RingSlint", "New"],
    ["Rust Bindings", "New"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "1565C0"
])

# Header
h1 = wordCell("Overview", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Details", [:bold = true, :color = "white", :bgColor = "4A148C", :align = "center", :verticalAlign = "center"])

# Row: Rich content cell
overview = new WordCell()
overview.addRun("Ring 1.26", [:bold = true, :size = 14, :color = "4A148C"])
overview.addLineBreak()
overview.addRun("Latest Release", [:italic = true, :color = "gray"])
overview.setVerticalAlign("center")
overview.setAlign("center")

details = new WordCell()
details.addRun("Ring 1.26 is the latest release with many improvements.", [])
details.addLineBreak()
details.addLineBreak()
details.addRun("Key Highlights:", [:bold = true])
details.addCellBulletList(["New closures support", "RingSlint UI framework", "Archive package", "Proc package"], bulletId)
details.addRun("Setup Steps:", [:bold = true])
details.addCellNumberedList(["Download from website", "Run installer", "Open Ring Notepad"], numId)
details.addRun("Useful Links:", [:bold = true])
details.addCellHyperlink("Ring Official Website", linkRing)
details.addCellHyperlink("Ring Documentation", linkDocs)
details.addRun("New Features Table:", [:bold = true])
details.addCellTable(nestedXml)

doc.addTable([
    [h1, h2],
    [overview, details]
], [
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "top"
])

if doc.save("demo_richcells2_mixed.docx")
    ? "  Created: demo_richcells2_mixed.docx"
else
    ? "  FAILED: demo_richcells2_mixed.docx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
