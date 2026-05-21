/*
	DOCXLib - Demo (RichCells4)
*/

load "docxlib.ring"

/*
    Row Span & Landscape Sections Demo
    Demonstrates vertical merge (row span) and mixed orientation sections.
*/

? "=============================================="
? "   DOCXLib - Row Span & Landscape Demo"
? "=============================================="

# =====================================================================
# Demo 1: Basic Row Span (Vertical Merge)
# =====================================================================
? "Demo 1: Basic row span..."

doc = new WordWriter()
doc.setTitle("Row Span Demo")

doc.addHeading("Vertical Cell Merge (Row Span)", 1)
doc.addParagraph("Cells can span multiple rows using setRowSpan() and wordMergeCell().", NULL)

# --- Table 1: Simple row span ---
doc.addHeading("Category Grouping", 2)

cat1 = new WordCell()
cat1.addRun("Electronics", [:bold = true, :size = 12, :color = "1565C0"])
cat1.setRowSpan(3)
cat1.setVerticalAlign("center")
cat1.setAlign("center")
cat1.setBgColor("E3F2FD")

cat2 = new WordCell()
cat2.addRun("Software", [:bold = true, :size = 12, :color = "2E7D32"])
cat2.setRowSpan(2)
cat2.setVerticalAlign("center")
cat2.setAlign("center")
cat2.setBgColor("E8F5E9")

h1 = wordCell("Category", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Product", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Price", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
h4 = wordCell("Stock", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])

doc.addTable([
    [h1, h2, h3, h4],
    [cat1, "Laptop Pro 15", "$1,299", "45"],
    [wordMergeCell(), "Wireless Mouse", "$29", "500"],
    [wordMergeCell(), "4K Monitor", "$449", "78"],
    [cat2, "Ring Language", "Free", "Unlimited"],
    [wordMergeCell(), "PWCT2", "Free", "Unlimited"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerHeight = 600,
    :headerVAlign = "center",
    :cellVAlign = "center",
    :colWidths = [3.5, 5, 3, 3]
])

if doc.save("demo_richcells4_basic.docx")
    ? "  Created: demo_richcells4_basic.docx"
else
    ? "  FAILED: demo_richcells4_basic.docx"
ok

# =====================================================================
# Demo 2: Row Span with Rich Content
# =====================================================================
? "Demo 2: Row span with rich content..."

doc = new WordWriter()
doc.setTitle("Row Span with Rich Content")

doc.addHeading("Row Span with Rich Cell Content", 1)

bulletId1 = doc.registerCellList(WORD_LIST_BULLET)
bulletId2 = doc.registerCellList(WORD_LIST_BULLET)

# Department spanning rows
dept1 = new WordCell()
dept1.addRun("Engineering", [:bold = true, :size = 14, :color = "white"])
dept1.setRowSpan(3)
dept1.setVerticalAlign("center")
dept1.setAlign("center")
dept1.setBgColor("1565C0")

dept2 = new WordCell()
dept2.addRun("Marketing", [:bold = true, :size = 14, :color = "white"])
dept2.setRowSpan(2)
dept2.setVerticalAlign("center")
dept2.setAlign("center")
dept2.setBgColor("E65100")

# Headers
h1 = wordCell("Department", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Employee", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Role", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h4 = wordCell("Skills", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])

# Engineering employees
emp1 = new WordCell()
emp1.addRun("Ahmed", [:bold = true])
emp1.setVerticalAlign("center")

role1 = wordCell("Senior Developer", [])
role1.setVerticalAlign("center")

skills1 = new WordCell()
skills1.addCellBulletList(["Ring", "C/C++", "Python"], bulletId1)

emp2 = wordCell("Sara", [:bold = true])
emp2.setVerticalAlign("center")
role2 = wordCell("QA Engineer", [])
role2.setVerticalAlign("center")
skills2 = wordCell("Testing, Automation", [])
skills2.setVerticalAlign("center")

emp3 = wordCell("Omar", [:bold = true])
emp3.setVerticalAlign("center")
role3 = wordCell("DevOps", [])
role3.setVerticalAlign("center")
skills3 = wordCell("Docker, CI/CD", [])
skills3.setVerticalAlign("center")

# Marketing employees
emp4 = wordCell("Fatima", [:bold = true])
emp4.setVerticalAlign("center")
role4 = wordCell("Content Lead", [])
role4.setVerticalAlign("center")

skills4 = new WordCell()
skills4.addCellBulletList(["SEO", "Copywriting", "Social Media"], bulletId2)

emp5 = wordCell("Khalid", [:bold = true])
emp5.setVerticalAlign("center")
role5 = wordCell("Designer", [])
role5.setVerticalAlign("center")
skills5 = wordCell("Figma, Photoshop", [])
skills5.setVerticalAlign("center")

doc.addTable([
    [h1, h2, h3, h4],
    [dept1, emp1, role1, skills1],
    [wordMergeCell(), emp2, role2, skills2],
    [wordMergeCell(), emp3, role3, skills3],
    [dept2, emp4, role4, skills4],
    [wordMergeCell(), emp5, role5, skills5]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "center",
    :colWidths = [3, 3, 3.5, 5]
])

if doc.save("demo_richcells4_rich.docx")
    ? "  Created: demo_richcells4_rich.docx"
else
    ? "  FAILED: demo_richcells4_rich.docx"
ok

# =====================================================================
# Demo 3: Combined Row Span + Column Span
# =====================================================================
? "Demo 3: Combined row span and column span..."

doc = new WordWriter()
doc.setTitle("Combined Row and Column Span")

doc.addHeading("Combined Row Span and Column Span", 1)
doc.addParagraph("Row spans and column spans can be used together in the same table.", NULL)

# Top-left corner spanning 2 rows and 2 columns
corner = new WordCell()
corner.addRun("Q4 2024", [:bold = true, :size = 16, :color = "white"])
corner.addLineBreak()
corner.addRun("Sales Report", [:italic = true, :size = 10, :color = "E0E0E0"])
corner.setRowSpan(2)
corner.setColSpan(2)
corner.setVerticalAlign("center")
corner.setAlign("center")
corner.setBgColor("1B5E20")

# Region headers
regionUS = wordCell("USA", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])
regionEU = wordCell("Europe", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])
regionAsia = wordCell("Asia", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])

# Month rows
oct = wordCell("October", [:bold = true, :bgColor = "E8F5E9"])
oct.setVerticalAlign("center")
nov = wordCell("November", [:bold = true, :bgColor = "E8F5E9"])
nov.setVerticalAlign("center")
dec = wordCell("December", [:bold = true, :bgColor = "E8F5E9"])
dec.setVerticalAlign("center")

# Total spanning row
total = new WordCell()
total.addRun("TOTAL", [:bold = true, :color = "white", :size = 12])
total.setColSpan(2)
total.setAlign("center")
total.setVerticalAlign("center")
total.setBgColor("37474F")

totalUS = wordCell("$450K", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
totalEU = wordCell("$320K", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])
totalAsia = wordCell("$280K", [:bold = true, :color = "white", :bgColor = "37474F", :align = "center", :verticalAlign = "center"])

# Merge continuation for the 2x2 corner
cornerCont = wordMergeCell()
cornerCont.setColSpan(2)

doc.addTable([
    [corner, "", regionUS, regionEU, regionAsia],
    [cornerCont, "", "", "", ""],
    [oct, "", "$150K", "$100K", "$90K"],
    [nov, "", "$140K", "$110K", "$95K"],
    [dec, "", "$160K", "$110K", "$95K"],
    [total, "", totalUS, totalEU, totalAsia]
], [
    :borderStyle = "single",
    :cellVAlign = "center",
    :colWidths = [2.5, 2.5, 3.5, 3.5, 3.5]
])

if doc.save("demo_richcells4_combined.docx")
    ? "  Created: demo_richcells4_combined.docx"
else
    ? "  FAILED: demo_richcells4_combined.docx"
ok

# =====================================================================
# Demo 4: Landscape Section
# =====================================================================
? "Demo 4: Landscape section..."

doc = new WordWriter()
doc.setTitle("Mixed Orientation Document")

doc.addHeading("Portrait Section", 1)
doc.addParagraph("This page is in portrait orientation. The next page will switch to landscape for a wide table.", NULL)
doc.addParagraph("Portrait orientation is ideal for regular text, paragraphs, and narrow tables.", NULL)

# --- Switch to landscape ---
doc.addLandscapeStart()

doc.addHeading("Landscape Section", 1)
doc.addParagraph("This page is in landscape orientation, providing extra width for wide tables and charts.", NULL)

# Wide table that benefits from landscape
doc.addTable([
    ["ID", "Project Name", "Lead", "Start Date", "End Date", "Budget", "Spent", "Status", "Priority", "Notes"],
    ["P-001", "Ring 1.26 Release", "Mahmoud", "2024-01-15", "2025-01-25", "$50K", "$48K", "Done", "High", "Major release with closures"],
    ["P-002", "PWCT2 Development", "Mahmoud", "2024-03-01", "2025-06-30", "$30K", "$18K", "Active", "High", "Visual programming platform"],
    ["P-003", "Ring Website Redesign", "Ahmed", "2024-06-01", "2024-12-31", "$10K", "$9K", "Done", "Medium", "Bootstrap 5 migration"],
    ["P-004", "Research Paper", "Mahmoud", "2024-09-01", "2025-03-30", "$5K", "$3K", "Review", "Medium", "Journal"],
    ["P-005", "Supernova Language", "Mahmoud", "2024-01-01", "2025-12-31", "$20K", "$8K", "Active", "Low", "Domain-specific VPL"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "1565C0",
    :evenRowBgColor = "E3F2FD",
    :headerHeight = 600,
    :headerVAlign = "center",
    :cellVAlign = "center",
    :colWidths = [1.5, 4, 2.5, 2.5, 2.5, 2, 2, 1.8, 1.8, 5]
])

# --- Return to portrait ---
doc.addLandscapeEnd()

doc.addHeading("Back to Portrait", 1)
doc.addParagraph("This page is back to portrait orientation. You can switch between portrait and landscape as many times as needed within a single document.", NULL)

if doc.save("demo_richcells4_landscape.docx")
    ? "  Created: demo_richcells4_landscape.docx"
else
    ? "  FAILED: demo_richcells4_landscape.docx"
ok

# =====================================================================
# Demo 5: Complete Report with All Features
# =====================================================================
? "Demo 5: Complete report combining all features..."

doc = new WordWriter()
doc.setTitle("Annual Report 2024")
doc.setAuthor("Mahmoud")
doc.setHeader("Annual Report 2024")
doc.showPageNumbers("center")

# --- Page 1: Title and summary (portrait) ---
doc.addHeading("Annual Report 2024", 1)
doc.addParagraph("A comprehensive overview of projects, financials, and team performance.", NULL)

doc.addHeading("Team Overview", 2)

bulletId = doc.registerCellList(WORD_LIST_BULLET)

dept1 = new WordCell()
dept1.addRun("R&D", [:bold = true, :size = 14, :color = "white"])
dept1.setRowSpan(3)
dept1.setVerticalAlign("center")
dept1.setAlign("center")
dept1.setBgColor("0D47A1")

dept2 = new WordCell()
dept2.addRun("QA", [:bold = true, :size = 14, :color = "white"])
dept2.setRowSpan(2)
dept2.setVerticalAlign("center")
dept2.setAlign("center")
dept2.setBgColor("E65100")

h1 = wordCell("Dept", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h2 = wordCell("Name", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h3 = wordCell("Role", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
h4 = wordCell("Focus Areas", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])

focus1 = new WordCell()
focus1.addCellBulletList(["Ring Language core", "Compiler optimization", "VM performance"], bulletId)

doc.addTable([
    [h1, h2, h3, h4],
    [dept1, "Mahmoud", "Lead Developer", focus1],
    [wordMergeCell(), "Ahmed", "Senior Developer", "Libraries & Extensions"],
    [wordMergeCell(), "Omar", "Junior Developer", "Documentation & Testing"],
    [dept2, "Sara", "QA Lead", "Test automation"],
    [wordMergeCell(), "Fatima", "QA Engineer", "Manual & integration testing"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerHeight = 600,
    :cellVAlign = "center",
    :colWidths = [2, 3, 4, 7.5]
])

# --- Page 2: Wide financial table (landscape) ---
doc.addLandscapeStart()

doc.addHeading("Financial Overview", 1)
doc.addParagraph("Quarterly breakdown of project budgets and expenditures across all departments.", NULL)

# Row span for department grouping
fin1 = new WordCell()
fin1.addRun("R&D", [:bold = true, :color = "white"])
fin1.setRowSpan(3)
fin1.setVerticalAlign("center")
fin1.setAlign("center")
fin1.setBgColor("0D47A1")

fin2 = new WordCell()
fin2.addRun("QA", [:bold = true, :color = "white"])
fin2.setRowSpan(2)
fin2.setVerticalAlign("center")
fin2.setAlign("center")
fin2.setBgColor("E65100")

fh1 = wordCell("Dept", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh2 = wordCell("Project", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh3 = wordCell("Q1 Budget", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh4 = wordCell("Q1 Actual", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh5 = wordCell("Q2 Budget", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh6 = wordCell("Q2 Actual", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh7 = wordCell("Q3 Budget", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh8 = wordCell("Q3 Actual", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh9 = wordCell("Q4 Budget", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
fh10 = wordCell("Q4 Actual", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])

doc.addTable([
    [fh1, fh2, fh3, fh4, fh5, fh6, fh7, fh8, fh9, fh10],
    [fin1, "Ring 1.26", "$12K", "$11K", "$13K", "$12K", "$13K", "$13K", "$12K", "$12K"],
    [wordMergeCell(), "PWCT2", "$7K", "$6K", "$8K", "$7K", "$8K", "$8K", "$7K", "$7K"],
    [wordMergeCell(), "Supernova", "$5K", "$4K", "$5K", "$5K", "$5K", "$4K", "$5K", "$5K"],
    [fin2, "Automation", "$3K", "$3K", "$3K", "$3K", "$4K", "$4K", "$4K", "$3K"],
    [wordMergeCell(), "Manual QA", "$2K", "$2K", "$2K", "$2K", "$2K", "$2K", "$3K", "$2K"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerHeight = 600,
    :headerVAlign = "center",
    :cellVAlign = "center",
    :evenRowBgColor = "F5F5F5",
    :colWidths = [2, 3, 2.3, 2.3, 2.3, 2.3, 2.3, 2.3, 2.3, 2.3]
])

doc.addLandscapeEnd()

# --- Page 3: Back to portrait ---
doc.addHeading("Summary & Next Steps", 1)

doc.addParagraph("All projects are on track. Key priorities for 2025:", NULL)

doc.addBulletList([
    "Complete Ring 1.26 release with new features",
    "Publish research papers",
    "Expand PWCT2 visual programming capabilities",
    "Grow community and documentation"
])

doc.addParagraph("Overall budget utilization: 93% - within acceptable range.", NULL)

if doc.save("demo_richcells4_report.docx")
    ? "  Created: demo_richcells4_report.docx"
else
    ? "  FAILED: demo_richcells4_report.docx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
