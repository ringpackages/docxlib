/*
	DOCXLib - Demo (RichCells)
*/

load "docxlib.ring"

/*
    Demonstrates the WordCell feature for full control 
    over table cell contents: formatted text, images, alignment,
    background colors, column spanning, vertical alignment, etc.
*/

? "=============================================="
? "   DOCXLib - Rich Table Cells Demos"
? "=============================================="
? ""

# =====================================================================
# Demo 1: Formatted Text in Cells
# =====================================================================
? "Demo 1: Table with formatted text in cells..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Formatted Text")
doc.setAuthor("Ring Programmer")

doc.addHeading("Formatted Text in Table Cells", 1)
doc.addParagraph("Each cell can have its own font, size, color, and style.", NULL)

tableData = [
    [
        wordCell("Name", [:bold = true, :align = "center"]),
        wordCell("Description", [:bold = true, :align = "center"]),
        wordCell("Status", [:bold = true, :align = "center"])
    ],
    [
        wordCell("Ring Language", [:bold = true, :color = "blue", :font = "Arial"]),
        wordCell("A practical general-purpose programming language", [:italic = true]),
        wordCell("Active", [:bold = true, :color = "green", :align = "center"])
    ],
    [
        wordCell("PWCT", [:bold = true, :color = "purple", :font = "Arial"]),
        wordCell("Programming Without Coding Technology", [:italic = true]),
        wordCell("Active", [:bold = true, :color = "green", :align = "center"])
    ],
    [
        wordCell("Supernova", [:bold = true, :color = "orange", :font = "Arial"]),
        wordCell("A domain-specific visual programming language", [:italic = true]),
        wordCell("In Development", [:bold = true, :color = "orange", :align = "center"])
    ]
]

doc.addTable(tableData, [:borderStyle = "single", :headerVAlign = "center", :cellVAlign = "center"])

if doc.save("demo_richcell_1_formatted_text.docx")
    ? "  Created: demo_richcell_1_formatted_text.docx"
else
    ? "  FAILED: demo_richcell_1_formatted_text.docx"
ok

# =====================================================================
# Demo 2: Multi-Run Rich Text in Cells
# =====================================================================
? "Demo 2: Table with multi-run rich text in cells..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Multi-Run Text")

doc.addHeading("Multiple Text Styles in One Cell", 1)
doc.addParagraph("A single cell can contain multiple text runs with different formatting.", NULL)

# Build cells with multiple runs
cell1 = new WordCell()
cell1.addRun("Feature", [:bold = true, :size = 14])
cell1.setAlign("center")
cell1.setBgColor("1565C0")

cell2 = new WordCell()
cell2.addRun("Details", [:bold = true, :size = 14])
cell2.setAlign("center")
cell2.setBgColor("1565C0")

cell3 = new WordCell()
cell3.addRun("Bold text, ", [:bold = true])
cell3.addRun("italic text, ", [:italic = true])
cell3.addRun("and ", [])
cell3.addRun("colored text", [:color = "red"])
cell3.addRun(" all in one cell!", [])

cell4 = new WordCell()
cell4.addRun("Mixed formatting: ", [])
cell4.addRun("Font sizes ", [:size = 8])
cell4.addRun("can vary ", [:size = 12])
cell4.addRun("within ", [:size = 16])
cell4.addRun("a cell", [:size = 20])

cell5 = new WordCell()
cell5.addRun("Ring", [:bold = true, :color = "blue", :size = 16, :font = "Georgia"])
cell5.addRun(" is ", [:size = 11])
cell5.addRun("awesome!", [:bold = true, :italic = true, :color = "green", :size = 14])

cell6 = new WordCell()
cell6.addRun("Underline", [:underline = true])
cell6.addRun(" + ", [])
cell6.addRun("Strikethrough", [:strike = true])
cell6.addRun(" + ", [])
cell6.addRun("Highlight", [:highlight = "yellow"])

tableData = [
    [cell1, cell2],
    [cell3, cell4],
    [cell5, cell6]
]

doc.addTable(tableData, [:borderStyle = "single"])

if doc.save("demo_richcell_2_multirun.docx")
    ? "  Created: demo_richcell_2_multirun.docx"
else
    ? "  FAILED: demo_richcell_2_multirun.docx"
ok

# =====================================================================
# Demo 3: Vertical Centering in Tables
# =====================================================================
? "Demo 3: Table with vertical centering..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Vertical Centering")

doc.addHeading("Vertical Centering in Tables", 1)

# --- Table 1: Using headerHeight + headerVAlign (plain strings) ---
doc.addHeading("Table-Level Vertical Centering (Plain Strings)", 2)
doc.addParagraph("Using headerHeight to make header taller, headerVAlign to center text:", NULL)

doc.addTable([
    ["Feature", "Ring", "Python", "Java"],
    ["Dynamic Typing", "YES", "YES", "NO"],
    ["Object-Oriented", "YES", "YES", "YES"],
    ["PWCT2", "YES", "NO", "NO"]
], [
    :headerRow = true, 
    :borderStyle = "single",
    :headerBgColor = "37474F",
    :headerHeight = 720,
    :headerVAlign = "center",
    :cellVAlign = "center"
])

# --- Table 2: Per-cell vertical alignment with tall rows ---
doc.addHeading("Per-Cell Vertical Alignment (Top / Center / Bottom)", 2)
doc.addParagraph("A tall reference cell shows the difference between top, center, and bottom:", NULL)

# Tall cell
tallCell = new WordCell()
tallCell.addRun("Line 1", [:bold = true])
tallCell.addLineBreak()
tallCell.addRun("Line 2", [])
tallCell.addLineBreak()
tallCell.addRun("Line 3", [])
tallCell.addLineBreak()
tallCell.addRun("Line 4", [])
tallCell.addLineBreak()
tallCell.addRun("Line 5", [])
tallCell.addLineBreak()
tallCell.addRun("Line 6", [:italic = true])
tallCell.setBgColor("F5F5F5")

topCell = new WordCell()
topCell.addRun("TOP", [:bold = true, :size = 14, :color = "red"])
topCell.setVerticalAlign("top")
topCell.setAlign("center")
topCell.setBgColor("FFEBEE")

centerCell = new WordCell()
centerCell.addRun("CENTER", [:bold = true, :size = 14, :color = "green"])
centerCell.setVerticalAlign("center")
centerCell.setAlign("center")
centerCell.setBgColor("E8F5E9")

bottomCell = new WordCell()
bottomCell.addRun("BOTTOM", [:bold = true, :size = 14, :color = "blue"])
bottomCell.setVerticalAlign("bottom")
bottomCell.setAlign("center")
bottomCell.setBgColor("E3F2FD")

# Headers
hRef = wordCell("Reference", [:bold = true, :color = "white", :bgColor = "263238", :align = "center", :verticalAlign = "center"])
hTop = wordCell("vAlign=top", [:bold = true, :color = "white", :bgColor = "C62828", :align = "center", :verticalAlign = "center"])
hCtr = wordCell("vAlign=center", [:bold = true, :color = "white", :bgColor = "2E7D32", :align = "center", :verticalAlign = "center"])
hBot = wordCell("vAlign=bottom", [:bold = true, :color = "white", :bgColor = "1565C0", :align = "center", :verticalAlign = "center"])

doc.addTable([
    [hRef, hTop, hCtr, hBot],
    [tallCell, topCell, centerCell, bottomCell]
], [:borderStyle = "single", :headerHeight = 600])

# --- Table 3: Styled table with headerHeight ---
doc.addHeading("Styled Table with Vertical Centering", 2)

doc.addTable([
    ["Product", "Category", "Price", "Stock"],
    ["Laptop Pro 15", "Electronics", "$1,299", "45"],
    ["Wireless Mouse", "Accessories", "$29", "500"],
    ["USB-C Hub", "Accessories", "$49", "200"],
    ["4K Monitor 27 inch", "Electronics", "$449", "78"]
], [
    :headerRow = true,
    :borderStyle = "single",
    :headerBgColor = "1565C0",
    :evenRowBgColor = "E3F2FD",
    :headerHeight = 700,
    :headerVAlign = "center",
    :cellVAlign = "center"
])

if doc.save("demo_richcell_3_alignment.docx")
    ? "  Created: demo_richcell_3_alignment.docx"
else
    ? "  FAILED: demo_richcell_3_alignment.docx"
ok

# =====================================================================
# Demo 4: Cell Background Colors
# =====================================================================
? "Demo 4: Table with custom cell background colors..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Colors")

doc.addHeading("Cell Background Colors", 1)
doc.addParagraph("Each cell can have its own background color, independent of row-level styling.", NULL)

tableData = [
    [
        wordCell("Red", [:bold = true, :color = "white", :bgColor = "FF0000", :align = "center"]),
        wordCell("Green", [:bold = true, :color = "white", :bgColor = "00AA00", :align = "center"]),
        wordCell("Blue", [:bold = true, :color = "white", :bgColor = "0000FF", :align = "center"])
    ],
    [
        wordCell("Orange", [:bold = true, :bgColor = "FFA500", :align = "center"]),
        wordCell("Purple", [:bold = true, :color = "white", :bgColor = "800080", :align = "center"]),
        wordCell("Teal", [:bold = true, :color = "white", :bgColor = "008080", :align = "center"])
    ],
    [
        wordCell("Light Yellow", [:bgColor = "FFFFCC", :align = "center"]),
        wordCell("Light Pink", [:bgColor = "FFCCCC", :align = "center"]),
        wordCell("Light Blue", [:bgColor = "CCE5FF", :align = "center"])
    ],
    [
        wordCell("Gold", [:bold = true, :bgColor = "FFD700", :align = "center"]),
        wordCell("Silver", [:bgColor = "C0C0C0", :align = "center"]),
        wordCell("Coral", [:color = "white", :bgColor = "FF6B6B", :align = "center"])
    ]
]

doc.addTable(tableData, [:borderStyle = "single"])

if doc.save("demo_richcell_4_colors.docx")
    ? "  Created: demo_richcell_4_colors.docx"
else
    ? "  FAILED: demo_richcell_4_colors.docx"
ok

# =====================================================================
# Demo 5: Column Spanning (Merged Cells)
# =====================================================================
? "Demo 5: Table with merged cells (column span)..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Column Spanning")

doc.addHeading("Column Spanning (Merged Cells)", 1)
doc.addParagraph("Cells can span multiple columns using setColSpan().", NULL)

# Title row spanning all 3 columns
titleCell = new WordCell()
titleCell.addRun("Quarterly Sales Report 2026", [:bold = true, :size = 16, :color = "white"])
titleCell.setAlign("center")
titleCell.setBgColor("1565C0")
titleCell.setColSpan(3)

# Header row
h1 = wordCell("Region", [:bold = true, :align = "center", :bgColor = "90CAF9"])
h2 = wordCell("Q1 Sales", [:bold = true, :align = "center", :bgColor = "90CAF9"])
h3 = wordCell("Q2 Sales", [:bold = true, :align = "center", :bgColor = "90CAF9"])

# Data rows
r1c1 = wordCell("North America", [:bold = true])
r1c2 = wordCell("$2.5M", [:align = "right"])
r1c3 = wordCell("$2.8M", [:align = "right", :color = "green"])

r2c1 = wordCell("Europe", [:bold = true])
r2c2 = wordCell("$1.8M", [:align = "right"])
r2c3 = wordCell("$2.1M", [:align = "right", :color = "green"])

r3c1 = wordCell("Asia Pacific", [:bold = true])
r3c2 = wordCell("$1.2M", [:align = "right"])
r3c3 = wordCell("$1.4M", [:align = "right", :color = "green"])

# Total row spanning
totalLabel = new WordCell()
totalLabel.addRun("Total Revenue", [:bold = true, :size = 12])
totalLabel.setAlign("right")
totalLabel.setBgColor("E3F2FD")
totalLabel.setColSpan(2)

totalValue = new WordCell()
totalValue.addRun("$11.8M", [:bold = true, :size = 14, :color = "1565C0"])
totalValue.setAlign("right")
totalValue.setBgColor("E3F2FD")

tableData = [
    [titleCell, "", ""],
    [h1, h2, h3],
    [r1c1, r1c2, r1c3],
    [r2c1, r2c2, r2c3],
    [r3c1, r3c2, r3c3],
    [totalLabel, "", totalValue]
]

doc.addTable(tableData, [:borderStyle = "single"])

if doc.save("demo_richcell_5_colspan.docx")
    ? "  Created: demo_richcell_5_colspan.docx"
else
    ? "  FAILED: demo_richcell_5_colspan.docx"
ok

# =====================================================================
# Demo 6: Line Breaks Within Cells
# =====================================================================
? "Demo 6: Table with line breaks in cells..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Line Breaks")

doc.addHeading("Line Breaks Within Cells", 1)
doc.addParagraph("Cells can contain multiple lines of text using addLineBreak().", NULL)

# Header
h1 = wordCell("Contact", [:bold = true, :align = "center", :bgColor = "2E7D32", :color = "white"])
h2 = wordCell("Details", [:bold = true, :align = "center", :bgColor = "2E7D32", :color = "white"])

# Multi-line cells
contact1 = new WordCell()
contact1.addRun("Ahmed Hassan", [:bold = true, :size = 12])
contact1.addLineBreak()
contact1.addRun("Senior Developer", [:italic = true, :color = "gray"])
contact1.addLineBreak()
contact1.addRun("ahmed@example.com", [:color = "blue", :underline = true])

details1 = new WordCell()
details1.addRun("Department: ", [:bold = true])
details1.addRun("Engineering", [])
details1.addLineBreak()
details1.addRun("Office: ", [:bold = true])
details1.addRun("Building A, Room 301", [])
details1.addLineBreak()
details1.addRun("Phone: ", [:bold = true])
details1.addRun("+966 50 123 4567", [])

contact2 = new WordCell()
contact2.addRun("Sara Ali", [:bold = true, :size = 12])
contact2.addLineBreak()
contact2.addRun("Project Manager", [:italic = true, :color = "gray"])
contact2.addLineBreak()
contact2.addRun("sara@example.com", [:color = "blue", :underline = true])

details2 = new WordCell()
details2.addRun("Department: ", [:bold = true])
details2.addRun("Management", [])
details2.addLineBreak()
details2.addRun("Office: ", [:bold = true])
details2.addRun("Building B, Room 205", [])
details2.addLineBreak()
details2.addRun("Phone: ", [:bold = true])
details2.addRun("+966 50 987 6543", [])

tableData = [
    [h1, h2],
    [contact1, details1],
    [contact2, details2]
]

doc.addTable(tableData, [:borderStyle = "single"])

if doc.save("demo_richcell_6_linebreaks.docx")
    ? "  Created: demo_richcell_6_linebreaks.docx"
else
    ? "  FAILED: demo_richcell_6_linebreaks.docx"
ok

# =====================================================================
# Demo 7: Cell Padding and Borders
# =====================================================================
? "Demo 7: Table with custom cell padding and borders..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Padding and Borders")

doc.addHeading("Cell Padding and Individual Borders", 1)
doc.addParagraph("Cells can have custom padding and individual border styles.", NULL)

# Cell with large padding
padded = new WordCell()
padded.addRun("This cell has large padding (0.5cm all around)", [:size = 10])
padded.setPadding(0.5, 0.5, 0.5, 0.5)
padded.setBgColor("FFF9C4")

# Cell with custom border
bordered = new WordCell()
bordered.addRun("This cell has a red dashed border", [:bold = true])
bordered.setBorder("dashed", "FF0000")
bordered.setAlign("center")

# Cell with double border
doubleBorder = new WordCell()
doubleBorder.addRun("Double blue border", [:bold = true, :color = "blue"])
doubleBorder.setBorder("double", "0000FF")
doubleBorder.setAlign("center")

# Normal cell for comparison
normal = wordCell("Normal cell (default)", NULL)

tableData = [
    [padded, bordered],
    [doubleBorder, normal]
]

doc.addTable(tableData, [:borderStyle = "single"])

if doc.save("demo_richcell_7_padding_borders.docx")
    ? "  Created: demo_richcell_7_padding_borders.docx"
else
    ? "  FAILED: demo_richcell_7_padding_borders.docx"
ok

# =====================================================================
# Demo 8: Images in Table Cells
# =====================================================================
? "Demo 8: Table with images in cells..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Images")

doc.addHeading("Images Inside Table Cells", 1)
doc.addParagraph("Images can be placed inside table cells alongside text.", NULL)

# Pre-register images with the document
img1Info = doc.registerImage("images/test1.png")
img2Info = doc.registerImage("images/test2.jpg")
img3Info = doc.registerImage("images/test3.bmp")

# Header
h1 = wordCell("Item", [:bold = true, :align = "center", :bgColor = "1565C0", :color = "white"])
h2 = wordCell("Preview", [:bold = true, :align = "center", :bgColor = "1565C0", :color = "white"])
h3 = wordCell("Details", [:bold = true, :align = "center", :bgColor = "1565C0", :color = "white"])

# Row 1
itemCell = new WordCell()
itemCell.addRun("Product A", [:bold = true, :size = 12])
itemCell.addLineBreak()
itemCell.addRun("SKU: PRD-001", [:size = 9, :color = "gray"])
itemCell.setVerticalAlign("center")

imageCell = new WordCell()
if len(img1Info[:relId]) > 0
    imageCell.addCellImage("images/test1.png", 2, 2)
    imageCell.setCellImageRel(1, img1Info[:relId], img1Info[:imageId])
else
    imageCell.addRun("(images/test1.png not found)", [:color = "gray", :size = 9])
ok
imageCell.setAlign("center")
imageCell.setVerticalAlign("center")

detailsCell = new WordCell()
detailsCell.addRun("Price: ", [:bold = true])
detailsCell.addRun("$49.99", [:color = "green", :bold = true])
detailsCell.addLineBreak()
detailsCell.addRun("Stock: ", [:bold = true])
detailsCell.addRun("In Stock (150 units)", [])
detailsCell.setVerticalAlign("center")

# Row 2
itemCell2 = new WordCell()
itemCell2.addRun("Product B", [:bold = true, :size = 12])
itemCell2.addLineBreak()
itemCell2.addRun("SKU: PRD-002", [:size = 9, :color = "gray"])
itemCell2.setVerticalAlign("center")

imageCell2 = new WordCell()
if len(img2Info[:relId]) > 0
    imageCell2.addCellImage("images/test2.jpg", 2, 2)
    imageCell2.setCellImageRel(1, img2Info[:relId], img2Info[:imageId])
else
    imageCell2.addRun("(images/test2.jpg not found)", [:color = "gray", :size = 9])
ok
imageCell2.setAlign("center")
imageCell2.setVerticalAlign("center")

detailsCell2 = new WordCell()
detailsCell2.addRun("Price: ", [:bold = true])
detailsCell2.addRun("$89.99", [:color = "green", :bold = true])
detailsCell2.addLineBreak()
detailsCell2.addRun("Stock: ", [:bold = true])
detailsCell2.addRun("Low Stock (12 units)", [:color = "red"])
detailsCell2.setVerticalAlign("center")

# Row 3
itemCell3 = new WordCell()
itemCell3.addRun("Product C", [:bold = true, :size = 12])
itemCell3.addLineBreak()
itemCell3.addRun("SKU: PRD-003", [:size = 9, :color = "gray"])
itemCell3.setVerticalAlign("center")

imageCell3 = new WordCell()
if len(img3Info[:relId]) > 0
    imageCell3.addCellImage("images/test3.bmp", 2, 2)
    imageCell3.setCellImageRel(1, img3Info[:relId], img3Info[:imageId])
else
    imageCell3.addRun("(images/test3.bmp not found)", [:color = "gray", :size = 9])
ok
imageCell3.setAlign("center")
imageCell3.setVerticalAlign("center")

detailsCell3 = new WordCell()
detailsCell3.addRun("Price: ", [:bold = true])
detailsCell3.addRun("$29.99", [:color = "green", :bold = true])
detailsCell3.addLineBreak()
detailsCell3.addRun("Stock: ", [:bold = true])
detailsCell3.addRun("In Stock (500 units)", [])
detailsCell3.setVerticalAlign("center")

tableData = [
    [h1, h2, h3],
    [itemCell, imageCell, detailsCell],
    [itemCell2, imageCell2, detailsCell2],
    [itemCell3, imageCell3, detailsCell3]
]

doc.addTable(tableData, [:borderStyle = "single", :cellVAlign = "center"])

if doc.save("demo_richcell_8_images.docx")
    ? "  Created: demo_richcell_8_images.docx"
else
    ? "  FAILED: demo_richcell_8_images.docx"
ok

# =====================================================================
# Demo 9: Mixed Plain Strings and WordCell Objects
# =====================================================================
? "Demo 9: Table mixing plain strings and WordCell objects..."

doc = new WordWriter()
doc.setTitle("Rich Table Cells - Mixed Content")

doc.addHeading("Mixing Plain Strings and WordCell Objects", 1)
doc.addParagraph("You can mix plain strings and WordCell objects in the same table. Existing code continues to work!", NULL)

tableData = [
    ["Name", "Role", wordCell("Status", [:bold = true, :align = "center", :bgColor = "333333", :color = "white"])],
    ["Ahmed", "Developer", wordCell("Active", [:bold = true, :color = "green", :align = "center", :bgColor = "E8F5E9"])],
    ["Sara", "Manager", wordCell("Active", [:bold = true, :color = "green", :align = "center", :bgColor = "E8F5E9"])],
    ["Omar", "Designer", wordCell("On Leave", [:bold = true, :color = "orange", :align = "center", :bgColor = "FFF3E0"])],
    ["Fatima", "QA Lead", wordCell("Inactive", [:color = "red", :align = "center", :bgColor = "FFEBEE"])]
]

doc.addTable(tableData, [:headerRow = true, :borderStyle = "single", :headerBgColor = "333333"])

if doc.save("demo_richcell_9_mixed.docx")
    ? "  Created: demo_richcell_9_mixed.docx"
else
    ? "  FAILED: demo_richcell_9_mixed.docx"
ok

# =====================================================================
# Demo 10: Professional Invoice with Rich Cells
# =====================================================================
? "Demo 10: Professional invoice using rich cells..."

doc = new WordWriter()
doc.setTitle("Invoice #2026-001")
doc.setAuthor("Mahmoud's Software Company")

# Invoice Header
doc.addHeading("INVOICE", 1)

doc.addRichParagraph([
    ["Invoice #: ", [:bold = true]],
    ["2026-001", [:color = "1565C0"]]
], NULL)
doc.addRichParagraph([
    ["Date: ", [:bold = true]],
    ["March 1, 2026", []]
], NULL)
doc.addRichParagraph([
    ["Due Date: ", [:bold = true]],
    ["March 31, 2026", [:color = "red"]]
], NULL)

doc.addHorizontalLine()

# Company Info Table (no borders)
fromCell = new WordCell()
fromCell.addRun("FROM:", [:bold = true, :size = 9, :color = "gray"])
fromCell.addLineBreak()
fromCell.addRun("Ring Software Solutions", [:bold = true, :size = 12])
fromCell.addLineBreak()
fromCell.addRun("123 Tech Street", [])
fromCell.addLineBreak()
fromCell.addRun("Riyadh, Saudi Arabia", [])
fromCell.addLineBreak()
fromCell.addRun("contact@ringsoftware.com", [:color = "blue"])

toCell = new WordCell()
toCell.addRun("TO:", [:bold = true, :size = 9, :color = "gray"])
toCell.addLineBreak()
toCell.addRun("ABC Corporation", [:bold = true, :size = 12])
toCell.addLineBreak()
toCell.addRun("456 Business Avenue", [])
toCell.addLineBreak()
toCell.addRun("Dubai, UAE", [])
toCell.addLineBreak()
toCell.addRun("billing@abccorp.com", [:color = "blue"])

doc.addTable([[fromCell, toCell]], [:borderStyle = "none"])

doc.addEmptyParagraph()

# Items Table
hItem = wordCell("Item", [:bold = true, :color = "white", :align = "center"])
hDesc = wordCell("Description", [:bold = true, :color = "white", :align = "center"])
hQty  = wordCell("Qty", [:bold = true, :color = "white", :align = "center"])
hRate = wordCell("Rate", [:bold = true, :color = "white", :align = "center"])
hTotal = wordCell("Total", [:bold = true, :color = "white", :align = "center"])

item1 = new WordCell()
item1.addRun("Ring Development", [:bold = true])
item1.addLineBreak()
item1.addRun("Custom library development", [:size = 9, :color = "gray"])

item2 = new WordCell()
item2.addRun("PWCT Training", [:bold = true])
item2.addLineBreak()
item2.addRun("5-day workshop", [:size = 9, :color = "gray"])

item3 = new WordCell()
item3.addRun("Technical Support", [:bold = true])
item3.addLineBreak()
item3.addRun("Monthly support package", [:size = 9, :color = "gray"])

# Subtotal row with span
subtotalLabel = new WordCell()
subtotalLabel.addRun("Subtotal", [:bold = true])
subtotalLabel.setAlign("right")
subtotalLabel.setColSpan(4)

subtotalValue = wordCell("$9,700.00", [:bold = true, :align = "right"])

# Tax row
taxLabel = new WordCell()
taxLabel.addRun("VAT (15%)", [])
taxLabel.setAlign("right")
taxLabel.setColSpan(4)

taxValue = wordCell("$1,455.00", [:align = "right"])

# Total row
totalLabel = new WordCell()
totalLabel.addRun("TOTAL DUE", [:bold = true, :size = 14])
totalLabel.setAlign("right")
totalLabel.setBgColor("1565C0")
totalLabel.setColSpan(4)

totalValue = new WordCell()
totalValue.addRun("$11,155.00", [:bold = true, :size = 14, :color = "white"])
totalValue.setAlign("right")
totalValue.setBgColor("1565C0")

tableData = [
    [hItem, hDesc, hQty, hRate, hTotal],
    [
        "1", item1,
        wordCell("40 hrs", [:align = "center"]),
        wordCell("$150.00", [:align = "right"]),
        wordCell("$6,000.00", [:align = "right"])
    ],
    [
        "2", item2,
        wordCell("1", [:align = "center"]),
        wordCell("$2,500.00", [:align = "right"]),
        wordCell("$2,500.00", [:align = "right"])
    ],
    [
        "3", item3,
        wordCell("1", [:align = "center"]),
        wordCell("$1,200.00", [:align = "right"]),
        wordCell("$1,200.00", [:align = "right"])
    ],
    [subtotalLabel, "", "", "", subtotalValue],
    [taxLabel, "", "", "", taxValue],
    [totalLabel, "", "", "", totalValue]
]

doc.addTable(tableData, [
    :borderStyle = "single", 
    :headerRow = true, 
    :headerBgColor = "1565C0",
    :evenRowBgColor = "F5F5F5",
    :headerVAlign = "center",
    :cellVAlign = "center"
])

doc.addEmptyParagraph()
doc.addParagraph("Payment Terms: Net 30 days", [:bold = true])
doc.addParagraph("Please make payment to: Ring Software Solutions, Bank Account: SA12 3456 7890 1234 5678 9012", [:size = 9, :color = "gray"])

doc.addEmptyParagraph()
doc.addParagraph("Thank you for your business!", [:italic = true, :align = "center", :size = 12])

if doc.save("demo_richcell_10_invoice.docx")
    ? "  Created: demo_richcell_10_invoice.docx"
else
    ? "  FAILED: demo_richcell_10_invoice.docx"
ok

# =====================================================================
# Demo 11: Color-Coded Dashboard Table
# =====================================================================
? "Demo 11: Color-coded dashboard table..."

doc = new WordWriter()
doc.setTitle("Project Dashboard")

doc.addHeading("Project Status Dashboard", 1)
doc.addParagraph("Real-time project tracking with color-coded status indicators.", NULL)

# Headers
hProject = wordCell("Project", [:bold = true, :color = "white", :align = "center"])
hLead    = wordCell("Lead", [:bold = true, :color = "white", :align = "center"])
hProg    = wordCell("Progress", [:bold = true, :color = "white", :align = "center"])
hStatus  = wordCell("Status", [:bold = true, :color = "white", :align = "center"])

# Project rows with status coloring
proj1Name = new WordCell()
proj1Name.addRun("Ring 1.26", [:bold = true, :size = 12])
proj1Name.addLineBreak()
proj1Name.addRun("New language features", [:size = 9, :color = "gray"])

proj1Status = wordCell("On Track", [:bold = true, :color = "white", :bgColor = "4CAF50", :align = "center"])

proj2Name = new WordCell()
proj2Name.addRun("RingSlint UI", [:bold = true, :size = 12])
proj2Name.addLineBreak()
proj2Name.addRun("UI toolkit integration", [:size = 9, :color = "gray"])

proj2Status = wordCell("On Track", [:bold = true, :color = "white", :bgColor = "4CAF50", :align = "center"])

proj3Name = new WordCell()
proj3Name.addRun("Rust Bindings", [:bold = true, :size = 12])
proj3Name.addLineBreak()
proj3Name.addRun("Language interop layer", [:size = 9, :color = "gray"])

proj3Status = wordCell("At Risk", [:bold = true, :color = "white", :bgColor = "FF9800", :align = "center"])

proj4Name = new WordCell()
proj4Name.addRun("Archive Package", [:bold = true, :size = 12])
proj4Name.addLineBreak()
proj4Name.addRun("Compression library", [:size = 9, :color = "gray"])

proj4Status = wordCell("Delayed", [:bold = true, :color = "white", :bgColor = "F44336", :align = "center"])

proj5Name = new WordCell()
proj5Name.addRun("Closures", [:bold = true, :size = 12])
proj5Name.addLineBreak()
proj5Name.addRun("Language core feature", [:size = 9, :color = "gray"])

proj5Status = wordCell("Complete", [:bold = true, :color = "white", :bgColor = "2196F3", :align = "center"])

tableData = [
    [hProject, hLead, hProg, hStatus],
    [proj1Name, "Mahmoud", wordCell("85%", [:align = "center", :bold = true]), proj1Status],
    [proj2Name, "Ahmed", wordCell("70%", [:align = "center", :bold = true]), proj2Status],
    [proj3Name, "Sara", wordCell("45%", [:align = "center", :bold = true]), proj3Status],
    [proj4Name, "Omar", wordCell("20%", [:align = "center", :bold = true]), proj4Status],
    [proj5Name, "Mahmoud", wordCell("100%", [:align = "center", :bold = true, :color = "green"]), proj5Status]
]

doc.addTable(tableData, [
    :borderStyle = "single",
    :headerRow = true,
    :headerBgColor = "263238",
    :headerVAlign = "center",
    :cellVAlign = "center"
])

if doc.save("demo_richcell_11_dashboard.docx")
    ? "  Created: demo_richcell_11_dashboard.docx"
else
    ? "  FAILED: demo_richcell_11_dashboard.docx"
ok

# =====================================================================
# Demo 12: Feature Comparison Table
# =====================================================================
? "Demo 12: Feature comparison table..."

doc = new WordWriter()
doc.setTitle("Feature Comparison")

doc.addHeading("Programming Language Feature Comparison", 1)

# Check mark and X mark using text
checkMark = "YES"
crossMark = "NO"

# Headers
hFeature = wordCell("Feature", [:bold = true, :color = "white", :align = "center"])
hRing    = wordCell("Ring", [:bold = true, :color = "white", :align = "center"])
hPython  = wordCell("Python", [:bold = true, :color = "white", :align = "center"])
hJava    = wordCell("Java", [:bold = true, :color = "white", :align = "center"])

# Feature rows
features = [
    ["Dynamic Typing", checkMark, checkMark, crossMark],
    ["Object-Oriented", checkMark, checkMark, checkMark],
    ["Functional Style", checkMark, checkMark, checkMark],
    ["PWCT2", checkMark, crossMark, crossMark],
    ["Declarative UI", checkMark, crossMark, crossMark],
    ["Embedded C Support", checkMark, crossMark, checkMark],
    ["Garbage Collection", checkMark, checkMark, checkMark],
    ["Natural Language Coding", checkMark, crossMark, crossMark]
]

tableData = [[hFeature, hRing, hPython, hJava]]

for feature in features
    featureCell = wordCell(feature[1], [:bold = true])
    
    cells = [featureCell]
    for i = 2 to 4
        if feature[i] = checkMark
            cells + wordCell(checkMark, [:bold = true, :color = "green", :align = "center", :bgColor = "E8F5E9"])
        else
            cells + wordCell(crossMark, [:color = "red", :align = "center", :bgColor = "FFEBEE"])
        ok
    next
    
    tableData + cells
next

doc.addTable(tableData, [
    :borderStyle = "single",
    :headerRow = true,
    :headerBgColor = "37474F",
    :headerVAlign = "center",
    :cellVAlign = "center"
])

if doc.save("demo_richcell_12_comparison.docx")
    ? "  Created: demo_richcell_12_comparison.docx"
else
    ? "  FAILED: demo_richcell_12_comparison.docx"
ok

# =====================================================================
# Summary
# =====================================================================

? ""
? "=============================================="
? "   All Rich Table Cell Demos Completed!"
? "=============================================="
? ""
? "Created files:"
? "  1.  demo_richcell_1_formatted_text.docx    - Formatted text in cells"
? "  2.  demo_richcell_2_multirun.docx          - Multiple text styles per cell"
? "  3.  demo_richcell_3_alignment.docx         - Horizontal & vertical alignment"
? "  4.  demo_richcell_4_colors.docx            - Custom cell background colors"
? "  5.  demo_richcell_5_colspan.docx           - Column spanning (merged cells)"
? "  6.  demo_richcell_6_linebreaks.docx        - Line breaks within cells"
? "  7.  demo_richcell_7_padding_borders.docx   - Cell padding & individual borders"
? "  8.  demo_richcell_8_images.docx            - Images inside table cells"
? "  9.  demo_richcell_9_mixed.docx             - Mixed plain strings & WordCell"
? "  10. demo_richcell_10_invoice.docx          - Professional invoice"
? "  11. demo_richcell_11_dashboard.docx        - Color-coded dashboard"
? "  12. demo_richcell_12_comparison.docx       - Feature comparison table"
? ""
