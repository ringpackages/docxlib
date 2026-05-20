/*
	DOCXLib - Demo (Paragraph Shading)
*/

load "docxlib.ring"

/*
    Demonstrates the next features:
      1. Table cell background color / shading     (via addTable headerBgColor / evenRowBgColor)
      2. Custom paragraph styles (defineStyle)
      3. Repeating header rows    (table option :repeatHeader)
      4. Different first-page header/footer
      5. Paragraph shading        (addShadedParagraph / [:bgColor] option)
      6. Floating text boxes      (addTextBox)
*/

doc = new WordWriter()

doc.setTitle("DOCXLib - Features Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")

# -------------------------------------------------
# Different first-page header/footer
# The cover page shows a bold title banner; all
# subsequent pages show the regular running header.
# -------------------------------------------------
doc.setFirstPageHeader("DOCXLib - Feature Showcase")
doc.setFirstPageFooter("Confidential - Do Not Distribute")
doc.setHeader("DOCXLib Documentation")
doc.showPageNumbers("right")

# -------------------------------------------------------
# COVER PAGE
# -------------------------------------------------------

doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)

doc.addParagraph("DOCXLib", [:align = "center", :bold = true, :size = 36, :color = "1F4E79"])
doc.addParagraph("Version 1.0", [:align = "center", :bold = true, :size = 20, :color = "2F75B6"])
doc.addParagraph("", NULL)
doc.addParagraph("New Features Demo", [:align = "center", :italic = true, :size = 16, :color = "595959"])
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)



# -------------------------------------------------
# Floating text box on cover page
# A styled announcement box placed in the center
# -------------------------------------------------
doc.addTextBox("New Features in This Sample",
    [:x = 4, :y = 10, :width = 10, :height = 2,
     :align = "center", :bold = true, :size = 14,
     :color = "FFFFFF", :bgColor = "2E75B6",
     :borderColor = "1F4E79", :borderSize = 12])

doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)

doc.addParagraph("Created with Ring Programming Language", [:align = "center", :italic = true, :size = 10, :color = "7F7F7F"])
doc.addParagraph("https://ring-lang.github.io", [:align = "center", :size = 10, :color = "0563C1"])

doc.addPageBreak()



# -------------------------------------------------------
# TABLE OF NEW FEATURES
# -------------------------------------------------------
doc.addHeading("Overview of New Features", 1)
doc.addParagraph("The following features were added in this sample:", NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["#", "Feature", "Key API"],
        ["1", "Table cell background color / shading",   "addTable + :headerBgColor / :evenRowBgColor / :bgColor"],
        ["2", "Custom paragraph styles",                  "defineStyle() + [:style = 'MyId']"],
        ["3", "Repeating table header rows",             "addTable + :repeatHeader = true"],
        ["4", "Different first-page header/footer",      "setFirstPageHeader() / setFirstPageFooter()"],
        ["5", "Paragraph shading",                        "addShadedParagraph() / [:bgColor = 'RRGGBB']"],
        ["6", "Floating text boxes",                     "addTextBox()"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "1F4E79",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "D6E4F0",
        :colWidths = [0.8, 6.5, 8.5]
    ]
)


# -------------------------------------------------------
# FEATURE 1 - TABLE CELL BACKGROUND COLOR
# -------------------------------------------------------
doc.addHeading("Feature 1 - Table Cell Background Color", 1)
doc.addParagraph("Tables now support per-cell background shading through three complementary options:", NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Option",          "Purpose",                      "Example Value"],
        [":headerBgColor",  "Fill color for the header row","2E75B6"],
        [":evenRowBgColor", "Zebra-stripe color for even rows", "D9E2F3"],
        [":bgColor on WordCell", "Per-cell individual fill",  "FFD700"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "2E75B6",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "DEEAF1",
        :colWidths = [4.5, 6, 4]
    ]
)

doc.addParagraph("", NULL)
doc.addParagraph("Example - product catalogue with colored tiers:", [:bold = true])
doc.addParagraph("", NULL)


# Product catalogue demo with per-cell colors via WordCell
headerCells = [
    new WordCell() { setText("Product", NULL) },
    new WordCell() { setText("Tier", NULL) },
    new WordCell() { setText("Price", NULL) },
    new WordCell() { setText("Stock", NULL) }
]

row1 = [
    new WordCell() { setText("Widget Pro X", NULL) },
    new WordCell() { setText("Gold", NULL) .setBgColor("FFD700") },
    new WordCell() { setText("$199.00", NULL) },
    new WordCell() { setText("In Stock", NULL) .setBgColor("C6EFCE") }
]
row2 = [
    new WordCell() { setText("Widget Standard", NULL) },
    new WordCell() { setText("Silver", NULL) .setBgColor("BFBFBF") },
    new WordCell() { setText("$99.00", NULL) },
    new WordCell() { setText("In Stock", NULL) .setBgColor("C6EFCE") }
]
row3 = [
    new WordCell() { setText("Widget Basic", NULL) },
    new WordCell() { setText("Bronze", [:color = "FFFFFF"]) .setBgColor("ED7D31") },
    new WordCell() { setText("$49.00", NULL) },
    new WordCell() { setText("Low Stock", NULL) .setBgColor("FFEB9C") }
]
row4 = [
    new WordCell() { setText("Widget Lite", NULL) },
    new WordCell() { setText("Basic", NULL) },
    new WordCell() { setText("$19.00", NULL) },
    new WordCell() { setText("Out of Stock", NULL) .setBgColor("FFC7CE") }
]

doc.addTable(
    [headerCells, row1, row2, row3, row4],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "404040",
        :headerTextColor = "FFFFFF",
        :colWidths = [5, 3, 3, 3]
    ]
)

doc.addPageBreak()

doc.addPageBreak()

# -------------------------------------------------------
# FEATURE 2 - CUSTOM PARAGRAPH STYLES
# -------------------------------------------------------
doc.addHeading("Feature 2 - Custom Paragraph Styles", 1)
doc.addParagraph("Use defineStyle() to declare reusable named styles. Apply them with [:style = 'StyleId'].", NULL)
doc.addParagraph("", NULL)

# Define styles
doc.defineStyle("CalloutNote",
    [:name = "Callout Note", :bold = false, :italic = true,
     :size = 11, :color = "1F4E79", :bgColor = "DEEAF1",
     :align = "both", :spaceBefore = 120, :spaceAfter = 120,
     :indent = 360, :keepLines = true])

doc.defineStyle("WarningBox",
    [:name = "Warning Box", :bold = true,
     :size = 11, :color = "7B2C2C", :bgColor = "FCE4E4",
     :align = "left", :spaceBefore = 120, :spaceAfter = 120,
     :indent = 360])

doc.defineStyle("SectionSummary",
    [:name = "Section Summary", :bold = false, :italic = false,
     :size = 12, :color = "375623", :bgColor = "E2EFDA",
     :align = "both", :spaceBefore = 80, :spaceAfter = 80,
     :lineSpacing = 1.4])

doc.defineStyle("TipBox",
    [:name = "Tip Box", :bold = false, :italic = false,
     :size = 11, :color = "375623",
     :align = "left", :spaceBefore = 80, :spaceAfter = 80,
     :indent = 360])

# Demo usage
doc.addParagraph("The following paragraphs each use a different custom style:", NULL)
doc.addParagraph("", NULL)

doc.addParagraph("CalloutNote style - Used for informational callouts inside documentation. The background and italic formatting draw the reader's eye without disrupting document flow.",
    [:style = "CalloutNote"])
doc.addParagraph("", NULL)

doc.addParagraph("WarningBox style - IMPORTANT: This style highlights critical warnings. The red background immediately signals the reader to pay attention to this content.",
    [:style = "WarningBox"])
doc.addParagraph("", NULL)

doc.addParagraph("SectionSummary style - This green-tinted summary style is ideal for end-of-section recaps, concluding remarks, or key takeaway paragraphs in long reports.",
    [:style = "SectionSummary"])
doc.addParagraph("", NULL)

doc.addParagraph("TipBox style - Tip: Custom styles are defined once at the top of your script and can be reused across any number of paragraphs throughout the entire document.",
    [:style = "TipBox"])

doc.addPageBreak()

# -------------------------------------------------------
# FEATURE 3 - REPEATING HEADER ROWS
# -------------------------------------------------------
doc.addHeading("Feature 3 - Repeating Table Header Rows", 1)
doc.addParagraph("When a table spans multiple pages, adding :repeatHeader = true causes Word to repeat the first (header) row at the top of every page automatically.", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("The 30-row table below will overflow to the next page. Open the document in Word and scroll to verify the blue header row repeats.", [:italic = true, :color = "595959"])
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("", NULL)
# Build a 30-row table that overflows a page
tableData = [["#", "Country", "Capital City", "Population (M)", "Area (km²)"]]
countries = [
    ["1",  "China",         "Beijing",         "1,412",   "9,597,000"],
    ["2",  "India",         "New Delhi",        "1,380",   "3,287,000"],
    ["3",  "USA",           "Washington D.C.",  "331",     "9,834,000"],
    ["4",  "Indonesia",     "Jakarta",          "274",     "1,905,000"],
    ["5",  "Pakistan",      "Islamabad",        "221",     "881,000"],
    ["6",  "Brazil",        "Brasília",         "213",     "8,516,000"],
    ["7",  "Nigeria",       "Abuja",            "211",     "924,000"],
    ["8",  "Bangladesh",    "Dhaka",            "166",     "147,000"],
    ["9",  "Russia",        "Moscow",           "146",     "17,098,000"],
    ["10", "Ethiopia",      "Addis Ababa",      "122",     "1,104,000"],
    ["11", "Mexico",        "Mexico City",      "129",     "1,964,000"],
    ["12", "Japan",         "Tokyo",            "125",     "378,000"],
    ["13", "Philippines",   "Manila",           "113",     "300,000"],
    ["14", "Egypt",         "Cairo",            "102",     "1,002,000"],
    ["15", "DR Congo",      "Kinshasa",         "100",     "2,345,000"],
    ["16", "Vietnam",       "Hanoi",            "97",      "331,000"],
    ["17", "Iran",          "Tehran",           "84",      "1,648,000"],
    ["18", "Turkey",        "Ankara",           "84",      "784,000"],
    ["19", "Germany",       "Berlin",           "83",      "357,000"],
    ["20", "Thailand",      "Bangkok",          "70",      "513,000"],
    ["21", "UK",            "London",           "67",      "244,000"],
    ["22", "France",        "Paris",            "67",      "551,000"],
    ["23", "Tanzania",      "Dodoma",           "60",      "945,000"],
    ["24", "South Africa",  "Pretoria",         "60",      "1,219,000"],
    ["25", "Myanmar",       "Naypyidaw",        "54",      "677,000"],
    ["26", "Kenya",         "Nairobi",          "54",      "580,000"],
    ["27", "South Korea",   "Seoul",            "52",      "100,000"],
    ["28", "Colombia",      "Bogotá",           "51",      "1,142,000"],
    ["29", "Spain",         "Madrid",           "47",      "506,000"],
    ["30", "Argentina",     "Buenos Aires",     "45",      "2,780,000"]
]

countriesLen = len(countries)
for ci = 1 to countriesLen
    tableData + countries[ci]
next

doc.addTable(tableData,
    [
        :headerRow = true,
        :repeatHeader = true,
        :borderStyle = "single",
        :headerBgColor = "1F4E79",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "EEF3F8",
        :colWidths = [0.8, 4, 4, 3.5, 3.5]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# FEATURE 4 - FIRST-PAGE DIFFERENT HEADER/FOOTER
# -------------------------------------------------------
doc.addHeading("Feature 4 - Different First-Page Header and Footer", 1)
doc.addParagraph("This document itself demonstrates Feature 4. Look at the header and footer areas:", NULL)
doc.addParagraph("", NULL)

# Use addTable with bold rows to explain what to look for
doc.addTable(
    [
        ["Page",           "Header shows",                                "Footer shows"],
        ["Page 1 (cover)", "DOCXLib v1.6.0 - Feature Showcase",      "Confidential - Do Not Distribute"],
        ["Pages 2+",       "DOCXLib v1.6.0 Documentation",           "Page N of M  (right-aligned)"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "404040",
        :headerTextColor = "FFFFFF",
        :colWidths = [3, 7, 5]
    ]
)

doc.addParagraph("", NULL)
doc.addParagraph("API used in this document:", [:bold = true])
doc.addParagraph("", NULL)

codeLines = [
    'doc.setFirstPageHeader("DOCXLib v1.6.0 - Feature Showcase")',
    'doc.setFirstPageFooter("Confidential - Do Not Distribute")',
    'doc.setHeader("DOCXLib v1.6.0 Documentation")',
    'doc.showPageNumbers("right")'
]

codeLinesLen = len(codeLines)
for ci = 1 to codeLinesLen
    doc.addParagraph(codeLines[ci],
        [:font = "Courier New", :size = 10, :bgColor = "F2F2F2",
         :spaceAfter = 40, :spaceBefore = 40, :indent = 360])
next

doc.addParagraph("", NULL)
doc.addParagraph("You can also call setFirstPageDifferent(true) without providing any first-page text to simply suppress the header/footer on page 1 (common for title pages).", [:italic = true, :color = "595959"])

doc.addPageBreak()

# -------------------------------------------------------
# FEATURE 5 - PARAGRAPH SHADING
# -------------------------------------------------------
doc.addHeading("Feature 5 - Paragraph Shading", 1)
doc.addParagraph("Individual paragraphs can now have a background fill color. There are two ways to apply it:", NULL)
doc.addParagraph("", NULL)

doc.addParagraph("Method 1 - addShadedParagraph(text, color, options)", [:bold = true])
doc.addParagraph("", NULL)
doc.addShadedParagraph("This paragraph uses addShadedParagraph() with color '1F4E79'. The background fills the full width of the paragraph text block.", "1F4E79",
    [:color = "FFFFFF", :bold = true, :spaceBefore = 60, :spaceAfter = 60, :indent = 360])

doc.addShadedParagraph("Warning: This is a yellow-highlighted paragraph. Great for callout boxes, tips, notes, and other highlighted content.", "FFF2CC",
    [:color = "7E6000", :spaceBefore = 60, :spaceAfter = 60, :indent = 360])

doc.addShadedParagraph("Success: This green shading works well for confirming positive actions, status indicators, or checklist completion messages.", "E2EFDA",
    [:color = "375623", :spaceBefore = 60, :spaceAfter = 60, :indent = 360])

doc.addShadedParagraph("Error: This red-shaded paragraph immediately draws attention to errors, failures, or critical missing information.", "FCE4E4",
    [:color = "7B2C2C", :bold = true, :spaceBefore = 60, :spaceAfter = 60, :indent = 360])

doc.addParagraph("", NULL)
doc.addParagraph("Method 2 - [:bgColor] option in addParagraph()", [:bold = true])
doc.addParagraph("", NULL)
doc.addParagraph("This paragraph uses the :bgColor key in a standard addParagraph() options list.",
    [:bgColor = "E9F5FB", :color = "1F4E79", :italic = true, :spaceBefore = 60, :spaceAfter = 60, :indent = 360])

doc.addParagraph("", NULL)
doc.addParagraph("Shading also works seamlessly with custom styles (Feature 2) by setting :bgColor in defineStyle().", [:italic = true, :color = "595959"])

doc.addPageBreak()

# -------------------------------------------------------
# FEATURE 6 - FLOATING TEXT BOXES
# -------------------------------------------------------
doc.addHeading("Feature 6 - Floating Text Boxes", 1)
doc.addParagraph("Text boxes are absolutely-positioned floating containers. They can appear anywhere on the page, overlapping or alongside regular content.", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("The boxes shown on this page are live - open in Word to see their exact positions.", [:italic = true, :color = "595959"])
doc.addParagraph("", NULL)

# Several demo text boxes at different positions
doc.addTextBox("SIDEBAR  This is a sidebar text box floating to the right of the main content area. Text boxes can hold any text content.",
    [:x = 11.5, :y = 5.5, :width = 5.5, :height = 5,
     :align = "left", :size = 10,
     :bgColor = "EEF3F8", :borderColor = "2E75B6", :borderSize = 8])

doc.addTextBox("NOTE",
    [:x = 0.5, :y = 6, :width = 2.5, :height = 1,
     :align = "center", :bold = true, :size = 12,
     :color = "FFFFFF", :bgColor = "ED7D31",
     :borderColor = "C55A11", :borderSize = 8])

doc.addParagraph("Option reference for addTextBox():", [:bold = true])
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Option",       "Default",    "Description"],
        [":x",           "2.0",        "Horizontal position from left margin (cm)"],
        [":y",           "5.0",        "Vertical position from top margin (cm)"],
        [":width",       "6.0",        "Box width in cm"],
        [":height",      "3.0",        "Box height in cm"],
        [":align",       "center",     "Text alignment inside box: left/center/right"],
        [":bold",        "false",      "Bold text inside the box"],
        [":italic",      "false",      "Italic text inside the box"],
        [":font",        "(default)",  "Font family name"],
        [":size",        "(default)",  "Font size in points"],
        [":color",       "(default)",  "Text color (name or hex)"],
        [":bgColor",     "FFFFFF",     "Box fill color"],
        [":borderColor", "4472C4",     "Border color"],
        [":borderSize",  "8",          "Border width in eighths of a point (8 = 1pt)"],
        [":noFill",      "false",      "Transparent background"],
        [":noBorder",    "false",      "No visible border"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "404040",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "F2F2F2",
        :colWidths = [3, 2.5, 10]
    ]
)

doc.addParagraph("", NULL)

# A no-fill bordered text box
doc.addTextBox("This text box has no fill (transparent background) but keeps its border.",
    [:x = 1, :y = 19, :width = 9, :height = 1.5,
     :align = "center", :italic = true, :size = 11,
     :noFill = true, :borderColor = "7030A0", :borderSize = 16])

# -------------------------------------------------------
# SAVE
# -------------------------------------------------------
filename = "demo_paragraphshading.docx"
result = doc.save(filename)

if result
    ? "OK - Document saved: " + filename
else
    ? "ERROR - Failed to save document"
ok
