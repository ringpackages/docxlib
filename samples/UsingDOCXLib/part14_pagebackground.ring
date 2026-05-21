/*
	DOCXLib - Demo (Page Background)
*/

load "docxlib.ring"

/*
    Demonstrates four features:
      A. Odd/Even page headers and footers
      B. Paragraph borders
      C. Section breaks with orientation change
      D. Page background color
*/

doc = new WordWriter()

doc.setTitle("DOCXLib - New Features Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")

# -------------------------------------------------
# Feature A: Odd/Even page headers/footers
# Odd pages (right-hand): chapter/section title on the right
# Even pages (left-hand): document title on the left
# -------------------------------------------------
doc.setHeader("Chapter 1 - DOCXLib Features")        # odd pages
doc.setEvenPageHeader("DOCXLib Documentation") # even pages
doc.showPageNumbers("right")
doc.setEvenPageFooter("Page - DOCXLib")

# -------------------------------------------------
# Feature D: Page background color
# A very light blue tint - visible in Word when
# "Print Background Colors" is on
# -------------------------------------------------
doc.setPageBackground("EBF3FB")

# -------------------------------------------------------
# PAGE 1 - Intro
# -------------------------------------------------------

doc.addHeading("DOCXLib - New Features", 1)
doc.addParagraph("This document demonstrates four features. Each feature section is clearly labeled and includes live examples.", NULL)
doc.addParagraph("", NULL)

# Overview table
doc.addTable(
    [
        ["Feature", "API"],
        ["A: Odd/Even page headers/footers", "setEvenPageHeader() / setEvenPageFooter()"],
        ["B: Paragraph borders",             "addBorderedParagraph() / [:_hasBorder]"],
        ["C: Section breaks + orientation",  "startLandscapeSection() / endLandscapeSection()"],
        ["D: Page background color",         "setPageBackground()"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "1F4E79",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "D9E8F5",
        :colWidths = [7, 9]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 2 - Feature A
# -------------------------------------------------------

doc.addHeading("Feature A - Odd/Even Page Headers and Footers", 1)
doc.addParagraph("In books and formal reports, right-hand (odd) pages typically show the chapter title while left-hand (even) pages show the document or book title. Word calls this Even and Odd Headers.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph("How to use: call setHeader() for odd pages and setEvenPageHeader() for even pages. Calling either method automatically enables even/odd mode.",
    [:borderStyle = "single", :borderColor = "2E75B6", :borderSize = 6,
     :bgColor = "DEEAF1", :spaceBefore = 80, :spaceAfter = 80, :indent = 240])

doc.addParagraph("", NULL)
doc.addParagraph("This document itself uses Feature A. Check the headers above as you page through it - odd pages show the chapter name on the right, even pages show the document title.", [:italic = true, :color = "595959"])
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Page type", "Header",                                  "Footer"],
        ["Odd pages",  "Chapter 1 - DOCXLib Features",       "Page N (right-aligned)"],
        ["Even pages", "DOCXLib Documentation",       "Page - DOCXLib (center)"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "2E75B6",
        :headerTextColor = "FFFFFF",
        :colWidths = [3, 7, 6]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 3 - Feature B
# -------------------------------------------------------

doc.addHeading("Feature B - Paragraph Borders", 1)
doc.addParagraph("Paragraph borders draw a rule on one or more sides of a paragraph. They are widely used for callout boxes, notes, warnings, and visual separators - without using a floating text box.", NULL)
doc.addParagraph("", NULL)

# Full box with default single border
doc.addBorderedParagraph("Default box border - all four sides, single style, 0.75pt, black. This is the simplest way to frame a paragraph.",
    [:spaceBefore = 80, :spaceAfter = 80, :indent = 240])

doc.addParagraph("", NULL)

# Colored thick box
doc.addBorderedParagraph("Thick blue border - borderStyle:thick, borderColor:2E75B6, borderSize:18 (2.25pt). Good for important callouts.",
    [:borderStyle = "thick", :borderColor = "2E75B6", :borderSize = 18,
     :bgColor = "EBF3FB", :spaceBefore = 80, :spaceAfter = 80, :indent = 240,
     :bold = true])

doc.addParagraph("", NULL)

# Top and bottom only (horizontal rules)
doc.addBorderedParagraph("Top and bottom borders only - sides:['top','bottom'], creates a horizontal rule effect above and below the text without side lines.",
    [:sides = ["top", "bottom"], :borderColor = "595959", :borderSize = 8,
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240, :align = "center", :italic = true])

doc.addParagraph("", NULL)

# Left side only (blockquote style)
doc.addBorderedParagraph("Left border only - sides:['left'], borderColor:ED7D31, borderSize:24. Creates a classic blockquote-style sidebar rule.",
    [:sides = ["left"], :borderColor = "ED7D31", :borderSize = 24,
     :bgColor = "FFF2CC", :spaceBefore = 80, :spaceAfter = 80, :indent = 360])

doc.addParagraph("", NULL)

# Wave border
doc.addBorderedParagraph("Wave border - borderStyle:wave, all sides. Decorative border useful for certificates, titles, or creative documents.",
    [:borderStyle = "wave", :borderColor = "7030A0", :borderSize = 10,
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240,
     :align = "center", :color = "7030A0", :bold = true])

doc.addParagraph("", NULL)

# Double border
doc.addBorderedParagraph("Double border - borderStyle:double, borderColor:1F4E79, borderSize:6. Traditional formal styling.",
    [:borderStyle = "double", :borderColor = "1F4E79", :borderSize = 6,
     :bgColor = "E2EFDA", :spaceBefore = 80, :spaceAfter = 80, :indent = 240])

doc.addParagraph("", NULL)

doc.addParagraph("API reference:", [:bold = true])
doc.addParagraph("addBorderedParagraph(text, options) accepts: :borderStyle, :borderColor, :borderSize, :borderSpace, :sides (list), plus all standard paragraph options.", [:size = 10, :color = "595959"])

doc.addPageBreak()

# -------------------------------------------------------
# PAGES 4-5 - Feature C (portrait → landscape → portrait)
# -------------------------------------------------------

doc.addHeading("Feature C - Section Breaks with Orientation Change", 1)
doc.addParagraph("Documents often need a single landscape page for a wide table or diagram, surrounded by portrait pages. DOCXLib v1.7.0 adds startLandscapeSection() and endLandscapeSection() to make this seamless.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph("The next page is landscape. After the wide table, the document returns to portrait automatically.",
    [:borderStyle = "single", :borderColor = "375623", :bgColor = "E2EFDA",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240])

# -- switch to landscape --
doc.addLandscapeStart()

doc.addHeading("Wide Data Table - Landscape Page", 2)
doc.addParagraph("This entire page is landscape. The table below would not fit on a portrait page.", [:italic = true, :color = "595959"])
doc.addParagraph("", NULL)

# Wide table with many columns - only fits on landscape
doc.addTable(
    [
        ["Country",      "Region",        "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Annual",   "Growth", "Rank"],
        ["Saudi Arabia", "Middle East",   "12,450",   "14,200",   "13,800",   "15,600",   "56,050",   "+12.4%", "1"],
        ["UAE",          "Middle East",   "9,800",    "11,200",   "10,900",   "12,400",   "44,300",   "+9.8%",  "2"],
        ["Egypt",        "North Africa",  "7,200",    "8,100",    "7,900",    "8,800",    "32,000",   "+8.2%",  "3"],
        ["Kuwait",       "Middle East",   "5,600",    "6,300",    "6,100",    "6,900",    "24,900",   "+7.5%",  "4"],
        ["Jordan",       "Middle East",   "4,100",    "4,700",    "4,500",    "5,100",    "18,400",   "+6.9%",  "5"],
        ["Morocco",      "North Africa",  "3,800",    "4,200",    "4,100",    "4,600",    "16,700",   "+5.8%",  "6"],
        ["Tunisia",      "North Africa",  "2,900",    "3,300",    "3,200",    "3,600",    "13,000",   "+4.9%",  "7"],
        ["Lebanon",      "Middle East",   "2,400",    "2,700",    "2,600",    "2,900",    "10,600",   "+3.2%",  "8"]
    ],
    [
        :headerRow = true,
        :repeatHeader = true,
        :borderStyle = "single",
        :headerBgColor = "1F4E79",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "EEF3F8",
        :colWidths = [3.5, 3.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 1.8]
    ]
)

# -- return to portrait --
doc.addLandscapeEnd()

doc.addHeading("Back to Portrait", 2)
doc.addParagraph("This page is portrait again. The document automatically resumed its original page dimensions after endLandscapeSection().", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph("Key API: startLandscapeSection() inserts a section break that switches to landscape. endLandscapeSection() inserts a break that returns to portrait. All content between the two calls is rendered on landscape pages.",
    [:borderStyle = "single", :borderColor = "2E75B6", :bgColor = "DEEAF1",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240])

doc.addPageBreak()

# -------------------------------------------------------
# LAST PAGE - Feature D
# -------------------------------------------------------

doc.addHeading("Feature D - Page Background Color", 1)
doc.addParagraph("The entire document has a light blue background (color EBF3FB) applied via setPageBackground(). This is visible in Word when View > Print Layout is active and print background colors is enabled.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph("setPageBackground() adds <w:background w:color='...'/> before <w:body> and sets <w:displayBackgroundShape/> in settings.xml so Word shows it on screen without any manual steps.",
    [:borderStyle = "single", :borderColor = "2E75B6", :bgColor = "D9E8F5",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240])

doc.addParagraph("", NULL)
doc.addParagraph("Common use cases:", [:bold = true])
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Use case",            "Suggested color", "Hex"],
        ["Light parchment",     "Warm off-white",  "FFF8DC"],
        ["Legal / formal",      "Pale grey",       "F2F2F2"],
        ["Technical docs",      "Light blue",      "EBF3FB"],
        ["Warning / alert",     "Light yellow",    "FFFDE7"],
        ["Ecological theme",    "Light green",     "F1F8E9"],
        ["Premium / dark mode", "Dark navy",       "1F3864"]
    ],
    [
        :headerRow = true,
        :borderStyle = "single",
        :headerBgColor = "404040",
        :headerTextColor = "FFFFFF",
        :evenRowBgColor = "F0F6FB",
        :colWidths = [5, 4.5, 3.5]
    ]
)

doc.addParagraph("", NULL)
doc.addParagraph("Note: passing NULL or an empty string to setPageBackground() removes the background color.", [:italic = true, :color = "595959"])

# -------------------------------------------------------
# SAVE
# -------------------------------------------------------
filename = "demo_pagebackground.docx"
result = doc.save(filename)

if result
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
