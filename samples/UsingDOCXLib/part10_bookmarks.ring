/*
	DOCXLib - Demo (Bookmarks)
*/

load "docxlib.ring"

/*
    Demonstrates three new features:
      1. Bookmarks and cross-references
      2. Tab stops (with leaders)
      3. Line numbers
*/

doc = new WordWriter()
doc.setTitle("DOCXLib - Features Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setHeader("DOCXLib - Features Demo")
doc.showPageNumbers("right")

# -------------------------------------------------------
# PAGE 1 — Cover / Overview
# -------------------------------------------------------

doc.addHeading("DOCXLib", 1)
doc.addParagraph("Three features were demonstrated in this sample.", NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["#", "Feature",              "Key Methods"],
        ["1",  "Bookmarks + cross-references", "addBookmarkedParagraph() / addCrossRef()"],
        ["2",  "Tab stops with leaders",       "addTabbedParagraph() with :tabStops"],
        ["3",  "Line numbers",                 "enableLineNumbers() / :suppressLineNumbers"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "1F4E79", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "EEF3F8", :colWidths = [1, 5.5, 8]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 2 — Feature 1: Bookmarks + Cross-References
# -------------------------------------------------------

doc.addBookmarkedParagraph("bm_feature1", "Feature 1 — Bookmarks and Cross-References", [:style = "Heading1"])

doc.addParagraph("A bookmark marks a named location in the document. A cross-reference inserts a field that automatically shows the page number where the bookmark lives — updating whenever Word recalculates fields (F9).", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: addBookmarkedParagraph(name, text, options) wraps a paragraph in a bookmark. addCrossRef(name, prefix) inserts a REF field that resolves to the bookmark's page.",
    [:borderStyle = "single", :borderColor = "2E75B6", :bgColor = "DEEAF1",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)
doc.addParagraph("", NULL)

# Section 1.1 — bookmark it
doc.addBookmarkedParagraph("bm_sec1_1", "1.1  Placing a Bookmark", [:style = "Heading2"])
doc.addParagraph("Use addBookmarkedParagraph() to mark any paragraph. The bookmark name must be unique across the document. Names may contain letters, digits and underscores — no spaces.", NULL)
doc.addParagraph("", NULL)

# Section 1.2 — bookmark it
doc.addBookmarkedParagraph("bm_sec1_2", "1.2  Inserting a Cross-Reference", [:style = "Heading2"])
doc.addParagraph("Use addCrossRef(bookmarkName, prefixText) anywhere after (or before) the bookmark. Word replaces the field with the live page number when you press F9 or print.", NULL)
doc.addParagraph("", NULL)

# Section 1.3 — cross-references back to earlier sections
doc.addBookmarkedParagraph("bm_sec1_3", "1.3  Live Examples", [:style = "Heading2"])
doc.addParagraph("The lines below are cross-references pointing back to sections on this page:", NULL)
doc.addParagraph("", NULL)

doc.addCrossRef("bm_sec1_1", "Section 1.1 starts")
doc.addCrossRef("bm_sec1_2", "Section 1.2 starts")
doc.addCrossRef("bm_feature2", "Feature 2 (Tab Stops) starts")
doc.addCrossRef("bm_feature3", "Feature 3 (Line Numbers) starts")

doc.addParagraph("", NULL)
doc.addParagraph("Press F9 in Word (or Ctrl+A then F9) to update all cross-reference fields to their live page numbers.", [:italic = true, :color = "595959"])

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 3 — Feature 2: Tab Stops
# -------------------------------------------------------

doc.addBookmarkedParagraph("bm_feature2", "Feature 2 — Tab Stops with Leaders", [:style = "Heading1"])

doc.addParagraph("Tab stops position text at precise horizontal locations within a paragraph. The optional leader fills the gap between the previous text and the tab stop with a repeated character (dots, dashes, underscores).", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: addTabbedParagraph(segments, options). Each item in segments is one column of text. Define :tabStops in options as a list of [:pos, :align, :leader].",
    [:borderStyle = "single", :borderColor = "375623", :bgColor = "E2EFDA",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)

doc.addParagraph("", NULL)
doc.addHeading("Example 1 — Table of Contents style (dot leader)", 2)
doc.addParagraph("Right-aligned page number with dot leader — classic TOC look:", NULL)
doc.addParagraph("", NULL)

# TOC-style entries: title on left, page number right-aligned with dots
tocTabs = [:tabStops = [[:pos = 8500, :align = "right", :leader = "dot"]]]
doc.addTabbedParagraph(["Introduction", "1"],             tocTabs)
doc.addTabbedParagraph(["Feature 1 — Bookmarks", "2"],    tocTabs)
doc.addTabbedParagraph(["Feature 2 — Tab Stops", "3"],    tocTabs)
doc.addTabbedParagraph(["Feature 3 — Line Numbers", "4"], tocTabs)
doc.addTabbedParagraph(["Conclusion", "5"],               tocTabs)

doc.addParagraph("", NULL)
doc.addHeading("Example 2 — Price list (dot leader)", 2)
doc.addParagraph("Product names on the left, prices right-aligned:", NULL)
doc.addParagraph("", NULL)

priceTabs = [:tabStops = [[:pos = 8500, :align = "right", :leader = "dot"]]]
doc.addTabbedParagraph(["Ring Language License",       "$0.00 (Open Source)"], priceTabs)
doc.addTabbedParagraph(["DOCXLib Library",         "$0.00 (Open Source)"], priceTabs)
doc.addTabbedParagraph(["PWCT2 Visual IDE",            "$0.00 (Open Source)"], priceTabs)
doc.addTabbedParagraph(["Supernova Language",          "$0.00 (Open Source)"], priceTabs)

doc.addParagraph("", NULL)
doc.addHeading("Example 3 — Multi-column layout (no leader)", 2)
doc.addParagraph("Three columns using left tab stops — simulates a simple table without borders:", NULL)
doc.addParagraph("", NULL)

# Header row (bold)
colTabs = [
    :bold = true,
    :tabStops = [
        [:pos = 3600, :align = "left", :leader = "none"],
        [:pos = 6800, :align = "left", :leader = "none"]
    ]
]
doc.addTabbedParagraph(["Country", "Capital", "Population"], colTabs)

# Data rows
colTabsNormal = [
    :tabStops = [
        [:pos = 3600, :align = "left", :leader = "none"],
        [:pos = 6800, :align = "left", :leader = "none"]
    ]
]
doc.addTabbedParagraph(["Saudi Arabia", "Riyadh",  "35 million"],  colTabsNormal)
doc.addTabbedParagraph(["Egypt",        "Cairo",   "104 million"], colTabsNormal)
doc.addTabbedParagraph(["UAE",          "Abu Dhabi","10 million"], colTabsNormal)
doc.addTabbedParagraph(["Kuwait",       "Kuwait City","4.7 million"], colTabsNormal)

doc.addParagraph("", NULL)
doc.addHeading("Example 4 — Underscore leader (signature lines)", 2)
doc.addParagraph("Common in forms and certificates:", NULL)
doc.addParagraph("", NULL)

sigTabs = [:tabStops = [[:pos = 7200, :align = "left", :leader = "underscore"]]]
doc.addTabbedParagraph(["Name: ", ""],      sigTabs)
doc.addTabbedParagraph(["Signature: ", ""], sigTabs)
doc.addTabbedParagraph(["Date: ", ""],      sigTabs)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 4 — Feature 3: Line Numbers
# -------------------------------------------------------

doc.addBookmarkedParagraph("bm_feature3", "Feature 3 — Line Numbers", [:style = "Heading1"])

doc.addParagraph("Line numbers are printed in the left margin beside each line of text. They are required by legal document standards and many academic journals.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: enableLineNumbers(options) activates line numbering. Options: :start (first number), :step (show every N lines), :distance (gap from text in twips), :restart (newPage | newSection | continuous). Use [:suppressLineNumbers = true] on a paragraph to exclude it from numbering.",
    [:borderStyle = "single", :borderColor = "7030A0", :bgColor = "EDE7F6",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)
doc.addParagraph("", NULL)

# Enable line numbers from here — step 5 so numbers show every 5 lines
doc.enableLineNumbers([:start = 1, :step = 5, :restart = "newPage"])

doc.addHeading("Numbered Section", 2)
doc.addParagraph("This paragraph has a line number. The document was configured with enableLineNumbers() using step:5, so numbers appear every 5 lines. This lets readers quickly locate specific passages without cluttering the margin with a number on every line.", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("Legal contracts, academic papers, court submissions, legislative drafts, and source code printouts all benefit from line numbers. The Ring programming language itself is often presented with numbered listings.", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("Cross-references in legal documents say things like 'see line 42' rather than 'see page 3, paragraph 2' — far more precise when the document may be reformatted. This is the primary reason courts and journals mandate line numbers.", NULL)
doc.addParagraph("", NULL)

# This paragraph suppresses its own line number (e.g. a caption or note)
doc.addParagraph("(This paragraph has suppressLineNumbers:true — it does not count toward line numbering and shows no line number in the margin.)",
    [:italic = true, :color = "595959", :suppressLineNumbers = true])

doc.addParagraph("", NULL)
doc.addParagraph("Numbering resumes on the next paragraph after the suppressed one. The count continues from where it left off before the suppressed line.", NULL)
doc.addParagraph("", NULL)
doc.addParagraph("The distance between the line number and the text body is controlled by :distance in twips (default 360 twips = 0.25 inch). Increase this if you need wider line numbers or if the document has a very narrow left margin.", NULL)

doc.addParagraph("", NULL)
doc.addHeading("Configuration Reference", 2)

doc.addTable(
    [
        ["Option",    "Type",    "Default",   "Description"],
        [":start",    "integer", "1",         "Line number of the first line"],
        [":step",     "integer", "1",         "Print number every N lines (5 = every 5th)"],
        [":distance", "integer", "360",       "Gap from left text edge in twips (720 = 0.5in)"],
        [":restart",  "string",  "newPage",   "newPage | newSection | continuous"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "4B0082", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "F3E5F5", :colWidths = [3, 2.5, 2.5, 6.5]
    ]
)

# -------------------------------------------------------
# SAVE
# -------------------------------------------------------
filename = "demo_bookmarks.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
