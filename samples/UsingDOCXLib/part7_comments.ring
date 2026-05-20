/*
	DOCXLib - Demo (Comments)
*/

load "docxlib.ring"

/*
    Demonstrates three features:
      1. Comments (annotations)
      2. SEQ captions (auto-numbered figures and tables)
      3. Mail merge fields
*/

doc = new WordWriter()
doc.setTitle("DOCXLib — Features Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setHeader("DOCXLib — Features Demo")
doc.showPageNumbers("right")

# -------------------------------------------------------
# PAGE 1 — Overview
# -------------------------------------------------------

doc.addHeading("DOCXLib — Three Features", 1)
doc.addParagraph("Three features are introduced in this sample.", NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["#", "Feature",                   "Key Methods"],
        ["1",  "Comments (annotations)",   "addCommentedParagraph() / addComment()"],
        ["2",  "SEQ captions",             "addFigureCaption() / addTableCaption()"],
        ["3",  "Mail merge fields",        "addMergeFieldParagraph() / addMergeTemplate()"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "1F4E79", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "EEF3F8", :colWidths = [1, 5, 8.5]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 2 — Feature 1: Comments
# -------------------------------------------------------

doc.addHeading("Feature 1 — Comments (Annotations)", 1)
doc.addParagraph("Comments let you embed reviewer notes into a document programmatically. Each comment is anchored to a specific paragraph and visible in Word's comment pane or as a balloon in the margin.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: addCommentedParagraph(text, commentText, options) attaches a comment to a paragraph. " +
    "Use :commentAuthor and :commentDate in options to set the reviewer name and timestamp.",
    [:borderStyle = "single", :borderColor = "2E75B6", :bgColor = "DEEAF1",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)
doc.addParagraph("", NULL)

doc.addHeading("Example Comments", 2)

doc.addCommentedParagraph(
    "The Ring programming language is an open-source project.",
    "Consider adding a citation or reference link here.",
    [:commentAuthor = "Editor", :commentDate = "2025-03-01T09:00:00Z"]
)

doc.addCommentedParagraph(
    "DOCXLib supports rich document generation including tables, images, headers, footers, watermarks, and comments.",
    "Great feature list",
    [:commentAuthor = "Reviewer", :commentDate = "2025-03-01T10:30:00Z"]
)

doc.addCommentedParagraph(
    "Performance benchmarks show that Ring generates a 50-page document in under 200ms on a standard desktop.",
    "Please verify this number against the latest benchmark run before publishing.",
    [:commentAuthor = "QA Team", :commentDate = "2025-03-02T14:15:00Z", :bold = true]
)

doc.addParagraph("", NULL)
doc.addParagraph("Each paragraph above has a comment attached. Open this file in Microsoft Word and switch to Review → Show Comments to see all three annotations in the comment pane.", [:italic = true, :color = "595959"])

doc.addParagraph("", NULL)
doc.addHeading("Multiple Reviewers Example", 2)

doc.addCommentedParagraph(
    "Section 2.1 — Architecture Overview",
    "This section needs a diagram. Use addImage() to insert the architecture PNG.",
    [:commentAuthor = "Mahmoud", :commentDate = "2025-03-03T08:00:00Z", :style = "Heading3"]
)

doc.addCommentedParagraph(
    "The virtual machine uses a stack-based design.",
    "Too technical for the introduction — move this paragraph to the implementation section.",
    [:commentAuthor = "Technical Writer", :commentDate = "2025-03-03T09:45:00Z"]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 3 — Feature 2: SEQ Captions
# -------------------------------------------------------

doc.addHeading("Feature 2 — SEQ Captions (Auto-Numbered)", 1)
doc.addParagraph("SEQ captions insert a Word SEQ field that automatically increments when the document is updated (F9). This produces 'Figure 1', 'Figure 2', 'Table 1', 'Table 2' labels that renumber themselves if items are added or removed.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: addFigureCaption(text) inserts a 'Figure N — text' caption. " +
    "addTableCaption(text) inserts a 'Table N — text' caption. " +
    "Both use Word's SEQ field internally — press F9 in Word to update all numbers.",
    [:borderStyle = "single", :borderColor = "375623", :bgColor = "E2EFDA",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)
doc.addParagraph("", NULL)

# --- Figure captions ---
doc.addHeading("Figure Captions", 2)
doc.addParagraph("The three grey boxes below represent image placeholders. Each is followed by an auto-numbered figure caption.", NULL)
doc.addParagraph("", NULL)

# Simulate figure with a shaded paragraph as placeholder
doc.addParagraph("[ Image placeholder: Ring language logo ]",
    [:align = "center", :bgColor = "D0D0D0", :italic = true,
     :spaceBefore = 120, :spaceAfter = 0])
doc.addFigureCaption("The Ring programming language logo")

doc.addParagraph("", NULL)

doc.addParagraph("[ Image placeholder: PWCT2 screenshot ]",
    [:align = "center", :bgColor = "D0D0D0", :italic = true,
     :spaceBefore = 120, :spaceAfter = 0])
doc.addFigureCaption("PWCT2 visual programming environment")

doc.addParagraph("", NULL)

doc.addParagraph("[ Image placeholder: Supernova IDE screenshot ]",
    [:align = "center", :bgColor = "D0D0D0", :italic = true,
     :spaceBefore = 120, :spaceAfter = 0])
doc.addFigureCaption("Supernova language integrated development environment")

doc.addParagraph("", NULL)

# --- Table captions ---
doc.addHeading("Table Captions", 2)
doc.addParagraph("Tables can also carry auto-numbered captions using addTableCaption().", NULL)
doc.addParagraph("", NULL)

doc.addTableCaption("Ring language release history")
doc.addTable(
    [
        ["Version", "Year", "Major Features"],
        ["1.0",  "2016", "Initial public release"],
        ["1.12", "2020", "List of features 1"],
        ["1.18", "2023", "List of features 2"],
        ["1.22", "2024", "List of features 3"],
        ["1.26", "2026", "List of features 4"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "375623", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "E2EFDA", :colWidths = [2.5, 2.5, 9.5]
    ]
)

doc.addParagraph("", NULL)

doc.addTableCaption("DOCXLib feature coverage by version")
doc.addTable(
    [
        ["Version", "Features Added"],
        ["0.1 – 0.5", "Paragraphs, headings, tables, images, headers, footers, lists"],
        ["0.6",       "Text boxes, first-page headers, rich cells, paragraph shading"],
        ["0.7",       "Odd/even headers, paragraph borders, landscape sections, page background"],
        ["0.8",       "Bookmarks, cross-references, tab stops with leaders, line numbers"],
        ["0.9",       "Comments, SEQ captions, mail merge fields"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "1F4E79", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "DEEAF1", :colWidths = [3, 11.5]
    ]
)

doc.addPageBreak()

# -------------------------------------------------------
# PAGE 4 — Feature 3: Mail Merge Fields
# -------------------------------------------------------

doc.addHeading("Feature 3 — Mail Merge Fields", 1)
doc.addParagraph("Mail merge fields are MERGEFIELD placeholders embedded in a template document. When Word merges the template with a data source (Excel, CSV, database), each placeholder is replaced with the actual data value for that record.", NULL)
doc.addParagraph("", NULL)

doc.addBorderedParagraph(
    "API: addMergeFieldParagraph(fields, options) adds a paragraph where strings are " +
    "literal text and lists [:field='Name'] are MERGEFIELD codes. " +
    "addMergeTemplate(lines) adds a series of such paragraphs at once.",
    [:borderStyle = "single", :borderColor = "7030A0", :bgColor = "EDE7F6",
     :spaceBefore = 80, :spaceAfter = 80, :indent = 240]
)
doc.addParagraph("", NULL)

# -- Example 1: Formal letter template --
doc.addHeading("Example 1 — Formal Letter Template", 2)
doc.addParagraph("A classic merge letter. Each «field» marker will be replaced by Word during merge.", NULL)
doc.addParagraph("", NULL)

doc.addMergeFieldParagraph([[:field="RecipientName"]], [:bold = true, :size = 13])
doc.addMergeFieldParagraph([[:field="CompanyName"]], NULL)
doc.addMergeFieldParagraph([[:field="Address"]], NULL)
doc.addMergeFieldParagraph([[:field="City"], ", ", [:field="Country"]], NULL)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph(["Date: ", [:field="LetterDate"]], NULL)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph(["Dear ", [:field="Salutation"], " ", [:field="LastName"], ","], NULL)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph(
    ["We are pleased to confirm your registration for the ", [:field="EventName"],
     " scheduled on ", [:field="EventDate"], " in ", [:field="EventCity"], "."],
    NULL
)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph(
    ["Your registration number is ", [:field="RegNumber"],
     " and your seat assignment is ", [:field="SeatNumber"], "."],
    NULL
)
doc.addParagraph("", NULL)
doc.addParagraph("Yours sincerely,", NULL)
doc.addParagraph("", NULL)
doc.addMergeFieldParagraph([[:field="SenderName"]], [:bold = true])
doc.addMergeFieldParagraph([[:field="SenderTitle"]], NULL)
doc.addMergeFieldParagraph([[:field="Organisation"]], NULL)

doc.addParagraph("", NULL)

# -- Example 2: Invoice template --
doc.addHeading("Example 2 — Invoice Template", 2)
doc.addParagraph("", NULL)

doc.addParagraph("INVOICE", [:bold = true, :size = 20, :align = "center"])
doc.addParagraph("", NULL)

doc.addMergeFieldParagraph(["Invoice No: ", [:field="InvoiceNumber"]], [:bold = true])
doc.addMergeFieldParagraph(["Date: ", [:field="InvoiceDate"]], NULL)
doc.addMergeFieldParagraph(["Due Date: ", [:field="DueDate"]], NULL)
doc.addParagraph("", NULL)

doc.addMergeFieldParagraph(["Bill To: ", [:field="ClientName"]], [:bold = true])
doc.addMergeFieldParagraph([[:field="ClientAddress"]], NULL)
doc.addMergeFieldParagraph([[:field="ClientCity"], " — ", [:field="ClientCountry"]], NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Description", "Qty", "Unit Price", "Total"],
        ["«ItemDescription»", "«Qty»", "«UnitPrice»", "«LineTotal»"],
        ["Subtotal", "", "", "«Subtotal»"],
        ["Tax («TaxRate»%)", "", "", "«TaxAmount»"],
        ["Total Due", "", "", "«TotalDue»"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "404040", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "F5F5F5", :colWidths = [7, 2, 3, 2.5]
    ]
)

doc.addParagraph("", NULL)
doc.addParagraph("To use this template:", [:bold = true])
doc.addParagraph("1. Open the document in Word.", NULL)
doc.addParagraph("2. Go to Mailings → Select Recipients → Use an Existing List and choose your data source (Excel or CSV).", NULL)
doc.addParagraph("3. Click Finish & Merge → Edit Individual Documents to generate one letter per record.", NULL)

# -------------------------------------------------------
# SAVE
# -------------------------------------------------------
filename = "demo_comments.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
