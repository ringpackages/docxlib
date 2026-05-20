/*
	DOCXLib - Demo (themes)
*/

load "docxlib.ring"

/*
    Document Themes
    8 built-in colour themes affecting headings and theme1.xml
*/

doc = new WordWriter()
doc.setTitle("Document Themes Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")

# Apply the Blue theme — headings and theme1.xml use this palette
doc.setTheme("Blue")

doc.setHeader("DOCXLib — Document Themes")
doc.showPageNumbers("right")

doc.addHeading("Feature 3 — Document Themes", 1)
doc.addParagraph("setTheme() applies a named colour palette to the document. Heading colours change automatically. A word/theme/theme1.xml file is written into the DOCX package.", NULL)
doc.addParagraph("", NULL)

doc.addHeading("Heading 1 — accent1 colour", 1)
doc.addHeading("Heading 2 — accent1 colour", 2)
doc.addHeading("Heading 3 — accent2 colour", 3)
doc.addHeading("Heading 4 — accent2 colour", 4)
doc.addHeading("Heading 5 — dk2 colour", 5)
doc.addHeading("Heading 6 — dk2 colour", 6)

doc.addParagraph("", NULL)
doc.addParagraph("This document uses the Blue theme. All six heading levels above pick their colour automatically from the theme palette.", NULL)
doc.addParagraph("", NULL)

doc.addTable(
    [
        ["Theme",   "Accent 1",        "Accent 2",       "DK2"],
        ["Office",  "#4472C4 Blue",     "#ED7D31 Orange", "#44546A Slate"],
        ["Blue",    "#1E6BB5 DeepBlue", "#2E75B6 MidBlue","#1F3864 Navy"],
        ["Dark",    "#88C0D0 IceBlue",  "#81A1C1 Steel",  "#C8D0D8 Silver"],
        ["Green",   "#375623 Forest",   "#70AD47 Lime",   "#375623 Forest"],
        ["Red",     "#C00000 DarkRed",  "#FF0000 Red",    "#7B0000 Maroon"],
        ["Purple",  "#7030A0 Purple",   "#9B59B6 Violet", "#3B0764 DeepPurple"],
        ["Teal",    "#00796B Teal",     "#009688 MedTeal","#004D40 DarkTeal"],
        ["Warm",    "#E65100 DeepOrange","#F57C00 Orange","#6D3B00 Brown"]
    ],
    [
        :headerRow = true, :borderStyle = "single",
        :headerBgColor = "1E6BB5", :headerTextColor = "FFFFFF",
        :evenRowBgColor = "DEE9F7", :colWidths = [2.5, 4.5, 4.5, 4.5]
    ]
)

filename = "demo_themes.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
