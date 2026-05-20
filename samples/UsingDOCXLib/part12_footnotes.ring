/*
	DOCXLib - Demo (Footnotes)
*/

load "docxlib.ring"

/*
    Demonstrates:
      1. Superscript and subscript
      2. Line spacing (single, 1.5x, double, exact)
      3. Nested / multi-level bullet and numbered lists
      4. Footnotes and endnotes
*/

# ============================================================================
# Demo 1: Superscript & Subscript
# ============================================================================
doc1 = new WordWriter()
doc1.setTitle("Demo 1 – Superscript & Subscript")
doc1.setAuthor("DOCXLib Demo")

doc1.addHeading("Superscript & Subscript Demo", 1)
doc1.addParagraph("Superscript and subscript are enabled via run options in any paragraph or rich paragraph.", NULL)

doc1.addHeading("Scientific Formulas", 2)

# Chemical formula using subscript
doc1.addRichParagraph([
    ["Water:  H", NULL],
    ["2", [:subscript = true]],
    ["O", NULL]
], NULL)

doc1.addRichParagraph([
    ["Carbon Dioxide:  CO", NULL],
    ["2", [:subscript = true]]
], NULL)

doc1.addRichParagraph([
    ["Sulfuric Acid:  H", NULL],
    ["2", [:subscript = true]],
    ["SO", NULL],
    ["4", [:subscript = true]]
], NULL)

doc1.addHeading("Mathematical Expressions", 2)

# Powers using superscript
doc1.addRichParagraph([
    ["Area of a circle:  A = ", [:bold = true]],
    ["π", NULL],
    ["r", NULL],
    ["2", [:superscript = true]]
], NULL)

doc1.addRichParagraph([
    ["Einstein's equation:  E = mc", [:bold = true]],
    ["2", [:superscript = true]]
], NULL)

doc1.addRichParagraph([
    ["Pythagorean theorem:  a", NULL],
    ["2", [:superscript = true]],
    [" + b", NULL],
    ["2", [:superscript = true]],
    [" = c", NULL],
    ["2", [:superscript = true]]
], NULL)

doc1.addHeading("Ordinal Numbers", 2)

doc1.addRichParagraph([
    ["1", NULL],
    ["st", [:superscript = true]],
    [" place,  2", NULL],
    ["nd", [:superscript = true]],
    [" place,  3", NULL],
    ["rd", [:superscript = true]],
    [" place", NULL]
], NULL)

doc1.addHeading("Mixed Styling", 2)

doc1.addRichParagraph([
    ["Bold", [:bold = true]],
    [" normal ", NULL],
    ["italic", [:italic = true]],
    [" normal X", NULL],
    ["n+1", [:superscript = true, :color = "red"]],
    [" and log", NULL],
    ["10", [:subscript = true]],
    ["(x)", NULL]
], NULL)

if doc1.save("demo_superscript_subscript.docx")
    ? "Created: demo_superscript_subscript.docx"
else
    ? "Error creating demo_superscript_subscript.docx"
ok

# ============================================================================
# Demo 2: Line Spacing
# ============================================================================
doc2 = new WordWriter()
doc2.setTitle("Demo 2 – Line Spacing")
doc2.setAuthor("DOCXLib Demo")

doc2.addHeading("Line Spacing Demo", 1)
doc2.addParagraph("Line spacing is controlled via :lineSpacing (multiplier) or :lineSpacingPt (exact points), with an optional :lineRule key.", NULL)

sampleText = "This paragraph demonstrates the configured line spacing. " +
             "Reading it carefully, you should notice the vertical gap between lines " +
             "is different from a standard single-spaced paragraph."

doc2.addHeading("Single Spacing  (1.0×)", 2)
doc2.addParagraph(sampleText, [:lineSpacing = 1.0])

doc2.addHeading("One-and-a-Half Spacing  (1.5×)", 2)
doc2.addParagraph(sampleText, [:lineSpacing = 1.5])

doc2.addHeading("Double Spacing  (2.0×)", 2)
doc2.addParagraph(sampleText, [:lineSpacing = 2.0])

doc2.addHeading("Exact 18pt Spacing", 2)
doc2.addParagraph(sampleText, [:lineSpacingPt = 18, :lineRule = "exact"])

doc2.addHeading("At-Least 20pt Spacing", 2)
doc2.addParagraph(sampleText, [:lineSpacingPt = 20, :lineRule = "atLeast"])

doc2.addHeading("Combined: double-spaced, centered, blue", 2)
doc2.addParagraph("This paragraph is double-spaced, centred, and coloured blue.",
    [:lineSpacing = 2.0, :align = "center", :color = "0000FF", :italic = true])

doc2.addHeading("Space Before & After + Line Spacing", 2)
doc2.addParagraph("Paragraph A – 240 twips before, 120 after, 1.5× line spacing.",
    [:spaceBefore = 240, :spaceAfter = 120, :lineSpacing = 1.5])
doc2.addParagraph("Paragraph B – same settings, demonstrating the gap above and below.",
    [:spaceBefore = 240, :spaceAfter = 120, :lineSpacing = 1.5])

if doc2.save("demo_line_spacing.docx")
    ? "Created: demo_line_spacing.docx"
else
    ? "Error creating demo_line_spacing.docx"
ok

# ============================================================================
# Demo 3: Nested / Multi-level Lists
# ============================================================================
doc3 = new WordWriter()
doc3.setTitle("Demo 3 – Nested Lists")
doc3.setAuthor("DOCXLib Demo")

doc3.addHeading("Nested Lists Demo", 1)
doc3.addParagraph("Multi-level lists use [text, level] pairs. Level 0 is the top, higher numbers indent further.", NULL)

doc3.addHeading("Nested Bullet List", 2)
doc3.addNestedBulletList([
    ["Fruits", 0],
    ["Tropical fruits", 1],
    ["Mango", 2],
    ["Papaya", 2],
    ["Dragonfruit", 2],
    ["Citrus fruits", 1],
    ["Orange", 2],
    ["Lemon", 2],
    ["Grapefruit", 2],
    ["Berries", 1],
    ["Strawberry", 2],
    ["Blueberry", 2],
    ["Vegetables", 0],
    ["Root vegetables", 1],
    ["Carrot", 2],
    ["Potato", 2],
    ["Leafy greens", 1],
    ["Spinach", 2],
    ["Kale", 2]
])

doc3.addPageBreak()

doc3.addHeading("Nested Numbered List", 2)
doc3.addNestedNumberedList([
    ["Project Phases", 0],
    ["Planning", 1],
    ["Define requirements", 2],
    ["Create project plan", 2],
    ["Assign resources", 2],
    ["Development", 1],
    ["Backend implementation", 2],
    ["Frontend implementation", 2],
    ["Database setup", 2],
    ["Testing", 1],
    ["Unit tests", 2],
    ["Integration tests", 2],
    ["User acceptance testing", 2],
    ["Deployment", 0],
    ["Staging environment", 1],
    ["Deploy to staging", 2],
    ["Smoke testing", 2],
    ["Production environment", 1],
    ["Final deployment", 2],
    ["Post-deployment monitoring", 2]
])

doc3.addHeading("Three-Level Mixed Nesting", 2)
doc3.addNestedBulletList([
    ["Level 0 – Top", 0],
    ["Level 1 – Sub", 1],
    ["Level 2 – Sub-sub", 2],
    ["Level 2 – another sub-sub", 2],
    ["Level 1 – another sub", 1],
    ["Level 0 – another top", 0]
])

if doc3.save("demo_nested_lists.docx")
    ? "Created: demo_nested_lists.docx"
else
    ? "Error creating demo_nested_lists.docx"
ok

# ============================================================================
# Demo 4: Footnotes & Endnotes
# ============================================================================
doc4 = new WordWriter()
doc4.setTitle("Demo 4 – Footnotes & Endnotes")
doc4.setAuthor("DOCXLib Demo")
doc4.showPageNumbers("center")

doc4.addHeading("Footnotes & Endnotes Demo", 1)
doc4.addParagraph("This document demonstrates footnotes (references at the bottom of each page) and endnotes (collected at the end of the document).", NULL)

doc4.addHeading("Basic Footnotes", 2)

doc4.addFootnote(
    "The Ring programming language was created by Mahmoud Fayed.",
    "Ring is a simple, yet powerful language for building applications. See https://ring-lang.net for details.",
    NULL
)

doc4.addFootnote(
    'Word documents use the Open XML format (OOXML), standardised as ISO/IEC 29500.',
    "ISO/IEC 29500:2016 – Information technology – Document description and processing languages – Office Open XML File Formats.",
    NULL
)

doc4.addFootnote(
    "Footnote markers appear as superscript numbers in the body text.",
    "The marker is placed immediately after the text it refers to, using the FootnoteReference character style.",
    NULL
)

doc4.addHeading("Footnotes with Formatting", 2)

doc4.addFootnote(
    "This sentence is bold and has a footnote attached.",
    "The bold formatting applies only to the preceding text run; the footnote marker inherits the FootnoteReference style.",
    [:bold = true]
)

doc4.addFootnote(
    "Centered paragraph with footnote.",
    "Alignment is a paragraph-level property and does not affect the footnote marker position.",
    [:align = "center"]
)

doc4.addHeading("Rich Paragraph with Multiple Footnotes", 2)
doc4.addParagraph("You can also place footnote references anywhere within a rich paragraph by registering footnotes first, then embedding the reference runs:", NULL)

fnA = doc4.registerFootnote("First footnote registered manually – appears at position A in the paragraph.")
fnB = doc4.registerFootnote("Second footnote registered manually – appears at position B in the paragraph.")
fnC = doc4.registerFootnote("Third footnote registered manually – appears at position C in the paragraph.")

doc4.addRichParagraph([
    ["Beginning of sentence", NULL],
    ["", [:footnoteId = fnA]],
    [", a middle part", NULL],
    ["", [:footnoteId = fnB]],
    [", and the end of the sentence", NULL],
    ["", [:footnoteId = fnC]],
    [".", NULL]
], NULL)

doc4.addPageBreak()

doc4.addHeading("Endnotes", 2)
doc4.addParagraph("Endnotes work the same way but their content appears at the very end of the document rather than at the bottom of each page.", NULL)

doc4.addEndnote(
    "The word 'docx' is short for Document XML.",
    "A .docx file is a ZIP archive containing a set of XML files that together describe the document structure, styles, and content.",
    NULL
)

doc4.addEndnote(
    "Endnotes are numbered independently from footnotes.",
    "In standard Word documents, footnotes and endnotes each maintain their own sequential counter starting at 1.",
    NULL
)

doc4.addHeading("Mixed Footnotes and Endnotes", 2)

en1 = doc4.registerEndnote("This is an endnote placed inside a rich paragraph using registerEndnote().")
fn4 = doc4.registerFootnote("This is a footnote placed alongside the endnote in the same paragraph.")

doc4.addRichParagraph([
    ["This sentence has both a footnote", NULL],
    ["", [:footnoteId = fn4]],
    [" and an endnote", NULL],
    ["", [:endnoteId = en1]],
    [" reference side by side.", NULL]
], NULL)

doc4.addHeading("Superscript Combined with Footnote", 2)

fn5 = doc4.registerFootnote("The formula E = mc² was published by Albert Einstein in 1905 in his paper on special relativity.")

doc4.addRichParagraph([
    ["The famous equation E = mc", NULL],
    ["2", [:superscript = true]],
    ["", [:footnoteId = fn5]],
    [" changed our understanding of energy and mass.", NULL]
], NULL)

if doc4.save("demo_footnotes_endnotes.docx")
    ? "Created: demo_footnotes_endnotes.docx"
else
    ? "Error creating demo_footnotes_endnotes.docx"
ok

# ============================================================================
# Demo 5: All four features combined in one document
# ============================================================================
doc5 = new WordWriter()
doc5.setTitle("Research Paper Sample")
doc5.setAuthor("DOCXLib Demo")
doc5.showPageNumbers("center")
doc5.setWatermarkText("DRAFT")

doc5.addHeading("A Brief Study of the Ring Language", 1)
doc5.addParagraph("Department of Computer Science – Technical Report 2025", [:align = "center", :italic = true])
doc5.addLineBreak()

doc5.addHeading("Abstract", 2)
doc5.addParagraph(
    "Ring is an open-source, multi-paradigm programming language that supports functional, " +
    "object-oriented, and procedural styles. This report surveys its features and examines " +
    "document-generation capabilities through the DOCXLib library.",
    [:lineSpacing = 1.5, :align = "both"]
)

doc5.addHeading("1. Introduction", 2)

fnIntro = doc5.registerFootnote("Ring was first publicly released in 2016. See ring-lang.net for the official documentation.")

doc5.addRichParagraph([
    ["The Ring programming language", [:bold = true]],
    ["", [:footnoteId = fnIntro]],
    [" was designed with simplicity in mind. Its syntax is straightforward, making it " +
     "accessible to beginners while retaining the power required by expert developers.", NULL]
], [:lineSpacing = 1.5, :align = "both"])

doc5.addHeading("2. Key Formula Support", 2)

fnFormula = doc5.registerFootnote("OOXML specifies superscript/subscript via the w:vertAlign element in run properties.")

doc5.addRichParagraph([
    ["The library correctly renders expressions such as H", NULL],
    ["2", [:subscript = true]],
    ["O, E = mc", NULL],
    ["2", [:superscript = true]],
    ["", [:footnoteId = fnFormula]],
    [", and x", NULL],
    ["n+1", [:superscript = true]],
    [" inside any Word paragraph.", NULL]
], [:lineSpacing = 1.5])

doc5.addHeading("3. Document Structure", 2)
doc5.addParagraph("A well-structured technical document typically follows the outline below:", [:lineSpacing = 1.5])

doc5.addNestedNumberedList([
    ["Front Matter", 0],
    ["Title page", 1],
    ["Abstract", 1],
    ["Table of contents", 1],
    ["Body", 0],
    ["Introduction", 1],
    ["Background", 1],
    ["Methodology", 1],
    ["Results", 1],
    ["Conclusion", 1],
    ["Back Matter", 0],
    ["References", 1],
    ["Appendices", 1],
    ["Index", 1]
])

doc5.addHeading("4. Conclusion", 2)

enConclusion = doc5.registerEndnote("Full source code for DOCXLib is available under the MIT license.")

doc5.addRichParagraph([
    ["DOCXLib now covers the most important Word document features required by " +
     "real-world applications, making it a solid foundation for document automation in Ring.", NULL],
    ["", [:endnoteId = enConclusion]]
], [:lineSpacing = 1.5, :align = "both"])

if doc5.save("demo_combined_research_paper.docx")
    ? "Created: demo_combined_research_paper.docx"
else
    ? "Error creating demo_combined_research_paper.docx"
ok

? ""
? "All demos created successfully."
? ""
