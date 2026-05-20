/*
	DOCXLib - Demo (Watermark)
*/

load "docxlib.ring"

/*
    Demonstrates text watermark support:
      1. Simple "DRAFT" watermark with default styling
      2. Custom text, color, opacity, rotation and font
      3. Watermark combined with header/footer text
*/


# -- Demo 1: Default "DRAFT" watermark ---------------------------------------
doc1 = new WordWriter()

doc1.setTitle("Watermark Demo 1 – Default DRAFT")
doc1.setWatermarkText("DRAFT")

doc1.addHeading("Invoice #1042", 1)
doc1.addParagraph("Date: 2026-05-01", NULL)
doc1.addParagraph("Bill To: Acme Corporation", NULL)
doc1.addParagraph("Amount Due: $4,500.00", NULL)
doc1.addParagraph("This document is still a draft and has not been finalised.", NULL)

if doc1.save("demo_watermark_draft.docx")
    ? "Created: demo_watermark_draft.docx"
else
    ? "Error creating demo_watermark_draft.docx"
ok

# -- Demo 2: Custom watermark (CONFIDENTIAL, red, 40% opacity) ---------------
doc2 = new WordWriter()

doc2.setTitle("Watermark Demo 2 – Confidential")
doc2.setWatermark("CONFIDENTIAL", [
    :color    = "FF4444",
    :opacity  = 40,
    :rotation = -45,
    :font     = "Arial",
    :size     = 80
])

doc2.addHeading("Q3 Financial Report", 1)
doc2.addParagraph("Revenue: $12,300,000", NULL)
doc2.addParagraph("Expenses: $8,750,000", NULL)
doc2.addParagraph("Net Profit: $3,550,000", NULL)
doc2.addParagraph("This report contains sensitive financial information.", NULL)

if doc2.save("demo_watermark_confidential.docx")
    ? "Created: demo_watermark_confidential.docx"
else
    ? "Error creating demo_watermark_confidential.docx"
ok

# -- Demo 3: Watermark + header text + footer + page numbers -----------------
doc3 = new WordWriter()

doc3.setTitle("Watermark Demo 3 – Combined")
doc3.setHeader("Acme Corp – Internal Use Only")
doc3.setFooter("Strictly Confidential")
doc3.showPageNumbers("center")
doc3.setWatermark("INTERNAL", [
    :color   = "AAAAAA",
    :opacity = 35,
    :font    = "Calibri"
])

doc3.addHeading("Project Alpha – Status Update", 1)
doc3.addParagraph("This memo combines a watermark with a custom header and footer.", NULL)
doc3.addParagraph("The watermark appears on every page alongside the regular header.", NULL)
doc3.addBulletList(["Milestone 1 – Complete", "Milestone 2 – In Progress", "Milestone 3 – Pending"])
doc3.addPageBreak()
doc3.addHeading("Appendix", 2)
doc3.addParagraph("Supporting data on page 2 also carries the watermark.", NULL)

if doc3.save("demo_watermark_combined.docx")
    ? "Created: demo_watermark_combined.docx"
else
    ? "Error creating demo_watermark_combined.docx"
ok

# -- Demo 4: Diagonal variation – 0° (horizontal watermark) ------------------
doc4 = new WordWriter()

doc4.setTitle("Watermark Demo 4 – Horizontal")
doc4.setWatermark("DO NOT COPY", [
    :color    = "CCCCCC",
    :opacity  = 60,
    :rotation = 0,
    :font     = "Times New Roman",
    :size     = 60
])

doc4.addHeading("Certificate of Completion", 1)
doc4.addParagraph("Awarded to: John Smith", [:bold = true])
doc4.addParagraph("Course: Advanced Ring Programming", NULL)
doc4.addParagraph("Date: May 2026", NULL)

if doc4.save("demo_watermark_horizontal.docx")
    ? "Created: demo_watermark_horizontal.docx"
else
    ? "Error creating demo_watermark_horizontal.docx"
ok

? ""
? "All watermark demos created successfully."
