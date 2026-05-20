/*
	DOCXLib - Demo (Watermark2)
*/

load "docxlib.ring"

/*
    Demonstrates setImageWatermark() with:
      1. PNG image watermark – default size and opacity
      2. JPG image watermark – custom size and lower opacity
      3. BMP image watermark – small size, full opacity
      4. Image watermark combined with header text and footer
      5. Image watermark combined with a text watermark (both at once)
      6. removeImageWatermark() – disable mid-build (doc without watermark)
*/


# -- Demo 1: PNG – default size (15 × 15 cm) and opacity (50%) ---------------
doc1 = new WordWriter()
doc1.setTitle("Image Watermark – PNG Default")
doc1.setAuthor("DOCXLib Demo")

doc1.setImageWatermark("images/test1.png", NULL)

doc1.addHeading("Image Watermark – PNG (Default Settings)", 1)
doc1.addParagraph("This document uses a PNG image as a watermark.", NULL)
doc1.addParagraph("Default settings: 15 cm × 15 cm, 50 % opacity, centred on every page.", NULL)
doc1.addParagraph("The watermark is anchored inside the header so it repeats automatically on every page.", NULL)
doc1.addLineBreak()
doc1.addParagraph("Body text sits on top of the watermark because the image shape has a negative z-index (-251653120), which places it behind the document text layer.", NULL)

doc1.addPageBreak()
doc1.addHeading("Page 2 – Watermark Repeats", 2)
doc1.addParagraph("The image watermark is visible here too – it is part of the page header, so Word renders it on every page automatically.", NULL)

if doc1.save("demo_watermark2_png_default.docx")
    ? "Created: demo_watermark2_png_default.docx"
else
    ? "Error creating demo_watermark2_png_default.docx"
ok


# -- Demo 2: JPG – custom size and opacity ------------------------------------
doc2 = new WordWriter()
doc2.setTitle("Image Watermark – JPG Custom")
doc2.setAuthor("DOCXLib Demo")

doc2.setImageWatermark("images/test2.jpg", [
    :width   = 10,    # 10 cm wide
    :height  = 10,    # 10 cm tall
    :opacity = 30     # 30 % opacity – very faint
])

doc2.addHeading("Image Watermark – JPG (Custom Size & Opacity)", 1)
doc2.addParagraph("This document uses a JPG image watermark with custom settings.", NULL)
doc2.addParagraph("Settings: 10 cm × 10 cm, 30 % opacity.", NULL)
doc2.addLineBreak()

doc2.addParagraph("A lower opacity value makes the watermark less intrusive, which is useful when the document content needs to remain easy to read.", NULL)

doc2.addHeading("Table with Image Watermark", 2)
doc2.addSimpleTable([
    ["Product",    "Qty",  "Unit Price", "Total"],
    ["Widget A",   "10",   "$5.00",      "$50.00"],
    ["Widget B",   "25",   "$3.50",      "$87.50"],
    ["Widget C",   "5",    "$12.00",     "$60.00"],
    ["",           "",     "Grand Total","$197.50"]
])

if doc2.save("demo_watermark2_jpg_custom.docx")
    ? "Created: demo_watermark2_jpg_custom.docx"
else
    ? "Error creating demo_watermark2_jpg_custom.docx"
ok


# -- Demo 3: BMP – small logo stamp, full opacity -----------------------------
doc3 = new WordWriter()
doc3.setTitle("Image Watermark – BMP Small Logo")
doc3.setAuthor("DOCXLib Demo")

doc3.setImageWatermark("images/test3.bmp", [
    :width   = 6,     # 6 cm wide – smaller stamp-style
    :height  = 6,
    :opacity = 100    # full opacity – stamp effect
])

doc3.addHeading("Image Watermark – BMP (Small Stamp, Full Opacity)", 1)
doc3.addParagraph("A BMP image is used here at full opacity (100 %) with a small 6 × 6 cm size.", NULL)
doc3.addParagraph("This produces a logo-stamp effect rather than a washed-out background watermark.", NULL)

doc3.addHeading("Use Cases", 2)
doc3.addBulletList([
    "Company logo stamped on internal memos",
    "Approval stamp on signed documents",
    "Brand mark on certificates"
])

if doc3.save("demo_watermark2_bmp_stamp.docx")
    ? "Created: demo_watermark2_bmp_stamp.docx"
else
    ? "Error creating demo_watermark2_bmp_stamp.docx"
ok


# -- Demo 4: Image watermark + header text + footer + page numbers -------------
doc4 = new WordWriter()
doc4.setTitle("Image Watermark + Header/Footer")
doc4.setAuthor("DOCXLib Demo")

doc4.setHeader("Acme Corporation – Confidential")
doc4.setFooter("Internal Use Only")
doc4.showPageNumbers("center")

doc4.setImageWatermark("images/test1.png", [
    :width   = 12,
    :height  = 12,
    :opacity = 40
])

doc4.addHeading("Combined: Image Watermark + Header + Footer", 1)
doc4.addParagraph("This document demonstrates that an image watermark coexists correctly with a custom header text, footer text, and page numbers.", NULL)
doc4.addParagraph("The header text 'Acme Corporation – Confidential' appears at the top, the image watermark floats behind the page content, and the footer shows page numbers.", NULL)

doc4.addPageBreak()
doc4.addHeading("Second Page", 2)
doc4.addParagraph("All three elements – header, watermark, and footer – repeat on this page too.", NULL)

if doc4.save("demo_watermark2_combined_header_footer.docx")
    ? "Created: demo_watermark2_combined_header_footer.docx"
else
    ? "Error creating demo_watermark2_combined_header_footer.docx"
ok


# -- Demo 5: Image watermark + text watermark simultaneously ------------------
doc5 = new WordWriter()
doc5.setTitle("Image + Text Watermark Together")
doc5.setAuthor("DOCXLib Demo")

# Text watermark at 25 % opacity so the image behind it is still visible
doc5.setWatermark("DRAFT", [
    :color   = "CC0000",
    :opacity = 25,
    :rotation = -45
])

# Image watermark at 20 % – very faint so the text watermark dominates
doc5.setImageWatermark("images/test2.jpg", [
    :width   = 14,
    :height  = 14,
    :opacity = 20
])

doc5.addHeading("Image Watermark + Text Watermark (Layered)", 1)
doc5.addParagraph("Both a text watermark ('DRAFT' in red) and an image watermark (images/test2.jpg, 20 % opacity) are active on this page.", NULL)
doc5.addParagraph("The image watermark uses z-index -251653120 and is emitted before the text watermark paragraph in the header XML, so it sits furthest behind the text layer.", NULL)
doc5.addParagraph("The text watermark sits in front of the image watermark but still behind the body text.", NULL)

if doc5.save("demo_watermark2_image_and_text.docx")
    ? "Created: demo_watermark2_image_and_text.docx"
else
    ? "Error creating demo_watermark2_image_and_text.docx"
ok


# -- Demo 6: removeImageWatermark() – build with watermark then remove it ------
doc6 = new WordWriter()
doc6.setTitle("removeImageWatermark Demo")
doc6.setAuthor("DOCXLib Demo")

# Enable image watermark …
doc6.setImageWatermark("images/test1.png", [:opacity = 50])

# … then remove it before saving
doc6.removeImageWatermark()

doc6.addHeading("removeImageWatermark() Demo", 1)
doc6.addParagraph("setImageWatermark() was called with images/test1.png, but removeImageWatermark() was called immediately after.", NULL)
doc6.addParagraph("The saved document should have NO watermark – a clean page.", NULL)

if doc6.save("demo_watermark2_removed.docx")
    ? "Created: demo_watermark2_removed.docx"
else
    ? "Error creating demo_watermark2_removed.docx"
ok


? ""
? "All image watermark demos created successfully."
