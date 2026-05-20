/*
	DOCXLib - Demo (BordersAndRTL)
*/

load "docxlib.ring"

/*
    Demonstrates:
      1. setSimplePageBorder()
      2. setPageBorder() with custom style, color, size
      3. Decorative / artistic border styles
      4. Per-side borders (top + bottom only, left only, ...)
      5. removePageBorder()
      6. RTL paragraph-level (mixed LTR/RTL in one document)
      7. RTL run-level (inline bidi with Arabic/Hebrew font)
      8. setDocumentRTL() – full document default RTL
      9. RTL + page border combined
*/


# ============================================================================
# Demo 1 – Simple page border (all four sides, single 3pt black line)
# ============================================================================
doc1 = new WordWriter()
doc1.setTitle("Page Borders – Simple")
doc1.setAuthor("RingWordLib Demo")

doc1.setSimplePageBorder()

doc1.addHeading("Simple Page Border", 1)
doc1.addParagraph("This document has a thin single-line black border on all four sides, applied using setSimplePageBorder() – the zero-configuration shortcut.", NULL)
doc1.addLineBreak()
doc1.addParagraph("The border is declared inside <w:sectPr> / <w:pgBorders> with offsetFrom='page', which means it is measured from the page edge rather than the text area.", NULL)

doc1.addHeading("Usage", 2)
doc1.addBulletList([
    "doc.setSimplePageBorder()   – one-line setup",
    "doc.removePageBorder()      – remove at any time",
    "doc.setPageBorder([...])    – full control over all options"
])

if doc1.save("demo_bordersandrtl_1_simple.docx")
    ? "Created: demo_bordersandrtl_1_simple.docx"
else
    ? "Error: demo_bordersandrtl_1_simple.docx"
ok

# ============================================================================
# Demo 2 – Custom border: thick, colored, further from edge
# ============================================================================
doc2 = new WordWriter()
doc2.setTitle("Page Borders – Custom Thick Color")
doc2.setAuthor("RingWordLib Demo")

doc2.setPageBorder([
    :style = "single",
    :color = "2E74B5",    # Blue
    :size  = 6,           # 6pt thick
    :space = 36           # 36pt from page edge
])

doc2.addHeading("Custom Page Border – Thick Blue", 1)
doc2.addParagraph("A 6-point thick blue border (color #2E74B5) set 36 points from the page edge.", NULL)
doc2.addParagraph("The :size option accepts points and is converted internally to OOXML eighths-of-a-point (1 pt = 8 units).", NULL)
doc2.addLineBreak()
doc2.addSimpleTable([
    ["Option",     "Value",   "Effect"],
    [":style",     "single",  "Solid line border"],
    [":color",     "2E74B5",  "Microsoft blue"],
    [":size",      "6",       "6pt = 48 eighths"],
    [":space",     "36",      "36pt from page edge"]
])

if doc2.save("demo_bordersandrtl_2_custom_color.docx")
    ? "Created: demo_bordersandrtl_2_custom_color.docx"
else
    ? "Error: demo_bordersandrtl_2_custom_color.docx"
ok

# ============================================================================
# Demo 3 – Decorative border styles
# ============================================================================
doc3 = new WordWriter()
doc3.setTitle("Page Borders – Decorative Styles")
doc3.setAuthor("RingWordLib Demo")

doc3.setPageBorder([
    :style = "wave",
    :color = "70AD47",   # Green
    :size  = 6,
    :space = 20
])

doc3.addHeading("Decorative Border – Wave", 1)
doc3.addParagraph("This document uses the 'wave' border style in green. Many decorative styles are available in the OOXML specification.", NULL)
doc3.addLineBreak()

doc3.addParagraph("Style name reference:", NULL)
doc3.addBulletList([
    "single           – standard thin line",
    "double           – two parallel lines",
    "triple           – three parallel lines",
    "thick            – single thick line",
    "dashed           – dashed line",
    "dotted           – dotted line",
    "wave             – wavy line",
    "doubleWave       – double wavy line",
    "thinThickSmallGap – thin outer, thick inner",
    "thickThinSmallGap – thick outer, thin inner",
    "threeDEmboss     – 3-D embossed effect",
    "threeDEngrave    – 3-D engraved effect",
    "inset            – inset shadow",
    "outset           – outset shadow"
])

if doc3.save("demo_bordersandrtl_3_wave.docx")
    ? "Created: demo_bordersandrtl_3_wave.docx"
else
    ? "Error: demo_bordersandrtl_3_wave.docx"
ok

# ============================================================================
# Demo 4 – Per-side borders (top and bottom only)
# ============================================================================
doc4 = new WordWriter()
doc4.setTitle("Page Borders – Top and Bottom Only")
doc4.setAuthor("RingWordLib Demo")

doc4.setPageBorder([
    :style = "double",
    :color = "C00000",   # Dark red
    :size  = 4,
    :space = 24,
    :sides = ["top", "bottom"]
])

doc4.addHeading("Per-Side Borders – Top & Bottom Only", 1)
doc4.addParagraph("Only the top and bottom edges have a double red border. The left and right edges are plain.", NULL)
doc4.addParagraph("Use the :sides list to specify any combination of 'top', 'left', 'bottom', 'right'.", NULL)
doc4.addLineBreak()

doc4.addHeading("Other useful combinations", 2)
doc4.addBulletList([
    '["top"]                   – header underline effect',
    '["top", "bottom"]         – newspaper-style rules',
    '["left"]                  – sidebar accent',
    '["top", "left", "right"]  – three-sided frame'
])

if doc4.save("demo_bordersandrtl_4_top_bottom.docx")
    ? "Created: demo_bordersandrtl_4_top_bottom.docx"
else
    ? "Error: demo_bordersandrtl_4_top_bottom.docx"
ok

# ============================================================================
# Demo 5 – removePageBorder (set then remove before saving)
# ============================================================================
doc5 = new WordWriter()
doc5.setTitle("Page Borders – Remove Demo")
doc5.setAuthor("RingWordLib Demo")

doc5.setSimplePageBorder()    # enabled ...
doc5.removePageBorder()       # ... then removed

doc5.addHeading("removePageBorder() Demo", 1)
doc5.addParagraph("setSimplePageBorder() was called, then removePageBorder() was called immediately after. The saved file has no page border.", NULL)

if doc5.save("demo_bordersandrtl_5_removed.docx")
    ? "Created: demo_bordersandrtl_5_removed.docx"
else
    ? "Error: demo_bordersandrtl_5_removed.docx"
ok

# ============================================================================
# Demo 6 – RTL paragraph-level (mixed LTR + RTL in one document)
# ============================================================================
doc6 = new WordWriter()
doc6.setTitle("RTL Demo – Paragraph Level")
doc6.setAuthor("RingWordLib Demo")

doc6.addHeading("Paragraph-Level RTL / BiDi", 1)
doc6.addParagraph("This document is LTR by default. Individual paragraphs can be switched to RTL using the :rtl option.", NULL)

doc6.addHeading("Normal LTR Paragraphs", 2)
doc6.addParagraph("This is a normal left-to-right paragraph in English.", NULL)
doc6.addParagraph("Second LTR paragraph. Text flows from left to right.", NULL)

doc6.addHeading("RTL Paragraphs (Arabic placeholder text)", 2)

# RTL paragraphs – using the convenience wrapper
doc6.addRTLParagraph("هذه فقرة باللغة العربية تسير من اليمين إلى اليسار.", NULL)
doc6.addRTLParagraph("البرمجة بلغة رينج سهلة وممتعة.", NULL)

# RTL paragraph – manually via addParagraph with options
doc6.addParagraph("النص العربي يُظهر الاتجاه الصحيح من اليمين إلى اليسار.",
    [:rtl = true, :align = "right"])

doc6.addHeading("Hebrew placeholder text", 2)
doc6.addRTLParagraph("זהו טקסט לדוגמה בעברית הנכתב מימין לשמאל.", NULL)

doc6.addHeading("Mixed BiDi in a Rich Paragraph", 2)
doc6.addParagraph("A rich paragraph can mix LTR and RTL runs. The Unicode Bidirectional Algorithm handles display direction automatically, but explicit run-level :rtl marks help Word render it correctly.", NULL)

# Rich paragraph with one LTR run followed by one RTL run
doc6.addRichParagraph([
    ["English text on the left  |  ", [:bold = true]],
    ["النص العربي على اليمين", [:rtl = true, :csFont = "Arial"]]
], [:rtl = true])

doc6.addHeading(":ltr Option – Force LTR Inside RTL Context", 2)
doc6.addParagraph("When the document or paragraph is RTL, use :ltr=true to force a specific paragraph back to left-to-right:", NULL)
doc6.addParagraph("This paragraph is explicitly LTR even if the document defaults were changed.", [:ltr = true, :align = "left"])

if doc6.save("demo_bordersandrtl_6_paragraph_level.docx")
    ? "Created: demo_bordersandrtl_6_paragraph_level.docx"
else
    ? "Error: demo_bordersandrtl_6_paragraph_level.docx"
ok

# ============================================================================
# Demo 7 – RTL run-level with complex-script font (:csFont)
# ============================================================================
doc7 = new WordWriter()
doc7.setTitle("RTL Demo – Run Level & csFont")
doc7.setAuthor("RingWordLib Demo")

doc7.addHeading("Run-Level RTL and Complex-Script Font", 1)
doc7.addParagraph("The :rtl and :csFont options work at the individual run level inside addRichParagraph, giving full control over mixed-script inline text.", NULL)

doc7.addHeading("Inline Arabic within English Sentence", 2)
doc7.addRichParagraph([
    ["The Arabic word for 'hello' is ", NULL],
    ["مرحبا", [:rtl = true, :csFont = "Arial", :bold = true]],
    [" and for 'thank you' is ", NULL],
    ["شكراً", [:rtl = true, :csFont = "Arial", :bold = true]],
    [".", NULL]
], NULL)

doc7.addHeading("Inline Hebrew within English Sentence", 2)
doc7.addRichParagraph([
    ["The Hebrew word for 'peace' is ", NULL],
    ["שָׁלוֹם", [:rtl = true, :csFont = "David", :bold = true, :color = "2E74B5"]],
    [" (shalom).", NULL]
], NULL)

doc7.addHeading("csFont without rtl – Complex-Script Font Only", 2)
doc7.addParagraph("Use :csFont alone (without :rtl) to set the complex-script font for a run that contains both LTR text and characters that need a specific CS font rendering.", NULL)
doc7.addRichParagraph([
    ["Normal font run then ", NULL],
    ["مرحبا (Arabic) rendered with Times New Roman CS", [:csFont = "Times New Roman"]]
], NULL)

doc7.addHeading("Combining :font and :csFont", 2)
doc7.addRichParagraph([
    ["You can specify both the LTR font and CS font in the same run: ", NULL],
    ["Hello مرحبا", [:font = "Georgia", :csFont = "Arial", :rtl = false]]
], NULL)

if doc7.save("demo_bordersandrtl_7_run_level.docx")
    ? "Created: demo_bordersandrtl_7_run_level.docx"
else
    ? "Error: demo_bordersandrtl_7_run_level.docx"
ok

# ============================================================================
# Demo 8 – setDocumentRTL() – full document RTL default
#           Every paragraph, heading, and list item is automatically RTL.
#           No :rtl=true needed on individual paragraphs.
# ============================================================================
doc8 = new WordWriter()
doc8.setTitle("مستند عربي بالكامل")
doc8.setAuthor("RingWordLib Demo")
doc8.showPageNumbers("right")
doc8.setPageBorder([
    :style = "single",
    :color = "1F4E79",
    :size  = 3,
    :space = 24
])

# One call flips the entire document to RTL
doc8.setDocumentRTL(true)

doc8.addHeading("مستند عربي بالكامل", 1)

doc8.addParagraph("هذا المستند تم تعيينه ليكون من اليمين إلى اليسار على مستوى الوثيقة بأكملها باستخدام setDocumentRTL(true). لاحظ أن جميع الفقرات تبدأ من الجانب الأيمن للصفحة.", NULL)

doc8.addParagraph("لا حاجة إلى إضافة الخيار :rtl = true في كل فقرة على حدة — فور تفعيل setDocumentRTL يصبح الاتجاه الافتراضي من اليمين إلى اليسار لجميع المحتوى.", NULL)

doc8.addHeading("قائمة بالنقاط", 2)
doc8.addBulletList([
    "العنصر الأول في القائمة",
    "العنصر الثاني في القائمة",
    "العنصر الثالث في القائمة"
])

doc8.addHeading("قائمة مرقمة", 2)
doc8.addNumberedList([
    "الخطوة الأولى",
    "الخطوة الثانية",
    "الخطوة الثالثة"
])

doc8.addHeading("فقرة محاذاة للوسط", 2)
doc8.addParagraph("هذه الفقرة محاذاة للوسط داخل مستند RTL.", [:align = "center"])

doc8.addHeading("إجبار فقرة على الاتجاه LTR داخل مستند RTL", 2)
doc8.addParagraph("This paragraph is forced LTR using :ltr=true — even inside a full RTL document.",
    [:ltr = true])

doc8.addHeading("نص مختلط (عربي + إنجليزي)", 2)
doc8.addRichParagraph([
    ["هذه الجملة تحتوي على ", NULL],
    ["English words", [:ltr = true, :bold = true]],
    [" داخل نص عربي.", NULL]
], NULL)

if doc8.save("demo_bordersandrtl_8_document_level.docx")
    ? "Created: demo_bordersandrtl_8_document_level.docx"
else
    ? "Error: demo_bordersandrtl_8_document_level.docx"
ok

# ============================================================================
# Demo 9 – RTL + page border combined
# ============================================================================
doc9 = new WordWriter()
doc9.setTitle("RTL + Page Border Combined")
doc9.setAuthor("RingWordLib Demo")

doc9.setDocumentRTL(true)
doc9.setPageBorder([
    :style = "double",
    :color = "1F5C2E",   # Dark green
    :size  = 4,
    :space = 20
])

doc9.addHeading("مستند عربي مع إطار الصفحة", 1)
doc9.addParagraph("هذا المستند يجمع بين الاتجاه من اليمين إلى اليسار وإطار الصفحة المزدوج باللون الأخضر الداكن.", NULL)
doc9.addParagraph("يمكن الجمع بين أي ميزات المكتبة بحرية تامة: الاتجاه، الإطارات، العلامات المائية، الهوامش المخصصة وغيرها.", NULL)

doc9.addHeading("English section inside RTL doc", 2)
doc9.addParagraph("This English paragraph is forced LTR using :ltr=true.",
    [:ltr = true, :align = "left"])
doc9.addRichParagraph([
    ["Mixed: English ", [:ltr = true]],
    ["عربي", [:rtl = true, :csFont = "Arial", :bold = true]],
    [" English again", [:ltr = true]]
], [:ltr = true])

if doc9.save("demo_bordersandrtl_9_combined_border.docx")
    ? "Created: demo_bordersandrtl_9_combined_border.docx"
else
    ? "Error: demo_bordersandrtl_9_combined_border.docx"
ok

# ============================================================================
# Summary
# ============================================================================
? ""
? "All demos created successfully."
? ""
? "API Quick Reference – new in v1.5.0:"
? ""
? "Page Borders:"
? "  doc.setSimplePageBorder()              # thin single black border, all sides"
? "  doc.setPageBorder(["
? "      :style = 'wave',   # border style name"
? "      :color = '2E74B5', # color name or hex"
? "      :size  = 4,        # thickness in points"
? "      :space = 24,       # distance from page edge in points"
? "      :sides = ['top','bottom']  # omit for all four sides"
? "  ])"
? "  doc.removePageBorder()"
? ""
? "RTL / BiDi:"
? "  doc.setDocumentRTL(true)               # full document RTL default"
? "  doc.addRTLParagraph(text, options)     # convenience RTL paragraph"
? "  doc.addParagraph(text, [:rtl=true])    # paragraph-level RTL"
? "  doc.addParagraph(text, [:ltr=true])    # force LTR inside RTL document"
? "  Rich paragraph run options:"
? "    [:rtl = true]                        # RTL run direction mark"
? "    [:csFont = 'Arial']                  # complex-script font for RTL text"
? "    [:font = 'X', :csFont = 'Y']         # both LTR and CS font"
