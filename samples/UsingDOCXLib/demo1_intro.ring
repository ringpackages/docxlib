/*
    DOCXLib Demo - Part 1
*/

load "docxlib.ring"

? "=============================================="
? "   DOCXLib Demo - Word Creation in Ring"
? "=============================================="
? ""

# Demo 1: Simple Document
? "Demo 1: Creating a simple document..."
doc = new WordWriter()
doc.setTitle("Simple Document")
doc.setAuthor("Ring Programmer")

doc.addHeading("Welcome to DOCXLib", 1)
doc.addParagraph("This is a simple paragraph created with DOCXLib.", NULL)
doc.addParagraph("You can easily create Word documents using Ring!", NULL)

if doc.save("demo1_simple.docx")
    ? "  Created: demo1_simple.docx"
else
    ? "  FAILED: demo1_simple.docx"
ok

# Demo 2: Text Formatting
? "Demo 2: Creating document with text formatting..."
doc = new WordWriter()
doc.setTitle("Formatted Document")

doc.addHeading("Text Formatting Examples", 1)

doc.bold("This text is bold.")
doc.italic("This text is italic.")
doc.underline("This text is underlined.")

doc.addParagraph("This is red text.", [:color = "red"])
doc.addParagraph("This is blue text.", [:color = "blue"])
doc.addParagraph("This is green text.", [:color = "00FF00"])

doc.addParagraph("Custom font and size", [:font = "Arial", :size = 16, :bold = true])
doc.addParagraph("Another style", [:font = "Times New Roman", :size = 14, :italic = true, :color = "purple"])

if doc.save("demo2_formatting.docx")
    ? "  Created: demo2_formatting.docx"
else
    ? "  FAILED: demo2_formatting.docx"
ok

# Demo 3: Headings and Structure
? "Demo 3: Creating document with headings..."
doc = new WordWriter()
doc.setTitle("Structured Document")

doc.addHeading("Main Title", 1)
doc.addParagraph("This is the introduction paragraph.", NULL)

doc.addHeading("Chapter 1: Getting Started", 2)
doc.addParagraph("Content for chapter 1 goes here.", NULL)

doc.addHeading("Section 1.1", 3)
doc.addParagraph("More detailed content in subsection.", NULL)

doc.addHeading("Section 1.2", 3)
doc.addParagraph("Another subsection with content.", NULL)

doc.addHeading("Chapter 2: Advanced Topics", 2)
doc.addParagraph("Content for chapter 2.", NULL)

if doc.save("demo3_headings.docx")
    ? "  Created: demo3_headings.docx"
else
    ? "  FAILED: demo3_headings.docx"
ok

# Demo 4: Lists
? "Demo 4: Creating document with lists..."
doc = new WordWriter()
doc.setTitle("Document with Lists")

doc.addHeading("Lists Demo", 1)

doc.addParagraph("Here is a bullet list:", NULL)
doc.addBulletList([
    "First item",
    "Second item",
    "Third item",
    "Fourth item"
])

doc.addParagraph("Here is a numbered list:", NULL)
doc.addNumberedList([
    "Step one",
    "Step two",
    "Step three",
    "Step four"
])

if doc.save("demo4_lists.docx")
    ? "  Created: demo4_lists.docx"
else
    ? "  FAILED: demo4_lists.docx"
ok

# Demo 5: Tables
? "Demo 5: Creating document with tables..."
doc = new WordWriter()
doc.setTitle("Document with Tables")

doc.addHeading("Tables Demo", 1)

doc.addParagraph("Here is a simple table:", NULL)

tableData = [
    ["Name", "Age", "City"],
    ["Ahmed", "25", "Cairo"],
    ["Sara", "30", "Riyadh"],
    ["Omar", "28", "Dubai"],
    ["Fatima", "22", "Amman"]
]
doc.addSimpleTable(tableData)

doc.addParagraph("Another table without header formatting:", NULL)
doc.addTable([
    ["Product", "Price", "Quantity"],
    ["Laptop", "$999", "5"],
    ["Mouse", "$25", "20"],
    ["Keyboard", "$75", "15"]
], [:borderStyle = "single"])

if doc.save("demo5_tables.docx")
    ? "  Created: demo5_tables.docx"
else
    ? "  FAILED: demo5_tables.docx"
ok

# Demo 6: Rich Text (Multiple formats in one paragraph)
? "Demo 6: Creating document with rich text..."
doc = new WordWriter()
doc.setTitle("Rich Text Document")

doc.addHeading("Rich Text Examples", 1)

doc.addRichParagraph([
    ["This is ", []],
    ["bold", [:bold = true]],
    [" and this is ", []],
    ["italic", [:italic = true]],
    [" and this is ", []],
    ["red", [:color = "red"]],
    [".", []]
], NULL)

doc.addRichParagraph([
    ["Welcome to ", [:size = 14]],
    ["DOCXLib", [:bold = true, :color = "blue", :size = 16]],
    [" - the ", [:size = 14]],
    ["best", [:italic = true, :underline = true]],
    [" Word library for Ring!", [:size = 14]]
], NULL)

if doc.save("demo6_richtext.docx")
    ? "  Created: demo6_richtext.docx"
else
    ? "  FAILED: demo6_richtext.docx"
ok

# Demo 7: Alignment
? "Demo 7: Creating document with alignments..."
doc = new WordWriter()
doc.setTitle("Alignment Demo")

doc.addHeading("Text Alignment Examples", 1)

doc.addParagraph("This text is left aligned (default).", NULL)
doc.centered("This text is centered.")
doc.rightAligned("This text is right aligned.")
doc.justified("This text is justified. It will stretch to fill the entire width of the page when there is enough text. Add more content here to see the justification effect properly.")

if doc.save("demo7_alignment.docx")
    ? "  Created: demo7_alignment.docx"
else
    ? "  FAILED: demo7_alignment.docx"
ok

# Demo 8: Page breaks and special elements
? "Demo 8: Creating document with page breaks..."
doc = new WordWriter()
doc.setTitle("Multi-page Document")

doc.addHeading("Page 1", 1)
doc.addParagraph("This is content on the first page.", NULL)
doc.addHorizontalLine()
doc.addParagraph("More content after a horizontal line.", NULL)

doc.addPageBreak()

doc.addHeading("Page 2", 1)
doc.addParagraph("This is content on the second page.", NULL)

doc.addPageBreak()

doc.addHeading("Page 3", 1)
doc.addParagraph("This is content on the third page.", NULL)

if doc.save("demo8_multipage.docx")
    ? "  Created: demo8_multipage.docx"
else
    ? "  FAILED: demo8_multipage.docx"
ok

# Demo 9: Hyperlinks
? "Demo 9: Creating document with hyperlinks..."
doc = new WordWriter()
doc.setTitle("Document with Links")

doc.addHeading("Hyperlinks Demo", 1)
doc.addParagraph("Click the links below to visit websites:", NULL)
doc.addHyperlink("Visit Ring Language Website", "https://ring-lang.github.io/")
doc.addHyperlink("Visit Google", "https://www.google.com/")
doc.addHyperlink("Visit GitHub", "https://github.com/")

if doc.save("demo9_hyperlinks.docx")
    ? "  Created: demo9_hyperlinks.docx"
else
    ? "  FAILED: demo9_hyperlinks.docx"
ok

# Demo 10: Complete professional document
? "Demo 10: Creating complete professional document..."
doc = new WordWriter()
doc.setTitle("Project Report")
doc.setAuthor("Development Team")
doc.setDefaultFont("Calibri", 11)

doc.addHeading("Project Status Report", 1)
doc.addParagraph("Date: January 2026", NULL)
doc.addParagraph("Prepared by: Development Team", NULL)
doc.addHorizontalLine()

doc.addHeading("Executive Summary", 2)
doc.addParagraph("This report provides an overview of the current project status, including completed milestones, ongoing tasks, and upcoming deliverables.", NULL)

doc.addHeading("Project Milestones", 2)
doc.addSimpleTable([
    ["Milestone", "Status", "Due Date"],
    ["Requirements Analysis", "Complete", "Dec 15"],
    ["Design Phase", "Complete", "Dec 30"],
    ["Development", "In Progress", "Jan 31"],
    ["Testing", "Pending", "Feb 15"],
    ["Deployment", "Pending", "Feb 28"]
])

doc.addHeading("Key Achievements", 2)
doc.addBulletList([
    "Completed system architecture design",
    "Implemented core functionality",
    "Established CI/CD pipeline",
    "Conducted code reviews"
])

doc.addHeading("Next Steps", 2)
doc.addNumberedList([
    "Complete remaining development tasks",
    "Begin integration testing",
    "Prepare user documentation",
    "Schedule training sessions"
])

doc.addHeading("Risks and Mitigation", 2)
doc.addRichParagraph([
    ["Risk: ", [:bold = true]],
    ["Resource availability may impact timeline.", []],
    [" Mitigation: ", [:bold = true]],
    ["Cross-training team members.", []]
], NULL)

doc.addPageBreak()

doc.addHeading("Appendix", 2)
doc.addParagraph("For more information, contact the project manager.", [:bold = true])
doc.addHyperlink("Project Documentation", "https://docs.example.com/project")

if doc.save("demo10_professional.docx")
    ? "  Created: demo10_professional.docx"
else
    ? "  FAILED: demo10_professional.docx"
ok

# Demo 11: Quick document creation
? "Demo 11: Using quick document function..."
quickWord("demo11_quick.docx", "This is a quick document created with a single function call!")
? "  Created: demo11_quick.docx"

quickWord("demo12_quick_list.docx", [
    "Paragraph 1: Introduction",
    "Paragraph 2: Main content goes here.",
    "Paragraph 3: Conclusion and summary."
])
? "  Created: demo12_quick_list.docx"

# Demo 13: Images
? "Demo 13: Creating document with images..."
doc = new WordWriter()
doc.setTitle("Document with Images")

doc.addHeading("Images Demo", 1)
doc.addParagraph("This document demonstrates how to add images.", NULL)

# Check for image files and add them
doc.addHeading("Test Images", 2)

if fexists("images/test1.png")
    doc.addParagraph("Here is images/test1.png:", NULL)
    doc.addImage("images/test1.png", 12, 8)  # 12cm x 8cm
else
    doc.addParagraph("(images/test1.png not found - skipping)", [:color = "gray"])
ok

if fexists("images/test2.jpg")
    doc.addParagraph("Here is images/test2.jpg (centered):", NULL)
    doc.addImageCentered("images/test2.jpg", 10, 7)  # 10cm x 7cm, centered
else
    doc.addParagraph("(images/test2.jpg not found - skipping)", [:color = "gray"])
ok

if fexists("images/test3.bmp")
    doc.addParagraph("Here is images/test3.bmp:", NULL)
    doc.addImage("images/test3.bmp", 8, 6)  # 8cm x 6cm
else
    doc.addParagraph("(images/test3.bmp not found - skipping)", [:color = "gray"])
ok

doc.addParagraph("End of images demo.", NULL)

if doc.save("demo13_images.docx")
    ? "  Created: demo13_images.docx"
else
    ? "  FAILED: demo13_images.docx"
ok

# Demo 14: Header and Footer
? "Demo 14: Creating document with header and footer..."
doc = new WordWriter()
doc.setTitle("Document with Header/Footer")

doc.setHeader("Company Name - Confidential Document")
doc.setFooter("© 2026 Company Name")

doc.addHeading("Header and Footer Demo", 1)
doc.addParagraph("This document has a header and footer on every page.", NULL)
doc.addParagraph("The header shows 'Company Name - Confidential Document'.", NULL)
doc.addParagraph("The footer shows '© 2026 Company Name'.", NULL)

doc.addPageBreak()

doc.addHeading("Page 2", 1)
doc.addParagraph("Notice how the header and footer appear on this page too.", NULL)

if doc.save("demo14_header_footer.docx")
    ? "  Created: demo14_header_footer.docx"
else
    ? "  FAILED: demo14_header_footer.docx"
ok

# Demo 15: Page Numbers
? "Demo 15: Creating document with page numbers..."
doc = new WordWriter()
doc.setTitle("Document with Page Numbers")

doc.setHeader("Project Documentation")
doc.showPageNumbers("center")

doc.addHeading("Page Numbers Demo", 1)
doc.addParagraph("This document has automatic page numbers in the footer.", NULL)
doc.addParagraph("Page numbers show 'Page X of Y' format.", NULL)

doc.addPageBreak()

doc.addHeading("Chapter 1: Introduction", 1)
doc.addParagraph("This is the second page of the document.", NULL)
doc.addParagraph("Check the footer to see the page number.", NULL)

doc.addPageBreak()

doc.addHeading("Chapter 2: Details", 1)
doc.addParagraph("This is the third page.", NULL)

if doc.save("demo15_page_numbers.docx")
    ? "  Created: demo15_page_numbers.docx"
else
    ? "  FAILED: demo15_page_numbers.docx"
ok

# Demo 16: Styled Tables with Colors
? "Demo 16: Creating document with styled tables..."
doc = new WordWriter()
doc.setTitle("Styled Tables Demo")

doc.addHeading("Styled Tables", 1)

doc.addParagraph("Table with blue header and zebra striping:", NULL)
doc.addStyledTable([
    ["Product", "Category", "Price", "Stock"],
    ["Laptop", "Electronics", "$999", "50"],
    ["Mouse", "Accessories", "$25", "200"],
    ["Keyboard", "Accessories", "$75", "150"],
    ["Monitor", "Electronics", "$350", "75"],
    ["Headphones", "Accessories", "$120", "100"]
], NULL, NULL)

doc.addParagraph("Table with custom colors (green header, light green rows):", NULL)
doc.addStyledTable([
    ["Name", "Department", "Role"],
    ["Ahmed", "Engineering", "Developer"],
    ["Sara", "Marketing", "Manager"],
    ["Omar", "Sales", "Representative"],
    ["Fatima", "HR", "Coordinator"]
], "2E7D32", "E8F5E9")

if doc.save("demo16_styled_tables.docx")
    ? "  Created: demo16_styled_tables.docx"
else
    ? "  FAILED: demo16_styled_tables.docx"
ok

# Demo 17: Table Border Customization
? "Demo 17: Creating document with custom table borders..."
doc = new WordWriter()
doc.setTitle("Table Borders Demo")

doc.addHeading("Table Border Styles", 1)

doc.addParagraph("Single border (default):", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "single", :headerRow = true])

doc.addParagraph("Double border:", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "double", :headerRow = true])

doc.addParagraph("Dashed border:", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "dashed", :headerRow = true])

doc.addParagraph("Dotted border:", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "dotted", :headerRow = true])

doc.addParagraph("Thick red border:", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "single", :borderSize = 12, :borderColor = "FF0000", :headerRow = true])

doc.addParagraph("No border:", NULL)
doc.addTable([
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"]
], [:borderStyle = "none"])

if doc.save("demo17_table_borders.docx")
    ? "  Created: demo17_table_borders.docx"
else
    ? "  FAILED: demo17_table_borders.docx"
ok

# Demo 18: Complete Professional Report
? "Demo 18: Creating complete professional report..."
doc = new WordWriter()
doc.setTitle("Quarterly Sales Report")
doc.setAuthor("Sales Department")
doc.setDefaultFont("Calibri", 11)

doc.setHeader("Q4 2025 Sales Report - Confidential")
doc.showPageNumbers("right")

doc.addHeading("Q4 2025 Sales Report", 1)
doc.addParagraph("Prepared by: Sales Department", [:italic = true])
doc.addParagraph("Date: January 15, 2026", [:italic = true])
doc.addHorizontalLine()

doc.addHeading("Executive Summary", 2)
doc.addParagraph("This report summarizes the sales performance for the fourth quarter of 2025. Overall, we achieved a 15% increase in revenue compared to Q3.", NULL)

doc.addHeading("Sales by Region", 2)
doc.addStyledTable([
    ["Region", "Q3 Sales", "Q4 Sales", "Growth"],
    ["North America", "$2.5M", "$2.9M", "+16%"],
    ["Europe", "$1.8M", "$2.1M", "+17%"],
    ["Asia Pacific", "$1.2M", "$1.4M", "+17%"],
    ["Latin America", "$0.5M", "$0.55M", "+10%"]
], "1565C0", "E3F2FD")

doc.addHeading("Top Products", 2)
doc.addStyledTable([
    ["Product", "Units Sold", "Revenue"],
    ["Enterprise Suite", "450", "$1.35M"],
    ["Professional Edition", "1,200", "$960K"],
    ["Basic Package", "3,500", "$525K"],
    ["Add-on Services", "800", "$320K"]
], "2E7D32", "E8F5E9")

doc.addPageBreak()

doc.addHeading("Key Achievements", 2)
doc.addBulletList([
    "Exceeded quarterly target by 8%",
    "Acquired 45 new enterprise clients",
    "Expanded into 3 new markets",
    "Reduced customer churn by 12%"
])

doc.addHeading("Challenges", 2)
doc.addBulletList([
    "Supply chain delays affected product availability",
    "Currency fluctuations impacted European margins",
    "Increased competition in Asia Pacific region"
])

doc.addHeading("Q1 2026 Outlook", 2)
doc.addParagraph("Based on current pipeline and market trends, we project continued growth in Q1 2025:", NULL)
doc.addNumberedList([
    "Target revenue: $7.5M (+10% vs Q4)",
    "Focus markets: North America and Europe",
    "New product launch scheduled for February",
    "Expanded marketing campaigns in APAC"
])

doc.addHeading("Recommendations", 2)
doc.addRichParagraph([
    ["Priority 1: ", [:bold = true]],
    ["Accelerate hiring in sales team to support growth.", []]
], NULL)
doc.addRichParagraph([
    ["Priority 2: ", [:bold = true]],
    ["Invest in customer success to reduce churn further.", []]
], NULL)
doc.addRichParagraph([
    ["Priority 3: ", [:bold = true]],
    ["Explore partnership opportunities in emerging markets.", []]
], NULL)

if doc.save("demo18_professional_report.docx")
    ? "  Created: demo18_professional_report.docx"
else
    ? "  FAILED: demo18_professional_report.docx"
ok

# Demo 19: Two-Column Layout
? "Demo 19: Creating document with two columns..."
doc = new WordWriter()
doc.setTitle("Two Column Document")
doc.setTwoColumns()

doc.addHeading("Two-Column Layout", 1)
doc.addParagraph("This document uses a two-column layout, which is common in newspapers, magazines, and academic journals. The text will flow from the bottom of the first column to the top of the second column.", NULL)

doc.addParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.", NULL)

doc.addParagraph("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident.", NULL)

doc.addColumnBreak()

doc.addHeading("Second Column", 2)
doc.addParagraph("After a column break, content continues in the next column. This is useful for controlling where content appears.", NULL)

doc.addParagraph("Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam.", NULL)

if doc.save("demo19_two_columns.docx")
    ? "  Created: demo19_two_columns.docx"
else
    ? "  FAILED: demo19_two_columns.docx"
ok

# Demo 20: Different Page Sizes
? "Demo 20: Creating documents with different page sizes..."

# A4 Document
doc = new WordWriter()
doc.setTitle("A4 Document")
doc.setPageSize("a4")

doc.addHeading("A4 Page Size", 1)
doc.addParagraph("This document uses A4 paper size (210 x 297 mm), which is the international standard for documents.", NULL)
doc.addParagraph("A4 is commonly used in Europe, Asia, and most countries outside North America.", NULL)

if doc.save("demo20a_a4_size.docx")
    ? "  Created: demo20a_a4_size.docx"
else
    ? "  FAILED: demo20a_a4_size.docx"
ok

# A5 Document (smaller, like a booklet)
doc = new WordWriter()
doc.setTitle("A5 Booklet")
doc.setPageSize("a5")

doc.addHeading("A5 Booklet", 1)
doc.addParagraph("This document uses A5 paper size (148 x 210 mm), which is half the size of A4.", NULL)
doc.addParagraph("A5 is perfect for booklets, flyers, and pocket-sized documents.", NULL)

if doc.save("demo20b_a5_size.docx")
    ? "  Created: demo20b_a5_size.docx"
else
    ? "  FAILED: demo20b_a5_size.docx"
ok

# Landscape Document
doc = new WordWriter()
doc.setTitle("Landscape Document")
doc.setPageSize("letter")
doc.setOrientation("landscape")

doc.addHeading("Landscape Orientation", 1)
doc.addParagraph("This document uses landscape orientation, which is wider than it is tall.", NULL)
doc.addParagraph("Landscape is ideal for wide tables, charts, presentations, and certificates.", NULL)

doc.addSimpleTable([
    ["January", "February", "March", "April", "May", "June"],
    ["$10,000", "$12,000", "$11,500", "$13,000", "$14,500", "$15,000"]
])

if doc.save("demo20c_landscape.docx")
    ? "  Created: demo20c_landscape.docx"
else
    ? "  FAILED: demo20c_landscape.docx"
ok

# Demo 21: Block Quotes and Captions
? "Demo 21: Creating document with block quotes and captions..."
doc = new WordWriter()
doc.setTitle("Quotes and Captions")

doc.addHeading("Using Block Quotes", 1)
doc.addParagraph("Block quotes are used to highlight important quotations or excerpts:", NULL)

doc.addBlockQuote("The only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. - Steve Jobs")

doc.addParagraph("Block quotes are indented and styled differently to stand out from regular text.", NULL)

doc.addBlockQuote("In the middle of difficulty lies opportunity. - Albert Einstein")

doc.addHeading("Figure Captions", 2)
doc.addParagraph("Captions are used to describe figures and tables:", NULL)
doc.addHorizontalLine()
doc.addCaption("Figure 1: Sample chart showing quarterly revenue growth")

doc.addParagraph("Tables can also have captions:", NULL)
doc.addSimpleTable([
    ["Quarter", "Revenue"],
    ["Q1", "$1.2M"],
    ["Q2", "$1.5M"]
])
doc.addCaption("Table 1: Quarterly revenue summary")

if doc.save("demo21_quotes_captions.docx")
    ? "  Created: demo21_quotes_captions.docx"
else
    ? "  FAILED: demo21_quotes_captions.docx"
ok

# Demo 22: Research Paper Format
? "Demo 22: Creating research paper format..."
doc = new WordWriter()
doc.setTitle("Research Paper")
doc.setAuthor("Dr. Ahmed Hassan, Prof. Sara Ali")
doc.setPageSize("a4")
doc.setTwoColumns()
doc.setNarrowMargins()
doc.showPageNumbers("center")

# Title (single column would be better but we simulate with centered text)
doc.addHeading("A Novel Approach to Natural Language Processing Using Ring Programming Language", 1)

doc.addParagraph("Dr. Ahmed Hassan¹, Prof. Sara Ali²", [:align = "center"])
doc.addParagraph("¹Department of Computer Science, King Saud University", [:align = "center", :size = 9])
doc.addParagraph("²Faculty of Engineering, Cairo University", [:align = "center", :size = 9])
doc.addHorizontalLine()

doc.addAbstract("This paper presents a novel approach to natural language processing using the Ring programming language. We demonstrate how Ring's unique features enable efficient text processing and analysis. Our experiments show significant improvements in processing speed compared to traditional approaches. The proposed method achieves 95% accuracy on standard benchmarks while reducing computational overhead by 40%.")

doc.addKeywords(["Natural Language Processing", "Ring Programming", "Text Analysis", "Machine Learning"])

doc.addHeading("1. Introduction", 2)
doc.addParagraph("Natural language processing (NLP) has become increasingly important in modern computing applications. From chatbots to sentiment analysis, NLP techniques are used across various domains.", NULL)

doc.addParagraph("In this paper, we explore the use of Ring programming language for NLP tasks. Ring offers several advantages including simplicity, flexibility, and good performance characteristics.", NULL)

doc.addColumnBreak()

doc.addHeading("2. Related Work", 2)
doc.addParagraph("Previous research has explored various programming languages for NLP applications. Python has been the dominant choice due to its extensive library ecosystem.", NULL)

doc.addBlockQuote("The choice of programming language significantly impacts both development time and runtime performance of NLP applications. - Smith et al., 2026")

doc.addHeading("3. Methodology", 2)
doc.addParagraph("Our approach consists of three main components: tokenization, feature extraction, and classification. Each component is implemented using Ring's object-oriented features.", NULL)

doc.addHeading("4. Results", 2)
doc.addStyledTable([
    ["Method", "Accuracy", "Speed"],
    ["Baseline", "87%", "1.0x"],
    ["Python NLP", "92%", "0.8x"],
    ["Our Method", "95%", "1.4x"]
], "1565C0", "E3F2FD")
doc.addCaption("Table 1: Comparison of different NLP approaches")

doc.addHeading("5. Conclusion", 2)
doc.addParagraph("We have demonstrated that Ring programming language is a viable option for NLP applications. Our method achieves state-of-the-art results while maintaining excellent performance.", NULL)

doc.addHeading("References", 2)
doc.addParagraph("[1] Smith, J. et al. (2023). Programming Languages for NLP. Journal of Computing.", [:size = 9])
doc.addParagraph("[2] Hassan, A. (2024). Ring Language Applications. Tech Report.", [:size = 9])

if doc.save("demo22_research_paper.docx")
    ? "  Created: demo22_research_paper.docx"
else
    ? "  FAILED: demo22_research_paper.docx"
ok

# Demo 23: Custom Page Settings
? "Demo 23: Creating document with custom page settings..."
doc = new WordWriter()
doc.setTitle("Custom Page Settings")

# Custom page size (15cm x 20cm)
doc.setCustomPageSize(15, 20)
doc.setMargins(1.5, 1.5, 2, 2)

doc.addHeading("Custom Page Size", 1)
doc.addParagraph("This document uses a custom page size of 15cm x 20cm with custom margins.", NULL)
doc.addParagraph("Custom page sizes are useful for special print jobs, booklets, or unique document requirements.", NULL)

if doc.save("demo23_custom_page.docx")
    ? "  Created: demo23_custom_page.docx"
else
    ? "  FAILED: demo23_custom_page.docx"
ok

# Demo 24: Three-Column Newsletter
? "Demo 24: Creating three-column newsletter..."
doc = new WordWriter()
doc.setTitle("Company Newsletter")
doc.setPageSize("letter")
doc.setThreeColumns()
doc.setNarrowMargins()
doc.setHeader("Monthly Newsletter - January 2026")

doc.addHeading("Company News", 1)

doc.addHeading("New Product Launch", 3)
doc.addParagraph("We are excited to announce our latest product release. After months of development, the new software is ready for customers.", NULL)

doc.addHeading("Team Updates", 3)
doc.addParagraph("Welcome to our new team members who joined this month. We're growing fast!", NULL)

doc.addColumnBreak()

doc.addHeading("Events", 2)

doc.addHeading("Annual Conference", 3)
doc.addParagraph("Don't miss our annual conference on February 15th. Register now to secure your spot.", NULL)

doc.addHeading("Training Sessions", 3)
doc.addParagraph("New training sessions are available. Check the intranet for the schedule.", NULL)

doc.addColumnBreak()

doc.addHeading("Tips & Tricks", 2)

doc.addHeading("Productivity Tip", 3)
doc.addParagraph("Use keyboard shortcuts to save time. Ctrl+S to save, Ctrl+C to copy.", NULL)

doc.addHeading("Health Corner", 3)
doc.addParagraph("Remember to take regular breaks. Stand up and stretch every hour.", NULL)

if doc.save("demo24_newsletter.docx")
    ? "  Created: demo24_newsletter.docx"
else
    ? "  FAILED: demo24_newsletter.docx"
ok

# Demo 25: Table of Contents
? "Demo 25: Creating document with table of contents..."
doc = new WordWriter()
doc.setTitle("Document with TOC")
doc.setHeader("User Manual")
doc.showPageNumbers("right")

doc.addHeading("User Manual", 1)
doc.addTableOfContents("Contents")
doc.addPageBreak()

doc.addHeading("Chapter 1: Getting Started", 1)
doc.addParagraph("This chapter covers the basics of getting started with the software.", NULL)
doc.addHeading("Installation", 2)
doc.addParagraph("Download the installer from our website and run the setup wizard.", NULL)
doc.addHeading("Configuration", 2)
doc.addParagraph("Configure the software according to your needs using the settings panel.", NULL)

doc.addPageBreak()

doc.addHeading("Chapter 2: Basic Features", 1)
doc.addParagraph("Learn about the basic features of the application.", NULL)
doc.addHeading("Creating Documents", 2)
doc.addParagraph("Click File > New to create a new document.", NULL)
doc.addHeading("Saving Documents", 2)
doc.addParagraph("Click File > Save or press Ctrl+S to save your work.", NULL)

doc.addPageBreak()

doc.addHeading("Chapter 3: Advanced Features", 1)
doc.addParagraph("Explore advanced features for power users.", NULL)
doc.addHeading("Automation", 2)
doc.addParagraph("Use macros to automate repetitive tasks.", NULL)
doc.addHeading("Customization", 2)
doc.addParagraph("Customize the interface to match your workflow.", NULL)

if doc.save("demo25_table_of_contents.docx")
    ? "  Created: demo25_table_of_contents.docx"
else
    ? "  FAILED: demo25_table_of_contents.docx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
? "Created files:"
? "  1. demo1_simple.docx"
? "  2. demo2_formatting.docx"
? "  3. demo3_headings.docx"
? "  4. demo4_lists.docx"
? "  5. demo5_tables.docx"
? "  6. demo6_richtext.docx"
? "  7. demo7_alignment.docx"
? "  8. demo8_multipage.docx"
? "  9. demo9_hyperlinks.docx"
? "  10. demo10_professional.docx"
? "  11. demo11_quick.docx"
? "  12. demo12_quick_list.docx"
? "  13. demo13_images.docx"
? "  14. demo14_header_footer.docx"
? "  15. demo15_page_numbers.docx"
? "  16. demo16_styled_tables.docx"
? "  17. demo17_table_borders.docx"
? "  18. demo18_professional_report.docx"
? "  19. demo19_two_columns.docx"
? "  20. demo20a_a4_size.docx, demo20b_a5_size.docx, demo20c_landscape.docx"
? "  21. demo21_quotes_captions.docx"
? "  22. demo22_research_paper.docx"
? "  23. demo23_custom_page.docx"
? "  24. demo24_newsletter.docx"
? "  25. demo25_table_of_contents.docx"
