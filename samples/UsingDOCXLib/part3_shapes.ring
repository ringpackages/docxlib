/*
	DOCXLib - Demo (Shapes)
*/

load "docxlib.ring"

doc = new WordWriter()
doc.setTitle("Drawing Shapes Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setHeader("DOCXLib - Drawing Shapes")
doc.showPageNumbers("right")

doc.addHeading("Drawing Shapes", 1)
doc.addParagraph("Inline DrawingML shapes with fill, border, and text.", NULL)
doc.addParagraph("", NULL)

# -- Rectangles --
doc.addHeading("Rectangles", 2)
doc.addRect([:width=6, :height=1.4, :fillColor="1E6BB5", :lineColor="0D4A8A",
             :text="Standard Rectangle", :textColor="FFFFFF"])
doc.addParagraph("", NULL)
doc.addRect([:width=6, :height=1.4, :fillColor="E2EFDA", :lineColor="375623",
             :lineWidth=2, :text="Rounded Rectangle", :textColor="375623", :rounded=true])
doc.addParagraph("", NULL)
doc.addRect([:width=6, :height=1.4, :noFill=true, :lineColor="C00000",
             :lineWidth=2, :text="No Fill - Border Only", :textColor="C00000"])

doc.addParagraph("", NULL)

# -- Ellipses --
doc.addHeading("Ellipses", 2)
doc.addEllipse([:width=6, :height=1.8, :fillColor="7030A0", :lineColor="4B0082",
                :text="Ellipse Shape", :textColor="FFFFFF"])
doc.addParagraph("", NULL)
doc.addEllipse([:width=2.5, :height=2.5, :fillColor="ED7D31", :lineColor="C55A11",
                :text="Circle", :textColor="FFFFFF", :textBold=true, :textSize=14])

doc.addParagraph("", NULL)

# -- Other shapes --
doc.addHeading("Diamond and Triangle", 2)
doc.addDiamond([:width=5, :height=2.2, :fillColor="FFC000", :lineColor="C09000",
                :text="Diamond", :textColor="000000"])
doc.addParagraph("", NULL)
doc.addShape("triangle", [:width=5, :height=2.2, :fillColor="C00000", :lineColor="8B0000",
                           :text="Triangle", :textColor="FFFFFF"])

doc.addParagraph("", NULL)

# -- Lines --
doc.addHeading("Horizontal Lines", 2)
doc.addLine([:width=14, :height=0.08, :fillColor="1E6BB5", :noBorder=true])
doc.addParagraph("", NULL)
doc.addLine([:width=14, :height=0.20, :fillColor="ED7D31", :noBorder=true])
doc.addParagraph("", NULL)
doc.addLine([:width=14, :height=0.08, :fillColor="595959", :noBorder=true])

filename = "demo_shapes.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
