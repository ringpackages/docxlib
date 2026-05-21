/*
	DOCXLib - Demo (Hello World)
*/

load "docxlib.ring"

? "Generate File: helloworld.docx" 

new WordWriter() {
	setTitle("My First Document")
	setAuthor("Your Name")
	setPageSize("a4")
	addHeading("Hello, DOCXLib!", 1)
	addParagraph("This document was created entirely from Ring code.", [:color = "0000FF"])
	addBulletList(["Feature A", "Feature B", "Feature C"])
	save("helloworld.docx")
}