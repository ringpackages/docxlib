/*
	DOCXLib - Demo (Comments)
*/

load "docxlib.ring"

/*
    Content Controls
    Checkboxes, dropdowns, text inputs
*/

doc = new WordWriter()
doc.setTitle("Content Controls Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setHeader("DOCXLib — Content Controls")
doc.showPageNumbers("right")

doc.addHeading("Feature 1 — Content Controls", 1)
doc.addParagraph("Interactive form elements embedded directly in Word.", NULL)
doc.addParagraph("", NULL)

# -- Checkboxes --
doc.addHeading("Checkboxes", 2)
doc.addCheckbox("I have read and accept the Terms and Conditions", true)
doc.addCheckbox("Subscribe to the Ring language newsletter", false)
doc.addCheckbox("Enable two-factor authentication", true)
doc.addCheckbox("Receive promotional emails", false)
doc.addCheckbox("Allow anonymous usage statistics", true)

doc.addParagraph("", NULL)

# -- Dropdowns --
doc.addHeading("Dropdown Lists", 2)
doc.addDropdown("Country",
    ["Saudi Arabia", "Egypt", "UAE", "Kuwait", "Jordan", "Morocco", "Other"],
    "Saudi Arabia")
doc.addParagraph("", NULL)
doc.addDropdown("Experience Level",
    ["Beginner", "Intermediate", "Advanced", "Expert"],
    "Intermediate")
doc.addParagraph("", NULL)
doc.addDropdown("Preferred Format",
    ["DOCX", "PDF", "HTML", "Markdown"],
    "DOCX")

doc.addParagraph("", NULL)

# -- Text Inputs --
doc.addHeading("Text Input Fields", 2)
doc.addTextInput("Full Name", "", "Enter your full name here")
doc.addParagraph("", NULL)
doc.addTextInput("Email Address", "", "name@example.com")
doc.addParagraph("", NULL)
doc.addTextInput("Organisation", "King Saud University", NULL)
doc.addParagraph("", NULL)
doc.addTextInput("Comments", "", "Type your comments here...")

filename = "demo_controls.docx"
if doc.save(filename)
    ? "OK - Saved: " + filename
else
    ? "ERROR - Failed to save"
ok
