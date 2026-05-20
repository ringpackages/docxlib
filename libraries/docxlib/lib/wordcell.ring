/*
    DOCXLib  
*/

# ============================================================================
# WordCell Class - Rich content for table cells
# ============================================================================

/*
    WordCell allows full control over table cell contents.
    Instead of passing a plain string to a table cell, pass a WordCell object.
*/

class WordCell

    aRunTexts           # List of text strings for each run
    nRunCount           # Number of runs
    aCellImgPaths       # Parallel arrays for cell images
    aCellImgWidths
    aCellImgHeights
    aCellImgRelIds      # Filled in by registerImage + setCellImageRel
    aCellImgIds
    nCellImageCount     # Number of cell images
    cAlign              # Text alignment: left, center, right, both
    cBgColor            # Cell background color (hex)
    cVerticalAlign      # Vertical alignment: top, center, bottom
    nColSpan            # Column span (merge cells horizontally)
    nRowSpan            # Row span (merge cells vertically)
    cMerge              # Vertical merge: "" = none, "restart" = start, "continue" = continuation
    cBorderColor        # Individual cell border color (all sides — legacy)
    cBorderStyle        # Individual cell border style (all sides — legacy)
    aBorderTop          # nil or [:style, :color, :size]
    aBorderBottom
    aBorderLeft
    aBorderRight
    aBorderInsideH      # inner horizontal (for top-left cell spanning)
    aBorderInsideV      # inner vertical
    cTextDir            # "btLr" (bottom-to-left-to-right = sideways up),
                        # "tbRl" (top-bottom right-to-left = sideways down),
                        # "lrTb" (normal, same as default/empty)
    nPaddingTop         # Cell padding in twips
    nPaddingBottom
    nPaddingLeft
    nPaddingRight
    nWidth              # Cell width in twips (0 = auto)

    # Per-run formatting (parallel arrays)
    aRunBold
    aRunItalic
    aRunUnderline
    aRunStrike
    aRunFont
    aRunSize
    aRunColor
    aRunHighlight

    # Cell lists (parallel arrays: items stored as ";" delimited strings)
    aCellListItems      # List of ";" delimited item strings
    aCellListTypes      # List of list types (1=bullet, 2=number)
    aCellListIds        # List of registered list IDs
    nCellListCount

    # Cell hyperlinks (parallel arrays)
    aCellLinkTexts      # Link display texts
    aCellLinkRelIds     # Registered relationship IDs
    nCellLinkCount

    # Cell nested tables
    nCellTableCount     # Number of nested tables
    aCellTableXml       # Pre-generated table XML strings

    # Cell block quotes and captions (parallel arrays)
    aCellBlockTexts     # Text content items (quote or caption text)
    aCellBlockTypes     # Type: "quote" or "caption"
    nCellBlockCount

    func init
        aRunTexts = []
        nRunCount = 0
        aCellImgPaths = []
        aCellImgWidths = []
        aCellImgHeights = []
        aCellImgRelIds = []
        aCellImgIds = []
        nCellImageCount = 0
        cAlign = ""
        cBgColor = ""
        cVerticalAlign = ""
        nColSpan = 0
        nRowSpan = 0
        cMerge = ""
        cBorderColor = ""
        cBorderStyle = ""
        aBorderTop    = []
        aBorderBottom = []
        aBorderLeft   = []
        aBorderRight  = []
        aBorderInsideH = []
        aBorderInsideV = []
        cTextDir = ""
        nPaddingTop = 0
        nPaddingBottom = 0
        nPaddingLeft = 0
        nPaddingRight = 0
        nWidth = 0
        aRunBold = []
        aRunItalic = []
        aRunUnderline = []
        aRunStrike = []
        aRunFont = []
        aRunSize = []
        aRunColor = []
        aRunHighlight = []
        aCellListItems = []
        aCellListTypes = []
        aCellListIds = []
        nCellListCount = 0
        aCellLinkTexts = []
        aCellLinkRelIds = []
        nCellLinkCount = 0
        nCellTableCount = 0
        aCellTableXml = []
        aCellBlockTexts = []
        aCellBlockTypes = []
        nCellBlockCount = 0
        return self

    # --- Text Content Methods ---

    func setText text, options
        /*
            Set cell text (replaces any existing content).
        */
        aRunTexts = []
        nRunCount = 0
        aRunBold = []
        aRunItalic = []
        aRunUnderline = []
        aRunStrike = []
        aRunFont = []
        aRunSize = []
        aRunColor = []
        aRunHighlight = []
        return addRun(text, options)

    func addText text, options
        return addRun(text, options)

    func addRun text, options
        /*
            Add a text run with optional formatting.
            options: [:bold, :italic, :underline, :strike, :font, :size, :color, :highlight]
        */
        if ! isList(options)
            options = []
        ok
        nRunCount = nRunCount + 1
        aRunTexts + text
        
        # Extract formatting from options into parallel arrays
        if isList(options) and len(options) > 0
            if options[:bold] = true  aRunBold + 1
            else  aRunBold + 0 ok

            if options[:italic] = true  aRunItalic + 1
            else  aRunItalic + 0 ok

            if options[:underline] = true  aRunUnderline + 1
            else  aRunUnderline + 0 ok

            if options[:strike] = true  aRunStrike + 1
            else  aRunStrike + 0 ok

            if options[:font] != NULL  aRunFont + options[:font]
            else  aRunFont + "" ok

            if options[:size] != NULL  aRunSize + options[:size]
            else  aRunSize + 0 ok

            if options[:color] != NULL  aRunColor + options[:color]
            else  aRunColor + "" ok

            if options[:highlight] != NULL  aRunHighlight + options[:highlight]
            else  aRunHighlight + "" ok
        else
            aRunBold + 0
            aRunItalic + 0
            aRunUnderline + 0
            aRunStrike + 0
            aRunFont + ""
            aRunSize + 0
            aRunColor + ""
            aRunHighlight + ""
        ok

        return self

    func addLineBreak
        /*
            Add a line break within the cell (new line, same cell).
        */
        return addRun("__LINEBREAK__", NULL)

    func getRunCount
        return nRunCount

    func getRunText idx
        return aRunTexts[idx]

    # --- Image in Cell ---

    func addCellImage imagePath, width, height
        /*
            Add an image inside this cell.
            imagePath: path to image file
            width: width in centimeters (default 3)
            height: height in centimeters (default 2)
            Note: You must call doc.registerImage(imagePath) FIRST,
            then pass the relId/imageId to setCellImageRel().
        */
        if width = NULL width = 3 ok
        if height = NULL height = 2 ok
        nCellImageCount = nCellImageCount + 1
        aCellImgPaths + imagePath
        aCellImgWidths + width
        aCellImgHeights + height
        aCellImgRelIds + ""
        aCellImgIds + 0
        return self

    func setCellImageRel imgIdx, relId, imageId
        /*
            Set the relationship ID and image ID for a cell image.
            Call after registerImage() to link the image.
        */
        aCellImgRelIds[imgIdx] = relId
        aCellImgIds[imgIdx] = imageId
        return self

    # --- Cell Formatting Methods ---

    func setAlign align
        /*
            Set text alignment: "left", "center", "right", "both" (justify)
        */
        cAlign = align
        return self

    # --- Lists in Cells ---

    func addCellBulletList items, listId
        /*
            Add a bullet list inside this cell.
            items: list of strings
            listId: from doc.registerCellList(WORD_LIST_BULLET)
        */
        nCellListCount = nCellListCount + 1
        # Store items as newline-delimited string (str2list splits by newline)
        cItems = ""
        itemsLen = len(items)
        for i = 1 to itemsLen
            if i > 1 cItems += nl ok
            cItems += items[i]
        next
        aCellListItems + cItems
        aCellListTypes + WORD_LIST_BULLET
        aCellListIds + listId
        return self

    func addCellNumberedList items, listId
        /*
            Add a numbered list inside this cell.
            items: list of strings
            listId: from doc.registerCellList(WORD_LIST_NUMBER)
        */
        nCellListCount = nCellListCount + 1
        cItems = ""
        itemsLen = len(items)
        for i = 1 to itemsLen
            if i > 1 cItems += nl ok
            cItems += items[i]
        next
        aCellListItems + cItems
        aCellListTypes + WORD_LIST_NUMBER
        aCellListIds + listId
        return self

    # --- Hyperlinks in Cells ---

    func addCellHyperlink text, relId
        /*
            Add a hyperlink inside this cell.
            text: display text
            relId: from doc.registerHyperlink(url)
        */
        nCellLinkCount = nCellLinkCount + 1
        aCellLinkTexts + text
        aCellLinkRelIds + relId
        return self

    # --- Nested Tables in Cells ---

    func addCellTable tableXml
        /*
            Add a nested table inside this cell.
            tableXml: pre-generated table XML from doc.generateNestedTable()
        */
        nCellTableCount = nCellTableCount + 1
        aCellTableXml + tableXml
        return self

    func addCellBlockQuote text
        /*
            Add a block quote inside this cell.
            Rendered with left indent and italic gray styling.
        */
        nCellBlockCount = nCellBlockCount + 1
        aCellBlockTexts + text
        aCellBlockTypes + "quote"
        return self

    func addCellCaption text
        /*
            Add a caption inside this cell.
            Rendered centered, italic, smaller text.
        */
        nCellBlockCount = nCellBlockCount + 1
        aCellBlockTexts + text
        aCellBlockTypes + "caption"
        return self

    func setBgColor color
        /*
            Set cell background color.
            Accepts color name (red, blue, etc.) or hex (FF0000, #FF0000)
        */
        cBgColor = wordColorToHex(color)
        return self

    func setVerticalAlign align
        /*
            Set vertical alignment: "top", "center", "bottom"
        */
        cVerticalAlign = align
        return self

    func setColSpan span
        /*
            Set column span (merge cells horizontally).
        */
        nColSpan = span
        return self

    func setRowSpan span
        /*
            Set row span (merge cells vertically).
            The cell will span 'span' rows. Place wordMergeCell() in
            the same column for the next (span-1) rows.
        */
        nRowSpan = span
        if span > 1
            cMerge = "restart"
        ok
        return self

    func setBorder style, color
        /*
            Set all four cell borders to the same style/color (legacy).
            style: "single", "double", "dashed", "dotted", "none"
            color: hex color or name
        */
        cBorderStyle = style
        if color != NULL
            cBorderColor = wordColorToHex(color)
        ok
        return self

    func setBorderSide side, style, color, size
        /*
            Set a single border side independently.
            side  : "top", "bottom", "left", "right", "insideH", "insideV"
            style : "single", "double", "dashed", "dotted", "thick",
                    "thinThickSmallGap", "none" (remove border)
            color : hex color, color name, or NULL for "auto"
            size  : border size in eighths-of-a-point (4 = 0.5pt, 8 = 1pt,
                    16 = 2pt, 24 = 3pt). NULL defaults to 4.
        */
        if color = NULL  color = "auto"
        else  color = wordColorToHex(color)  ok
        if size = NULL  size = 4  ok
        bdrEntry = [:style=style, :color=color, :size=size]
        if side = "top"       aBorderTop     = bdrEntry  ok
        if side = "bottom"    aBorderBottom  = bdrEntry  ok
        if side = "left"      aBorderLeft    = bdrEntry  ok
        if side = "right"     aBorderRight   = bdrEntry  ok
        if side = "insideH"   aBorderInsideH = bdrEntry  ok
        if side = "insideV"   aBorderInsideV = bdrEntry  ok
        return self

    func setTopBorder style, color, size
        return setBorderSide("top", style, color, size)

    func setBottomBorder style, color, size
        return setBorderSide("bottom", style, color, size)

    func setLeftBorder style, color, size
        return setBorderSide("left", style, color, size)

    func setRightBorder style, color, size
        return setBorderSide("right", style, color, size)

    func setNoBorder
        /*  Remove all borders from this cell.  */
        bdrNone = [:style="none", :color="auto", :size=0]
        aBorderTop     = bdrNone
        aBorderBottom  = bdrNone
        aBorderLeft    = bdrNone
        aBorderRight   = bdrNone
        return self

    func setTextDir dir
        /*
            Set cell text direction.
            dir: "btLr"  — text runs bottom-to-top (rotated 90° CCW)
                 "tbRl"  — text runs top-to-bottom (rotated 90° CW)
                 "lrTb"  — normal horizontal (default, clears override)
        */
        cTextDir = dir
        return self

    func setPadding topCm, bottomCm, leftCm, rightCm
        /*
            Set cell padding in centimeters.
        */
        nPaddingTop = floor(topCm * 567)
        nPaddingBottom = floor(bottomCm * 567)
        nPaddingLeft = floor(leftCm * 567)
        nPaddingRight = floor(rightCm * 567)
        return self

    func setWidth widthCm
        /*
            Set explicit cell width in centimeters.
        */
        nWidth = floor(widthCm * 567)
        return self

