/*
    RingWordLib
*/

# ============================================================================
# WordWriter Class
# ============================================================================

class WordWriter

    # Document content storage
    aContent            # List of content elements
    aRelationships      # Document relationships
    aNumbering          # Numbering definitions for lists
    nRelId              # Relationship ID counter
    nListId             # List ID counter
    cAuthor             # Document author
    cTitle              # Document title
    
    # Default formatting
    cDefaultFont
    nDefaultSize
    
    # Image storage
    aImages             # List of images [relId, filename, data, contentType]
    nImageId            # Image ID counter
    
    # Cell image mapping (parallel arrays, indexed by registration order)
    aCellImgMapPaths    # image paths registered for cell images
    aCellImgMapRelIds   # corresponding relIds
    aCellImgMapImgIds   # corresponding imageIds
    aCellImgMapWidths   # widths in cm
    aCellImgMapHeights  # heights in cm
    
    # Header/Footer
    cHeaderText         # Header text
    cFooterText         # Footer text
    bShowPageNumbers    # Show page numbers in footer
    cPageNumberAlign    # Page number alignment (left, center, right)
    cHeaderRelId        # Header relationship ID
    cFooterRelId        # Footer relationship ID
    
    # Watermark – text
    bWatermark          # Whether a text watermark is enabled
    cWatermarkText      # Watermark display text
    cWatermarkColor     # Watermark fill color (hex, default light grey)
    nWatermarkOpacity   # Opacity 0-100 (default 50)
    nWatermarkRotation  # Rotation in degrees (default -45)
    cWatermarkFont      # Font family (default Arial)
    nWatermarkSize      # Font size in points (default 72)
    
    # Watermark – image
    bImgWatermark       # Whether an image watermark is enabled
    cImgWatermarkPath   # Path to the watermark image file
    cImgWatermarkRelId  # Relationship ID assigned to the watermark image (set at save time)
    cImgWatermarkFile   # Filename inside word/media/ (set at save time)
    nImgWatermarkWidth  # Display width in cm (default 15)
    nImgWatermarkHeight # Display height in cm (default 15)
    nImgWatermarkOpacity # Opacity 0-100 (default 50)
    
    # Footnotes
    aFootnotes          # List of [:id, :text] for each footnote
    nFootnoteId         # Sequential footnote ID counter (starts at 1)
    
    # Endnotes
    aEndnotes           # List of [:id, :text] for each endnote
    nEndnoteId          # Sequential endnote ID counter (starts at 1)
    
    # Page settings
    nPageWidth          # Page width in twips
    nPageHeight         # Page height in twips
    nMarginTop          # Top margin in twips
    nMarginBottom       # Bottom margin in twips
    nMarginLeft         # Left margin in twips
    nMarginRight        # Right margin in twips
    nColumns            # Number of columns (1-3)
    nColumnSpace        # Space between columns in twips
    cOrientation        # Page orientation: "portrait" or "landscape"
    
    # Page borders
    bPageBorder         # Whether page borders are enabled
    cPageBorderStyle    # Border style: "single","double","dashed","wave",etc.
    cPageBorderColor    # Border color hex (default "000000")
    nPageBorderSize     # Border width in eighths of a point (default 24 = 3pt)
    nPageBorderSpace    # Distance from page edge in points (default 24)
    cPageBorderOffsetFrom  # "page" or "text" (default "page")
    aPageBorderSides    # Which sides: list of "top","left","bottom","right"
    
    # RTL / BiDi
    bDocumentRTL        # Document-level RTL (affects Normal style + sectPr)
    
    # Custom styles
    aCustomStyles       # List of user-defined paragraph styles

    # First-page different header/footer
    cFirstPageHeaderText
    cFirstPageFooterText
    bFirstPageDifferent
    cFirstPageHeaderRelId
    cFirstPageFooterRelId

    # Text boxes
    nTextBoxId          # Unique ID counter for text box drawings

    # Odd/Even page headers/footers
    bEvenAndOddHeaders      # Enable distinct even-page header/footer
    cEvenPageHeaderText     # Header text for even pages
    cEvenPageFooterText     # Footer text for even pages
    cEvenPageHeaderRelId
    cEvenPageFooterRelId

    # Paragraph borders
    # (no global state needed — applied per paragraph via options)

    # Section breaks with orientation
    # (no global state needed — emitted inline as content items)

    # Page background color
    cPageBgColor            # Page background fill color (hex, NULL = none)

    # Bookmarks and cross-references
    nBookmarkId             # Auto-incrementing bookmark ID counter
    aBookmarks              # Registry: [:id, :name] for cross-ref lookups

    # Tab stops
    # (no global state — applied per paragraph via options)

    # Line numbers
    bLineNumbers            # Enable line numbering for the document
    nLineNumberStart        # Starting line number (default 1)
    nLineNumberStep         # Show number every N lines (default 1)
    nLineNumberDistance     # Distance from text margin in twips (default 360 = 0.25in)
    cLineNumberRestart      # "newPage", "newSection", "continuous"

    # Comments (annotations)
    aComments               # List of comment definitions [:id, :author, :date, :text]
    nCommentId              # Auto-incrementing comment ID counter

    # SEQ captions (auto-numbered figures and tables)
    nFigureSeq              # Running figure counter (for fallback display)
    nTableSeq               # Running table counter  (for fallback display)
    aFigureCaptions         # Ordered list of figure caption texts 
    aTableCaptions          # Ordered list of table caption texts  
    aSeqCaptions            # Generic registry: list of [:label, :num, :text] 

    # Mail merge fields
    # (no global state — each field emitted inline as a MERGEFIELD)

    # Content controls (SDT)
    nSdtId              # Auto-incrementing SDT ID counter

    # Drawing shapes (DrawingML)
    nShapeId            # Auto-incrementing shape ID counter

    # Document themes
    cThemeName          # Active theme name (NULL = no theme)
    aThemeColors        # [:dk1,:lt1,:dk2,:lt2,:acc1..:acc6] hex colors

    # Rich table cells in merge templates: mergeCell(), mergeFillWordCell();
    #            {{FIELD}} tokens inside WordCell runs are filled on merge
    # Mail merge template engine: setMergeTemplate(), mergeRecord(), mergeAll();
    #            Chart data tables: :showDataTable option
    # Chart style presets: setChartDefaults(), clearChartDefaults();
    #            Conditional table formatting: :conditionalRules option
    # Chart formatting: :yAxisMin/Max, :xAxisMin/Max, :yAxisFormat, :xAxisFormat,
    #            :yAxisTitle, :xAxisTitle (category charts), :showGridlines;
    #            pie/doughnut multi-series warning
    # Native OOXML Charts
    aCharts             # List of chart definitions [:chartId,:relId,:type,:title,:categories,:series,:options]
    nChartId            # Auto-incrementing chart ID counter (1-based)
    nDocPrId            # Shared docPr id counter for drawings (images + charts)
    aChartDefaults      # Default chart options merged into every addChart() call 
    aMergeTemplate      # Template definition list for mail merge 
    aCustomProps        # List of [:name,:value] custom document properties
    
    func init
        aContent = []
        aRelationships = []
        aNumbering = []
        aImages = []
        nRelId = 10  # Start at 10 to avoid conflicts with document relationships (rId1-rId4)
        nListId = 1
        nImageId = 0
        aCellImgMapPaths = []
        aCellImgMapRelIds = []
        aCellImgMapImgIds = []
        aCellImgMapWidths = []
        aCellImgMapHeights = []
        cAuthor = "RingWordLib"
        cTitle = "Document"
        cDefaultFont = "Calibri"
        nDefaultSize = 11
        cHeaderText = ""
        cFooterText = ""
        bShowPageNumbers = false
        cPageNumberAlign = "center"
        cHeaderRelId = ""
        cFooterRelId = ""
        bWatermark = false
        cWatermarkText = "DRAFT"
        cWatermarkColor = "C0C0C0"
        nWatermarkOpacity = 50
        nWatermarkRotation = -45
        cWatermarkFont = "Arial"
        nWatermarkSize = 72
        bImgWatermark = false
        cImgWatermarkPath = ""
        cImgWatermarkRelId = ""
        cImgWatermarkFile = ""
        nImgWatermarkWidth = 15
        nImgWatermarkHeight = 15
        nImgWatermarkOpacity = 50
        aFootnotes = []
        nFootnoteId = 0
        aEndnotes = []
        nEndnoteId = 0
        # Default page settings (Letter size, 1 inch margins)
        nPageWidth = 12240      # 8.5 inches in twips
        nPageHeight = 15840     # 11 inches in twips
        nMarginTop = 1440       # 1 inch
        nMarginBottom = 1440
        nMarginLeft = 1440
        nMarginRight = 1440
        nColumns = 1
        nColumnSpace = 720      # 0.5 inch
        cOrientation = "portrait"
        # Page borders
        bPageBorder = false
        cPageBorderStyle = "single"
        cPageBorderColor = "000000"
        nPageBorderSize = 24    # 3pt (24 eighths of a point)
        nPageBorderSpace = 24   # 24pt from page edge
        cPageBorderOffsetFrom = "page"
        aPageBorderSides = ["top", "left", "bottom", "right"]
        # RTL
        bDocumentRTL = false
        # Custom styles 
        aCustomStyles = []
        # First-page header/footer 
        cFirstPageHeaderText = ""
        cFirstPageFooterText = ""
        bFirstPageDifferent = false
        cFirstPageHeaderRelId = ""
        cFirstPageFooterRelId = ""
        # Text box ID 
        nTextBoxId = 0
        # Odd/even headers
        bEvenAndOddHeaders = false
        cEvenPageHeaderText = ""
        cEvenPageFooterText = ""
        cEvenPageHeaderRelId = ""
        cEvenPageFooterRelId = ""
        # Page background
        cPageBgColor = ""
        # Bookmarks
        nBookmarkId = 1
        aBookmarks = []
        # Line numbers
        bLineNumbers = false
        nLineNumberStart = 1
        nLineNumberStep = 1
        nLineNumberDistance = 360
        cLineNumberRestart = "newPage"
        # Comments
        aComments = []
        nCommentId = 0
        # SEQ captions
        nFigureSeq = 0
        nTableSeq  = 0
        # Caption registries for pre-populated TOF
        aFigureCaptions = []
        aTableCaptions  = []
        aSeqCaptions    = []
        # Content controls
        nSdtId = 1
        # Drawing shapes
        nShapeId = 1
        # Document theme
        cThemeName   = ""
        aThemeColors = []
        # Charts
        aCharts        = []
        aCustomProps   = []
        nChartId       = 0
        nDocPrId       = 0
        # Chart style presets
        aChartDefaults = []
        # Mail merge template
        aMergeTemplate = []
        return self
    
    # ========================================================================
    # Document Properties
    # ========================================================================
    
    func setAuthor author
        cAuthor = author
        return self
    
    func setTitle title
        cTitle = title
        return self
    
    func setDefaultFont fontName, fontSize
        cDefaultFont = fontName
        if fontSize != NULL
            nDefaultSize = fontSize
        ok
        return self
    
    # ========================================================================
    # Page Settings
    # ========================================================================
    
    func setPageSize size
        /*
            Set page size using predefined sizes
            Supported: "letter", "legal", "a4", "a5", "a3", "b5", "executive"
        */
        size = lower(size)
        
        switch size
            on "letter"
                nPageWidth = 12240      # 8.5 x 11 inches
                nPageHeight = 15840
            on "legal"
                nPageWidth = 12240      # 8.5 x 14 inches
                nPageHeight = 20160
            on "a4"
                nPageWidth = 11906      # 210 x 297 mm
                nPageHeight = 16838
            on "a5"
                nPageWidth = 8391       # 148 x 210 mm
                nPageHeight = 11906
            on "a3"
                nPageWidth = 16838      # 297 x 420 mm
                nPageHeight = 23811
            on "b5"
                nPageWidth = 10318      # 182 x 257 mm
                nPageHeight = 14570
            on "executive"
                nPageWidth = 10800      # 7.5 x 10 inches
                nPageHeight = 14400
        off
        
        return self
    
    func setCustomPageSize widthCm, heightCm
        /*
            Set custom page size in centimeters
            1 cm = 567 twips (approximately)
        */
        nPageWidth = floor(widthCm * 567)
        nPageHeight = floor(heightCm * 567)
        return self
    
    func setOrientation orientation
        /*
            Set page orientation: "portrait" or "landscape"
        */
        cOrientation = lower(orientation)
        if cOrientation = "landscape"
            # Swap width and height
            temp = nPageWidth
            nPageWidth = nPageHeight
            nPageHeight = temp
        ok
        return self
    
    func setMargins topCm, bottomCm, leftCm, rightCm
        /*
            Set page margins in centimeters
        */
        nMarginTop = floor(topCm * 567)
        nMarginBottom = floor(bottomCm * 567)
        nMarginLeft = floor(leftCm * 567)
        nMarginRight = floor(rightCm * 567)
        return self
    
    func setMarginsInches top, bottom, left, right
        /*
            Set page margins in inches
            1 inch = 1440 twips
        */
        nMarginTop = floor(top * 1440)
        nMarginBottom = floor(bottom * 1440)
        nMarginLeft = floor(left * 1440)
        nMarginRight = floor(right * 1440)
        return self
    
    func setNarrowMargins
        # 0.5 inch margins
        nMarginTop = 720
        nMarginBottom = 720
        nMarginLeft = 720
        nMarginRight = 720
        return self
    
    func setWideMargins
        # 1.5 inch left/right margins
        nMarginTop = 1440
        nMarginBottom = 1440
        nMarginLeft = 2160
        nMarginRight = 2160
        return self
    
    func setColumns numColumns, spaceCm
        /*
            Set number of columns (1-3)
            spaceCm: space between columns in centimeters (default 0.5)
        */
        if numColumns < 1 numColumns = 1 ok
        if numColumns > 3 numColumns = 3 ok
        nColumns = numColumns
        if spaceCm != NULL
            nColumnSpace = floor(spaceCm * 567)
        ok
        return self
    
    func setTwoColumns
        nColumns = 2
        nColumnSpace = 720  # 0.5 inch
        return self
    
    func setThreeColumns
        nColumns = 3
        nColumnSpace = 540  # ~0.38 inch
        return self
    
    # ========================================================================
    # Page Borders
    # ========================================================================
    
    func setPageBorder options
        /*
            Enable a page border. options is a list with keys:
              :style      – border style (default "single")
                            Many values: "single","double","triple","thick",
                            "dashed","dotted","wave","doubleWave",
                            "thinThickSmallGap","thickThinSmallGap",
                            "threeDEmboss","threeDEngrave","inset","outset"
              :color      – color name or hex (default "000000")
              :size       – thickness in points (default 3)
              :space      – distance from page edge in points (default 24)
              :offsetFrom – "page" or "text" (default "page")
              :sides      – list of sides, e.g. ["top","bottom"]
                            Default: all four sides
        */
        bPageBorder = true
        if !isList(options)
            return self
        ok
        if options[:style] != NULL
            cPageBorderStyle = options[:style]
        ok
        if options[:color] != NULL
            cPageBorderColor = wordColorToHex(options[:color])
        ok
        if options[:size] != NULL
            nPageBorderSize = floor(options[:size] * 8)   # pt → eighths-of-pt
        ok
        if options[:space] != NULL
            nPageBorderSpace = options[:space]
        ok
        if options[:offsetFrom] != NULL
            cPageBorderOffsetFrom = options[:offsetFrom]
        ok
        if options[:sides] != NULL
            aPageBorderSides = options[:sides]
        ok
        return self

    func setSimplePageBorder
        /* Quick shortcut: thin single black border on all four sides. */
        return setPageBorder([
            :style = "single",
            :color = "000000",
            :size  = 3,
            :space = 24
        ])

    func removePageBorder
        /* Disable page borders. */
        bPageBorder = false
        return self

    # ========================================================================
    # RTL / BiDi
    # ========================================================================

    func setDocumentRTL bEnable
        /*
            Enable or disable document-level RTL.
            When true the Normal style and the final section properties are marked
            BiDi so Word treats the whole document as right-to-left by default.
        */
        bDocumentRTL = bEnable
        return self

    func addRTLParagraph text, options
        /*
            Convenience: add a paragraph pre-configured for RTL text.
            Merges :rtl=true and right-alignment into options automatically.
        */
        if !isList(options)
            options = []
        ok
        options[:rtl] = true
        if options[:align] = NULL
            options[:align] = "right"
        ok
        return addParagraph(text, options)

    # ========================================================================
    # Custom Paragraph Styles
    # ========================================================================

    func defineStyle styleId, options
        /*
            Define a reusable named paragraph style.
            styleId : unique identifier, used with [:style = "MyId"] in addParagraph
            options : list with keys:
                :name        - display name (defaults to styleId)
                :basedOn     - parent style id (default "Normal")
                :bold        - true/false
                :italic      - true/false
                :underline   - true/false
                :font        - font name
                :size        - font size in points
                :color       - text color (name or hex)
                :bgColor     - paragraph shading fill color (hex)
                :align       - left/center/right/both
                :spaceBefore - space before in twips (1 pt = 20 twips)
                :spaceAfter  - space after in twips
                :lineSpacing - multiplier (1.0, 1.5, 2.0)
                :indent      - left indent in twips
                :keepNext    - true: keep paragraph with the next one
                :keepLines   - true: keep all lines together (no page break inside)
        */
        if !isList(options)
            options = []
        ok
        # Extract option values into locals (Ring cannot use subscripts inside list literals)
        stName        = options[:name]
        stBasedOn     = options[:basedOn]
        stBold        = options[:bold]
        stItalic      = options[:italic]
        stUnderline   = options[:underline]
        stFont        = options[:font]
        stSize        = options[:size]
        stColor       = options[:color]
        stBgColor     = options[:bgColor]
        stAlign       = options[:align]
        stSpaceBefore = options[:spaceBefore]
        stSpaceAfter  = options[:spaceAfter]
        stLineSpacing = options[:lineSpacing]
        stIndent      = options[:indent]
        stKeepNext    = options[:keepNext]
        stKeepLines   = options[:keepLines]
        style = [
            :id          = styleId,
            :name        = stName,
            :basedOn     = stBasedOn,
            :bold        = stBold,
            :italic      = stItalic,
            :underline   = stUnderline,
            :font        = stFont,
            :size        = stSize,
            :color       = stColor,
            :bgColor     = stBgColor,
            :align       = stAlign,
            :spaceBefore = stSpaceBefore,
            :spaceAfter  = stSpaceAfter,
            :lineSpacing = stLineSpacing,
            :indent      = stIndent,
            :keepNext    = stKeepNext,
            :keepLines   = stKeepLines
        ]
        aCustomStyles + style
        return self

    # ========================================================================
    # First-Page Different Header/Footer
    # ========================================================================

    func setFirstPageHeader text
        /*
            Set a different header for the first page only.
            The default header (setHeader) applies to all subsequent pages.
            Automatically enables first-page-different mode.
        */
        cFirstPageHeaderText = text
        bFirstPageDifferent = true
        return self

    func setFirstPageFooter text
        /*
            Set a different footer for the first page only.
            The default footer (setFooter/showPageNumbers) applies to subsequent pages.
            Automatically enables first-page-different mode.
        */
        cFirstPageFooterText = text
        bFirstPageDifferent = true
        return self

    func setFirstPageDifferent bEnable
        /*
            Explicitly control first-page-different mode.
            When true, the first page gets its own header/footer slot.
            Pass empty strings to setFirstPageHeader/Footer for a blank first page.
        */
        bFirstPageDifferent = bEnable
        return self

    # ========================================================================
    # Odd/Even Page Headers and Footers
    # ========================================================================

    func setEvenPageHeader text
        /*
            Set the header for even-numbered pages (2, 4, 6 ...).
            The default header (setHeader) is used for odd pages (1, 3, 5 ...).
            Automatically enables even/odd header mode.
        */
        cEvenPageHeaderText = text
        bEvenAndOddHeaders = true
        return self

    func setEvenPageFooter text
        /*
            Set the footer for even-numbered pages.
            The default footer / page numbers apply to odd pages.
            Automatically enables even/odd header mode.
        */
        cEvenPageFooterText = text
        bEvenAndOddHeaders = true
        return self

    func setEvenAndOddHeaders bEnable
        /*
            Explicitly control even/odd header mode without setting text.
            Useful to enable the mode so that odd pages show the default
            header while even pages are intentionally blank.
        */
        bEvenAndOddHeaders = bEnable
        return self

    # ========================================================================
    # Paragraph Borders
    # ========================================================================

    func addBorderedParagraph text, options
        /*
            Add a paragraph with a visible border around it.
            options keys (all optional):
                :borderStyle  - border style (default "single")
                                values: single, double, dashed, dotted, thick,
                                        wave, thinThickSmallGap, etc.
                :borderColor  - hex color (default "000000")
                :borderSize   - width in eighths of a point (default 6 = 0.75pt)
                :borderSpace  - space between text and border in points (default 4)
                :sides        - list of sides to draw, e.g. ["top","bottom"]
                                default: all four sides (box)
                + any standard addParagraph options (bold, italic, align, etc.)
        */
        if !isList(options) options = [] ok
        options[:_hasBorder] = true
        return addParagraph(text, options)

    # ========================================================================
    # Page Background Color
    # ========================================================================

    func setPageBackground color
        /*
            Set a solid background fill color for all pages in the document.
            color : color name ("lightblue") or hex string ("BDD7EE")
            Pass NULL or "" to remove the background.
            Note: Word also requires <w:displayBackgroundShape/> in settings.xml
            for the background to be visible on screen (automatically added).
        */
        if color = NULL or color = ""
            cPageBgColor = ""
        else
            cPageBgColor = wordColorToHex(color)
        ok
        return self

    # ========================================================================
    # Bookmarks and Cross-References
    # ========================================================================

    func addBookmarkedParagraph bookmarkName, text, options
        /*
            Add a paragraph wrapped in a named bookmark.
            bookmarkName : unique string identifier (e.g. "sec_intro")
            text         : paragraph text
            options      : same as addParagraph options
            Returns self for chaining.
            Use addCrossRef(bookmarkName) in another paragraph to link back.
        */
        if !isList(options) options = [] ok
        options[:_bookmark] = bookmarkName
        return addParagraph(text, options)

    func addCrossRef bookmarkName, displayText
        /*
            Add an inline cross-reference to a named bookmark.
            bookmarkName : name used in addBookmarkedParagraph()
            displayText  : text shown before the page number,
                           e.g. "see page" → outputs "see page 5"
                           Pass NULL to show only the page number.
            Appends a paragraph containing:
               "<displayText> <page-number-of-bookmark>"
        */
        if displayText = NULL displayText = "" ok
        item = [
            :type        = "crossref",
            :bookmark    = bookmarkName,
            :displayText = displayText
        ]
        aContent + item
        return self

    func addBookmark bookmarkName
        /*
            Insert a zero-width bookmark anchor at the current position
            (inside the next paragraph).  Use this to bookmark a spot that
            already has its own addParagraph() call and you just want the
            anchor placed there.
        */
        item = [
            :type     = "bookmarkanchor",
            :bookmark = bookmarkName
        ]
        aContent + item
        return self

    # ========================================================================
    # Tab Stops
    # ========================================================================

    func addTabbedParagraph segments, options
        /*
            Add a paragraph whose runs are separated by tab characters,
            with custom tab stop positions.

            segments : list of text strings, one per tab-stop column.
                       They are joined with \t characters.
            options  : standard paragraph options PLUS:
                :tabStops — list of tab stop definitions, each a list:
                    [:pos, :align, :leader]
                    :pos    — position in twips from left margin (720 = 0.5in)
                    :align  — "left" | "center" | "right" | "decimal" | "bar"
                    :leader — "none" | "dot" | "hyphen" | "underscore" (default "none")
                    Example: [[:pos=4320, :align="right", :leader="dot"]]

            Example — TOC-style entry:
                doc.addTabbedParagraph(
                    ["Introduction", "1"],
                    [:tabStops = [[:pos=8640, :align="right", :leader="dot"]]]
                )
        */
        if !isList(options) options = [] ok
        item = [
            :type     = "tabbedparagraph",
            :segments = segments,
            :options  = options
        ]
        aContent + item
        return self

    # ========================================================================
    # Line Numbers
    # ========================================================================

    func enableLineNumbers options
        /*
            Enable automatic line numbering for the document.
            options (all optional):
                :start    — first line number (default 1)
                :step     — show number every N lines (default 1, use 5 for every 5th)
                :distance — gap from text in twips (default 360 = 0.25 inch)
                :restart  — "newPage" | "newSection" | "continuous" (default "newPage")
        */
        bLineNumbers = true
        if isList(options)
            if options[:start]    != NULL  nLineNumberStart    = options[:start]    ok
            if options[:step]     != NULL  nLineNumberStep     = options[:step]     ok
            if options[:distance] != NULL  nLineNumberDistance = options[:distance] ok
            if options[:restart]  != NULL  cLineNumberRestart  = options[:restart]  ok
        ok
        return self

    func disableLineNumbers
        bLineNumbers = false
        return self

    func suppressLineNumbers options
        /*
            Suppress line numbering on a specific paragraph by passing
            [:suppressLineNumbers = true] in that paragraph's options.
            This method is just documentation — the option is handled
            inside generateParagraphProperties().
        */
        return self

    # ========================================================================
    # Comments (Annotations)
    # ========================================================================

    func addCommentedParagraph text, commentText, options
        /*
            Add a paragraph with an annotation comment attached to it.
            text        : the paragraph text
            commentText : the reviewer comment visible in Word's comment pane
            options     : standard paragraph options; also accepts:
                :commentAuthor — author name shown in comment balloon (default "Author")
                :commentDate   — ISO date string, e.g. "2025-01-15T10:30:00Z"
                                 defaults to a fixed placeholder date
        */
        if !isList(options) options = [] ok
        author = options[:commentAuthor]
        if author = NULL author = "Author" ok
        cDate = options[:commentDate]
        if cDate = NULL cDate = "2025-01-01T00:00:00Z" ok
        nCommentId++
        cmId = nCommentId
        aComments + [:id = cmId, :author = author, :date = cDate, :text = commentText]
        options[:_commentId] = cmId
        return addParagraph(text, options)

    func addComment text, commentText, commentAuthor
        /*
            Convenience alias: add a paragraph with a comment.
            commentAuthor is optional (defaults to "Author").
        */
        opts = []
        if commentAuthor != NULL opts[:commentAuthor] = commentAuthor ok
        return addCommentedParagraph(text, commentText, opts)

    # ========================================================================
    # SEQ Captions (auto-numbered figures and tables)
    # ========================================================================

    func addFigureCaption text
        /*
            Add a centered "Figure N — text" caption below an image.
            The SEQ field auto-increments in Word; a local counter tracks
            the display number for the placeholder.
            text : caption text after the figure number, e.g. "Ring architecture"
        */
        nFigureSeq++
        aFigureCaptions + text
        aSeqCaptions + [:label="Figure", :num=nFigureSeq, :text=text]
        item = [
            :type    = "seqcaption",
            :seqType = "Figure",
            :seqNum  = nFigureSeq,
            :text    = text
        ]
        aContent + item
        return self

    func addTableCaption text
        /*
            Add a centered "Table N — text" caption above or below a table.
            text : caption text after the table number
        */
        nTableSeq++
        aTableCaptions + text
        aSeqCaptions + [:label="Table", :num=nTableSeq, :text=text]
        item = [
            :type    = "seqcaption",
            :seqType = "Table",
            :seqNum  = nTableSeq,
            :text    = text
        ]
        aContent + item
        return self

    # ========================================================================
    # Table of Figures / Table of Tables
    # ========================================================================

    func addTableOfFigures title
        /*
            Insert an auto-generated Table of Figures.
            Word collects all "Figure N — text" SEQ captions and builds
            a numbered list with page references, exactly like a TOC.
            title : heading shown above the list (e.g. "List of Figures")
                    pass NULL to omit the heading
            Update with F9 in Word to populate page numbers.
        */
        if title = NULL  title = ""  ok
        item = [
            :type     = "tof",
            :seqType  = "Figure",
            :title    = title
        ]
        aContent + item
        return self

    func addTableOfTables title
        /*
            Insert an auto-generated Table of Tables.
            Word collects all "Table N — text" SEQ captions.
            title : heading shown above the list (e.g. "List of Tables")
                    pass NULL to omit the heading
        */
        if title = NULL  title = ""  ok
        item = [
            :type     = "tof",
            :seqType  = "Table",
            :title    = title
        ]
        aContent + item
        return self

    func addTableOfSeq seqLabel, title
        /*
            Generic version: insert a table of any SEQ caption type.
            seqLabel : the label used in addFigureCaption / addTableCaption,
                       e.g. "Figure", "Table", "Equation", "Chart"
            title    : heading text, or NULL to omit
        */
        if title = NULL  title = ""  ok
        item = [
            :type     = "tof",
            :seqType  = seqLabel,
            :title    = title
        ]
        aContent + item
        return self

    # ========================================================================
    # Mail Merge Fields
    # ========================================================================

    func addMergeFieldParagraph fields, options
        /*
            Add a paragraph containing one or more MERGEFIELD placeholders.
            fields  : a list of strings — each is a merge field name.
                      Literal text between fields can be included using
                      the :prefix / :suffix in individual field specs,
                      OR by passing a mixed list where strings are literal
                      text and lists are field specs:
                        ["Dear ", "FirstName", " ", "LastName", ","]
            options : standard paragraph options (bold, align, etc.)

            Simple example — single-field paragraph:
                doc.addMergeFieldParagraph(["FullName"], [:bold = true])

            Multi-field with literal text:
                doc.addMergeFieldParagraph(
                    ["Dear ", [:field="FirstName"], " ", [:field="LastName"], ","],
                    NULL
                )
        */
        if !isList(options) options = [] ok
        item = [
            :type    = "mergeparagraph",
            :fields  = fields,
            :options = options
        ]
        aContent + item
        return self

    func addMergeTemplate lines
        /*
            Add a series of merge-field paragraphs from a template list.
            lines : list of lists — each inner list is passed as fields
                    to addMergeFieldParagraph with NULL options.
        */
        for line in lines
            addMergeFieldParagraph(line, NULL)
        next
        return self

    # ========================================================================
    # Mail Merge Template Engine
    # ========================================================================

    func setMergeTemplate templateDef
        /*
            Register a document template for repeated mail-merge output.
            templateDef is a list of items. Each item is one of:

            String:
              "Dear {{FirstName}} {{LastName}},"
              — rendered as a plain paragraph with {{FIELD}} placeholders filled.

            Associative list:
              [:type = "heading",    :level = 2,   :text = "Invoice #{{InvoiceNum}}"]
              [:type = "paragraph",  :text  = "...", :options = [:bold = true]]
              [:type = "table",      :data  = [...], :options = [...]]
                  table :data cells may contain {{FIELD}} tokens.
              [:type = "pagebreak"]
              [:type = "emptyline"]

            Call mergeRecord(data) or mergeAll(dataList) to produce output.
            data is an associative list: [:FirstName = "Alice", :LastName = "Smith"]
        */
        aMergeTemplate = templateDef
        return self

    func clearMergeTemplate
        /*  Remove the registered merge template.  */
        aMergeTemplate = []
        return self

    func mergeFillString str, data
        /*
            Replace all {{FIELDNAME}} tokens in str with
            the corresponding value from data.
            Matching is case-insensitive: {{FirstName}}, {{FIRSTNAME}},
            and {{firstname}} all match a key stored as "firstname".
            (Ring lowercases associative-list keys, so case-insensitive
            search is required for tokens written in mixed case.)
            Unknown fields are left as {{FIELDNAME}}.
            Returns the filled string.
        */
        result = "" + str
        if !isList(data)  return result  ok
        dataLen = len(data)
        for i = 1 to dataLen
            key   = data[i][1]
            val   = "" + data[i][2]
            token = "{{" + key + "}}"
            # Case-insensitive replace: search on lowered copy, replace in original
            lowerResult = lower(result)
            lowerToken  = lower(token)
            tLen        = len(lowerToken)
            pos = substr(lowerResult, lowerToken)
            while pos > 0
                result      = left(result, pos-1) + val + substr(result, pos + tLen)
                lowerResult = lower(result)
                pos         = substr(lowerResult, lowerToken)
            end
        next
        return result

    func mergeFillTable tableData, data
        /*
            Walk a table data array and fill {{FIELD}} tokens in
            every cell.  Plain strings are filled with mergeFillString().
            WordCell objects are cloned and filled with mergeFillWordCell()
            so all run-level formatting (bold, color, italic, etc.) is
            preserved while their text content is substituted.
            Returns a new (filled) copy of the table data.
        */
        filled = []
        for rowIn in tableData
            filledRow = []
            for cellIn in rowIn
                if isObject(cellIn)
                    filledRow + mergeFillWordCell(cellIn, data)
                else
                    filledRow + mergeFillString("" + cellIn, data)
                ok
            next
            filled + filledRow
        next
        return filled

    func mergeFillWordCell cellObj, data
        /*
            Clone a WordCell, filling {{FIELD}} tokens in every
            text run.  Cell formatting (bgColor, align, bold, color per run,
            etc.) is preserved exactly.  Returns the filled clone.
        */
        clone = new WordCell()
        # Copy cell-level properties
        clone.cAlign        = cellObj.cAlign
        clone.cBgColor      = cellObj.cBgColor
        clone.cVerticalAlign = cellObj.cVerticalAlign
        clone.nColSpan      = cellObj.nColSpan
        clone.nRowSpan      = cellObj.nRowSpan
        clone.cMerge        = cellObj.cMerge
        clone.cBorderColor  = cellObj.cBorderColor
        clone.cBorderStyle  = cellObj.cBorderStyle
        clone.aBorderTop    = cellObj.aBorderTop
        clone.aBorderBottom = cellObj.aBorderBottom
        clone.aBorderLeft   = cellObj.aBorderLeft
        clone.aBorderRight  = cellObj.aBorderRight
        clone.aBorderInsideH = cellObj.aBorderInsideH
        clone.aBorderInsideV = cellObj.aBorderInsideV
        clone.cTextDir       = cellObj.cTextDir
        clone.nPaddingTop    = cellObj.nPaddingTop
        clone.nPaddingBottom = cellObj.nPaddingBottom
        clone.nPaddingLeft   = cellObj.nPaddingLeft
        clone.nPaddingRight  = cellObj.nPaddingRight
        clone.nWidth         = cellObj.nWidth
        # Copy and fill each run
        runCount = cellObj.nRunCount
        for ri = 1 to runCount
            rawText = "" + cellObj.aRunTexts[ri]
            filledText = mergeFillString(rawText, data)
            clone.aRunTexts + filledText
            clone.nRunCount  = clone.nRunCount + 1
            clone.aRunBold      + cellObj.aRunBold[ri]
            clone.aRunItalic    + cellObj.aRunItalic[ri]
            clone.aRunUnderline + cellObj.aRunUnderline[ri]
            clone.aRunStrike    + cellObj.aRunStrike[ri]
            clone.aRunFont      + cellObj.aRunFont[ri]
            clone.aRunSize      + cellObj.aRunSize[ri]
            clone.aRunColor     + cellObj.aRunColor[ri]
            clone.aRunHighlight + cellObj.aRunHighlight[ri]
        next
        return clone

    func mergeRecord data
        /*
            Render the registered template once, substituting data
            into every {{FIELD}} token, and append the result to the document.
            data : associative list  [:FirstName = "Alice", :Salary = "75000"]
        */
        if !isList(aMergeTemplate) or len(aMergeTemplate) = 0
            ? "RingWordLib WARNING: mergeRecord() called but no template is registered."
            ? "                     Call setMergeTemplate() first."
            return self
        ok

        tmplLen = len(aMergeTemplate)
        for i = 1 to tmplLen
            item = aMergeTemplate[i]
            if isString(item)
                # Plain string — fill tokens and add as paragraph
                addParagraph(mergeFillString(item, data), NULL)
            elseif isList(item)
                itemType = item[:type]
                if itemType = NULL  itemType = ""  ok

                if itemType = "heading"
                    level = item[:level]
                    text  = item[:text]
                    if level = NULL  level = 1  ok
                    if text  = NULL  text  = "" ok
                    addHeading(mergeFillString(text, data), level)

                elseif itemType = "paragraph"
                    text = item[:text]
                    if text = NULL  text = ""  ok
                    opts = item[:options]
                    if !isList(opts)  opts = []  ok
                    addParagraph(mergeFillString(text, data), opts)

                elseif itemType = "table"
                    tblData = item[:data]
                    tblOpts = item[:options]
                    if !isList(tblOpts)  tblOpts = []  ok
                    if isList(tblData)
                        addTable(mergeFillTable(tblData, data), tblOpts)
                    ok

                elseif itemType = "pagebreak"
                    addPageBreak()

                elseif itemType = "emptyline"
                    addEmptyParagraph()

                ok
            ok
        next
        return self

    func mergeAll dataList, separator
        /*
            Render the registered template for every record in
            dataList, inserting separator between records.
            separator : "pagebreak" (default) | "emptyline" | "line" | "none"

            dataList : list of associative lists
              [
                [:Name = "Alice", :Score = "92"],
                [:Name = "Bob",   :Score = "78"]
              ]
        */
        if !isList(dataList) or len(dataList) = 0  return self  ok
        if separator = NULL  separator = "pagebreak"  ok

        listLen = len(dataList)
        for i = 1 to listLen
            mergeRecord(dataList[i])
            if i < listLen
                if separator = "pagebreak"
                    addPageBreak()
                elseif separator = "emptyline"
                    addEmptyParagraph()
                    addEmptyParagraph()
                elseif separator = "line"
                    addHorizontalLine()
                ok
            ok
        next
        return self

    # ========================================================================
    # Content Controls (Structured Document Tags)
    # ========================================================================

    func addCheckbox label, bChecked
        /*
            Insert a Word checkbox content control followed by a label.
            label    : text displayed to the right of the checkbox
            bChecked : true = checked, false = unchecked
        */
        sdtId = nSdtId  nSdtId++
        checkedChar   = "&#10003;"   # ✓ (checked symbol)
        uncheckedChar = "&#9744;"    # ☐ (unchecked box symbol)
        item = [
            :type      = "sdt_checkbox",
            :id        = sdtId,
            :label     = label,
            :checked   = bChecked,
            :checkedChar   = checkedChar,
            :uncheckedChar = uncheckedChar
        ]
        aContent + item
        return self

    func addDropdown label, choices, defaultChoice
        /*
            Insert a dropdown list content control.
            label         : heading paragraph above the control
            choices       : list of strings shown in the dropdown
            defaultChoice : initial displayed value (shown before selection)
                            pass NULL to show the first choice
        */
        sdtId = nSdtId  nSdtId++
        if defaultChoice = NULL and isList(choices) and len(choices) > 0
            defaultChoice = choices[1]
        ok
        item = [
            :type    = "sdt_dropdown",
            :id      = sdtId,
            :label   = label,
            :choices = choices,
            :default = defaultChoice
        ]
        aContent + item
        return self

    func addTextInput label, defaultText, placeholder
        /*
            Insert a plain-text content control (text input field).
            label       : heading paragraph above the control
            defaultText : pre-filled text value
            placeholder : grey hint text shown when the field is empty
                          pass NULL to use "Click here to enter text."
        */
        sdtId = nSdtId  nSdtId++
        if placeholder = NULL placeholder = "Click here to enter text." ok
        item = [
            :type        = "sdt_text",
            :id          = sdtId,
            :label       = label,
            :default     = defaultText,
            :placeholder = placeholder
        ]
        aContent + item
        return self

    # ========================================================================
    # Drawing Shapes (DrawingML)
    # ========================================================================

    func addShape shapeType, options
        /*
            Insert an inline DrawingML shape.
            shapeType : "rect"    — rectangle / rounded rectangle
                        "ellipse" — circle or oval
                        "line"    — horizontal rule
                        "diamond" — diamond / rhombus
                        "triangle"— right-pointing triangle
            options (all optional):
                :width      — width  in cm (default 5)
                :height     — height in cm (default 3); ignored for "line"
                :fillColor  — fill hex color or name (default "4472C4")
                :lineColor  — border/stroke hex color (default "2E4E7E")
                :lineWidth  — border width in pt (default 1)
                :text       — text inside the shape
                :textColor  — text hex color (default "FFFFFF")
                :textBold   — true/false (default true)
                :textSize   — font size in pt (default 11)
                :align      — paragraph alignment "left","center","right" (default "center")
                :rounded    — true = rounded corners for rect (default false)
                :noFill     — true = transparent fill
                :noBorder   — true = no border line
        */
        if !isList(options) options = [] ok
        shpId = nShapeId  nShapeId++
        item = [
            :type      = "drawshape",
            :shapeType = shapeType,
            :id        = shpId,
            :options   = options
        ]
        aContent + item
        return self

    func addRect options
        return addShape("rect", options)

    func addEllipse options
        return addShape("ellipse", options)

    func addLine options
        return addShape("line", options)

    func addDiamond options
        return addShape("diamond", options)

    # ========================================================================
    # Document Themes
    # ========================================================================

    func setTheme themeName
        /*
            Apply a built-in colour theme to the document.
            Available themes:
                "Office"   — classic blue/orange Word theme
                "Blue"     — all-blue professional theme
                "Dark"     — dark navy/slate corporate theme
                "Green"    — emerald green theme
                "Red"      — crimson red theme
                "Purple"   — violet/purple theme
                "Teal"     — teal/cyan theme
                "Warm"     — warm orange/gold theme
            The theme affects:
              • Heading colours in styles.xml
              • Accent colours used by tables (headerBgColor uses "accent1" etc.)
              • A theme1.xml file is written into word/theme/
        */
        cThemeName = themeName
        # Define color palettes: [dk1, lt1, dk2, lt2, acc1, acc2, acc3, acc4, acc5, acc6]
        switch themeName
            on "Office"
                aThemeColors = ["000000","FFFFFF","44546A","E7E6E6",
                                "4472C4","ED7D31","A9D18E","FFC000","5A96CB","70AD47"]
            on "Blue"
                aThemeColors = ["000000","FFFFFF","1F3864","DEE9F7",
                                "1E6BB5","2E75B6","4A86C8","9DC3E6","4472C4","70ACE4"]
            on "Dark"
                aThemeColors = ["FFFFFF","1F2226","C8D0D8","2E3440",
                                "88C0D0","81A1C1","5E81AC","B48EAD","A3BE8C","EBCB8B"]
            on "Green"
                aThemeColors = ["000000","FFFFFF","375623","E2EFDA",
                                "375623","70AD47","A9D18E","548235","C6E0B4","E2EFDA"]
            on "Red"
                aThemeColors = ["000000","FFFFFF","7B0000","FDECEA",
                                "C00000","FF0000","E74C3C","C0392B","E57373","FFCDD2"]
            on "Purple"
                aThemeColors = ["000000","FFFFFF","3B0764","EDE7F6",
                                "7030A0","9B59B6","8E44AD","AB47BC","CE93D8","E1BEE7"]
            on "Teal"
                aThemeColors = ["000000","FFFFFF","004D40","E0F2F1",
                                "00796B","009688","26A69A","4DB6AC","80CBC4","B2DFDB"]
            on "Warm"
                aThemeColors = ["000000","FFFFFF","6D3B00","FFF3E0",
                                "E65100","F57C00","FB8C00","FFA726","FFCC02","FFE082"]
            other
                # Default to Office if unknown
                cThemeName = "Office"
                aThemeColors = ["000000","FFFFFF","44546A","E7E6E6",
                                "4472C4","ED7D31","A9D18E","FFC000","5A96CB","70AD47"]
        off
        return self

    func getThemeAccent n
        /*
            Return the hex color of theme accent N (1–6).
            Used internally by writeStyles to colour headings.
            Returns "" if no theme is set.
        */
        if !isList(aThemeColors) or len(aThemeColors) < (4 + n)
            return ""
        ok
        return aThemeColors[4 + n]

    func setThemeColors accent1, accent2
        /*
            Set the two most visible heading accent colors directly, without
            choosing a full named theme.  Used by WordReader.toWriter() to
            preserve custom/non-standard theme colors from the source document.
            accent1 is used for Heading1/H2, accent2 for Heading3/H4.
        */
        if !isList(aThemeColors) or len(aThemeColors) < 6
            aThemeColors = ["000000","FFFFFF","44546A","E7E6E6",
                            accent1, accent2, "A9D18E","FFC000","5A96CB","70AD47"]
        else
            aThemeColors[5] = accent1
            aThemeColors[6] = accent2
        ok
        return self

    # ========================================================================
    # Paragraph Shading (standalone helper)
    # ========================================================================

    func addShadedParagraph text, bgColor, options
        /*
            Add a paragraph with a colored background shading.
            bgColor : fill color hex or name (e.g. "FFD700", "lightblue")
            options : same as addParagraph options (bold, italic, align, etc.)
        */
        if !isList(options)
            options = []
        ok
        options[:bgColor] = wordColorToHex(bgColor)
        return addParagraph(text, options)

    # ========================================================================
    # Text Boxes
    # ========================================================================

    func addTextBox text, options
        /*
            Add an absolutely-positioned floating text box.
            text   : the text content of the box
            options: list with keys:
                :x           - horizontal position from left margin in cm (default 2)
                :y           - vertical position from top margin in cm  (default 5)
                :width       - box width in cm  (default 6)
                :height      - box height in cm (default 3)
                :align       - text alignment inside box: left/center/right (default center)
                :bold        - true/false
                :italic      - true/false
                :font        - font name
                :size        - font size in points
                :color       - text color (name or hex)
                :bgColor     - fill color (name or hex, default "FFFFFF")
                :borderColor - border color (name or hex, default "4472C4")
                :borderSize  - border width in eighths of a point (default 8 = 1pt)
                :noFill      - true: transparent / no fill
                :noBorder    - true: no visible border
        */
        if !isList(options)
            options = []
        ok
        nTextBoxId = nTextBoxId + 1
        # Extract option values into locals (Ring cannot use subscripts inside list literals)
        tbX           = options[:x]
        tbY           = options[:y]
        tbWidth       = options[:width]
        tbHeight      = options[:height]
        tbAlign       = options[:align]
        tbBold        = options[:bold]
        tbItalic      = options[:italic]
        tbFont        = options[:font]
        tbSize        = options[:size]
        tbColor       = options[:color]
        tbBgColor     = options[:bgColor]
        tbBorderColor = options[:borderColor]
        tbBorderSize  = options[:borderSize]
        tbNoFill      = options[:noFill]
        tbNoBorder    = options[:noBorder]
        tb = [
            :type        = "textbox",
            :text        = text,
            :x           = tbX,
            :y           = tbY,
            :width       = tbWidth,
            :height      = tbHeight,
            :align       = tbAlign,
            :bold        = tbBold,
            :italic      = tbItalic,
            :font        = tbFont,
            :size        = tbSize,
            :color       = tbColor,
            :bgColor     = tbBgColor,
            :borderColor = tbBorderColor,
            :borderSize  = tbBorderSize,
            :noFill      = tbNoFill,
            :noBorder    = tbNoBorder,
            :tbId        = nTextBoxId
        ]
        aContent + tb
        return self

    # ========================================================================
    # Header and Footer
    # ========================================================================

    func setHeader text
        /*
            Set the document header text
        */
        cHeaderText = text
        return self
    
    func setFooter text
        /*
            Set the document footer text
        */
        cFooterText = text
        return self
    
    func showPageNumbers align
        /*
            Enable page numbers in footer
            align: "left", "center", "right" (default: center)
        */
        bShowPageNumbers = true
        if align != NULL
            cPageNumberAlign = align
        ok
        return self
    
    func setHeaderFooter headerText, footerText
        /*
            Set both header and footer at once
        */
        cHeaderText = headerText
        cFooterText = footerText
        return self
    
    func setWatermark text, options
        /*
            Enable a diagonal text watermark on every page.
            text   : watermark string  (e.g. "CONFIDENTIAL", "DRAFT")
            options: optional list with keys:
                     :color    – fill color name or hex  (default "C0C0C0")
                     :opacity  – 0-100                  (default 50)
                     :rotation – degrees, negative=CW   (default -45)
                     :font     – font family             (default "Arial")
                     :size     – font size in points     (default 72)
        */
        bWatermark = true
        if text != NULL and len(text) > 0
            cWatermarkText = text
        ok
        if isList(options)
            if options[:color] != NULL
                cWatermarkColor = wordColorToHex(options[:color])
            ok
            if options[:opacity] != NULL
                nWatermarkOpacity = options[:opacity]
            ok
            if options[:rotation] != NULL
                nWatermarkRotation = options[:rotation]
            ok
            if options[:font] != NULL
                cWatermarkFont = options[:font]
            ok
            if options[:size] != NULL
                nWatermarkSize = options[:size]
            ok
        ok
        return self

    func setWatermarkText text
        /*
            Shortcut: enable a watermark with default grey styling.
        */
        return setWatermark(text, NULL)

    func removeWatermark
        /*
            Disable the text watermark.
        */
        bWatermark = false
        return self

    func setImageWatermark imagePath, options
        /*
            Enable an image watermark on every page.
            imagePath: path to image file (png, jpg, jpeg, bmp, gif)
            options  : optional list with keys:
                       :width   – display width in cm   (default 15)
                       :height  – display height in cm  (default 15)
                       :opacity – 0-100                 (default 50)
        */
        if !fexists(imagePath)
            ? "Warning: Watermark image file not found: " + imagePath
            return self
        ok
        bImgWatermark = true
        cImgWatermarkPath = imagePath
        if isList(options)
            if options[:width] != NULL
                nImgWatermarkWidth = options[:width]
            ok
            if options[:height] != NULL
                nImgWatermarkHeight = options[:height]
            ok
            if options[:opacity] != NULL
                nImgWatermarkOpacity = options[:opacity]
            ok
        ok
        return self

    func removeImageWatermark
        /*
            Disable the image watermark.
        */
        bImgWatermark = false
        cImgWatermarkPath = ""
        return self
    
    # ========================================================================
    # Text and Paragraph Methods
    # ========================================================================
    
    func addParagraph text, options
        /*
            Add a paragraph with optional formatting
            options is a list with optional keys:
            - bold, italic, underline, strike
            - font, size, color
            - align (left, center, right, both)
        */
        if !isList(options)
            options = []
        ok
        
        para = [
            :type = "paragraph",
            :text = text,
            :options = options
        ]
        aContent + para
        return self
    
    func addText text, options
        # Alias for addParagraph
        return addParagraph(text, options)
    
    func addHeading text, level
        /*
            Add a heading (level 1-9)
        */
        if level = NULL level = 1 ok
        if level < 1 level = 1 ok
        if level > 9 level = 9 ok
        
        heading = [
            :type = "heading",
            :text = text,
            :level = level
        ]
        aContent + heading
        return self
    
    func addTitle text
        return addHeading(text, 1)
    
    func addSubtitle text
        return addHeading(text, 2)
    
    # ========================================================================
    # Formatting Helpers
    # ========================================================================
    
    func bold text
        return addParagraph(text, [:bold = true])
    
    func italic text
        return addParagraph(text, [:italic = true])
    
    func underline text
        return addParagraph(text, [:underline = true])
    
    func colored text, color
        return addParagraph(text, [:color = color])
    
    func styled text, fontName, fontSize, color
        options = []
        if fontName != NULL options[:font] = fontName ok
        if fontSize != NULL options[:size] = fontSize ok
        if color != NULL options[:color] = color ok
        return addParagraph(text, options)
    
    func centered text
        return addParagraph(text, [:align = WORD_ALIGN_CENTER])
    
    func rightAligned text
        return addParagraph(text, [:align = WORD_ALIGN_RIGHT])
    
    func justified text
        return addParagraph(text, [:align = WORD_ALIGN_JUSTIFY])
    
    # ========================================================================
    # Rich Text (Multiple Runs in One Paragraph)
    # ========================================================================
    
    func addRichParagraph runs, paragraphOptions
        /*
            Add a paragraph with multiple formatted text runs
            runs is a list of [text, options] pairs
            Example:
            doc.addRichParagraph([
                ["Hello ", [:bold = true]],
                ["World", [:italic = true, :color = "red"]]
            ])
        */
        if paragraphOptions = NULL
            paragraphOptions = []
        ok
        
        para = [
            :type = "richparagraph",
            :runs = runs,
            :options = paragraphOptions
        ]
        aContent + para
        return self
    
    # ========================================================================
    # Lists
    # ========================================================================
    
    func addBulletList items
        /*
            Add a bullet list
            items is a list of strings
        */
        listDef = [
            :type = "list",
            :listType = WORD_LIST_BULLET,
            :items = items,
            :listId = nListId
        ]
        aContent + listDef
        
        # Add numbering definition
        aNumbering + [:id = nListId, :type = WORD_LIST_BULLET]
        nListId++
        
        return self
    
    func addNumberedList items
        /*
            Add a numbered list
            items is a list of strings
        */
        listDef = [
            :type = "list",
            :listType = WORD_LIST_NUMBER,
            :items = items,
            :listId = nListId
        ]
        aContent + listDef
        
        aNumbering + [:id = nListId, :type = WORD_LIST_NUMBER]
        nListId++
        
        return self
    
    # ========================================================================
    # Tables
    # ========================================================================
    
    func addTable data, options
        /*
            Add a table
            data is a 2D list: [[row1col1, row1col2], [row2col1, row2col2]]
            options can include:
            - headerRow: true/false (first row is header)
            - borderStyle: "single", "double", "dashed", "dotted", "none"
            - borderColor: hex color for borders (e.g., "000000")
            - borderSize: border thickness (default 4, in eighths of a point)
            - headerBgColor: background color for header row (e.g., "4472C4")
            - evenRowBgColor: background color for even rows (zebra striping)
            - cellPadding: cell padding in twips
            - conditionalRules: list of rules for cell-level conditional formatting
                Each rule is an associative list with keys:
                  :col        column index (1-based); 0 = all columns
                  :condition  "lt"|"lte"|"gt"|"gte"|"eq"|"neq"|"between"|"contains"
                  :value      threshold (number or string)
                  :value2     upper bound for "between"
                  :bgColor    hex background colour when rule matches
                  :textColor  hex text colour when rule matches (optional)
                  :bold       true/false when rule matches (optional)
                Rules are evaluated in order; later rules override earlier ones.
                Header row is never subject to conditional formatting.
                Example:
                  :conditionalRules = [
                    [:col=2, :condition="lt", :value=0,  :bgColor="FDECEA", :textColor="C00000"],
                    [:col=2, :condition="gte",:value=90, :bgColor="E2EFDA", :textColor="375623"],
                    [:col=0, :condition="contains", :value="FAIL", :bgColor="FFF2CC", :bold=true]
                  ]
        */
        if !isList(options)
            options = []
        ok
        
        tbl = [
            :type = "table",
            :data = data,
            :options = options
        ]
        aContent + tbl
        return self
    
    func addSimpleTable data
        # Add a simple table with borders and header row
        return addTable(data, [:headerRow = true, :borderStyle = "single"])
    
    func addStyledTable data, headerBgColor, evenRowBgColor
        /*
            Add a styled table with colored header and zebra striping
        */
        if headerBgColor = NULL headerBgColor = "4472C4" ok
        if evenRowBgColor = NULL evenRowBgColor = "D9E2F3" ok
        return addTable(data, [
            :headerRow = true, 
            :borderStyle = "single",
            :headerBgColor = headerBgColor,
            :evenRowBgColor = evenRowBgColor
        ])
    
    # ========================================================================
    # Page Elements
    # ========================================================================
    
    func addPageBreak
        aContent + [:type = "pagebreak"]
        return self
    
    func addLineBreak
        aContent + [:type = "linebreak"]
        return self
    
    func addEmptyParagraph
        return addParagraph("", NULL)
    
    func addHorizontalLine
        # Add a paragraph with a bottom border to simulate a horizontal line
        aContent + [:type = "horizontalline"]
        return self
    
    func addColumnBreak
        # Force content to next column (in multi-column layout)
        aContent + [:type = "columnbreak"]
        return self
    
    func addSectionBreak breakType, options
        /*
            Add a section break.
            breakType: "nextPage", "continuous", "evenPage", "oddPage"
            options (optional list):
              :numColumns    - number of columns in the section (default 1)
              :columnSpaceCm - space between columns in cm (default 1.25 cm = 720 twips)
        */
        if breakType = NULL breakType = "nextPage" ok
        item = [:type = "sectionbreak", :breakType = breakType]
        if isList(options)
            nc = options[:numColumns]
            cs = options[:columnSpaceCm]
            if nc != NULL and isNumber(nc) and nc > 1
                item[:numColumns] = nc
            ok
            if cs != NULL and isNumber(cs) and cs > 0
                item[:columnSpaceTwips] = number(cs * 567)
            ok
        ok
        aContent + item
        return self

    func addLandscapeStart
        /*
            Start a landscape-oriented section.
            Inserts a section break and changes orientation to landscape.
            Call addLandscapeEnd() to return to portrait.
        */
        aContent + [:type = "landscapestart"]
        return self

    func addLandscapeEnd
        /*
            End a landscape-oriented section, returning to portrait.
        */
        aContent + [:type = "landscapeend"]
        return self
    
    func addBlockQuote text
        /*
            Add an indented block quote
        */
        aContent + [:type = "blockquote", :text = text]
        return self
    
    func addCaption text
        /*
            Add a centered, italic caption (for figures/tables)
        */
        aContent + [:type = "caption", :text = text]
        return self
    
    func addAbstract text
        /*
            Add an abstract section (common in research papers)
        */
        aContent + [:type = "abstract", :text = text]
        return self
    
    func addKeywords keywords
        /*
            Add keywords section (for research papers)
            keywords can be a string or list of strings
        */
        if isList(keywords)
            keywordStr = ""
            kwLen = len(keywords)
            for i = 1 to kwLen
                if i > 1 keywordStr += ", " ok
                keywordStr += keywords[i]
            next
        else
            keywordStr = keywords
        ok
        aContent + [:type = "keywords", :text = keywordStr]
        return self
    
    func registerFootnote footnoteContent
        /*
            Register a footnote and return its numeric ID.
            footnoteContent: plain string  — simple note text
                             list of [text, opts] pairs — rich formatted runs
                             e.g. [["Normal ", []], ["bold part", [:bold=true]], [" end.", []]]
            Use the returned ID in a rich paragraph run via [:footnoteId = id].
        */
        nFootnoteId = nFootnoteId + 1
        aFootnotes + [:id = nFootnoteId, :content = footnoteContent]
        return nFootnoteId

    func registerEndnote endnoteContent
        /*
            Register an endnote and return its numeric ID.
            endnoteContent: plain string or list of [text, opts] run pairs (same as registerFootnote).
            Use the returned ID in a rich paragraph run via [:endnoteId = id].
        */
        nEndnoteId = nEndnoteId + 1
        aEndnotes + [:id = nEndnoteId, :content = endnoteContent]
        return nEndnoteId

    func addFootnote text, footnoteContent, options
        /*
            Add a paragraph whose text is followed by a footnote superscript marker.
            text           : paragraph text before the marker
            footnoteContent: plain string body, or list of [text, opts] run pairs for
                             rich formatting (bold, italic, color, size, etc.) inside the note
            options        : optional run/paragraph formatting for the main paragraph
        */
        fnId = registerFootnote(footnoteContent)
        if !isList(options) options = [] ok
        runs = [
            [text, options],
            ["", [:footnoteId = fnId]]
        ]
        return addRichParagraph(runs, options)

    func addEndnote text, endnoteContent, options
        /*
            Add a paragraph whose text is followed by an endnote superscript marker.
            text          : paragraph text before the marker
            endnoteContent: plain string body, or list of [text, opts] run pairs for
                            rich formatting inside the note
            options       : optional run/paragraph formatting for the main paragraph
        */
        enId = registerEndnote(endnoteContent)
        if !isList(options) options = [] ok
        runs = [
            [text, options],
            ["", [:endnoteId = enId]]
        ]
        return addRichParagraph(runs, options)

    func addNestedList items, listType, startAt
        /*
            Add a multi-level (nested) list.
            items   : list of [text, level] pairs, level is 0-based (0 = top)
                      e.g. [["Item 1", 0], ["Sub item", 1], ["Item 2", 0]]
            listType: WORD_LIST_BULLET or WORD_LIST_NUMBER
            startAt : starting number for level 0 (default 1)
        */
        if listType = NULL  listType = WORD_LIST_BULLET  ok
        if startAt  = NULL  startAt  = 1                 ok
        if !isNumber(startAt) or startAt < 1  startAt = 1  ok
        listDef = [
            :type = "nestedlist",
            :listType = listType,
            :items = items,
            :listId = nListId
        ]
        aContent + listDef
        aNumbering + [:id = nListId, :type = listType, :multilevel = true, :startAt = startAt]
        nListId++
        return self

    func addNestedNumberedListAt items, startAt
        /*  Add a numbered list starting at a custom value (e.g. restart mid-document).  */
        if startAt = NULL  startAt = 1  ok
        return addNestedList(items, WORD_LIST_NUMBER, startAt)

    func addNestedBulletList items
        /*
            Add a nested bullet list.
            items: list of [text, level] pairs
        */
        return addNestedList(items, WORD_LIST_BULLET, 1)

    func addNestedNumberedList items
        /*
            Add a nested numbered list.
            items: list of [text, level] pairs
        */
        return addNestedList(items, WORD_LIST_NUMBER, 1)
    
    func addTableOfContents title
        /*
            Add a table of contents placeholder
            Word will auto-generate it based on headings
        */
        if title = NULL title = "Table of Contents" ok
        aContent + [:type = "toc", :title = title]
        return self
    
    # ========================================================================
    # Document Fields  (DATE, AUTHOR, TITLE, FILENAME, etc.)
    # ========================================================================

    func addField fieldType, cachedValue, options
        if !isList(options)  options = []  ok
        item = [:type="field", :fieldType=fieldType,
                :cachedValue=cachedValue, :options=options]
        aContent + item
        return self

    # ========================================================================
    # Hyperlinks
    # ========================================================================

    func addHyperlink text, url
        nRelId++
        relId = "rId" + nRelId
        
        aRelationships + [
            :id = relId,
            :type = "hyperlink",
            :target = url
        ]
        
        link = [
            :type = "hyperlink",
            :text = text,
            :relId = relId
        ]
        aContent + link
        return self
    
    # ========================================================================
    # Images
    # ========================================================================
    

    func addFloatingImage imagePath, width, height, options
        /*
            Add a floating image anchored to the page (wp:anchor).
            imagePath: path to image file (png, jpg, jpeg, gif, bmp)
            width    : width in cm  (default 8)
            height   : height in cm (default 6)
            options  : list with optional keys:
              :wrapType  "wrapSquare"|"wrapTight"|"wrapThrough"|
                         "wrapTopAndBottom"|"wrapNone"  (default "wrapSquare")
              :wrapSide  "bothSides"|"left"|"right"|"largest" (default "bothSides")
              :posX      horizontal position in cm from left margin (default 1)
              :posY      vertical position in cm below paragraph (default 0)
              :distCm    uniform distance-from-text in cm (default 0.3)
        */
        if width  = NULL  width  = 8   ok
        if height = NULL  height = 6   ok
        if !isList(options)  options = []  ok

        if !fexists(imagePath)
            ? "Warning: Floating image file not found: " + imagePath
            return self
        ok
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read floating image: " + imagePath
            return self
        ok

        ext = lower(wordGetImageExtension(imagePath))
        contentType = "image/png"
        switch ext
            on "png"   contentType = "image/png"
            on "jpg"   contentType = "image/jpeg"
            on "jpeg"  contentType = "image/jpeg"
            on "gif"   contentType = "image/gif"
            on "bmp"   contentType = "image/bmp"
        off

        nRelId++
        nImageId++
        relId = "rId" + nRelId
        imageFilename = "image" + nImageId + "." + ext

        aImages + [
            :relId       = relId,
            :filename    = imageFilename,
            :data        = imageData,
            :contentType = contentType
        ]
        aRelationships + [
            :id     = relId,
            :type   = "image",
            :target = "media/" + imageFilename
        ]

        wrapType = options[:wrapType]
        if wrapType = NULL  wrapType = "wrapSquare"  ok
        wrapSide = options[:wrapSide]
        if wrapSide = NULL  wrapSide = "bothSides"   ok
        posXcm   = options[:posX]
        if posXcm  = NULL  posXcm  = 1.0  ok
        posYcm   = options[:posY]
        if posYcm  = NULL  posYcm  = 0.0  ok
        distCm   = options[:distCm]
        if distCm  = NULL  distCm  = 0.3  ok

        fltAltIn = options[:altText]
        if fltAltIn = NULL  fltAltIn = ""  ok
        fimg = [
            :type      = "floatingimage",
            :relId     = relId,
            :imageId   = nImageId,
            :width     = floor(width  * 360000),
            :height    = floor(height * 360000),
            :posX      = floor(posXcm  * 360000),
            :posY      = floor(posYcm  * 360000),
            :distEmu   = floor(distCm  * 360000),
            :wrapType  = wrapType,
            :wrapSide  = wrapSide,
            :altText   = fltAltIn
        ]
        aContent + fimg
        return self

    func addNumberedHeading text, level, numId
        /*
            Add a heading paragraph that carries outline numbering (w:numPr).
            text  : heading text
            level : 1-9
            numId : numbering definition id (usually 1 for the first numbering
                    definition in numbering.xml; defaults to 1)
            The reader will detect block[:numbered]=true on this heading.
        */
        if level = NULL  level = 1  ok
        if level < 1  level = 1  ok
        if level > 9  level = 9  ok
        if numId = NULL  numId = 1  ok
        ilvl = level - 1   # outline ilvl is 0-based

        nh = [
            :type   = "numberedheading",
            :text   = text,
            :level  = level,
            :numId  = numId,
            :ilvl   = ilvl
        ]
        aContent + nh
        return self

    func addImage imagePath, width, height
        /*
            Add an inline image (left-aligned).
            imagePath : path to the image file (png, jpg, jpeg, gif, bmp)
            width     : display width in centimetres (default 10)
            height    : display height in centimetres (default 7.5)
            For crop / altText options use addImageWithOptions().
        */
        return addImageWithOptions(imagePath, width, height, [])

    func addImageWithOptions imagePath, width, height, options
        /*
            Add an inline image with full options support.
            options list keys:
              :altText    = "description"   — accessibility alt text
              :cropLeft   = percent (0-100) — trim left edge
              :cropRight  = percent         — trim right edge
              :cropTop    = percent         — trim top edge
              :cropBottom = percent         — trim bottom edge
        */
        if !isList(options)  options = []  ok
        if width = NULL width = 10 ok
        if height = NULL height = 7.5 ok
        
        # Read image file
        if !fexists(imagePath)
            ? "Warning: Image file not found: " + imagePath
            return self
        ok
        
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read image file: " + imagePath
            return self
        ok
        
        # Determine content type from extension
        ext = lower(wordGetImageExtension(imagePath))
        contentType = "image/png"  # default
        
        switch ext
            on "png"
                contentType = "image/png"
            on "jpg"
                contentType = "image/jpeg"
            on "jpeg"
                contentType = "image/jpeg"
            on "gif"
                contentType = "image/gif"
            on "bmp"
                contentType = "image/bmp"
        off
        
        # Create relationship
        nRelId++
        nImageId++
        relId = "rId" + nRelId
        imageFilename = "image" + nImageId + "." + ext
        
        # Store image data
        aImages + [
            :relId = relId,
            :filename = imageFilename,
            :data = imageData,
            :contentType = contentType
        ]
        
        # Add relationship
        aRelationships + [
            :id = relId,
            :type = "image",
            :target = "media/" + imageFilename
        ]
        
        # Convert cm to EMUs (914400 EMUs per inch, 2.54 cm per inch)
        # EMUs = cm * 914400 / 2.54 = cm * 360000
        widthEmu = floor(width * 360000)
        heightEmu = floor(height * 360000)

        # Alt text for accessibility (wp:docPr descr attribute)
        imgAlt = options[:altText]
        if imgAlt = NULL  imgAlt = ""  ok

        # Image crop — values are percentages 0–100
        # OOXML a:srcRect uses thousandths-of-a-percent (1% = 1000 units)
        imgCropL = options[:cropLeft]
        imgCropR = options[:cropRight]
        imgCropT = options[:cropTop]
        imgCropB = options[:cropBottom]

        # Add to content
        img = [
            :type    = "image",
            :relId   = relId,
            :width   = widthEmu,
            :height  = heightEmu,
            :imageId = nImageId,
            :altText = imgAlt,
            :cropL   = imgCropL,
            :cropR   = imgCropR,
            :cropT   = imgCropT,
            :cropB   = imgCropB
        ]
        aContent + img
        return self
    
    func addImageCentered imagePath, width, height
        /*
            Add a centred inline image.
            For crop / altText options use addImageCenteredWithOptions().
        */
        return addImageCenteredWithOptions(imagePath, width, height, [])

    func addImageCenteredWithOptions imagePath, width, height, options
        /*
            Add a centred inline image with full options support.
            options list keys: :altText, :cropLeft, :cropRight, :cropTop, :cropBottom
        */
        if !isList(options)  options = []  ok
        if width = NULL width = 10 ok
        if height = NULL height = 7.5 ok

        if !fexists(imagePath)
            ? "Warning: Image file not found: " + imagePath
            return self
        ok
        
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read image file: " + imagePath
            return self
        ok
        
        ext = lower(wordGetImageExtension(imagePath))
        contentType = "image/png"
        
        switch ext
            on "png"
                contentType = "image/png"
            on "jpg"
                contentType = "image/jpeg"
            on "jpeg"
                contentType = "image/jpeg"
            on "gif"
                contentType = "image/gif"
            on "bmp"
                contentType = "image/bmp"
        off
        
        nRelId++
        nImageId++
        relId = "rId" + nRelId
        imageFilename = "image" + nImageId + "." + ext
        
        aImages + [
            :relId = relId,
            :filename = imageFilename,
            :data = imageData,
            :contentType = contentType
        ]
        
        aRelationships + [
            :id = relId,
            :type = "image",
            :target = "media/" + imageFilename
        ]
        
        widthEmu = floor(width * 360000)
        heightEmu = floor(height * 360000)

        imgAlt2 = options[:altText]
        if imgAlt2 = NULL  imgAlt2 = ""  ok

        imgCropL2 = options[:cropLeft]
        imgCropR2 = options[:cropRight]
        imgCropT2 = options[:cropTop]
        imgCropB2 = options[:cropBottom]

        img = [
            :type    = "image",
            :relId   = relId,
            :width   = widthEmu,
            :height  = heightEmu,
            :imageId = nImageId,
            :centered = true,
            :altText = imgAlt2,
            :cropL   = imgCropL2,
            :cropR   = imgCropR2,
            :cropT   = imgCropT2,
            :cropB   = imgCropB2
        ]
        aContent + img
        return self

    # =========================================================================
    # Native OOXML Chart Support
    # =========================================================================

    func addChart chartType, title, categories, series, options
        /*
            Embed a native OOXML chart — no images, no Python, pure Ring XML.

            chartType  : "column" | "bar" | "line" | "area" | "pie" | "doughnut"
            title      : chart heading string (pass "" for no title)
            categories : list of category labels  e.g. ["Q1","Q2","Q3","Q4"]
            series     : list of series, each a list with keys:
                           :name   = series name string
                           :values = list of numeric values
                           :color  = (optional) hex colour e.g. "4472C4"
                         Example:
                           [ [:name="Revenue", :values=[120,200,180,240]],
                             [:name="Costs",   :values=[80, 140,120,160]] ]
            options    : (optional) list with any of:
                           :widthCm        display width  in cm  (default 14)
                           :heightCm       display height in cm  (default 9)
                           :centered       true/false            (default true)
                           :showLegend     true/false            (default true)
                           :legendPos      "r"|"b"|"t"|"l"       (default "r")
                           :showDataLabels true/false            (default false)
                           :grouping       "clustered"|"stacked"|"percent"
                                           (bar/column/area, default "clustered")
                           :smooth         true/false line-chart smooth curves
                                           (default false)
                           :colors         list of hex overrides e.g.
                                           ["FF0000","00AA00","0000FF"]
        */
        nChartId++
        nRelId++
        nDocPrId++

        # Merge chart defaults (defaults < per-call options)
        if isList(aChartDefaults) and len(aChartDefaults) > 0
            options = mergeChartOptions(aChartDefaults, options)
        ok

        widthCm   = 14
        heightCm  = 9
        bCentered = true

        if isList(options)
            if options[:widthCm]  != NULL  widthCm  = options[:widthCm]  ok
            if options[:heightCm] != NULL  heightCm = options[:heightCm] ok
            if options[:centered] != NULL  bCentered = options[:centered] ok
        ok

        relId = "rId" + nRelId

        aCharts + [
            :chartId    = nChartId,
            :relId      = relId,
            :type       = lower(chartType),
            :title      = title,
            :categories = categories,
            :series     = series,
            :options    = options
        ]

        aRelationships + [
            :id      = relId,
            :type    = "chart",
            :chartId = nChartId
        ]

        widthEmu  = floor(widthCm  * 360000)
        heightEmu = floor(heightCm * 360000)

        aContent + [
            :type     = "chart",
            :relId    = relId,
            :chartId  = nChartId,
            :docPrId  = nDocPrId,
            :width    = widthEmu,
            :height   = heightEmu,
            :centered = bCentered
        ]
        return self

    func addColumnChart title, categories, series, options
        /*  Vertical bars (most common chart type).  */
        return addChart("column", title, categories, series, options)

    func addBarChart title, categories, series, options
        /*  Horizontal bars — good for long category labels.  */
        return addChart("bar", title, categories, series, options)

    func addLineChart title, categories, series, options
        /*  Line chart — ideal for trends over time.  */
        return addChart("line", title, categories, series, options)

    func addAreaChart title, categories, series, options
        /*  Area chart — like a line chart but with filled areas.  */
        return addChart("area", title, categories, series, options)

    func addPieChart title, categories, series, options
        /*  Pie chart — uses only the first series; categories become slices.  */
        return addChart("pie", title, categories, series, options)

    func addDoughnutChart title, categories, series, options
        return addChart("doughnut", title, categories, series, options)

    func addScatterChart title, series, options
        /*
            XY Scatter chart — both axes are numeric; no string category labels.
            series items use :xValues and :yValues instead of :values.
            Extra options: :markerStyle, :markerSize, :lines, :smooth,
                           :xAxisTitle, :yAxisTitle
            Pass NULL for categories (unused by scatter charts).
        */
        return addChart("scatter", title, NULL, series, options)

    func addBubbleChart title, series, options
        /*
            Bubble chart — like scatter with a third dimension encoded as bubble size.
            series items use :xValues, :yValues, and :sizes.
            Extra options: same as addScatterChart() plus :bubble3D (default false).
            Pass NULL for categories (unused by bubble charts).
        */
        return addChart("bubble", title, NULL, series, options)

    func setChartDefaults options
        /*
            Set default chart options applied to all subsequent charts.
            Any option explicitly provided in the per-chart call overrides the default.

            Example — define a corporate style applied to every chart:
                doc.setChartDefaults([
                    :widthCm       = 14,
                    :heightCm      = 8,
                    :legendPos     = "b",
                    :showGridlines = false,
                    :yAxisFormat   = "$#,##0",
                    :yAxisMin      = 0
                ])

            Options are merged: per-chart value wins over default.
            Call clearChartDefaults() to remove all defaults.
        */
        if isList(options)
            aChartDefaults = options
        ok
        return self

    func clearChartDefaults
        /*  Remove all chart defaults set by setChartDefaults().  */
        aChartDefaults = []
        return self

    func mergeChartOptions base, overrides
        /*
            Internal: merge two options lists.
            Keys present in overrides take priority over base.
            Returns a new merged list.
        */
        if !isList(base)       base      = []  ok
        if !isList(overrides)  overrides = []  ok
        merged = base
        # All keys in Ring associative lists are lower-cased strings internally.
        # We iterate over overrides and set each key in merged.
        overrideLen = len(overrides)
        for i = 1 to overrideLen
            key = overrides[i][1]
            val = overrides[i][2]
            merged[key] = val
        next
        return merged

    func registerImage imagePath
        /*
            Register an image with the document (for use in table cells).
            Returns a list: [:relId = "rIdX", :imageId = N]
            Does NOT add the image to document content flow.
        */
        if !fexists(imagePath)
            ? "Warning: Image file not found: " + imagePath
            return [:relId = "", :imageId = 0]
        ok
        
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read image file: " + imagePath
            return [:relId = "", :imageId = 0]
        ok
        
        ext = lower(wordGetImageExtension(imagePath))
        contentType = "image/png"
        switch ext
            on "png"  contentType = "image/png"
            on "jpg"  contentType = "image/jpeg"
            on "jpeg" contentType = "image/jpeg"
            on "gif"  contentType = "image/gif"
            on "bmp"  contentType = "image/bmp"
        off
        
        nRelId++
        nImageId++
        relId = "rId" + nRelId
        imageFilename = "image" + nImageId + "." + ext
        
        aImages + [
            :relId = relId,
            :filename = imageFilename,
            :data = imageData,
            :contentType = contentType
        ]
        
        aRelationships + [
            :id = relId,
            :type = "image",
            :target = "media/" + imageFilename
        ]
        
        return [:relId = relId, :imageId = nImageId]

    func registerCellList listType
        /*
            Register a numbering definition for use in a cell list.
            Returns the listId to pass to WordCell.addCellBulletList/addCellNumberedList
        */
        listId = nListId
        aNumbering + [:id = nListId, :type = listType]
        nListId++
        return listId

    func registerHyperlink url
        /*
            Register a hyperlink relationship for use in a cell.
            Returns the relId to pass to WordCell.addCellHyperlink
        */
        nRelId++
        relId = "rId" + nRelId
        aRelationships + [
            :id = relId,
            :type = "hyperlink",
            :target = url
        ]
        return relId

    func generateNestedTable data, options
        /*
            Generate table XML for embedding inside a cell.
            Returns XML string (does not add to document content).
        */
        return generateTable(data, options)
    
    func addCustomProperty name, value
        aCustomProps + [:name=name, :value=value]

    # ========================================================================
    # Save Document
    # ========================================================================
    
    func save filename
        sep = wordGetSep()
        
        # Create temp directory
        tempDir = filename + "_temp" + sep
        wordMakeDir(tempDir)
        wordMakeDir(tempDir + "_rels")
        wordMakeDir(tempDir + "word")
        wordMakeDir(tempDir + "word" + sep + "_rels")
        wordMakeDir(tempDir + "docProps")
        
        # Create media folder if we have images or an image watermark
        if len(aImages) > 0 or bImgWatermark
            wordMakeDir(tempDir + "word" + sep + "media")
        ok

        # Create theme folder if theme is active 
        if len(cThemeName) > 0
            wordMakeDir(tempDir + "word" + sep + "theme")
        ok

        # Create charts folder if we have charts
        if len(aCharts) > 0
            wordMakeDir(tempDir + "word" + sep + "charts")
            wordMakeDir(tempDir + "word" + sep + "charts" + sep + "_rels")
        ok
        
        # Register image watermark file into the media folder now so it is
        # available when writeHeader() and writeHeaderRels() are called.
        if bImgWatermark and len(cImgWatermarkPath) > 0
            imgData = read(cImgWatermarkPath)
            if len(imgData) = 0
                ? "Warning: Could not read watermark image: " + cImgWatermarkPath
                bImgWatermark = false
            else
                ext = lower(wordGetImageExtension(cImgWatermarkPath))
                cImgWatermarkFile = "watermark." + ext
                # Write directly to media; NOT via aImages (header owns this relationship)
                write(tempDir + "word" + sep + "media" + sep + cImgWatermarkFile, imgData)
                # Assign a header-local relationship ID (always rId1 in header1.xml.rels)
                cImgWatermarkRelId = "rId1"
            ok
        ok
        
        # Write all XML files
        writeContentTypes(tempDir)
        writeRels(tempDir)
        writeDocumentRels(tempDir)
        writeDocument(tempDir)
        writeStyles(tempDir)
        writeSettings(tempDir)
        writeFontTable(tempDir)
        writeCore(tempDir)
        writeApp(tempDir)

        if len(aCustomProps) > 0
            writeCustomProps(tempDir)
        ok

        # Write theme if active 
        if len(cThemeName) > 0
            writeTheme(tempDir)
        ok
        
        if len(aNumbering) > 0
            writeNumbering(tempDir)
        ok
        
        # Write header if needed (header text, text watermark, or image watermark)
        if len(cHeaderText) > 0 or bWatermark or bImgWatermark
            writeHeader(tempDir)
        ok
        
        # Write first-page header if enabled
        if bFirstPageDifferent
            writeFirstPageHeader(tempDir)
        ok
        
        # Write header relationships file when image watermark is active
        if bImgWatermark
            writeHeaderRels(tempDir)
        ok
        
        # Write footer if set
        if len(cFooterText) > 0 or bShowPageNumbers
            writeFooter(tempDir)
        ok

        # Write first-page footer if enabled 
        if bFirstPageDifferent
            writeFirstPageFooter(tempDir)
        ok
        
        # Write even-page header/footer if enabled
        if bEvenAndOddHeaders
            writeEvenPageHeader(tempDir)
            writeEvenPageFooter(tempDir)
        ok
        
        # Write footnotes if any
        if len(aFootnotes) > 0
            writeFootnotes(tempDir)
        ok
        
        # Write endnotes if any
        if len(aEndnotes) > 0
            writeEndnotes(tempDir)
        ok

        # Write comments if any 
        if len(aComments) > 0
            writeComments(tempDir)
        ok

        # Write chart XML files
        if len(aCharts) > 0
            for ch in aCharts
                writeChart(tempDir, ch)
                writeChartRels(tempDir, ch)
            next
        ok
        
        # Create ZIP file
        result = createZip(tempDir, filename)
        
        # Clean up temp directory
        if wordIsWindows()
            system('rmdir /s /q "' + tempDir + '" 2>nul')
        else
            system("rm -rf '" + tempDir + "'")
        ok
        
        return result
    
    # ========================================================================
    # Private: XML Generation Methods
    # ========================================================================
    
    private
    
    func writeContentTypes tempDir
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        c += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        c += '<Default Extension="xml" ContentType="application/xml"/>'
        
        # Add image extensions if we have images
        hasPng = false
        hasJpg = false
        hasGif = false
        hasBmp = false
        
        for img in aImages
            ct = img[:contentType]
            if ct = "image/png" and !hasPng
                c += '<Default Extension="png" ContentType="image/png"/>'
                hasPng = true
            ok
            if ct = "image/jpeg" and !hasJpg
                c += '<Default Extension="jpg" ContentType="image/jpeg"/>'
                c += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                hasJpg = true
            ok
            if ct = "image/gif" and !hasGif
                c += '<Default Extension="gif" ContentType="image/gif"/>'
                hasGif = true
            ok
            if ct = "image/bmp" and !hasBmp
                c += '<Default Extension="bmp" ContentType="image/bmp"/>'
                hasBmp = true
            ok
        next
        
        # Register the watermark image extension if not already covered above
        if bImgWatermark and len(cImgWatermarkFile) > 0
            wmExt = lower(wordGetImageExtension(cImgWatermarkFile))
            if wmExt = "png" and !hasPng
                c += '<Default Extension="png" ContentType="image/png"/>'
            ok
            if (wmExt = "jpg" or wmExt = "jpeg") and !hasJpg
                c += '<Default Extension="jpg" ContentType="image/jpeg"/>'
                c += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
            ok
            if wmExt = "gif" and !hasGif
                c += '<Default Extension="gif" ContentType="image/gif"/>'
            ok
            if wmExt = "bmp" and !hasBmp
                c += '<Default Extension="bmp" ContentType="image/bmp"/>'
            ok
        ok
        
        c += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        c += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        c += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        c += '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        c += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        c += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        if len(aCustomProps) > 0
            c += '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
        ok
        
        if len(aNumbering) > 0
            c += '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        ok
        
        # Header and footer content types
        if len(cHeaderText) > 0 or bWatermark or bImgWatermark
            c += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        ok
        if bFirstPageDifferent
            c += '<Override PartName="/word/header_fp.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        ok
        if len(cFooterText) > 0 or bShowPageNumbers
            c += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        ok
        if bFirstPageDifferent
            c += '<Override PartName="/word/footer_fp.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        ok
        if bEvenAndOddHeaders
            c += '<Override PartName="/word/header_even.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            c += '<Override PartName="/word/footer_even.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        ok
        
        # Footnotes and endnotes content types
        if len(aFootnotes) > 0
            c += '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        ok
        if len(aEndnotes) > 0
            c += '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        ok
        if len(aComments) > 0
            c += '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        ok
        if len(cThemeName) > 0
            c += '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        ok

        # Chart content types (one Override per chart)
        for ch in aCharts
            c += '<Override PartName="/word/charts/chart' + ch[:chartId] + '.xml"'
            c += ' ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        next
        
        c += '</Types>'
        
        write(tempDir + "[Content_Types].xml", c)
    
    func writeRels tempDir
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        c += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        if len(aCustomProps) > 0
            c += '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        ok
        c += '</Relationships>'
        
        write(tempDir + "_rels" + sep + ".rels", c)
    
    func writeDocumentRels tempDir
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        c += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'

        nextRel = 4

        # Theme relationship — must be registered before numbering
        if len(cThemeName) > 0
            c += '<Relationship Id="rId' + nextRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
            nextRel++
        ok
        
        if len(aNumbering) > 0
            c += '<Relationship Id="rId' + nextRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
            nextRel++
        ok
        
        # Header relationship (needed for header text, text watermark, or image watermark)
        if len(cHeaderText) > 0 or bWatermark or bImgWatermark
            cHeaderRelId = "rId" + nextRel
            c += '<Relationship Id="' + cHeaderRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            nextRel++
        ok
        
        # First-page header relationship
        if bFirstPageDifferent
            cFirstPageHeaderRelId = "rId" + nextRel
            c += '<Relationship Id="' + cFirstPageHeaderRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_fp.xml"/>'
            nextRel++
        ok
        
        # Footer relationship
        if len(cFooterText) > 0 or bShowPageNumbers
            cFooterRelId = "rId" + nextRel
            c += '<Relationship Id="' + cFooterRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
            nextRel++
        ok
        
        # First-page footer relationship
        if bFirstPageDifferent
            cFirstPageFooterRelId = "rId" + nextRel
            c += '<Relationship Id="' + cFirstPageFooterRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_fp.xml"/>'
            nextRel++
        ok

        # even-page header/footer relationships
        if bEvenAndOddHeaders
            cEvenPageHeaderRelId = "rId" + nextRel
            c += '<Relationship Id="' + cEvenPageHeaderRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_even.xml"/>'
            nextRel++
            cEvenPageFooterRelId = "rId" + nextRel
            c += '<Relationship Id="' + cEvenPageFooterRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_even.xml"/>'
            nextRel++
        ok
        
        # Footnotes relationship
        if len(aFootnotes) > 0
            c += '<Relationship Id="rId' + nextRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
            nextRel++
        ok
        
        # Endnotes relationship
        if len(aEndnotes) > 0
            c += '<Relationship Id="rId' + nextRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
            nextRel++
        ok

        # Comments relationship 
        if len(aComments) > 0
            c += '<Relationship Id="rId' + nextRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
            nextRel++
        ok
        
        # Add hyperlink and image relationships
        for rel in aRelationships
            if rel[:type] = "hyperlink"
                c += '<Relationship Id="' + rel[:id] + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + wordXmlEsc(rel[:target]) + '" TargetMode="External"/>'
            ok
            if rel[:type] = "image"
                c += '<Relationship Id="' + rel[:id] + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel[:target] + '"/>'
            ok
            # chart relationships
            if rel[:type] = "chart"
                c += '<Relationship Id="' + rel[:id] + '"'
                c += ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"'
                c += ' Target="charts/chart' + rel[:chartId] + '.xml"/>'
            ok
        next
        
        c += '</Relationships>'
        
        write(tempDir + "word" + sep + "_rels" + sep + "document.xml.rels", c)
    
    func writeDocument tempDir
        sep = wordGetSep()
        
        xml_ = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        xml_ += '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        xml_ += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        xml_ += 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        xml_ += 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        xml_ += 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        xml_ += 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        xml_ += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        xml_ += 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        xml_ += 'xmlns:v="urn:schemas-microsoft-com:vml" '
        xml_ += 'xmlns:o="urn:schemas-microsoft-com:office:office" '
        xml_ += 'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        xml_ += 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        xml_ += 'mc:Ignorable="w14 wp14">'

        # Page background color
        if len(cPageBgColor) > 0
            xml_ += '<w:background w:color="' + cPageBgColor + '"/>'
        ok

        xml_ += '<w:body>'
        
        # VML shape type definition required for v:rect text boxes
        if nTextBoxId > 0
            xml_ += '<w:p><w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr>'
            xml_ += '<w:r><w:rPr><w:noProof/></w:rPr>'
            xml_ += '<w:pict>'
            xml_ += '<v:shapetype id="_x0000_t202" coordsize="21600,21600"'
            xml_ += ' o:spt="202" path="m,l,21600r21600,l21600,xe">'
            xml_ += '<v:stroke joinstyle="miter"/>'
            xml_ += '<v:path gradientshapeok="t" o:connecttype="rect"/>'
            xml_ += '</v:shapetype>'
            xml_ += '</w:pict>'
            xml_ += '</w:r>'
            xml_ += '</w:p>'
        ok
        for item in aContent
            itemType = item[:type]
            
            switch itemType
                on "paragraph"
                    xml_ += generateParagraph(item[:text], item[:options])
                
                on "heading"
                    xml_ += generateHeading(item[:text], item[:level])
                
                on "richparagraph"
                    xml_ += generateRichParagraph(item[:runs], item[:options])
                
                on "list"
                    xml_ += generateList(item)
                
                on "nestedlist"
                    xml_ += generateNestedList(item)
                
                on "table"
                    xml_ += generateTable(item[:data], item[:options])
                
                on "pagebreak"
                    xml_ += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                
                on "linebreak"
                    xml_ += '<w:p><w:r><w:br/></w:r></w:p>'
                
                on "horizontalline"
                    xml_ += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
                
                on "hyperlink"
                    xml_ += generateHyperlink(item[:text], item[:relId])
                
                on "image"
                    xml_ += generateImage(item)
                
                on "columnbreak"
                    xml_ += '<w:p><w:r><w:br w:type="column"/></w:r></w:p>'
                
                on "sectionbreak"
                    xml_ += '<w:p><w:pPr><w:sectPr>'
                    xml_ += generateSectPrHeaderFooterRefs()
                    xml_ += '<w:type w:val="' + item[:breakType] + '"/>'
                    sbNC = item[:numColumns]
                    sbCS = item[:columnSpaceTwips]
                    if sbNC != NULL and isNumber(sbNC) and sbNC > 1
                        if sbCS = NULL or !isNumber(sbCS) or sbCS <= 0  sbCS = 720  ok
                        xml_ += '<w:cols w:num="' + string(sbNC) + '" w:space="' + string(sbCS) + '"/>'
                    ok
                    xml_ += '</w:sectPr></w:pPr></w:p>'
                
                on "landscapestart"
                    # Close the current portrait section
                    xml_ += '<w:p><w:pPr><w:sectPr>'
                    xml_ += generateSectPrHeaderFooterRefs()
                    xml_ += '<w:type w:val="nextPage"/>'
                    xml_ += '<w:pgSz w:w="' + nPageWidth + '" w:h="' + nPageHeight + '"/>'
                    xml_ += '<w:pgMar w:top="' + nMarginTop + '" w:right="' + nMarginRight + '" w:bottom="' + nMarginBottom + '" w:left="' + nMarginLeft + '" w:header="720" w:footer="720" w:gutter="0"/>'
                    if nColumns > 1
                        xml_ += '<w:cols w:num="' + nColumns + '" w:space="' + nColumnSpace + '"/>'
                    ok
                    xml_ += '</w:sectPr></w:pPr></w:p>'
                
                on "landscapeend"
                    # Close the landscape section (swap width/height for landscape)
                    xml_ += '<w:p><w:pPr><w:sectPr>'
                    xml_ += generateSectPrHeaderFooterRefs()
                    xml_ += '<w:type w:val="nextPage"/>'
                    xml_ += '<w:pgSz w:w="' + nPageHeight + '" w:h="' + nPageWidth + '" w:orient="landscape"/>'
                    xml_ += '<w:pgMar w:top="' + nMarginLeft + '" w:right="' + nMarginTop + '" w:bottom="' + nMarginRight + '" w:left="' + nMarginBottom + '" w:header="720" w:footer="720" w:gutter="0"/>'
                    xml_ += '</w:sectPr></w:pPr></w:p>'
                
                on "blockquote"
                    xml_ += '<w:p><w:pPr><w:ind w:left="720" w:right="720"/><w:pBdr>'
                    xml_ += '<w:left w:val="single" w:sz="18" w:space="4" w:color="CCCCCC"/>'
                    xml_ += '</w:pBdr></w:pPr>'
                    xml_ += '<w:r><w:rPr><w:i/></w:rPr><w:t>' + wordXmlEsc(item[:text]) + '</w:t></w:r></w:p>'
                
                on "caption"
                    xml_ += '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                    xml_ += '<w:r><w:rPr><w:i/><w:sz w:val="20"/></w:rPr><w:t>' + wordXmlEsc(item[:text]) + '</w:t></w:r></w:p>'
                
                on "abstract"
                    xml_ += '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                    xml_ += '<w:r><w:rPr><w:b/></w:rPr><w:t>Abstract</w:t></w:r></w:p>'
                    xml_ += '<w:p><w:pPr><w:ind w:left="720" w:right="720"/><w:jc w:val="both"/></w:pPr>'
                    xml_ += '<w:r><w:t>' + wordXmlEsc(item[:text]) + '</w:t></w:r></w:p>'
                
                on "keywords"
                    xml_ += '<w:p><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr>'
                    xml_ += '<w:r><w:rPr><w:b/></w:rPr><w:t>Keywords: </w:t></w:r>'
                    xml_ += '<w:r><w:rPr><w:i/></w:rPr><w:t>' + wordXmlEsc(item[:text]) + '</w:t></w:r></w:p>'
                
                on "toc"
                    xml_ += '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>'
                    xml_ += '<w:r><w:t>' + wordXmlEsc(item[:title]) + '</w:t></w:r></w:p>'
                    xml_ += '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    xml_ += '<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>'
                    xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    xml_ += '<w:r><w:t>Press F9 to update table of contents</w:t></w:r>'
                    xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

                on "textbox"
                    xml_ += generateTextBox(item)

                on "crossref"
                    xml_ += generateCrossRef(item)

                on "bookmarkanchor"
                    xml_ += generateBookmarkAnchor(item)

                on "tabbedparagraph"
                    xml_ += generateTabbedParagraph(item)

                on "seqcaption"
                    xml_ += generateSeqCaption(item)

                on "floatingimage"
                    xml_ += generateFloatingImage(item)

                on "numberedheading"
                    xml_ += generateNumberedHeading(item)

                on "mergeparagraph"
                    xml_ += generateMergeParagraph(item)

                on "sdt_checkbox"
                    xml_ += generateSdtCheckbox(item)

                on "sdt_dropdown"
                    xml_ += generateSdtDropdown(item)

                on "sdt_text"
                    xml_ += generateSdtText(item)

                on "drawshape"
                    xml_ += generateDrawShape(item)

                on "tof"
                    xml_ += generateTableOfFigures(item)

                on "chart"
                    xml_ += generateChartDrawing(item)

                on "field"
                    xml_ += generateField(item[:fieldType], item[:cachedValue], item[:options])
            off
        next
        
        # Section properties (page size, margins, header/footer, columns)
        xml_ += '<w:sectPr>'
        
        # Header/footer references (shared helper — same refs used in inline sectPr elements)
        xml_ += generateSectPrHeaderFooterRefs()
        
        # Page size
        xml_ += '<w:pgSz w:w="' + nPageWidth + '" w:h="' + nPageHeight + '"'
        if cOrientation = "landscape"
            xml_ += ' w:orient="landscape"'
        ok
        xml_ += '/>'
        
        # Page margins
        xml_ += '<w:pgMar w:top="' + nMarginTop + '" w:right="' + nMarginRight + '" w:bottom="' + nMarginBottom + '" w:left="' + nMarginLeft + '" w:header="720" w:footer="720" w:gutter="0"/>'
        
        # Page borders (must come after pgMar in the schema)
        if bPageBorder
            xml_ += generatePageBordersXml()
        ok
        
        # Columns
        if nColumns > 1
            xml_ += '<w:cols w:num="' + nColumns + '" w:space="' + nColumnSpace + '"/>'
        else
            xml_ += '<w:cols w:space="720"/>'
        ok
        
        # Document-level BiDi
        if bDocumentRTL
            xml_ += '<w:bidi/>'
        ok

        # Line numbers
        if bLineNumbers
            xml_ += '<w:lnNumType w:countBy="' + nLineNumberStep + '"'
            xml_ += ' w:start="' + nLineNumberStart + '"'
            xml_ += ' w:distance="' + nLineNumberDistance + '"'
            xml_ += ' w:restart="' + cLineNumberRestart + '"/>'
        ok
        
        xml_ += '</w:sectPr>'
        
        xml_ += '</w:body>'
        xml_ += '</w:document>'
        
        write(tempDir + "word" + sep + "document.xml", xml_)
    
    func generateSectPrHeaderFooterRefs
        /*
            Returns the header/footer reference XML that belongs inside any
            <w:sectPr> — both the final one and every inline one (landscape
            section breaks).  Centralising this ensures all sections share
            the same headers and footers.
        */
        xml_ = ""
        # Default (odd-page) header
        if len(cHeaderRelId) > 0
            xml_ += '<w:headerReference w:type="default" r:id="' + cHeaderRelId + '"/>'
        ok
        # First-page header
        if bFirstPageDifferent and len(cFirstPageHeaderRelId) > 0
            xml_ += '<w:headerReference w:type="first" r:id="' + cFirstPageHeaderRelId + '"/>'
        ok
        # Even-page header
        if bEvenAndOddHeaders and len(cEvenPageHeaderRelId) > 0
            xml_ += '<w:headerReference w:type="even" r:id="' + cEvenPageHeaderRelId + '"/>'
        ok
        # Default (odd-page) footer
        if len(cFooterRelId) > 0
            xml_ += '<w:footerReference w:type="default" r:id="' + cFooterRelId + '"/>'
        ok
        # First-page footer
        if bFirstPageDifferent and len(cFirstPageFooterRelId) > 0
            xml_ += '<w:footerReference w:type="first" r:id="' + cFirstPageFooterRelId + '"/>'
        ok
        # Even-page footer
        if bEvenAndOddHeaders and len(cEvenPageFooterRelId) > 0
            xml_ += '<w:footerReference w:type="even" r:id="' + cEvenPageFooterRelId + '"/>'
        ok
        # titlePg: required when first-page header/footer differ
        if bFirstPageDifferent
            xml_ += '<w:titlePg/>'
        ok
        return xml_

    func generateParagraph text, options
        # wrap in bookmark if requested
        bmName = NULL
        if isList(options)
            bmName = options[:_bookmark]
        ok

        # attach comment if requested
        cmId = NULL
        if isList(options)
            cmId = options[:_commentId]
        ok

        xml_ = '<w:p>'
        
        # Check for custom style reference
        styleId = NULL
        if isList(options)
            styleId = options[:style]
        ok
        
        # Paragraph properties
        pPr = generateParagraphProperties(options)
        if styleId != NULL
            xml_ += '<w:pPr><w:pStyle w:val="' + styleId + '"/>' + pPr + '</w:pPr>'
        elseif len(pPr) > 0
            xml_ += '<w:pPr>' + pPr + '</w:pPr>'
        ok

        # Open bookmark anchor inside the paragraph
        if bmName != NULL
            bmId = nBookmarkId
            nBookmarkId++
            aBookmarks + [:id = bmId, :name = bmName]
            xml_ += '<w:bookmarkStart w:id="' + bmId + '" w:name="' + wordXmlEsc(bmName) + '"/>'
        ok

        # Open comment range
        if cmId != NULL
            xml_ += '<w:commentRangeStart w:id="' + cmId + '"/>'
        ok
        
        # Text run
        xml_ += '<w:r>'
        
        # Run properties
        rPr = generateRunProperties(options)
        if len(rPr) > 0
            xml_ += '<w:rPr>' + rPr + '</w:rPr>'
        ok
        
        # Text content
        if text != NULL and len(text) > 0
            # Preserve spaces
            if left(text, 1) = " " or right(text, 1) = " "
                xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(text) + '</w:t>'
            else
                xml_ += '<w:t>' + wordXmlEsc(text) + '</w:t>'
            ok
        ok
        
        xml_ += '</w:r>'

        # Close comment range and add comment reference run
        if cmId != NULL
            xml_ += '<w:commentRangeEnd w:id="' + cmId + '"/>'
            xml_ += '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
            xml_ += '<w:commentReference w:id="' + cmId + '"/></w:r>'
        ok

        # Close bookmark anchor
        if bmName != NULL
            xml_ += '<w:bookmarkEnd w:id="' + bmId + '"/>'
        ok

        xml_ += '</w:p>'
        
        return xml_
    
    func generateHeading text, level
        xml_ = '<w:p>'
        if bDocumentRTL
            # Must set jc explicitly — Word ignores Normal style jc inheritance in RTL mode
            xml_ += '<w:pPr><w:pStyle w:val="Heading' + level + '"/><w:jc w:val="both"/></w:pPr>'
        else
            xml_ += '<w:pPr><w:pStyle w:val="Heading' + level + '"/></w:pPr>'
        ok
        xml_ += '<w:r><w:t>' + wordXmlEsc(text) + '</w:t></w:r>'
        xml_ += '</w:p>'
        return xml_
    
    func generateRichParagraph runs, paragraphOptions
        xml_ = '<w:p>'
        
        # Paragraph properties
        pPr = generateParagraphProperties(paragraphOptions)
        if len(pPr) > 0
            xml_ += '<w:pPr>' + pPr + '</w:pPr>'
        ok
        
        # Multiple runs
        runsLen = len(runs)
        for i = 1 to runsLen
            run = runs[i]
            text = run[1]
            runOptions = []
            if len(run) > 1
                runOptions = run[2]
            ok
            
            # Footnote reference run (no text content)
            if isList(runOptions) and runOptions[:footnoteId] != NULL
                xml_ += generateFootnoteRef(runOptions[:footnoteId])
                loop
            ok
            
            # Endnote reference run (no text content)
            if isList(runOptions) and runOptions[:endnoteId] != NULL
                xml_ += generateEndnoteRef(runOptions[:endnoteId])
                loop
            ok
            
            # Normal text run
            xml_ += '<w:r>'
            
            rPr = generateRunProperties(runOptions)
            if len(rPr) > 0
                xml_ += '<w:rPr>' + rPr + '</w:rPr>'
            ok
            
            if left(text, 1) = " " or right(text, 1) = " "
                xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(text) + '</w:t>'
            else
                xml_ += '<w:t>' + wordXmlEsc(text) + '</w:t>'
            ok
            
            xml_ += '</w:r>'
        next
        
        xml_ += '</w:p>'
        return xml_
    
    func generateParagraphProperties options
        if !isList(options)
            options = []
        ok
        
        # When bDocumentRTL we must always emit explicit <w:jc> — Word does NOT
        # propagate jc from Normal style to individual paragraphs in RTL mode.
        if len(options) = 0 and !bDocumentRTL
            return ""
        ok
        
        xml_ = ""

        # Paragraph style (must be first element in <w:pPr>)
        if options[:style] != NULL
            xml_ += '<w:pStyle w:val="' + options[:style] + '"/>'
        ok
        
        # BiDi direction (schema requires <w:bidi> BEFORE <w:jc>)
        # pPrDefault already carries <w:bidi/> for the whole document —
        # only explicit :rtl/:ltr per-paragraph overrides need it here.
        if options[:ltr] = true
            # Force LTR override inside an RTL document
            xml_ += '<w:bidi w:val="0"/>'
            alignDefault = "left"
        elseif options[:rtl] = true
            # Paragraph-level RTL in an otherwise-LTR document
            xml_ += '<w:bidi/>'
            alignDefault = "both"
        elseif bDocumentRTL
            # RTL document default — must be explicit on every paragraph
            alignDefault = "both"
        else
            alignDefault = NULL
        ok
        
        # Alignment: explicit :align wins; otherwise direction default
        alignVal = options[:align]
        if alignVal = NULL
            alignVal = alignDefault
        ok
        if alignVal != NULL
            xml_ += '<w:jc w:val="' + alignVal + '"/>'
        ok
        
        # Spacing (before/after + optional line spacing combined in one element)
        hasBefore    = options[:spaceBefore] != NULL
        hasAfter     = options[:spaceAfter]  != NULL
        hasLineSpacing = options[:lineSpacing] != NULL     # multiplier: 1.0, 1.5, 2.0
        hasLineSpacingPt = options[:lineSpacingPt] != NULL # exact/at-least value in points
        hasLineRule  = options[:lineRule] != NULL          # "auto", "exact", "atLeast"
        
        if hasBefore or hasAfter or hasLineSpacing or hasLineSpacingPt
            xml_ += '<w:spacing'
            if hasBefore
                xml_ += ' w:before="' + options[:spaceBefore] + '"'
            ok
            if hasAfter
                xml_ += ' w:after="' + options[:spaceAfter] + '"'
            ok
            if hasLineSpacingPt
                # Exact or at-least in points → convert to twips (1 pt = 20 twips)
                lineVal = floor(options[:lineSpacingPt] * 20)
                lineRule = "exact"
                if hasLineRule
                    lineRule = options[:lineRule]
                ok
                xml_ += ' w:line="' + lineVal + '" w:lineRule="' + lineRule + '"'
            elseif hasLineSpacing
                # Multiplier-based: Word uses 240 as 1× (single spacing)
                lineVal = floor(options[:lineSpacing] * 240)
                xml_ += ' w:line="' + lineVal + '" w:lineRule="auto"'
            ok
            xml_ += '/>'
        ok
        
        # Indentation (left, right, firstLine, hanging)
        hasIndent = options[:indent] != NULL or options[:indentRight] != NULL or
                    options[:indentFirstLine] != NULL or options[:indentHanging] != NULL
        if hasIndent
            indXml = "<w:ind"
            if options[:indent]          != NULL  indXml += ' w:left="'      + options[:indent]          + '"'  ok
            if options[:indentRight]     != NULL  indXml += ' w:right="'     + options[:indentRight]     + '"'  ok
            if options[:indentFirstLine] != NULL  indXml += ' w:firstLine="' + options[:indentFirstLine] + '"'  ok
            if options[:indentHanging]   != NULL  indXml += ' w:hanging="'   + options[:indentHanging]   + '"'  ok
            xml_ += indXml + "/>"
        ok
        
        # Paragraph shading 
        if options[:bgColor] != NULL
            xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + wordColorToHex(options[:bgColor]) + '"/>'
        ok

        # Paragraph borders
        if options[:_hasBorder] = true
            bdrStyle = options[:borderStyle]
            if bdrStyle = NULL bdrStyle = "single" ok
            bdrColor = options[:borderColor]
            if bdrColor = NULL bdrColor = "000000" ok
            bdrColor = wordColorToHex(bdrColor)
            bdrSize = options[:borderSize]
            if bdrSize = NULL bdrSize = 6 ok
            bdrSpace = options[:borderSpace]
            if bdrSpace = NULL bdrSpace = 4 ok
            bdrSides = options[:sides]
            if !isList(bdrSides)
                bdrSides = ["top", "left", "bottom", "right"]
            ok
            bdrAttr = ' w:val="' + bdrStyle + '" w:sz="' + bdrSize + '" w:space="' + bdrSpace + '" w:color="' + bdrColor + '"'
            xml_ += '<w:pBdr>'
            sidesLen = len(bdrSides)
            for si = 1 to sidesLen
                xml_ += '<w:' + bdrSides[si] + bdrAttr + '/>'
            next
            xml_ += '</w:pBdr>'
        ok

        # Keep together flags
        if options[:keepNext] = true
            xml_ += '<w:keepNext/>'
        ok
        if options[:keepLines] = true
            xml_ += '<w:keepLines/>'
        ok

        # suppress line numbers on this paragraph
        if options[:suppressLineNumbers] = true
            xml_ += '<w:suppressLineNumbers/>'
        ok

        # tab stops
        tabList = options[:tabStops]
        if isList(tabList) and len(tabList) > 0
            xml_ += '<w:tabs>'
            tabListLen = len(tabList)
            for ti = 1 to tabListLen
                tab = tabList[ti]
                tabPos   = tab[:pos]
                tabAlign = tab[:align]
                tabLead  = tab[:leader]
                if tabPos   = NULL  tabPos   = 720    ok
                if tabAlign = NULL  tabAlign = "left" ok
                if tabLead  = NULL  tabLead  = "none" ok
                # Map leader names to Word values
                if tabLead = "dot"        tabLead = "dot"        ok
                if tabLead = "hyphen"     tabLead = "hyphen"     ok
                if tabLead = "underscore" tabLead = "underscore" ok
                xml_ += '<w:tab w:val="' + tabAlign + '" w:leader="' + tabLead + '" w:pos="' + tabPos + '"/>'
            next
            xml_ += '</w:tabs>'
        ok
        
        # Outline level (for TOC depth / navigation pane)
        if options[:outlineLvl] != NULL
            ovl = options[:outlineLvl]
            if isNumber(ovl) and ovl >= 0 and ovl <= 8
                xml_ += '<w:outlineLvl w:val="' + string(floor(ovl)) + '"/>'
            ok
        ok
        
        # Contextual spacing (suppress spacing between same-style paragraphs)
        if options[:contextualSpacing] = true
            xml_ += '<w:contextualSpacing/>'
        ok

        # Widow/orphan control
        # Word default = ON (protect last/first line). Pass false to disable.
        if options[:widowControl] = false
            xml_ += '<w:widowControl w:val="0"/>'
        ok

        # Suppress automatic hyphenation on this paragraph
        if options[:noHyphenate] = true
            xml_ += '<w:suppressAutoHyphens/>'
        ok

        return xml_
    
    func generateRunProperties options
        if !isList(options)
            options = []
        ok
        
        if len(options) = 0
            return ""
        ok
        
        xml_ = ""
        
        # Bold
        if options[:bold] = true
            xml_ += '<w:b/>'
        ok
        
        # Italic
        if options[:italic] = true
            xml_ += '<w:i/>'
        ok
        
        # Underline
        if options[:underline] = true
            xml_ += '<w:u w:val="single"/>'
        ok
        
        # Strikethrough / double strikethrough
        if options[:strike] = true
            xml_ += '<w:strike/>'
        ok
        if options[:dstrike] = true
            xml_ += '<w:dstrike/>'
        ok

        # All-caps / small-caps
        if options[:caps] = true
            xml_ += '<w:caps/>'
        ok
        if options[:smallCaps] = true
            xml_ += '<w:smallCaps/>'
        ok

        # Hidden text
        if options[:vanish] = true
            xml_ += '<w:vanish/>'
        ok
        
        # Font name – support both LTR and complex-script (RTL) fonts
        if options[:font] != NULL
            if options[:csFont] != NULL
                xml_ += '<w:rFonts w:ascii="' + options[:font] + '" w:hAnsi="' + options[:font] + '" w:cs="' + options[:csFont] + '"/>'
            else
                xml_ += '<w:rFonts w:ascii="' + options[:font] + '" w:hAnsi="' + options[:font] + '"/>'
            ok
        elseif options[:csFont] != NULL
            # Complex-script font only
            xml_ += '<w:rFonts w:cs="' + options[:csFont] + '"/>'
        ok
        
        # Font size (in half-points)
        if options[:size] != NULL
            halfPoints = options[:size] * 2
            xml_ += '<w:sz w:val="' + halfPoints + '"/>'
            xml_ += '<w:szCs w:val="' + halfPoints + '"/>'
        ok
        
        # Color
        if options[:color] != NULL
            xml_ += '<w:color w:val="' + wordColorToHex(options[:color]) + '"/>'
        ok
        
        # Highlight/background color
        if options[:highlight] != NULL
            xml_ += '<w:highlight w:val="' + options[:highlight] + '"/>'
        ok
        
        # Superscript / subscript
        if options[:superscript] = true
            xml_ += '<w:vertAlign w:val="superscript"/>'
        ok
        if options[:subscript] = true
            xml_ += '<w:vertAlign w:val="subscript"/>'
        ok
        
        # RTL run direction mark — only for explicit per-run overrides.
        # When bDocumentRTL, Normal style's <w:rtl/> already covers all runs.
        if options[:ltr] = true
            xml_ += '<w:rtl w:val="0"/>'
        elseif options[:rtl] = true
            xml_ += '<w:rtl/>'
        ok
        
        # Run border (w:bdr) — e.g. for boxed/outlined text runs
        if options[:runBorderStyle] != NULL and len(options[:runBorderStyle]) > 0
            bdrStyle = options[:runBorderStyle]
            bdrColor = options[:runBorderColor]
            bdrSize  = options[:runBorderSize]
            if bdrColor = NULL or len(bdrColor) = 0  bdrColor = "auto"  ok
            if bdrSize  = NULL  bdrSize = 4  ok
            xml_ += '<w:bdr w:val="' + bdrStyle + '" w:sz="' + string(bdrSize) + '" w:space="0" w:color="' + bdrColor + '"/>'
        ok

        # Named character style (w:rStyle) — preserves semantic char styles like
        # "Strong", "Emphasis", "Book Title", custom styles, etc.
        # Must appear first in rPr per OOXML spec — prepend to already-built xml_.
        if options[:charStyle] != NULL and len(options[:charStyle]) > 0
            xml_ = '<w:rStyle w:val="' + wordXmlEsc(options[:charStyle]) + '"/>' + xml_
        ok

        # Run language (w:lang) — preserves per-run spell-check/hyphenation language
        if options[:lang] != NULL and len(options[:lang]) > 0
            xml_ += '<w:lang w:val="' + wordXmlEsc(options[:lang]) + '"/>'
        ok

        return xml_
    
    func generateList item
        xml_ = ""
        listId = item[:listId]
        items = item[:items]
        itemsCount = len(items)
        
        for i = 1 to itemsCount
            itemText = items[i]
            xml_ += '<w:p>'
            xml_ += '<w:pPr>'
            xml_ += '<w:pStyle w:val="ListParagraph"/>'
            xml_ += '<w:numPr>'
            xml_ += '<w:ilvl w:val="0"/>'
            xml_ += '<w:numId w:val="' + listId + '"/>'
            xml_ += '</w:numPr>'
            if bDocumentRTL
                # Must be explicit — Word ignores Normal style jc inheritance in RTL mode
                xml_ += '<w:jc w:val="both"/>'
            ok
            xml_ += '</w:pPr>'
            xml_ += '<w:r><w:t>' + wordXmlEsc(itemText) + '</w:t></w:r>'
            xml_ += '</w:p>'
        next
        
        return xml_
    
    func generateNestedList item
        xml_ = ""
        listId = item[:listId]
        items = item[:items]
        itemsCount = len(items)
        
        for i = 1 to itemsCount
            entry = items[i]
            itemText  = entry[1]
            itemLevel = entry[2]
            if itemLevel < 0 itemLevel = 0 ok
            if itemLevel > 8 itemLevel = 8 ok
            
            xml_ += '<w:p>'
            xml_ += '<w:pPr>'
            xml_ += '<w:pStyle w:val="ListParagraph"/>'
            xml_ += '<w:numPr>'
            xml_ += '<w:ilvl w:val="' + itemLevel + '"/>'
            xml_ += '<w:numId w:val="' + listId + '"/>'
            xml_ += '</w:numPr>'
            if bDocumentRTL
                xml_ += '<w:jc w:val="both"/>'
            ok
            xml_ += '</w:pPr>'
            xml_ += '<w:r><w:t>' + wordXmlEsc(itemText) + '</w:t></w:r>'
            xml_ += '</w:p>'
        next
        
        return xml_

    func generateFootnoteRef fnId
        /*
            Generates a single superscript run that references a footnote.
            Used inline within a paragraph.
        */
        xml_ = '<w:r>'
        xml_ += '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
        xml_ += '<w:footnoteReference w:id="' + fnId + '"/>'
        xml_ += '</w:r>'
        return xml_

    func generateEndnoteRef enId
        /*
            Generates a single superscript run that references an endnote.
            Used inline within a paragraph.
        */
        xml_ = '<w:r>'
        xml_ += '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
        xml_ += '<w:endnoteReference w:id="' + enId + '"/>'
        xml_ += '</w:r>'
        return xml_
    
    func registerCellImages data
        /*
            Scan table data for WordCell objects containing images.
            Register each image with the document's relationship and image systems.
            Stores mapping in WordWriter's aCellImgMap* arrays.
        */
        numRows = len(data)
        for rowIdx = 1 to numRows
            row = data[rowIdx]
            if !isList(row) loop ok
            rowLen = len(row)
            for colIdx = 1 to rowLen
                if !isObject(row[colIdx]) loop ok
                cellData = row[colIdx]
                imgCount = cellData.nCellImageCount
                if !isNumber(imgCount) loop ok
                if imgCount < 1 loop ok
                for imgIdx = 1 to imgCount
                    cImgPath = cellData.aCellImgPaths[imgIdx]
                    
                    if !fexists(cImgPath)
                        ? "Warning: Cell image file not found: " + cImgPath
                        loop
                    ok
                    
                    imageData = read(cImgPath)
                    if len(imageData) = 0
                        ? "Warning: Could not read cell image file: " + cImgPath
                        loop
                    ok
                    
                    # Determine content type
                    ext = lower(wordGetImageExtension(cImgPath))
                    contentType = "image/png"
                    switch ext
                        on "png"  contentType = "image/png"
                        on "jpg"  contentType = "image/jpeg"
                        on "jpeg" contentType = "image/jpeg"
                        on "gif"  contentType = "image/gif"
                        on "bmp"  contentType = "image/bmp"
                    off
                    
                    # Register with document
                    nRelId = nRelId + 1
                    nImageId = nImageId + 1
                    relId = "rId" + nRelId
                    imageFilename = "image" + nImageId + "." + ext
                    
                    aImages + [
                        :relId = relId,
                        :filename = imageFilename,
                        :data = imageData,
                        :contentType = contentType
                    ]
                    
                    aRelationships + [
                        :id = relId,
                        :type = "image",
                        :target = "media/" + imageFilename
                    ]
                    
                    # Store mapping in WordWriter for lookup during rendering
                    aCellImgMapPaths + cImgPath
                    aCellImgMapRelIds + relId
                    aCellImgMapImgIds + nImageId
                    aCellImgMapWidths + cellData.aCellImgWidths[imgIdx]
                    aCellImgMapHeights + cellData.aCellImgHeights[imgIdx]
                next
            next
        next

    func generateTable data, options
        if !isList(options)
            options = []
        ok
        
        hasHeader = false
        if options[:headerRow] = true
            hasHeader = true
        ok
        
        borderStyle = "single"
        if options[:borderStyle] != NULL
            borderStyle = options[:borderStyle]
        ok
        
        borderColor = "auto"
        if options[:borderColor] != NULL
            borderColor = options[:borderColor]
        ok
        
        borderSize = "4"
        if options[:borderSize] != NULL
            borderSize = "" + options[:borderSize]
        ok
        
        headerBgColor = ""
        if options[:headerBgColor] != NULL
            headerBgColor = options[:headerBgColor]
        ok
        
        evenRowBgColor = ""
        if options[:evenRowBgColor] != NULL
            evenRowBgColor = options[:evenRowBgColor]
        ok
        
        # per-row background colors — list indexed by row (1-based)
        # e.g. :rowBgColors = ["FF0000", "", "00FF00"] — empty string = no override
        aRowBgColors = []
        if isList(options[:rowBgColors])
            aRowBgColors = options[:rowBgColors]
        ok
        
        # Conditional formatting rules
        aConditionalRules = []
        if isList(options[:conditionalRules])
            aConditionalRules = options[:conditionalRules]
        ok

        # Per-row height overrides
        # :rowHeights = [1.5, 0, 2.0, 0]  — cm per row; 0 means auto (skip)
        # :rowHRule   = "atLeast" (default) or "exact"
        aRowHeightsCm = []
        if isList(options[:rowHeights])
            aRowHeightsCm = options[:rowHeights]
        ok
        cRowHRule = "atLeast"
        if options[:rowHRule] != NULL and len(options[:rowHRule]) > 0
            cRowHRule = options[:rowHRule]
        ok

        # repeat header row on every page when table spans pages
        bRepeatHeader = false
        if options[:repeatHeader] = true
            bRepeatHeader = true
        ok

        # Header text color
        headerTextColor = "FFFFFF"
        if options[:headerTextColor] != NULL
            headerTextColor = wordColorToHex(options[:headerTextColor])
        ok

        # Vertical alignment for header and cells
        headerVAlign = ""
        if options[:headerVAlign] != NULL
            headerVAlign = options[:headerVAlign]
        elseif hasHeader
            headerVAlign = "center"
        ok
        cellVAlign = ""
        if options[:cellVAlign] != NULL
            cellVAlign = options[:cellVAlign]
        ok

        # Header row height in twips (0 = auto)
        headerHeight = 0
        if options[:headerHeight] != NULL
            headerHeight = options[:headerHeight]
        elseif hasHeader and len(headerBgColor) > 0
            # Default: give styled headers some height for visual comfort
            headerHeight = 600   # ~0.33 inch / ~0.85 cm
        ok

        # Cell padding for headers (in twips)
        headerPadding = 0
        if options[:headerPadding] != NULL
            headerPadding = options[:headerPadding]
        ok
        
        # Calculate column count
        numCols = 0
        numRows = len(data)

        # Compute true column count by scanning every row and summing colspans.
        # A plain string counts as 1; a WordCell with nColSpan > 1 counts as nColSpan.
        # This handles tables whose first row is a single full-width spanning header.
        for scanRow = 1 to numRows
            rowCells = data[scanRow]
            rowEffective = 0
            for scanCell in rowCells
                span = 1
                if isObject(scanCell)
                    if isNumber(scanCell.nColSpan) and scanCell.nColSpan > 1
                        span = scanCell.nColSpan
                    ok
                ok
                rowEffective = rowEffective + span
            next
            if rowEffective > numCols
                numCols = rowEffective
            ok
        next
        if numCols = 0  numCols = 1  ok

        # Column widths - support custom widths in cm
        # Options: :colWidths = [3, 5, 2] (in cm) or :colWidths = "auto" (equal)
        # Total page width = 9360 twips (6.5 inches)
        hasCustomWidths = false
        aColWidths = []
        if isList(options[:colWidths])
            hasCustomWidths = true
            # Convert cm values to twips (1 cm = 567 twips)
            widthsList = options[:colWidths]
            wLen = len(widthsList)
            for wi = 1 to wLen
                if wi <= numCols
                    aColWidths + floor(widthsList[wi] * 567)
                ok
            next
            # If fewer widths than columns, fill remaining with equal share
            if wLen < numCols
                usedWidth = 0
                uLen = len(aColWidths)
                for wi = 1 to uLen
                    usedWidth = usedWidth + aColWidths[wi]
                next
                remainingWidth = 9360 - usedWidth
                remainingCols = numCols - uLen
                if remainingCols > 0
                    equalShare = floor(remainingWidth / remainingCols)
                    for wi = 1 to remainingCols
                        aColWidths + equalShare
                    next
                ok
            ok
        else
            # Equal width columns
            equalWidth = floor(9360 / numCols)
            for wi = 1 to numCols
                aColWidths + equalWidth
            next
        ok
        
        xml_ = '<w:tbl>'
        
        # Table properties
        xml_ += '<w:tblPr>'
        # Table style: use caller-supplied style or fall back to TableGrid
        useTblStyle = options[:tblStyle]
        if useTblStyle = NULL or len(useTblStyle) = 0  useTblStyle = "TableGrid"  ok
        xml_ += '<w:tblStyle w:val="' + useTblStyle + '"/>'

        # Table width  — :tblWidthTwips (raw, from round-trip) takes priority;
        #                 else :tblWidth in cm; else auto.
        tblWxml = ""
        tblWtTwips = options[:tblWidthTwips]
        tblWtType  = options[:tblWidthType]
        if isNumber(tblWtTwips) and tblWtTwips > 0
            if tblWtType = NULL or len(tblWtType) = 0  tblWtType = "dxa"  ok
            tblWxml = '<w:tblW w:w="' + floor(tblWtTwips) + '" w:type="' + tblWtType + '"/>'
        else
            tblWcm = options[:tblWidth]
            if isNumber(tblWcm) and tblWcm > 0
                tblWtTwips2 = floor(tblWcm * 567)
                tblWxml = '<w:tblW w:w="' + tblWtTwips2 + '" w:type="dxa"/>'
            ok
        ok
        if len(tblWxml) = 0
            tblWxml = '<w:tblW w:w="0" w:type="auto"/>'
        ok
        xml_ += tblWxml

        # Table alignment (jc): :tblAlign = "left" | "center" | "right"
        tblJcv = options[:tblAlign]
        if tblJcv != NULL and len(tblJcv) > 0
            xml_ += '<w:jc w:val="' + lower(tblJcv) + '"/>'
        ok

        # Table indent: :tblIndentTwips (raw) or :tblIndent in cm
        tblIndRaw = options[:tblIndentTwips]
        if isNumber(tblIndRaw) and tblIndRaw > 0
            xml_ += '<w:tblInd w:w="' + floor(tblIndRaw) + '" w:type="dxa"/>'
        else
            tblIndCm = options[:tblIndent]
            if isNumber(tblIndCm) and tblIndCm > 0
                xml_ += '<w:tblInd w:w="' + floor(tblIndCm * 567) + '" w:type="dxa"/>'
            ok
        ok
        
        if borderStyle != "none"
            xml_ += '<w:tblBorders>'
            xml_ += '<w:top w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '<w:left w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '<w:bottom w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '<w:right w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '<w:insideH w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '<w:insideV w:val="' + borderStyle + '" w:sz="' + borderSize + '" w:space="0" w:color="' + borderColor + '"/>'
            xml_ += '</w:tblBorders>'
        ok
        
        # F3: tblLook — use options[:tblLook] map if provided, else defaults
        tblLookOpts = options[:tblLook]
        if isList(tblLookOpts)
            tlFR3 = tblLookOpts[:firstRow]
            tlLR3 = tblLookOpts[:lastRow]
            tlFC3 = tblLookOpts[:firstCol]
            tlLC3 = tblLookOpts[:lastCol]
            tlNH3 = tblLookOpts[:noHBand]
            tlNV3 = tblLookOpts[:noVBand]
            if !isNumber(tlFR3)  tlFR3 = 1  ok
            if !isNumber(tlLR3)  tlLR3 = 0  ok
            if !isNumber(tlFC3)  tlFC3 = 1  ok
            if !isNumber(tlLC3)  tlLC3 = 0  ok
            if !isNumber(tlNH3)  tlNH3 = 0  ok
            if !isNumber(tlNV3)  tlNV3 = 1  ok
            xml_ += '<w:tblLook w:firstRow="' + tlFR3 + '" w:lastRow="'  + tlLR3 +
                    '" w:firstColumn="' + tlFC3 + '" w:lastColumn="' + tlLC3 +
                    '" w:noHBand="' + tlNH3 + '" w:noVBand="' + tlNV3 + '"/>'
        else
            xml_ += '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        ok
        xml_ += '</w:tblPr>'
        
        # Grid columns
        xml_ += '<w:tblGrid>'
        for col = 1 to numCols
            xml_ += '<w:gridCol w:w="' + aColWidths[col] + '"/>'
        next
        xml_ += '</w:tblGrid>'
        
        # Rows
        for rowIdx = 1 to numRows
            row = data[rowIdx]
            isHeader = (rowIdx = 1 and hasHeader)
            isEvenRow = (rowIdx % 2 = 0)
            rowLen = len(row)
            
            xml_ += '<w:tr>'
            
            # Row properties - height + repeat header
            trPr = ""
            if isHeader and headerHeight > 0
                trPr += '<w:trHeight w:val="' + headerHeight + '" w:hRule="atLeast"/>'
            ok
            # per-row height override (takes priority over headerHeight for header row)
            if rowIdx <= len(aRowHeightsCm)
                rhCm = aRowHeightsCm[rowIdx]
                if isNumber(rhCm) and rhCm > 0
                    rhTwips = floor(rhCm * 567)
                    # Replace any headerHeight trHeight already added
                    trPr = '<w:trHeight w:val="' + string(rhTwips) + '" w:hRule="' + cRowHRule + '"/>'
                ok
            ok
            if isHeader and bRepeatHeader
                trPr += '<w:tblHeader/>'
            ok
            if len(trPr) > 0
                xml_ += '<w:trPr>' + trPr + '</w:trPr>'
            ok
            
            # Track which logical column we are in (accounting for colspans already placed)
            logicalCol = 0
            for colIdx = 1 to rowLen
                cellData = row[colIdx]

                # Determine colspan of this cell
                cellSpan = 1
                if isObject(cellData)
                    if isNumber(cellData.nColSpan) and cellData.nColSpan > 1
                        cellSpan = cellData.nColSpan
                    ok
                ok
                logicalCol = logicalCol + 1

                # Get column width: sum widths of all spanned columns
                colWidth = 0
                for spanOff = 0 to cellSpan - 1
                    wIdx = logicalCol + spanOff
                    if wIdx <= len(aColWidths)
                        colWidth = colWidth + aColWidths[wIdx]
                    ok
                next
                logicalCol = logicalCol + cellSpan - 1   # advance past spanned columns

                # Check if cell is a WordCell object or a plain string
                if isObject(cellData)
                    # WordCell object - render rich content (pass row color)
                    rowColor = ""
                    if rowIdx <= len(aRowBgColors)
                        rowColor = aRowBgColors[rowIdx]
                    ok
                    xml_ += generateRichCell(cellData, colWidth, isHeader, headerBgColor, headerTextColor, isEvenRow, evenRowBgColor, headerVAlign, cellVAlign, rowColor)
                else
                    # Plain string - render with optional conditional formatting
                    cellText = "" + cellData

                    # Evaluate conditional rules (skip header row)
                    cfBg        = ""
                    cfTextColor = ""
                    cfBold      = false
                    if !isHeader and len(aConditionalRules) > 0
                        rulesLen = len(aConditionalRules)
                        for rIdx = 1 to rulesLen
                            rule    = aConditionalRules[rIdx]
                            ruleCol = rule[:col]
                            if ruleCol = NULL  ruleCol = 0  ok
                            # :col = 0 means all columns
                            if ruleCol = 0 or ruleCol = colIdx
                                if evalTableCondition(cellText, rule)
                                    if rule[:bgColor]    != NULL  cfBg        = rule[:bgColor]    ok
                                    if rule[:textColor]  != NULL  cfTextColor = rule[:textColor]  ok
                                    if rule[:bold]       != NULL  cfBold      = rule[:bold]       ok
                                ok
                            ok
                        next
                    ok

                    xml_ += '<w:tc>'
                    xml_ += '<w:tcPr>'
                    xml_ += '<w:tcW w:w="' + colWidth + '" w:type="dxa"/>'

                    # Cell background: conditional > per-row > header > even-row stripe
                    rowColor = ""
                    if rowIdx <= len(aRowBgColors)
                        rowColor = aRowBgColors[rowIdx]
                    ok
                    if len(cfBg) > 0
                        xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + wordColorToHex(cfBg) + '"/>'
                    elseif isHeader and len(headerBgColor) > 0
                        xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + headerBgColor + '"/>'
                    elseif len(rowColor) > 0
                        xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + wordColorToHex(rowColor) + '"/>'
                    elseif isEvenRow and len(evenRowBgColor) > 0
                        xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + evenRowBgColor + '"/>'
                    ok

                    # Vertical alignment
                    if isHeader and len(headerVAlign) > 0
                        xml_ += '<w:vAlign w:val="' + headerVAlign + '"/>'
                    elseif len(cellVAlign) > 0
                        xml_ += '<w:vAlign w:val="' + cellVAlign + '"/>'
                    ok

                    xml_ += '</w:tcPr>'
                    xml_ += '<w:p>'

                    if isHeader
                        xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
                        if len(headerBgColor) > 0
                            xml_ += '<w:r><w:rPr><w:b/><w:color w:val="' + headerTextColor + '"/></w:rPr><w:t>' + wordXmlEsc(cellText) + '</w:t></w:r>'
                        else
                            xml_ += '<w:r><w:rPr><w:b/></w:rPr><w:t>' + wordXmlEsc(cellText) + '</w:t></w:r>'
                        ok
                    else
                        xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                        # Build run properties for conditional formatting
                        rPr = ""
                        if cfBold  rPr += '<w:b/><w:bCs/>'  ok
                        if len(cfTextColor) > 0  rPr += '<w:color w:val="' + wordColorToHex(cfTextColor) + '"/>'  ok
                        if len(rPr) > 0
                            xml_ += '<w:r><w:rPr>' + rPr + '</w:rPr><w:t>' + wordXmlEsc(cellText) + '</w:t></w:r>'
                        else
                            xml_ += '<w:r><w:t>' + wordXmlEsc(cellText) + '</w:t></w:r>'
                        ok
                    ok
                    
                    xml_ += '</w:p>'
                    xml_ += '</w:tc>'
                ok
            next
            
            xml_ += '</w:tr>'
        next
        
        xml_ += '</w:tbl>'
        
        # Add paragraph after table
        xml_ += '<w:p/>'
        
        return xml_

    func safeNumberFromCell str
        /*
            Safely parse a cell string as a number.
            Strips common financial decoration ($, commas, %, +) before parsing.
            Returns 0 when the string is not numeric rather than throwing.
        */
        s = trim("" + str)
        s = substr(s, "$", "")
        s = substr(s, ",", "")
        s = substr(s, "%", "")
        s = trim(s)
        if len(s) = 0  return 0  ok

        # Validate: optional leading sign, then digits with at most one dot
        bOk    = true
        dotCnt = 0
        start  = 1
        if s[1] = "-" or s[1] = "+"  start = 2  ok
        sLen = len(s)
        if start > sLen  return 0  ok
        for k = start to sLen
            ch = s[k]
            if ch = "."
                dotCnt++
                if dotCnt > 1  bOk = false  ok
            elseif ascii(ch) < ascii("0") or ascii(ch) > ascii("9")
                bOk = false
            ok
        next
        if bOk  return number(s)  ok
        return 0

    func evalTableCondition cellText, rule
        /*
            Evaluate one conditional formatting rule against a cell value.
            Returns true if the rule matches (formatting should be applied).

            rule keys:
              :condition  "lt"|"lte"|"gt"|"gte"|"eq"|"neq"|"between"|"contains"
              :value      threshold — numeric type triggers numeric comparison;
                          string type triggers case-insensitive string comparison
              :value2     upper bound for "between"

            Numeric comparison uses safeNumberFromCell() which strips $, commas,
            % signs before parsing — so "$1,234" and "38%" compare correctly.
        */
        cond = lower("" + rule[:condition])
        if cond = NULL or len(cond) = 0  return false  ok

        threshold = rule[:value]
        if threshold = NULL  return false  ok

        # Use numeric comparison only when the threshold is a number literal
        bNumeric     = isNumber(threshold)
        numThreshold = 0
        numCell      = 0
        if bNumeric
            numThreshold = threshold
            numCell      = safeNumberFromCell(cellText)
        ok

        if cond = "lt"
            if bNumeric  return numCell < numThreshold  ok
            return lower(cellText) < lower("" + threshold)
        ok
        if cond = "lte"
            if bNumeric  return numCell <= numThreshold  ok
            return lower(cellText) <= lower("" + threshold)
        ok
        if cond = "gt"
            if bNumeric  return numCell > numThreshold  ok
            return lower(cellText) > lower("" + threshold)
        ok
        if cond = "gte"
            if bNumeric  return numCell >= numThreshold  ok
            return lower(cellText) >= lower("" + threshold)
        ok
        if cond = "eq"
            if bNumeric  return numCell = numThreshold  ok
            return lower(cellText) = lower("" + threshold)
        ok
        if cond = "neq"
            if bNumeric  return numCell != numThreshold  ok
            return lower(cellText) != lower("" + threshold)
        ok
        if cond = "between"
            threshold2 = rule[:value2]
            if threshold2 = NULL  return false  ok
            if bNumeric
                numT2 = safeNumberFromCell("" + threshold2)
                return numCell >= numThreshold and numCell <= numT2
            ok
            return lower(cellText) >= lower("" + threshold) and
                   lower(cellText) <= lower("" + threshold2)
        ok
        if cond = "contains"
            return substr(lower(cellText), lower("" + threshold)) > 0
        ok
        return false

    func generateRichCell cellObj, colWidth, isHeader, headerBgColor, headerTextColor, isEvenRow, evenRowBgColor, headerVAlign, cellVAlign, rowColor
        /*
            Generate XML for a WordCell object inside a table cell.
            Supports: rich text runs, images, alignment, background, vertical align,
            column span, individual borders, padding.
        */
        xml_ = '<w:tc>'
        
        # --- Cell Properties ---
        xml_ += '<w:tcPr>'
        
        # Width
        if isNumber(cellObj.nWidth) and cellObj.nWidth > 0
            xml_ += '<w:tcW w:w="' + cellObj.nWidth + '" w:type="dxa"/>'
        else
            xml_ += '<w:tcW w:w="' + colWidth + '" w:type="dxa"/>'
        ok
        
        # Column span
        if isNumber(cellObj.nColSpan) and cellObj.nColSpan > 1
            xml_ += '<w:gridSpan w:val="' + cellObj.nColSpan + '"/>'
        ok
        
        # Vertical merge (row span)
        if cellObj.cMerge = "restart"
            xml_ += '<w:vMerge w:val="restart"/>'
        elseif cellObj.cMerge = "continue"
            xml_ += '<w:vMerge/>'
        ok
        
        # Background color priority: cell-level > header > per-row > even-row stripe
        if !isString(rowColor) rowColor = "" ok
        bgColor = ""
        if len(cellObj.cBgColor) > 0
            bgColor = cellObj.cBgColor
        elseif isHeader and len(headerBgColor) > 0
            bgColor = headerBgColor
        elseif len(rowColor) > 0
            bgColor = wordColorToHex(rowColor)
        elseif isEvenRow and len(evenRowBgColor) > 0
            bgColor = evenRowBgColor
        ok
        if len(bgColor) > 0
            xml_ += '<w:shd w:val="clear" w:color="auto" w:fill="' + bgColor + '"/>'
        ok
        
        # Vertical alignment: cell-level overrides table-level
        vAlign = ""
        if len(cellObj.cVerticalAlign) > 0
            vAlign = cellObj.cVerticalAlign
        elseif isHeader and len(headerVAlign) > 0
            vAlign = headerVAlign
        elseif len(cellVAlign) > 0
            vAlign = cellVAlign
        ok
        if len(vAlign) > 0
            xml_ += '<w:vAlign w:val="' + vAlign + '"/>'
        ok
        
        # Cell borders: per-side overrides take priority over legacy all-sides
        hasPerSide = (isList(cellObj.aBorderTop)    and len(cellObj.aBorderTop)    > 0) or
                     (isList(cellObj.aBorderBottom) and len(cellObj.aBorderBottom) > 0) or
                     (isList(cellObj.aBorderLeft)   and len(cellObj.aBorderLeft)   > 0) or
                     (isList(cellObj.aBorderRight)  and len(cellObj.aBorderRight)  > 0) or
                     (isList(cellObj.aBorderInsideH) and len(cellObj.aBorderInsideH) > 0) or
                     (isList(cellObj.aBorderInsideV) and len(cellObj.aBorderInsideV) > 0)
        hasLegacy  = len(cellObj.cBorderStyle) > 0
        if hasPerSide or hasLegacy
            xml_ += '<w:tcBorders>'
            # Helper: emit one border side
            # Per-side entry wins over legacy; legacy applies to all four sides
            aSides = [
                ["top",     cellObj.aBorderTop],
                ["left",    cellObj.aBorderLeft],
                ["bottom",  cellObj.aBorderBottom],
                ["right",   cellObj.aBorderRight],
                ["insideH", cellObj.aBorderInsideH],
                ["insideV", cellObj.aBorderInsideV]
            ]
            legacyStyle = cellObj.cBorderStyle
            legacyColor = cellObj.cBorderColor
            if len(legacyColor) = 0  legacyColor = "auto"  ok
            for aSide in aSides
                sideName = aSide[1]
                sideData = aSide[2]
                if isList(sideData) and len(sideData) > 0
                    sStyle = sideData[:style]
                    sColor = sideData[:color]
                    sSize  = sideData[:size]
                    if sStyle = NULL  sStyle = "single"  ok
                    if sColor = NULL  sColor = "auto"    ok
                    if sSize  = NULL  sSize  = 4         ok
                    if sStyle = "none"
                        xml_ += '<w:' + sideName + ' w:val="none"/>'
                    else
                        xml_ += '<w:' + sideName + ' w:val="' + sStyle + '" w:sz="' + string(sSize) + '" w:space="0" w:color="' + sColor + '"/>'
                    ok
                elseif hasLegacy
                    # insideH/insideV don't get legacy fallback
                    if sideName != "insideH" and sideName != "insideV"
                        xml_ += '<w:' + sideName + ' w:val="' + legacyStyle + '" w:sz="4" w:space="0" w:color="' + legacyColor + '"/>'
                    ok
                ok
            next
            xml_ += '</w:tcBorders>'
        ok

        # Cell text direction
        if len(cellObj.cTextDir) > 0 and cellObj.cTextDir != "lrTb"
            xml_ += '<w:textDirection w:val="' + cellObj.cTextDir + '"/>'
        ok
        
        # Cell margins/padding
        if (isNumber(cellObj.nPaddingTop)    and cellObj.nPaddingTop > 0)    or
           (isNumber(cellObj.nPaddingBottom) and cellObj.nPaddingBottom > 0) or
           (isNumber(cellObj.nPaddingLeft)   and cellObj.nPaddingLeft > 0)   or
           (isNumber(cellObj.nPaddingRight)  and cellObj.nPaddingRight > 0)
            xml_ += '<w:tcMar>'
            if isNumber(cellObj.nPaddingTop) and cellObj.nPaddingTop > 0
                xml_ += '<w:top w:w="' + cellObj.nPaddingTop + '" w:type="dxa"/>'
            ok
            if isNumber(cellObj.nPaddingLeft) and cellObj.nPaddingLeft > 0
                xml_ += '<w:left w:w="' + cellObj.nPaddingLeft + '" w:type="dxa"/>'
            ok
            if isNumber(cellObj.nPaddingBottom) and cellObj.nPaddingBottom > 0
                xml_ += '<w:bottom w:w="' + cellObj.nPaddingBottom + '" w:type="dxa"/>'
            ok
            if isNumber(cellObj.nPaddingRight) and cellObj.nPaddingRight > 0
                xml_ += '<w:right w:w="' + cellObj.nPaddingRight + '" w:type="dxa"/>'
            ok
            xml_ += '</w:tcMar>'
        ok
        
        xml_ += '</w:tcPr>'
        
        # --- Cell Content ---
        # Determine alignment
        cellAlign = ""
        if len(cellObj.cAlign) > 0
            cellAlign = cellObj.cAlign
        elseif isHeader
            cellAlign = "center"
        ok
        
        # Check if we have any images in the cell
        nCellImgs = cellObj.nCellImageCount
        if !isNumber(nCellImgs) nCellImgs = 0 ok
        hasImages = nCellImgs > 0
        
        # Generate text paragraph
        runCount = cellObj.nRunCount
        if !isNumber(runCount) runCount = 0 ok
        if runCount > 0 or !hasImages
            xml_ += '<w:p>'
            
            # Paragraph properties (alignment + zero spacing for vAlign)
            if len(cellAlign) > 0
                xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="' + cellAlign + '"/></w:pPr>'
            else
                xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            ok
            
            # Text runs
            for runIdx = 1 to runCount
                runText = cellObj.getRunText(runIdx)
                
                # Check for line break marker
                if runText = "__LINEBREAK__"
                    xml_ += '<w:r><w:br/></w:r>'
                    loop
                ok
                
                xml_ += '<w:r>'
                
                # Build run properties from parallel arrays
                rPr = ""
                
                # Check if this run has any formatting
                hasFormatting = false
                if cellObj.aRunBold[runIdx] = 1 hasFormatting = true ok
                if cellObj.aRunItalic[runIdx] = 1 hasFormatting = true ok
                if cellObj.aRunUnderline[runIdx] = 1 hasFormatting = true ok
                if cellObj.aRunStrike[runIdx] = 1 hasFormatting = true ok
                if len(cellObj.aRunFont[runIdx]) > 0 hasFormatting = true ok
                if cellObj.aRunSize[runIdx] > 0 hasFormatting = true ok
                if len(cellObj.aRunColor[runIdx]) > 0 hasFormatting = true ok
                if len(cellObj.aRunHighlight[runIdx]) > 0 hasFormatting = true ok
                
                # For header rows without explicit formatting, add bold + white
                if isHeader and !hasFormatting and len(headerBgColor) > 0
                    rPr += '<w:b/>'
                    rPr += '<w:color w:val="' + headerTextColor + '"/>'
                elseif isHeader and !hasFormatting
                    rPr += '<w:b/>'
                else
                    # Build from parallel arrays
                    if cellObj.aRunBold[runIdx] = 1
                        rPr += '<w:b/>'
                    ok
                    if cellObj.aRunItalic[runIdx] = 1
                        rPr += '<w:i/>'
                    ok
                    if cellObj.aRunUnderline[runIdx] = 1
                        rPr += '<w:u w:val="single"/>'
                    ok
                    if cellObj.aRunStrike[runIdx] = 1
                        rPr += '<w:strike/>'
                    ok
                    if len(cellObj.aRunFont[runIdx]) > 0
                        rPr += '<w:rFonts w:ascii="' + cellObj.aRunFont[runIdx] + '" w:hAnsi="' + cellObj.aRunFont[runIdx] + '"/>'
                    ok
                    if cellObj.aRunSize[runIdx] > 0
                        halfPoints = cellObj.aRunSize[runIdx] * 2
                        rPr += '<w:sz w:val="' + halfPoints + '"/>'
                        rPr += '<w:szCs w:val="' + halfPoints + '"/>'
                    ok
                    if len(cellObj.aRunColor[runIdx]) > 0
                        rPr += '<w:color w:val="' + wordColorToHex(cellObj.aRunColor[runIdx]) + '"/>'
                    ok
                    if len(cellObj.aRunHighlight[runIdx]) > 0
                        rPr += '<w:highlight w:val="' + cellObj.aRunHighlight[runIdx] + '"/>'
                    ok
                ok
                
                if len(rPr) > 0
                    xml_ += '<w:rPr>' + rPr + '</w:rPr>'
                ok
                
                # Text content
                if runText != NULL and len(runText) > 0
                    if left(runText, 1) = " " or right(runText, 1) = " "
                        xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(runText) + '</w:t>'
                    else
                        xml_ += '<w:t>' + wordXmlEsc(runText) + '</w:t>'
                    ok
                ok
                
                xml_ += '</w:r>'
            next
            
            xml_ += '</w:p>'
        ok
        
        # Generate image paragraphs
        for imgIdx = 1 to nCellImgs
            relId = cellObj.aCellImgRelIds[imgIdx]
            imgId = cellObj.aCellImgIds[imgIdx]
            widthEmu = floor(cellObj.aCellImgWidths[imgIdx] * 360000)
            heightEmu = floor(cellObj.aCellImgHeights[imgIdx] * 360000)
            
            if len(relId) > 0 and imgId > 0
                xml_ += '<w:p>'
                if len(cellAlign) > 0
                    xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="' + cellAlign + '"/></w:pPr>'
                else
                    xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                ok
                xml_ += '<w:r>'
                xml_ += '<w:drawing>'
                xml_ += '<wp:inline distT="0" distB="0" distL="0" distR="0">'
                xml_ += '<wp:extent cx="' + widthEmu + '" cy="' + heightEmu + '"/>'
                xml_ += '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
                xml_ += '<wp:docPr id="' + imgId + '" name="Picture ' + imgId + '"/>'
                xml_ += '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                xml_ += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                xml_ += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                xml_ += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                xml_ += '<pic:nvPicPr>'
                xml_ += '<pic:cNvPr id="' + imgId + '" name="Picture ' + imgId + '"/>'
                xml_ += '<pic:cNvPicPr/>'
                xml_ += '</pic:nvPicPr>'
                xml_ += '<pic:blipFill>'
                xml_ += '<a:blip r:embed="' + relId + '"/>'
                xml_ += '<a:stretch><a:fillRect/></a:stretch>'
                xml_ += '</pic:blipFill>'
                xml_ += '<pic:spPr>'
                xml_ += '<a:xfrm>'
                xml_ += '<a:off x="0" y="0"/>'
                xml_ += '<a:ext cx="' + widthEmu + '" cy="' + heightEmu + '"/>'
                xml_ += '</a:xfrm>'
                xml_ += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                xml_ += '</pic:spPr>'
                xml_ += '</pic:pic>'
                xml_ += '</a:graphicData>'
                xml_ += '</a:graphic>'
                xml_ += '</wp:inline>'
                xml_ += '</w:drawing>'
                xml_ += '</w:r>'
                xml_ += '</w:p>'
            ok
        next
        
        # Generate lists inside the cell
        xml_ += generateCellLists(cellObj)
        
        # Generate hyperlinks inside the cell
        xml_ += generateCellHyperlinks(cellObj)
        
        # Generate nested tables inside the cell
        xml_ += generateCellTables(cellObj)
        
        # Generate block quotes and captions inside the cell
        xml_ += generateCellBlocks(cellObj)
        
        xml_ += '</w:tc>'
        return xml_

    func generateCellLists cellObj
        /*
            Generate XML for lists inside a cell.
        */
        xml_ = ""
        nLists = cellObj.nCellListCount
        if !isNumber(nLists) return c ok
        for li = 1 to nLists
            listId = cellObj.aCellListIds[li]
            cItems = cellObj.aCellListItems[li]
            aItems = str2list(cItems)
            itemsLen = len(aItems)
            for ii = 1 to itemsLen
                xml_ += '<w:p>'
                xml_ += '<w:pPr>'
                xml_ += '<w:pStyle w:val="ListParagraph"/>'
                xml_ += '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                xml_ += '<w:numPr>'
                xml_ += '<w:ilvl w:val="0"/>'
                xml_ += '<w:numId w:val="' + listId + '"/>'
                xml_ += '</w:numPr>'
                xml_ += '</w:pPr>'
                xml_ += '<w:r><w:t>' + wordXmlEsc(aItems[ii]) + '</w:t></w:r>'
                xml_ += '</w:p>'
            next
        next
        return xml_

    func generateCellHyperlinks cellObj
        /*
            Generate XML for hyperlinks inside a cell.
        */
        xml_ = ""
        nLinks = cellObj.nCellLinkCount
        if !isNumber(nLinks) return c ok
        for hi = 1 to nLinks
            relId = cellObj.aCellLinkRelIds[hi]
            text = cellObj.aCellLinkTexts[hi]
            xml_ += '<w:p>'
            xml_ += '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            xml_ += '<w:hyperlink r:id="' + relId + '">'
            xml_ += '<w:r>'
            xml_ += '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
            xml_ += '<w:t>' + wordXmlEsc(text) + '</w:t>'
            xml_ += '</w:r>'
            xml_ += '</w:hyperlink>'
            xml_ += '</w:p>'
        next
        return xml_

    func generateCellTables cellObj
        /*
            Generate XML for nested tables inside a cell.
        */
        xml_ = ""
        nTables = cellObj.nCellTableCount
        if !isNumber(nTables) return c ok
        for ti = 1 to nTables
            xml_ += cellObj.aCellTableXml[ti]
            # Word requires a paragraph after a nested table inside a cell
            xml_ += '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'
        next
        return xml_

    func generateCellBlocks cellObj
        /*
            Generate XML for block quotes and captions inside a cell.
        */
        xml_ = ""
        nBlocks = cellObj.nCellBlockCount
        if !isNumber(nBlocks) return c ok
        for bi = 1 to nBlocks
            bText = cellObj.aCellBlockTexts[bi]
            bType = cellObj.aCellBlockTypes[bi]
            if bType = "quote"
                # Block quote: left indent, italic, gray, left border effect
                xml_ += '<w:p>'
                xml_ += '<w:pPr>'
                xml_ += '<w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'
                xml_ += '<w:ind w:left="284"/>'
                xml_ += '<w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="BBBBBB"/></w:pBdr>'
                xml_ += '</w:pPr>'
                xml_ += '<w:r><w:rPr><w:i/><w:color w:val="555555"/></w:rPr>'
                xml_ += '<w:t>' + wordXmlEsc(bText) + '</w:t>'
                xml_ += '</w:r></w:p>'
            elseif bType = "caption"
                # Caption: centered, italic, smaller, gray
                xml_ += '<w:p>'
                xml_ += '<w:pPr>'
                xml_ += '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>'
                xml_ += '<w:jc w:val="center"/>'
                xml_ += '</w:pPr>'
                xml_ += '<w:r><w:rPr><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="666666"/></w:rPr>'
                xml_ += '<w:t>' + wordXmlEsc(bText) + '</w:t>'
                xml_ += '</w:r></w:p>'
            ok
        next
        return xml_

    func generateField fieldType, cachedValue, options
        if fieldType = NULL  fieldType = ""  ok
        if cachedValue = NULL  cachedValue = ""  ok
        if !isList(options)  options = []  ok
        pPr = generateParagraphProperties(options)
        xml_ = '<w:p>' + pPr
        xml_ += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        xml_ += '<w:r><w:instrText xml:space="preserve"> ' + fieldType + ' </w:instrText></w:r>'
        xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        if len(cachedValue) > 0
            xml_ += '<w:r><w:t xml:space="preserve">' + wordXmlEsc(cachedValue) + '</w:t></w:r>'
        ok
        xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        xml_ += '</w:p>'
        return xml_

    func generateHyperlink text, relId
        xml_ = '<w:p>'
        xml_ += '<w:hyperlink r:id="' + relId + '">'
        xml_ += '<w:r>'
        xml_ += '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
        xml_ += '<w:t>' + wordXmlEsc(text) + '</w:t>'
        xml_ += '</w:r>'
        xml_ += '</w:hyperlink>'
        xml_ += '</w:p>'
        return xml_
    
    func generateSrcRect item
        /*
            Build <a:srcRect> XML from crop percentages stored on image item.
            OOXML unit = thousandths-of-a-percent (1% = 1000).
            Returns empty string when no crop is set.
        */
        cL = item[:cropL]
        cR = item[:cropR]
        cT = item[:cropT]
        cB = item[:cropB]
        hasCrop = (cL != NULL and isNumber(cL) and cL > 0) or
                  (cR != NULL and isNumber(cR) and cR > 0) or
                  (cT != NULL and isNumber(cT) and cT > 0) or
                  (cB != NULL and isNumber(cB) and cB > 0)
        if !hasCrop  return ""  ok
        if cL = NULL or !isNumber(cL)  cL = 0  ok
        if cR = NULL or !isNumber(cR)  cR = 0  ok
        if cT = NULL or !isNumber(cT)  cT = 0  ok
        if cB = NULL or !isNumber(cB)  cB = 0  ok
        # Convert % to OOXML thousandths-of-a-percent
        lV = string(floor(cL * 1000))
        rV = string(floor(cR * 1000))
        tV = string(floor(cT * 1000))
        bV = string(floor(cB * 1000))
        return '<a:srcRect l="' + lV + '" r="' + rV + '" t="' + tV + '" b="' + bV + '"/>'

    func generateImage item
        relId   = item[:relId]
        width   = item[:width]
        height  = item[:height]
        imageId = item[:imageId]
        centered = false
        if item[:centered] = true
            centered = true
        ok

        srcRectXml = generateSrcRect(item)

        xml_ = '<w:p>'
        
        # Center alignment if requested
        if centered
            xml_ += '<w:pPr><w:jc w:val="center"/></w:pPr>'
        ok
        
        xml_ += '<w:r>'
        xml_ += '<w:drawing>'
        xml_ += '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        xml_ += '<wp:extent cx="' + width + '" cy="' + height + '"/>'
        xml_ += '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        imgAltG = item[:altText]
        if imgAltG = NULL  imgAltG = ""  ok
        if len(imgAltG) > 0
            xml_ += '<wp:docPr id="' + imageId + '" name="Picture ' + imageId + '" descr="' + wordXmlEsc(imgAltG) + '"/>'
        else
            xml_ += '<wp:docPr id="' + imageId + '" name="Picture ' + imageId + '"/>'
        ok
        xml_ += '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        xml_ += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        xml_ += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        xml_ += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        xml_ += '<pic:nvPicPr>'
        xml_ += '<pic:cNvPr id="' + imageId + '" name="Picture ' + imageId + '"/>'
        xml_ += '<pic:cNvPicPr/>'
        xml_ += '</pic:nvPicPr>'
        xml_ += '<pic:blipFill>'
        xml_ += '<a:blip r:embed="' + relId + '"/>'
        if len(srcRectXml) > 0
            xml_ += srcRectXml
            xml_ += '<a:stretch><a:fillRect/></a:stretch>'
        else
            xml_ += '<a:stretch><a:fillRect/></a:stretch>'
        ok
        xml_ += '</pic:blipFill>'
        xml_ += '<pic:spPr>'
        xml_ += '<a:xfrm>'
        xml_ += '<a:off x="0" y="0"/>'
        xml_ += '<a:ext cx="' + width + '" cy="' + height + '"/>'
        xml_ += '</a:xfrm>'
        xml_ += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        xml_ += '</pic:spPr>'
        xml_ += '</pic:pic>'
        xml_ += '</a:graphicData>'
        xml_ += '</a:graphic>'
        xml_ += '</wp:inline>'
        xml_ += '</w:drawing>'
        xml_ += '</w:r>'
        xml_ += '</w:p>'
        
        return xml_
    
    func generatePageBordersXml
        /*
            Generates <w:pgBorders> XML for insertion inside <w:sectPr>.
            Respects aPageBorderSides to draw only the requested sides.
        */
        bdrAttr  = ' w:val="' + cPageBorderStyle + '"'
        bdrAttr += ' w:sz="' + nPageBorderSize + '"'
        bdrAttr += ' w:space="' + nPageBorderSpace + '"'
        bdrAttr += ' w:color="' + cPageBorderColor + '"'
        
        xml_ = '<w:pgBorders w:offsetFrom="' + cPageBorderOffsetFrom + '">'
        
        sidesLen = len(aPageBorderSides)
        for si = 1 to sidesLen
            side = lower(aPageBorderSides[si])
            if side = "top" or side = "left" or side = "bottom" or side = "right"
                xml_ += '<w:' + side + bdrAttr + '/>'
            ok
        next
        
        xml_ += '</w:pgBorders>'
        return xml_

    func writeStyles tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        
        # Document defaults
        c += '<w:docDefaults>'
        c += '<w:rPrDefault><w:rPr>'
        c += '<w:rFonts w:ascii="' + cDefaultFont + '" w:hAnsi="' + cDefaultFont + '" w:eastAsia="' + cDefaultFont + '" w:cs="' + cDefaultFont + '"/>'
        c += '<w:sz w:val="' + (nDefaultSize * 2) + '"/>'
        c += '<w:szCs w:val="' + (nDefaultSize * 2) + '"/>'
        # Include w:bidi language so Word's engine activates RTL text shaping
        if bDocumentRTL
            c += '<w:lang w:val="en-US" w:bidi="ar-SA"/>'
        else
            c += '<w:lang w:val="en-US"/>'
        ok
        c += '</w:rPr></w:rPrDefault>'
        c += '<w:pPrDefault><w:pPr>'
        if bDocumentRTL
            c += '<w:bidi/>'   # MUST be first — before spacing — this is the doc-level RTL master switch
        ok
        c += '<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>'
        c += '</w:pPr></w:pPrDefault>'
        c += '</w:docDefaults>'
        
        # Normal style
        c += '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        c += '<w:name w:val="Normal"/>'
        c += '<w:qFormat/>'
        if bDocumentRTL
            # pPrDefault already carries <w:bidi/> — Normal inherits it.
            # Only set alignment + run direction here.
            c += '<w:pPr><w:jc w:val="right"/></w:pPr>'
            c += '<w:rPr><w:rtl/></w:rPr>'
        ok
        c += '</w:style>'
        
        # Heading styles — no bidi/jc/rtl needed; all inherited from pPrDefault + Normal
        for i = 1 to 6
            sizes = [32, 26, 24, 22, 20, 18]
            # Heading color: use theme accent1 for H1/H2, accent2 for H3/H4, accent3 for H5/H6
            # Fall back to a fixed dark color if no theme
            hColor = ""
            if isList(aThemeColors) and len(aThemeColors) >= 5
                if i <= 2
                    hColor = aThemeColors[5]   # accent1
                elseif i <= 4
                    hColor = aThemeColors[6]   # accent2
                else
                    hColor = aThemeColors[3]   # dk2
                ok
            ok
            c += '<w:style w:type="paragraph" w:styleId="Heading' + i + '">'
            c += '<w:name w:val="heading ' + i + '"/>'
            c += '<w:basedOn w:val="Normal"/>'
            c += '<w:next w:val="Normal"/>'
            c += '<w:qFormat/>'
            c += '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="' + (i-1) + '"/></w:pPr>'
            rPrH = '<w:b/><w:sz w:val="' + (sizes[i] * 2) + '"/><w:szCs w:val="' + (sizes[i] * 2) + '"/>'
            if len(hColor) > 0
                rPrH += '<w:color w:val="' + hColor + '"/>'
            ok
            c += '<w:rPr>' + rPrH + '</w:rPr>'
            c += '</w:style>'
        next
        
        # List paragraph style — ind direction handled in numbering.xml, not here
        c += '<w:style w:type="paragraph" w:styleId="ListParagraph">'
        c += '<w:name w:val="List Paragraph"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:qFormat/>'
        if bDocumentRTL
            c += '<w:pPr><w:ind w:right="720"/><w:contextualSpacing/></w:pPr>'
        else
            c += '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>'
        ok
        c += '</w:style>'
        
        # Table Grid style
        c += '<w:style w:type="table" w:styleId="TableGrid">'
        c += '<w:name w:val="Table Grid"/>'
        c += '<w:basedOn w:val="TableNormal"/>'
        c += '<w:tblPr>'
        c += '<w:tblBorders>'
        c += '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        c += '</w:tblBorders>'
        c += '</w:tblPr>'
        c += '</w:style>'
        
        # Table Normal style
        c += '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
        c += '<w:name w:val="Normal Table"/>'
        c += '<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        c += '</w:style>'
        
        # Hyperlink style
        c += '<w:style w:type="character" w:styleId="Hyperlink">'
        c += '<w:name w:val="Hyperlink"/>'
        c += '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
        c += '</w:style>'
        
        # Footnote Text paragraph style
        c += '<w:style w:type="paragraph" w:styleId="FootnoteText">'
        c += '<w:name w:val="footnote text"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        c += '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        c += '</w:style>'
        
        # Footnote Reference character style
        c += '<w:style w:type="character" w:styleId="FootnoteReference">'
        c += '<w:name w:val="footnote reference"/>'
        c += '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        c += '</w:style>'
        
        # Endnote Text paragraph style
        c += '<w:style w:type="paragraph" w:styleId="EndnoteText">'
        c += '<w:name w:val="endnote text"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        c += '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        c += '</w:style>'
        
        # Endnote Reference character style
        c += '<w:style w:type="character" w:styleId="EndnoteReference">'
        c += '<w:name w:val="endnote reference"/>'
        c += '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        c += '</w:style>'
        
        # Header paragraph style (required by header XML files)
        c += '<w:style w:type="paragraph" w:styleId="Header">'
        c += '<w:name w:val="header"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:pPr><w:jc w:val="center"/><w:tabs>'
        c += '<w:tab w:val="center" w:pos="4680"/>'
        c += '<w:tab w:val="right" w:pos="9360"/>'
        c += '</w:tabs></w:pPr>'
        c += '</w:style>'
        
        # Footer paragraph style (required by footer XML files)
        c += '<w:style w:type="paragraph" w:styleId="Footer">'
        c += '<w:name w:val="footer"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:pPr><w:jc w:val="center"/><w:tabs>'
        c += '<w:tab w:val="center" w:pos="4680"/>'
        c += '<w:tab w:val="right" w:pos="9360"/>'
        c += '</w:tabs></w:pPr>'
        c += '</w:style>'

        # TOC Heading style (used by addTOC)
        c += '<w:style w:type="paragraph" w:styleId="TOCHeading">'
        c += '<w:name w:val="TOC Heading"/>'
        c += '<w:basedOn w:val="Heading1"/>'
        c += '<w:qFormat/>'
        c += '</w:style>'

        # CommentText paragraph style (used by comments.xml)
        c += '<w:style w:type="paragraph" w:styleId="CommentText">'
        c += '<w:name w:val="annotation text"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:pPr><w:spacing w:after="0"/></w:pPr>'
        c += '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        c += '</w:style>'

        # CommentReference character style (inline balloon marker)
        c += '<w:style w:type="character" w:styleId="CommentReference">'
        c += '<w:name w:val="annotation reference"/>'
        c += '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        c += '</w:style>'

        # PlaceholderText character style (used by SDT content controls)
        c += '<w:style w:type="character" w:styleId="PlaceholderText">'
        c += '<w:name w:val="Placeholder Text"/>'
        c += '<w:rPr><w:color w:val="808080"/></w:rPr>'
        c += '</w:style>'

        # TableOfFigures paragraph style (used by addTableOfFigures/Tables)
        c += '<w:style w:type="paragraph" w:styleId="TableOfFigures">'
        c += '<w:name w:val="table of figures"/>'
        c += '<w:basedOn w:val="Normal"/>'
        c += '<w:next w:val="Normal"/>'
        c += '<w:pPr>'
        c += '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>'
        c += '<w:spacing w:after="80"/>'
        c += '</w:pPr>'
        c += '<w:rPr><w:b/></w:rPr>'
        c += '</w:style>'
        
        # User-defined custom styles
        csLen = len(aCustomStyles)
        for si = 1 to csLen
            st = aCustomStyles[si]
            # Use display name = id if :name not given
            stName = st[:name]
            if stName = NULL stName = st[:id] ok
            stBase = st[:basedOn]
            if stBase = NULL stBase = "Normal" ok
            
            c += '<w:style w:type="paragraph" w:styleId="' + st[:id] + '">'
            c += '<w:name w:val="' + wordXmlEsc(stName) + '"/>'
            c += '<w:basedOn w:val="' + stBase + '"/>'
            c += '<w:qFormat/>'
            
            # Paragraph properties
            pPrContent = ""
            if st[:align] != NULL
                pPrContent += '<w:jc w:val="' + st[:align] + '"/>'
            ok
            if st[:spaceBefore] != NULL or st[:spaceAfter] != NULL or st[:lineSpacing] != NULL
                pPrContent += '<w:spacing'
                if st[:spaceBefore] != NULL
                    pPrContent += ' w:before="' + st[:spaceBefore] + '"'
                ok
                if st[:spaceAfter] != NULL
                    pPrContent += ' w:after="' + st[:spaceAfter] + '"'
                ok
                if st[:lineSpacing] != NULL
                    pPrContent += ' w:line="' + floor(st[:lineSpacing] * 240) + '" w:lineRule="auto"'
                ok
                pPrContent += '/>'
            ok
            if st[:indent] != NULL
                pPrContent += '<w:ind w:left="' + st[:indent] + '"/>'
            ok
            if st[:bgColor] != NULL
                pPrContent += '<w:shd w:val="clear" w:color="auto" w:fill="' + wordColorToHex(st[:bgColor]) + '"/>'
            ok
            if st[:keepNext] = true
                pPrContent += '<w:keepNext/>'
            ok
            if st[:keepLines] = true
                pPrContent += '<w:keepLines/>'
            ok
            if len(pPrContent) > 0
                c += '<w:pPr>' + pPrContent + '</w:pPr>'
            ok
            
            # Run properties
            rPrContent = ""
            if st[:bold] = true
                rPrContent += '<w:b/>'
            ok
            if st[:italic] = true
                rPrContent += '<w:i/>'
            ok
            if st[:underline] = true
                rPrContent += '<w:u w:val="single"/>'
            ok
            if st[:font] != NULL
                rPrContent += '<w:rFonts w:ascii="' + st[:font] + '" w:hAnsi="' + st[:font] + '"/>'
            ok
            if st[:size] != NULL
                rPrContent += '<w:sz w:val="' + (st[:size] * 2) + '"/>'
                rPrContent += '<w:szCs w:val="' + (st[:size] * 2) + '"/>'
            ok
            if st[:color] != NULL
                rPrContent += '<w:color w:val="' + wordColorToHex(st[:color]) + '"/>'
            ok
            if len(rPrContent) > 0
                c += '<w:rPr>' + rPrContent + '</w:rPr>'
            ok
            
            c += '</w:style>'
        next
        
        c += '</w:styles>'
        
        write(tempDir + "word" + sep + "styles.xml", c)
    
    func writeNumbering tempDir
        sep = wordGetSep()
        
        # Direction-dependent values
        if bDocumentRTL
            lvlJc = "right"
        else
            lvlJc = "left"
        ok
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        
        # Abstract numbering definitions
        numCount = len(aNumbering)
        for ni = 1 to numCount
            numDef = aNumbering[ni]
            absId  = numDef[:id]
            isMulti = numDef[:multilevel] = true
            
            c += '<w:abstractNum w:abstractNumId="' + absId + '">'
            c += '<w:multiLevelType w:val="hybridMultilevel"/>'
            
            if isMulti
                # Define 9 levels (0-8) for nested lists
                for lvl = 0 to 8
                    c += '<w:lvl w:ilvl="' + lvl + '">'
                    # use startAt for level 0 restart support
                    if lvl = 0 and numDef[:startAt] != NULL and isNumber(numDef[:startAt]) and numDef[:startAt] > 1
                        c += '<w:start w:val="' + string(numDef[:startAt]) + '"/>'
                    else
                        c += '<w:start w:val="1"/>'
                    ok
                    indent  = 720 + lvl * 360   # 720 base + 360 per level
                    hanging = 360
                    
                    if bDocumentRTL
                        indXml = '<w:ind w:right="' + indent + '" w:hanging="' + hanging + '"/>'
                    else
                        indXml = '<w:ind w:left="' + indent + '" w:hanging="' + hanging + '"/>'
                    ok
                    
                    if numDef[:type] = WORD_LIST_BULLET
                        # Cycle bullet characters: ●  ○  ■  ●  ○  ■ ...
                        bulletMod = lvl % 3
                        if bulletMod = 0
                            c += '<w:numFmt w:val="bullet"/>'
                            c += '<w:lvlText w:val="&#61623;"/>'
                            c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                            c += '<w:pPr>' + indXml + '</w:pPr>'
                            c += '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
                        elseif bulletMod = 1
                            c += '<w:numFmt w:val="bullet"/>'
                            c += '<w:lvlText w:val="o"/>'
                            c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                            c += '<w:pPr>' + indXml + '</w:pPr>'
                            c += '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>'
                        else
                            c += '<w:numFmt w:val="bullet"/>'
                            c += '<w:lvlText w:val="&#61607;"/>'
                            c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                            c += '<w:pPr>' + indXml + '</w:pPr>'
                            c += '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>'
                        ok
                    else
                        # Numbered: decimal / lowerLetter / lowerRoman cycling
                        numMod = lvl % 3
                        if numMod = 0
                            c += '<w:numFmt w:val="decimal"/>'
                            c += '<w:lvlText w:val="%' + (lvl + 1) + '."/>'
                        elseif numMod = 1
                            c += '<w:numFmt w:val="lowerLetter"/>'
                            c += '<w:lvlText w:val="%' + (lvl + 1) + '."/>'
                        else
                            c += '<w:numFmt w:val="lowerRoman"/>'
                            c += '<w:lvlText w:val="%' + (lvl + 1) + '."/>'
                        ok
                        c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                        c += '<w:pPr>' + indXml + '</w:pPr>'
                    ok
                    
                    c += '</w:lvl>'
                next
            else
                # Single-level (original behaviour) – level 0 only
                c += '<w:lvl w:ilvl="0">'
                c += '<w:start w:val="1"/>'
                
                if bDocumentRTL
                    indXml = '<w:ind w:right="720" w:hanging="360"/>'
                else
                    indXml = '<w:ind w:left="720" w:hanging="360"/>'
                ok
                
                if numDef[:type] = WORD_LIST_BULLET
                    c += '<w:numFmt w:val="bullet"/>'
                    c += '<w:lvlText w:val="&#61623;"/>'
                    c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                    c += '<w:pPr>' + indXml + '</w:pPr>'
                    c += '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
                else
                    c += '<w:numFmt w:val="decimal"/>'
                    c += '<w:lvlText w:val="%1."/>'
                    c += '<w:lvlJc w:val="' + lvlJc + '"/>'
                    c += '<w:pPr>' + indXml + '</w:pPr>'
                ok
                
                c += '</w:lvl>'
            ok
            
            c += '</w:abstractNum>'
        next
        
        # Numbering instances
        for ni = 1 to numCount
            numDef = aNumbering[ni]
            c += '<w:num w:numId="' + numDef[:id] + '">'
            c += '<w:abstractNumId w:val="' + numDef[:id] + '"/>'
            c += '</w:num>'
        next
        
        c += '</w:numbering>'
        
        write(tempDir + "word" + sep + "numbering.xml", c)
    
    func writeSettings tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        c += '<w:zoom w:percent="100"/>'
        c += '<w:defaultTabStop w:val="720"/>'
        c += '<w:characterSpacingControl w:val="doNotCompress"/>'
        # Tell Word the complex-script (BiDi) language when document is RTL.
        # Without this Word ignores <w:bidi/> markers and keeps LTR rendering.
        if bDocumentRTL
            c += '<w:themeFontLang w:val="en-US" w:bidi="ar-SA"/>'
        ok
        c += '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        # enable distinct even/odd headers
        if bEvenAndOddHeaders
            c += '<w:evenAndOddHeaders/>'
        ok
        # make background color visible on-screen in Word
        if len(cPageBgColor) > 0
            c += '<w:displayBackgroundShape/>'
        ok
        c += '</w:settings>'
        
        write(tempDir + "word" + sep + "settings.xml", c)
    
    func writeFootnotes tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        
        # Required separator entries (Word expects these)
        c += '<w:footnote w:type="separator" w:id="-1">'
        c += '<w:p><w:r><w:separator/></w:r></w:p>'
        c += '</w:footnote>'
        c += '<w:footnote w:type="continuationSeparator" w:id="0">'
        c += '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
        c += '</w:footnote>'
        
        fnCount = len(aFootnotes)
        for i = 1 to fnCount
            fn = aFootnotes[i]
            c += '<w:footnote w:id="' + fn[:id] + '">'
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            c += '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            # support plain string or rich run list
            fnContent = fn[:content]
            if fnContent = NULL  fnContent = fn[:text]  ok   # backward compat
            if isList(fnContent)
                # Rich runs: list of [text, opts] pairs
                for fnRun in fnContent
                    if isList(fnRun) and len(fnRun) >= 2
                        fnRunText = fnRun[1]
                        fnRunOpts = fnRun[2]
                        if !isList(fnRunOpts)  fnRunOpts = []  ok
                        fnRprXml = generateRunProperties(fnRunOpts)
                        c += '<w:r>'
                        if len(fnRprXml) > 0
                            c += '<w:rPr>' + fnRprXml + '</w:rPr>'
                        ok
                        c += '<w:t xml:space="preserve">' + wordXmlEsc("" + fnRunText) + '</w:t>'
                        c += '</w:r>'
                    ok
                next
            else
                # Plain string
                c += '<w:r><w:t xml:space="preserve"> ' + wordXmlEsc("" + fnContent) + '</w:t></w:r>'
            ok
            c += '</w:p>'
            c += '</w:footnote>'
        next
        
        c += '</w:footnotes>'
        write(tempDir + "word" + sep + "footnotes.xml", c)

    func writeEndnotes tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        
        # Required separator entries
        c += '<w:endnote w:type="separator" w:id="-1">'
        c += '<w:p><w:r><w:separator/></w:r></w:p>'
        c += '</w:endnote>'
        c += '<w:endnote w:type="continuationSeparator" w:id="0">'
        c += '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
        c += '</w:endnote>'
        
        enCount = len(aEndnotes)
        for i = 1 to enCount
            en = aEndnotes[i]
            c += '<w:endnote w:id="' + en[:id] + '">'
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>'
            c += '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>'
            # support plain string or rich run list
            enContent = en[:content]
            if enContent = NULL  enContent = en[:text]  ok   # backward compat
            if isList(enContent)
                for enRun in enContent
                    if isList(enRun) and len(enRun) >= 2
                        enRunText = enRun[1]
                        enRunOpts = enRun[2]
                        if !isList(enRunOpts)  enRunOpts = []  ok
                        enRprXml = generateRunProperties(enRunOpts)
                        c += '<w:r>'
                        if len(enRprXml) > 0
                            c += '<w:rPr>' + enRprXml + '</w:rPr>'
                        ok
                        c += '<w:t xml:space="preserve">' + wordXmlEsc("" + enRunText) + '</w:t>'
                        c += '</w:r>'
                    ok
                next
            else
                c += '<w:r><w:t xml:space="preserve"> ' + wordXmlEsc("" + enContent) + '</w:t></w:r>'
            ok
            c += '</w:p>'
            c += '</w:endnote>'
        next
        
        c += '</w:endnotes>'
        write(tempDir + "word" + sep + "endnotes.xml", c)
    
    func writeFontTable tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        c += '<w:font w:name="' + cDefaultFont + '"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
        c += '<w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        c += '<w:font w:name="Symbol"><w:panose1 w:val="05050102010706020507"/><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        c += '</w:fonts>'
        
        write(tempDir + "word" + sep + "fontTable.xml", c)
    
    func writeCore tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        c += 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        c += 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        c += '<dc:title>' + wordXmlEsc(cTitle) + '</dc:title>'
        c += '<dc:creator>' + wordXmlEsc(cAuthor) + '</dc:creator>'
        c += '<cp:lastModifiedBy>' + wordXmlEsc(cAuthor) + '</cp:lastModifiedBy>'
        c += '<cp:revision>1</cp:revision>'
        c += '<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>'
        c += '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>'
        c += '</cp:coreProperties>'
        
        write(tempDir + "docProps" + sep + "core.xml", c)
    
    func writeApp tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        c += '<Application>RingWordLib</Application>'
        c += '<AppVersion>1.0</AppVersion>'
        c += '<Pages>1</Pages>'
        c += '<Words>0</Words>'
        c += '<Characters>0</Characters>'
        c += '<Lines>1</Lines>'
        c += '<Paragraphs>1</Paragraphs>'
        c += '</Properties>'
        
        write(tempDir + "docProps" + sep + "app.xml", c)
    
    func writeCustomProps tempDir
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
        c += ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        pid = 2
        for cp in aCustomProps
            c += '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            c += ' pid="' + string(pid) + '" name="' + wordXmlEsc(cp[:name]) + '">'
            c += '<vt:lpwstr>' + wordXmlEsc(cp[:value]) + '</vt:lpwstr>'
            c += '</property>'
            pid++
        next
        c += '</Properties>'
        write(tempDir + "docProps" + sep + "custom.xml", c)

    func writeHeader tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        c += 'xmlns:v="urn:schemas-microsoft-com:vml" '
        c += 'xmlns:o="urn:schemas-microsoft-com:office:office" '
        c += 'xmlns:w10="urn:schemas-microsoft-com:office:word">'
        
        # Image watermark paragraph (sits furthest behind everything)
        if bImgWatermark
            c += generateImageWatermarkXml()
        ok
        
        # Text watermark paragraph (sits behind header text)
        if bWatermark
            c += generateWatermarkXml()
        ok
        
        # Regular header text paragraph
        if len(cHeaderText) > 0
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>'
            c += '<w:r><w:t>' + wordXmlEsc(cHeaderText) + '</w:t></w:r>'
            c += '</w:p>'
        ok
        
        # Watermark(s) without header text still need one empty paragraph
        if (bWatermark or bImgWatermark) and len(cHeaderText) = 0
            c += '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>'
        ok
        
        c += '</w:hdr>'
        write(tempDir + "word" + sep + "header1.xml", c)

    func generateWatermarkXml
        /*
            Generates the VML paragraph that renders the diagonal text watermark.
            Anchored absolutely in the header so it repeats on every page.
        */
        opacityInt = floor(nWatermarkOpacity)
        if opacityInt > 100 opacityInt = 100 ok
        if opacityInt < 0   opacityInt = 0   ok
        opacityFrac = opacityInt / 100
        opacityWhole = floor(opacityFrac)
        opacityDec   = floor((opacityFrac - opacityWhole) * 100)
        opacityStr = "" + opacityWhole + "." + right("00" + opacityDec, 2)
        if opacityInt = 100 opacityStr = "1" ok
        
        rotStr    = "" + nWatermarkRotation
        fillColor = "#" + cWatermarkColor
        
        shapeStyle  = "position:absolute;"
        shapeStyle += "left:0;margin-left:0;margin-top:0;"
        shapeStyle += "width:468pt;height:351pt;"
        shapeStyle += "rotation:" + rotStr + ";"
        shapeStyle += "z-index:-251653120;"
        shapeStyle += "mso-position-horizontal:center;"
        shapeStyle += "mso-position-vertical:center"
        
        textStyle  = 'font-family:&quot;' + cWatermarkFont + '&quot;;'
        textStyle += 'font-size:1pt;font-weight:bold;'
        textStyle += 'color:' + fillColor
        
        xml_ = '<w:p>'
        xml_ += '<w:pPr><w:pStyle w:val="Header"/></w:pPr>'
        xml_ += '<w:r><w:rPr><w:noProof/></w:rPr>'
        xml_ += '<w:pict>'
        xml_ += '<v:shape id="_WaterMark_1" o:spid="_x0000_s1025"'
        xml_ += ' type="#_x0000_t136"'
        xml_ += ' style="' + shapeStyle + '"'
        xml_ += ' fillcolor="' + fillColor + '"'
        xml_ += ' stroked="f" filled="t" o:allowoverlap="t">'
        xml_ += '<v:fill on="t" opacity="' + opacityStr + '" focussize="0,0"/>'
        xml_ += '<v:path textpathok="t" o:connecttype="none"/>'
        xml_ += '<v:textpath on="t" string="' + wordXmlEsc(cWatermarkText) + '"'
        xml_ += ' style="' + textStyle + '"/>'
        xml_ += '<w10:wrap type="none"/>'
        xml_ += '<w10:anchorlock/>'
        xml_ += '</v:shape>'
        xml_ += '</w:pict>'
        xml_ += '</w:r>'
        xml_ += '</w:p>'
        
        return xml_

    func generateImageWatermarkXml
        /*
            Generates the VML paragraph that places an image watermark behind all
            page content.  The image is referenced via cImgWatermarkRelId which
            points to an entry in word/_rels/header1.xml.rels.

            Opacity is expressed as a decimal CSS value (0.0 – 1.0) in the VML
            shape style.  Width/height are converted from cm to pt (1 cm = 28.35 pt).
        */
        # Convert opacity 0-100 → decimal string "0.xx"
        opInt = floor(nImgWatermarkOpacity)
        if opInt > 100 opInt = 100 ok
        if opInt < 0   opInt = 0   ok
        opWhole = floor(opInt / 100)
        opDec   = floor((opInt / 100 - opWhole) * 100)
        opStr   = "" + opWhole + "." + right("00" + opDec, 2)
        if opInt = 100 opStr = "1" ok
        
        # Convert cm to pt for VML style (1 cm ≈ 28.35 pt)
        wPt = floor(nImgWatermarkWidth  * 2835 / 100)
        hPt = floor(nImgWatermarkHeight * 2835 / 100)
        
        shapeStyle  = "position:absolute;"
        shapeStyle += "left:0;margin-left:0;margin-top:0;"
        shapeStyle += "width:"  + wPt + "pt;"
        shapeStyle += "height:" + hPt + "pt;"
        shapeStyle += "z-index:-251653120;"
        shapeStyle += "mso-position-horizontal:center;"
        shapeStyle += "mso-position-vertical:center;"
        shapeStyle += "opacity:" + opStr
        
        xml_ = '<w:p>'
        xml_ += '<w:pPr><w:pStyle w:val="Header"/></w:pPr>'
        xml_ += '<w:r><w:rPr><w:noProof/></w:rPr>'
        xml_ += '<w:pict>'
        xml_ += '<v:shape id="_WaterMarkImg_1" o:spid="_x0000_s1026"'
        xml_ += ' type="#_x0000_t75"'
        xml_ += ' style="' + shapeStyle + '"'
        xml_ += ' o:allowoverlap="t">'
        xml_ += '<v:imagedata r:id="' + cImgWatermarkRelId + '" o:title="watermark"/>'
        xml_ += '<w10:wrap type="none"/>'
        xml_ += '<w10:anchorlock/>'
        xml_ += '</v:shape>'
        xml_ += '</w:pict>'
        xml_ += '</w:r>'
        xml_ += '</w:p>'
        
        return xml_

    func writeHeaderRels tempDir
        /*
            Writes word/_rels/header1.xml.rels which maps the watermark image
            relationship ID to its file path inside word/media/.
            This file is only created when an image watermark is active.
        */
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="' + cImgWatermarkRelId + '"'
        c += ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        c += ' Target="media/' + cImgWatermarkFile + '"/>'
        c += '</Relationships>'
        
        write(tempDir + "word" + sep + "_rels" + sep + "header1.xml.rels", c)

    func writeFooter tempDir
        sep = wordGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        c += '<w:p>'
        c += '<w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="' + cPageNumberAlign + '"/></w:pPr>'
        
        # Footer text
        if len(cFooterText) > 0
            c += '<w:r><w:t>' + wordXmlEsc(cFooterText) + '</w:t></w:r>'
        ok
        
        # Page numbers
        if bShowPageNumbers
            if len(cFooterText) > 0
                c += '<w:r><w:t xml:space="preserve"> - </w:t></w:r>'
            ok
            c += '<w:r><w:t xml:space="preserve">Page </w:t></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            c += '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            c += '<w:r><w:t>1</w:t></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            c += '<w:r><w:t xml:space="preserve"> of </w:t></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            c += '<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            c += '<w:r><w:t>1</w:t></w:r>'
            c += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        ok
        
        c += '</w:p>'
        c += '</w:ftr>'
        
        write(tempDir + "word" + sep + "footer1.xml", c)
    
    func writeFirstPageHeader tempDir
        /*
            Write header_fp.xml — the first-page-only header.
            If cFirstPageHeaderText is empty, writes a blank header (no text).
        */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        if len(cFirstPageHeaderText) > 0
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>'
            c += '<w:r><w:t>' + wordXmlEsc(cFirstPageHeaderText) + '</w:t></w:r>'
            c += '</w:p>'
        else
            c += '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>'
        ok
        c += '</w:hdr>'
        write(tempDir + "word" + sep + "header_fp.xml", c)

    func writeFirstPageFooter tempDir
        /*
            Write footer_fp.xml — the first-page-only footer.
            If cFirstPageFooterText is empty, writes a blank footer.
        */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        if len(cFirstPageFooterText) > 0
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
            c += '<w:r><w:t>' + wordXmlEsc(cFirstPageFooterText) + '</w:t></w:r>'
            c += '</w:p>'
        else
            c += '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
        ok
        c += '</w:ftr>'
        write(tempDir + "word" + sep + "footer_fp.xml", c)

    func writeEvenPageHeader tempDir
        /*
            Write header_even.xml — the even-page header.
            If cEvenPageHeaderText is empty, writes a blank header.
        */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        if len(cEvenPageHeaderText) > 0
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>'
            c += '<w:r><w:t>' + wordXmlEsc(cEvenPageHeaderText) + '</w:t></w:r>'
            c += '</w:p>'
        else
            c += '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>'
        ok
        c += '</w:hdr>'
        write(tempDir + "word" + sep + "header_even.xml", c)

    func writeEvenPageFooter tempDir
        /*
            Write footer_even.xml — the even-page footer.
            If cEvenPageFooterText is empty, writes a blank footer.
        */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        if len(cEvenPageFooterText) > 0
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
            c += '<w:r><w:t>' + wordXmlEsc(cEvenPageFooterText) + '</w:t></w:r>'
            c += '</w:p>'
        else
            c += '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
        ok
        c += '</w:ftr>'
        write(tempDir + "word" + sep + "footer_even.xml", c)

    func generateCrossRef item
        /*
            Emits a paragraph with optional prefix text followed by a
            REF field that inserts the page number of the named bookmark.
        */
        bmName   = item[:bookmark]
        dispText = item[:displayText]
        xml_ = '<w:p><w:r>'
        if dispText != NULL and len(dispText) > 0
            xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(dispText) + ' </w:t>'
        ok
        xml_ += '</w:r>'
        # REF field: \p = "on page N" style, \h = hyperlink
        xml_ += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        xml_ += '<w:r><w:instrText xml:space="preserve"> REF ' + wordXmlEsc(bmName) + ' \p \h </w:instrText></w:r>'
        xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        xml_ += '<w:r><w:t>on page ?</w:t></w:r>'
        xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        xml_ += '</w:p>'
        return xml_

    func generateBookmarkAnchor item
        /*
            Emits a zero-width paragraph containing only a bookmark
            start/end pair — useful when you want to mark a spot that
            already has its own content paragraph.
        */
        bmName = item[:bookmark]
        bmId   = nBookmarkId
        nBookmarkId++
        aBookmarks + [:id = bmId, :name = bmName]
        xml_ = '<w:p>'
        xml_ += '<w:bookmarkStart w:id="' + bmId + '" w:name="' + wordXmlEsc(bmName) + '"/>'
        xml_ += '<w:r><w:t></w:t></w:r>'
        xml_ += '<w:bookmarkEnd w:id="' + bmId + '"/>'
        xml_ += '</w:p>'
        return xml_

    func generateTabbedParagraph item
        /*
            Emits a paragraph whose segments are separated by tab characters,
            with the custom tab stops defined in item[:options][:tabStops].
        */
        segments = item[:segments]
        options  = item[:options]
        if !isList(options)  options  = [] ok
        if !isList(segments) segments = [] ok

        xml_ = '<w:p>'

        # Paragraph properties (includes tab stop definitions)
        pPr = generateParagraphProperties(options)
        styleId = options[:style]
        if styleId != NULL
            xml_ += '<w:pPr><w:pStyle w:val="' + styleId + '"/>' + pPr + '</w:pPr>'
        elseif len(pPr) > 0
            xml_ += '<w:pPr>' + pPr + '</w:pPr>'
        ok

        rPr = generateRunProperties(options)

        # Emit each segment separated by <w:tab/>
        segCount = len(segments)
        for si = 1 to segCount
            seg = segments[si]
            xml_ += '<w:r>'
            if len(rPr) > 0
                xml_ += '<w:rPr>' + rPr + '</w:rPr>'
            ok
            if si > 1
                xml_ += '<w:tab/>'
            ok
            if seg != NULL and len(seg) > 0
                if left(seg,1) = " " or right(seg,1) = " "
                    xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(seg) + '</w:t>'
                else
                    xml_ += '<w:t>' + wordXmlEsc(seg) + '</w:t>'
                ok
            ok
            xml_ += '</w:r>'
        next

        xml_ += '</w:p>'
        return xml_

    func generateSeqCaption item
        /*
            Emits a centered paragraph with a SEQ field for auto-numbering.
            Falls back to a local counter for the placeholder text.
            Format: "Figure 1 — caption text"
        */
        seqType = item[:seqType]   # "Figure" or "Table"
        seqNum  = item[:seqNum]
        txt     = item[:text]

        xml_ = '<w:p>'
        xml_ += '<w:pPr><w:jc w:val="center"/>'
        xml_ += '<w:pStyle w:val="Caption"/>'
        xml_ += '</w:pPr>'
        # Literal label
        xml_ += '<w:r><w:rPr><w:b/></w:rPr>'
        xml_ += '<w:t xml:space="preserve">' + seqType + ' </w:t></w:r>'
        # SEQ field
        xml_ += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        xml_ += '<w:r><w:instrText xml:space="preserve"> SEQ ' + seqType + ' \* ARABIC </w:instrText></w:r>'
        xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        xml_ += '<w:r><w:rPr><w:b/></w:rPr><w:t>' + seqNum + '</w:t></w:r>'
        xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        # Separator and caption text
        if txt != NULL and len(txt) > 0
            xml_ += '<w:r><w:t xml:space="preserve"> — ' + wordXmlEsc(txt) + '</w:t></w:r>'
        ok
        xml_ += '</w:p>'
        return xml_


    func generateFloatingImage item
        relId    = item[:relId]
        imgId    = item[:imageId]
        widthEmu = item[:width]
        heightEmu= item[:height]
        posX     = item[:posX]
        posY     = item[:posY]
        distEmu  = item[:distEmu]
        wrapType = item[:wrapType]
        wrapSide = item[:wrapSide]
        imgIdStr = "" + imgId

        # Wrap element XML
        wrapXml = ""
        if wrapType = "wrapSquare"
            wrapXml = '<wp:wrapSquare wrapText="' + wrapSide + '"/>'
        elseif wrapType = "wrapTight"
            wrapXml = '<wp:wrapTight wrapText="' + wrapSide + '"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
        elseif wrapType = "wrapThrough"
            wrapXml = '<wp:wrapThrough wrapText="' + wrapSide + '"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>'
        elseif wrapType = "wrapTopAndBottom"
            wrapXml = '<wp:wrapTopAndBottom/>'
        else
            # wrapNone or any unrecognised value
            wrapXml = '<wp:wrapNone/>'
        ok

        xml_ = '<w:p><w:r><w:drawing>'
        xml_ += '<wp:anchor'
        xml_ += ' distT="' + distEmu + '"'
        xml_ += ' distB="' + distEmu + '"'
        xml_ += ' distL="' + distEmu + '"'
        xml_ += ' distR="' + distEmu + '"'
        xml_ += ' simplePos="0" relativeHeight="251658240"'
        xml_ += ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        xml_ += '<wp:simplePos x="0" y="0"/>'
        xml_ += '<wp:positionH relativeFrom="margin">'
        xml_ +=   '<wp:posOffset>' + posX + '</wp:posOffset>'
        xml_ += '</wp:positionH>'
        xml_ += '<wp:positionV relativeFrom="paragraph">'
        xml_ +=   '<wp:posOffset>' + posY + '</wp:posOffset>'
        xml_ += '</wp:positionV>'
        xml_ += '<wp:extent cx="' + widthEmu + '" cy="' + heightEmu + '"/>'
        xml_ += '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        xml_ += wrapXml
        fltAlt = item[:altText]
        if fltAlt = NULL  fltAlt = ""  ok
        fltAltAttr = ""
        if len(fltAlt) > 0
            fltAltAttr = ' descr="' + wordXmlEsc(fltAlt) + '"'
        ok
        xml_ += '<wp:docPr id="' + imgIdStr + '" name="FloatImg' + imgIdStr + '"' + fltAltAttr + '/>'
        xml_ += '<wp:cNvGraphicFramePr>'
        xml_ +=   '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        xml_ += '</wp:cNvGraphicFramePr>'
        xml_ += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        xml_ +=   '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        xml_ +=     '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        xml_ +=       '<pic:nvPicPr>'
        xml_ +=         '<pic:cNvPr id="' + imgIdStr + '" name="FloatImg' + imgIdStr + '"/>'
        xml_ +=         '<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>'
        xml_ +=       '</pic:nvPicPr>'
        xml_ +=       '<pic:blipFill>'
        xml_ +=         '<a:blip r:embed="' + relId + '"/>'
        fltSrcRect = generateSrcRect(item)
        if len(fltSrcRect) > 0
            xml_ +=     fltSrcRect
        ok
        xml_ +=         '<a:stretch><a:fillRect/></a:stretch>'
        xml_ +=       '</pic:blipFill>'
        xml_ +=       '<pic:spPr>'
        xml_ +=         '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + widthEmu + '" cy="' + heightEmu + '"/></a:xfrm>'
        xml_ +=         '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        xml_ +=       '</pic:spPr>'
        xml_ +=     '</pic:pic>'
        xml_ +=   '</a:graphicData>'
        xml_ += '</a:graphic>'
        xml_ += '</wp:anchor>'
        xml_ += '</w:drawing></w:r></w:p>'
        return xml_

    func generateNumberedHeading item
        text  = item[:text]
        level = item[:level]
        numId = item[:numId]
        ilvl  = item[:ilvl]
        xml_ = '<w:p>'
        xml_ += '<w:pPr>'
        xml_ += '<w:pStyle w:val="Heading' + level + '"/>'
        xml_ += '<w:numPr>'
        xml_ += '<w:ilvl w:val="' + ilvl + '"/>'
        xml_ += '<w:numId w:val="' + numId + '"/>'
        xml_ += '</w:numPr>'
        xml_ += '</w:pPr>'
        xml_ += '<w:r><w:t>' + wordXmlEsc(text) + '</w:t></w:r>'
        xml_ += '</w:p>'
        return xml_

    func generateMergeParagraph item
        /*
            Emits a paragraph containing a mix of literal text runs and
            MERGEFIELD field codes.
            Each element in item[:fields] is either:
              - a string  → emitted as a literal text run
              - a list    → must have :field key → emitted as MERGEFIELD
        */
        fields  = item[:fields]
        options = item[:options]
        if !isList(options)  options  = [] ok
        if !isList(fields)   fields   = [] ok

        xml_ = '<w:p>'
        pPr = generateParagraphProperties(options)
        styleId = options[:style]
        if styleId != NULL
            xml_ += '<w:pPr><w:pStyle w:val="' + styleId + '"/>' + pPr + '</w:pPr>'
        elseif len(pPr) > 0
            xml_ += '<w:pPr>' + pPr + '</w:pPr>'
        ok

        rPr = generateRunProperties(options)
        rPrXml = ""
        if len(rPr) > 0  rPrXml = '<w:rPr>' + rPr + '</w:rPr>'  ok

        fieldsLen = len(fields)
        for fi = 1 to fieldsLen
            seg = fields[fi]
            if isList(seg)
                # MERGEFIELD
                fname = seg[:field]
                if fname = NULL  fname = "Field"  ok
                xml_ += '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="begin"/></w:r>'
                xml_ += '<w:r>' + rPrXml + '<w:instrText xml:space="preserve"> MERGEFIELD ' + wordXmlEsc(fname) + ' </w:instrText></w:r>'
                xml_ += '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="separate"/></w:r>'
                xml_ += '<w:r>' + rPrXml + '<w:t>«' + wordXmlEsc(fname) + '»</w:t></w:r>'
                xml_ += '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="end"/></w:r>'
            else
                # Literal text
                if seg != NULL and len(seg) > 0
                    xml_ += '<w:r>' + rPrXml
                    if left(seg,1) = " " or right(seg,1) = " "
                        xml_ += '<w:t xml:space="preserve">' + wordXmlEsc(seg) + '</w:t>'
                    else
                        xml_ += '<w:t>' + wordXmlEsc(seg) + '</w:t>'
                    ok
                    xml_ += '</w:r>'
                ok
            ok
        next

        xml_ += '</w:p>'
        return xml_

    func writeComments tempDir
        /*
            Write word/comments.xml containing all comments registered
            via addCommentedParagraph().
        */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'

        for cm in aComments
            c += '<w:comment w:id="' + cm[:id] + '"'
            c += ' w:author="' + wordXmlEsc(cm[:author]) + '"'
            c += ' w:date="' + cm[:date] + '"'
            c += ' w:initials="' + upper(left(cm[:author], 1)) + '">'
            c += '<w:p>'
            c += '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
            c += '<w:r>'
            c += '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
            c += '<w:annotationRef/>'
            c += '</w:r>'
            c += '<w:r><w:t>' + wordXmlEsc(cm[:text]) + '</w:t></w:r>'
            c += '</w:p>'
            c += '</w:comment>'
        next

        c += '</w:comments>'
        write(tempDir + "word" + sep + "comments.xml", c)

    # Content Control generators

    func generateSdtCheckbox item
        /*  Checkbox SDT using the legacy checkbox character approach.
            Word displays a checked (✓) or unchecked (☐) character inside
            a locked run, which behaves as an interactive checkbox.  */
        sdtId  = item[:id]
        label  = item[:label]
        isChk  = item[:checked]
        dispCh = ""
        if isChk = true
            dispCh = item[:checkedChar]
        else
            dispCh = item[:uncheckedChar]
        ok

        xml_ = '<w:p>'
        xml_ += '<w:sdt>'
        xml_ += '<w:sdtPr>'
        xml_ += '<w:id w:val="' + sdtId + '"/>'
        xml_ += '<w:tag w:val="checkbox"/>'
        xml_ += '<w:alias w:val="' + wordXmlEsc(label) + '"/>'
        xml_ += '<w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        chkVal = "0"
        if isChk = true  chkVal = "1"  ok
        xml_ += '<w14:checked w14:val="' + chkVal + '"/>'
        xml_ += '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        xml_ += '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        xml_ += '</w14:checkbox>'
        xml_ += '</w:sdtPr>'
        xml_ += '<w:sdtContent>'
        xml_ += '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:cs="MS Gothic"/></w:rPr>'
        xml_ += '<w:t>' + dispCh + '</w:t></w:r>'
        xml_ += '</w:sdtContent>'
        xml_ += '</w:sdt>'
        # Label is OUTSIDE the SDT — Word replaces sdtContent when toggling,
        # so anything inside would be lost. The label run here is permanent.
        if label != NULL and len(label) > 0
            xml_ += '<w:r><w:t xml:space="preserve"> ' + wordXmlEsc(label) + '</w:t></w:r>'
        ok
        xml_ += '</w:p>'
        return xml_

    func generateSdtDropdown item
        sdtId   = item[:id]
        label   = item[:label]
        choices = item[:choices]
        defVal  = item[:default]
        if !isList(choices) choices = [] ok
        if defVal = NULL and len(choices) > 0 defVal = choices[1] ok

        xml_ = ''
        # Label paragraph
        if label != NULL and len(label) > 0
            xml_ += '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>' + wordXmlEsc(label) + '</w:t></w:r></w:p>'
        ok
        # SDT paragraph
        xml_ += '<w:p>'
        xml_ += '<w:sdt>'
        xml_ += '<w:sdtPr>'
        xml_ += '<w:id w:val="' + sdtId + '"/>'
        xml_ += '<w:tag w:val="dropdown"/>'
        xml_ += '<w:alias w:val="' + wordXmlEsc(label) + '"/>'
        xml_ += '<w:dropDownList>'
        for ch in choices
            xml_ += '<w:listItem w:displayText="' + wordXmlEsc(ch) + '" w:value="' + wordXmlEsc(ch) + '"/>'
        next
        xml_ += '</w:dropDownList>'
        xml_ += '</w:sdtPr>'
        xml_ += '<w:sdtContent>'
        xml_ += '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
        xml_ += '<w:t>' + wordXmlEsc(defVal) + '</w:t></w:r>'
        xml_ += '</w:sdtContent>'
        xml_ += '</w:sdt>'
        xml_ += '</w:p>'
        return xml_

    func generateSdtText item
        sdtId  = item[:id]
        label  = item[:label]
        defTxt = item[:default]
        phTxt  = item[:placeholder]
        if defTxt = NULL defTxt = "" ok
        if phTxt  = NULL phTxt  = "Click here to enter text." ok

        xml_ = ''
        # Label paragraph
        if label != NULL and len(label) > 0
            xml_ += '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>' + wordXmlEsc(label) + '</w:t></w:r></w:p>'
        ok
        # SDT paragraph
        xml_ += '<w:p>'
        xml_ += '<w:sdt>'
        xml_ += '<w:sdtPr>'
        xml_ += '<w:id w:val="' + sdtId + '"/>'
        xml_ += '<w:tag w:val="textinput"/>'
        xml_ += '<w:alias w:val="' + wordXmlEsc(label) + '"/>'
        xml_ += '<w:showingPlcHdr/>'
        xml_ += '<w:text/>'
        xml_ += '</w:sdtPr>'
        xml_ += '<w:sdtContent>'
        if len(defTxt) > 0
            xml_ += '<w:r><w:t>' + wordXmlEsc(defTxt) + '</w:t></w:r>'
        else
            xml_ += '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
            xml_ += '<w:t>' + wordXmlEsc(phTxt) + '</w:t></w:r>'
        ok
        xml_ += '</w:sdtContent>'
        xml_ += '</w:sdt>'
        xml_ += '</w:p>'
        return xml_

    # Drawing Shape generator 

    func generateDrawShape item
        /*
            Emits an inline DrawingML shape using <a:sp> inside <wp:inline>.
            Supports: rect, ellipse, line, diamond, triangle.
        */
        opts      = item[:options]
        shpType   = item[:shapeType]
        shpId     = item[:id]
        if !isList(opts) opts = [] ok

        # Dimensions: cm → EMU (1 cm = 360000 EMU)
        wCm = opts[:width]
        hCm = opts[:height]
        if wCm = NULL wCm = 5   ok
        if hCm = NULL hCm = 3   ok
        if shpType = "line" and hCm = 3  hCm = 0.1  ok
        wEmu = floor(wCm * 360000)
        hEmu = floor(hCm * 360000)

        # Colors
        fillColor = opts[:fillColor]
        lineColor = opts[:lineColor]
        lineWidth = opts[:lineWidth]
        if fillColor = NULL fillColor = "4472C4" ok
        if lineColor = NULL lineColor = "2E4E7E" ok
        if lineWidth = NULL lineWidth = 1        ok
        fillColor = wordColorToHex(fillColor)
        lineColor = wordColorToHex(lineColor)
        lineWidthEmu = floor(lineWidth * 12700)  # pt → EMU

        # Text options
        shpText   = opts[:text]
        txtColor  = opts[:textColor]
        txtBold   = opts[:textBold]
        txtSize   = opts[:textSize]
        txtAlign  = opts[:align]
        if txtColor = NULL txtColor = "FFFFFF" ok
        if txtBold  = NULL txtBold  = true     ok
        if txtSize  = NULL txtSize  = 11       ok
        if txtAlign = NULL txtAlign = "center" ok
        txtColor = wordColorToHex(txtColor)
        # DrawingML alignment values differ from WordprocessingML
        dmlAlign = "ctr"
        if txtAlign = "left"  dmlAlign = "l" ok
        if txtAlign = "right" dmlAlign = "r" ok

        # Map shape type to DrawingML preset geometry
        switch shpType
            on "rect"     prst = "rect"
            on "ellipse"  prst = "ellipse"
            on "line"     prst = "rect"   # horizontal rule: thin rect, not DrawingML line connector
            on "diamond"  prst = "diamond"
            on "triangle" prst = "rtTriangle"
            other         prst = "rect"
        off
        if opts[:rounded] = true prst = "roundRect" ok

        # Build the XML using mc:AlternateContent as Word requires for wps shapes
        xml_ = '<w:p><w:pPr><w:jc w:val="' + txtAlign + '"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
        xml_ += '<w:r><w:rPr><w:noProof/></w:rPr>'
        xml_ += '<mc:AlternateContent>'
        xml_ += '<mc:Choice Requires="wps">'
        xml_ += '<w:drawing>'
        xml_ += '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        xml_ += '<wp:extent cx="' + wEmu + '" cy="' + hEmu + '"/>'
        xml_ += '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        xml_ += '<wp:docPr id="' + shpId + '" name="Shape ' + shpId + '"/>'
        xml_ += '<wp:cNvGraphicFramePr/>'
        xml_ += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        xml_ += '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        xml_ += '<wps:wsp>'
        xml_ += '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        xml_ += '<wps:spPr>'
        xml_ += '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + wEmu + '" cy="' + hEmu + '"/></a:xfrm>'
        xml_ += '<a:prstGeom prst="' + prst + '"><a:avLst/></a:prstGeom>'
        # Fill
        if opts[:noFill] = true
            xml_ += '<a:noFill/>'
        else
            xml_ += '<a:solidFill><a:srgbClr val="' + fillColor + '"/></a:solidFill>'
        ok
        # Line
        if opts[:noBorder] = true
            xml_ += '<a:ln><a:noFill/></a:ln>'
        else
            xml_ += '<a:ln w="' + lineWidthEmu + '">'
            xml_ += '<a:solidFill><a:srgbClr val="' + lineColor + '"/></a:solidFill>'
            xml_ += '</a:ln>'
        ok
        xml_ += '</wps:spPr>'
        # Text body
        if shpText != NULL and len(shpText) > 0
            xml_ += '<wps:txbx>'
            xml_ += '<w:txbxContent>'
            xml_ += '<w:p>'
            xml_ += '<w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
            xml_ += '<w:r><w:rPr>'
            if txtBold = true  xml_ += '<w:b/>'  ok
            xml_ += '<w:color w:val="' + txtColor + '"/>'
            xml_ += '<w:sz w:val="' + (txtSize * 2) + '"/>'
            xml_ += '<w:szCs w:val="' + (txtSize * 2) + '"/>'
            xml_ += '</w:rPr>'
            xml_ += '<w:t>' + wordXmlEsc(shpText) + '</w:t>'
            xml_ += '</w:r></w:p>'
            xml_ += '</w:txbxContent>'
            xml_ += '</wps:txbx>'
            xml_ += '<wps:bodyPr anchor="ctr"/>'
        else
            xml_ += '<wps:bodyPr/>'
        ok
        xml_ += '</wps:wsp>'
        xml_ += '</a:graphicData>'
        xml_ += '</a:graphic>'
        xml_ += '</wp:inline>'
        xml_ += '</w:drawing>'
        xml_ += '</mc:Choice>'
        # VML fallback for older Word versions
        wPt = floor(wCm * 28.35)
        hPt = floor(hCm * 28.35)
        xml_ += '<mc:Fallback>'
        xml_ += '<w:pict>'
        xml_ += '<v:rect style="width:' + wPt + 'pt;height:' + hPt + 'pt">'
        xml_ += '<v:fill color="#' + fillColor + '"/>'
        xml_ += '</v:rect>'
        xml_ += '</w:pict>'
        xml_ += '</mc:Fallback>'
        xml_ += '</mc:AlternateContent>'
        xml_ += '</w:r></w:p>'
        return xml_

    # Theme writer

    func writeTheme tempDir
        /*
            Write word/theme/theme1.xml using the active theme color palette.
        */
        sep = wordGetSep()
        dk1  = aThemeColors[1]
        lt1  = aThemeColors[2]
        dk2  = aThemeColors[3]
        lt2  = aThemeColors[4]
        acc1 = aThemeColors[5]
        acc2 = aThemeColors[6]
        acc3 = aThemeColors[7]
        acc4 = aThemeColors[8]
        acc5 = aThemeColors[9]
        acc6 = aThemeColors[10]

        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="' + cThemeName + '">'
        c += '<a:themeElements>'
        c += '<a:clrScheme name="' + cThemeName + '">'
        c += '<a:dk1><a:srgbClr val="' + dk1 + '"/></a:dk1>'
        c += '<a:lt1><a:srgbClr val="' + lt1 + '"/></a:lt1>'
        c += '<a:dk2><a:srgbClr val="' + dk2 + '"/></a:dk2>'
        c += '<a:lt2><a:srgbClr val="' + lt2 + '"/></a:lt2>'
        c += '<a:accent1><a:srgbClr val="' + acc1 + '"/></a:accent1>'
        c += '<a:accent2><a:srgbClr val="' + acc2 + '"/></a:accent2>'
        c += '<a:accent3><a:srgbClr val="' + acc3 + '"/></a:accent3>'
        c += '<a:accent4><a:srgbClr val="' + acc4 + '"/></a:accent4>'
        c += '<a:accent5><a:srgbClr val="' + acc5 + '"/></a:accent5>'
        c += '<a:accent6><a:srgbClr val="' + acc6 + '"/></a:accent6>'
        c += '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        c += '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        c += '</a:clrScheme>'
        c += '<a:fontScheme name="' + cThemeName + '">'
        c += '<a:majorFont><a:latin typeface="' + cDefaultFont + '"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        c += '<a:minorFont><a:latin typeface="' + cDefaultFont + '"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        c += '</a:fontScheme>'
        c += '<a:fmtScheme name="' + cThemeName + '">'
        c += '<a:fillStyleLst>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '</a:fillStyleLst>'
        c += '<a:lnStyleLst>'
        c += '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        c += '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        c += '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        c += '</a:lnStyleLst>'
        c += '<a:effectStyleLst>'
        c += '<a:effectStyle><a:effectLst/></a:effectStyle>'
        c += '<a:effectStyle><a:effectLst/></a:effectStyle>'
        c += '<a:effectStyle><a:effectLst/></a:effectStyle>'
        c += '</a:effectStyleLst>'
        c += '<a:bgFillStyleLst>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        c += '</a:bgFillStyleLst>'
        c += '</a:fmtScheme>'
        c += '</a:themeElements>'
        c += '</a:theme>'

        write(tempDir + "word" + sep + "theme" + sep + "theme1.xml", c)

    # Table of Figures / Tables generator

    func generateTableOfFigures item
        /*
            Emits a TOC \c field pre-populated with the captions already
            registered in aSeqCaptions. Entries are visible immediately on
            open — no F9 needed. Word re-computes page numbers on F9.

            Field structure (multi-paragraph):
              Para 1 : fldChar-begin + instrText + fldChar-separate + entry-1-text
              Para 2…N : entry text
              Para N : fldChar-end appended to last entry paragraph
        */
        seqType = item[:seqType]
        title   = item[:title]
        xml_ = ""

        # Optional heading
        if title != NULL and len(title) > 0
            xml_ += '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>'
            xml_ += '<w:r><w:t>' + wordXmlEsc(title) + '</w:t></w:r></w:p>'
        ok

        # Collect matching captions from registry
        aEntries = []
        nAll = len(aSeqCaptions)
        for ci = 1 to nAll
            cap = aSeqCaptions[ci]
            if cap[:label] = seqType
                aEntries + cap
            ok
        next

        pStyle = '<w:pPr><w:pStyle w:val="TableOfFigures"/></w:pPr>'

        if len(aEntries) = 0
            # No captions — single paragraph with placeholder
            xml_ += '<w:p>' + pStyle
            xml_ += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            xml_ += '<w:r><w:instrText xml:space="preserve"> TOC \c "' + wordXmlEsc(seqType) + '" </w:instrText></w:r>'
            xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            xml_ += '<w:r><w:rPr><w:color w:val="808080"/></w:rPr>'
            xml_ += '<w:t>Press F9 to update</w:t></w:r>'
            xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            xml_ += '</w:p>'
        else
            nEnt = len(aEntries)
            for ei = 1 to nEnt
                ent     = aEntries[ei]
                entText = wordXmlEsc(seqType) + " " + ent[:num] + " &#x2014; " + wordXmlEsc(ent[:text])
                xml_ += '<w:p>' + pStyle
                if ei = 1
                    # First paragraph carries the field begin + instrText + separate
                    xml_ += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    xml_ += '<w:r><w:instrText xml:space="preserve"> TOC \c "' + wordXmlEsc(seqType) + '" </w:instrText></w:r>'
                    xml_ += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                ok
                xml_ += '<w:r><w:t>' + entText + '</w:t></w:r>'
                xml_ += '<w:r><w:tab/></w:r>'
                xml_ += '<w:r><w:t>1</w:t></w:r>'
                if ei = nEnt
                    # Last paragraph carries the field end
                    xml_ += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                ok
                xml_ += '</w:p>'
            next
        ok
        return xml_

    func generateTextBox item
        /*
            Generates a VML floating text box paragraph.
            Uses v:rect + v:textbox (same engine as watermarks) for universal
            Word compatibility — no mc:AlternateContent wrapper needed.
            Position/size in cm converted to pt (1 cm = 28.35 pt).
        */
        # Defaults
        tbX = 2.0
        if item[:x] != NULL tbX = item[:x] ok
        tbY = 5.0
        if item[:y] != NULL tbY = item[:y] ok
        tbW = 6.0
        if item[:width] != NULL tbW = item[:width] ok
        tbH = 3.0
        if item[:height] != NULL tbH = item[:height] ok
        tbAlign = "center"
        if item[:align] != NULL tbAlign = item[:align] ok
        tbBold = false
        if item[:bold] = true tbBold = true ok
        tbItalic = false
        if item[:italic] = true tbItalic = true ok
        tbFont = ""
        if item[:font] != NULL tbFont = item[:font] ok
        tbSize = 0
        if item[:size] != NULL tbSize = item[:size] ok
        tbColor = "000000"
        if item[:color] != NULL tbColor = wordColorToHex(item[:color]) ok
        tbBgColor = "FFFFFF"
        if item[:bgColor] != NULL tbBgColor = wordColorToHex(item[:bgColor]) ok
        tbBorderColor = "4472C4"
        if item[:borderColor] != NULL tbBorderColor = wordColorToHex(item[:borderColor]) ok
        tbBorderSize = 1
        if item[:borderSize] != NULL
            # convert eighths-of-pt to pt
            tbBorderSize = item[:borderSize] / 8
        ok
        tbNoFill = false
        if item[:noFill] = true tbNoFill = true ok
        tbNoBorder = false
        if item[:noBorder] = true tbNoBorder = true ok
        tbId = item[:tbId]
        tbText = item[:text]
        if tbText = NULL tbText = "" ok

        # Convert cm to pt (1 cm = 28.35 pt)
        xPt = floor(tbX * 2835 / 100)
        yPt = floor(tbY * 2835 / 100)
        wPt = floor(tbW * 2835 / 100)
        hPt = floor(tbH * 2835 / 100)

        # VML shape style
        shapeStyle  = "position:absolute;"
        shapeStyle += "left:" + xPt + "pt;"
        shapeStyle += "top:" + yPt + "pt;"
        shapeStyle += "width:" + wPt + "pt;"
        shapeStyle += "height:" + hPt + "pt;"
        shapeStyle += "z-index:251659264;"
        shapeStyle += "mso-wrap-style:none;"
        shapeStyle += "mso-position-horizontal:absolute;"
        shapeStyle += "mso-position-vertical:absolute"

        # Fill color
        if tbNoFill
            fillAttr = 'filled="f"'
        else
            fillAttr = 'filled="t" fillcolor="#' + tbBgColor + '"'
        ok

        # Border/stroke
        if tbNoBorder
            strokeAttr = 'stroked="f"'
        else
            strokeAttr = 'stroked="t" strokecolor="#' + tbBorderColor + '" strokeweight="' + tbBorderSize + 'pt"'
        ok

        # Run properties for text inside box
        wrPrXml = ""
        if tbBold    wrPrXml += '<w:b/>' ok
        if tbItalic  wrPrXml += '<w:i/>' ok
        if len(tbFont) > 0
            wrPrXml += '<w:rFonts w:ascii="' + tbFont + '" w:hAnsi="' + tbFont + '"/>'
        ok
        if tbSize > 0
            wrPrXml += '<w:sz w:val="' + (tbSize * 2) + '"/>'
            wrPrXml += '<w:szCs w:val="' + (tbSize * 2) + '"/>'
        ok
        wrPrXml += '<w:color w:val="' + tbColor + '"/>'

        xml_ = '<w:p>'
        xml_ += '<w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr>'
        xml_ += '<w:r><w:rPr><w:noProof/></w:rPr>'
        xml_ += '<w:pict>'
        xml_ += '<v:rect id="_TextBox_' + tbId + '" type="#_x0000_t202"'
        xml_ += ' style="' + shapeStyle + '"'
        xml_ += ' ' + fillAttr
        xml_ += ' ' + strokeAttr + '>'
        xml_ += '<v:textbox>'
        xml_ += '<w:txbxContent>'
        xml_ += '<w:p>'
        xml_ += '<w:pPr><w:jc w:val="' + tbAlign + '"/>'
        xml_ += '<w:spacing w:before="0" w:after="0"/>'
        xml_ += '</w:pPr>'
        xml_ += '<w:r>'
        if len(wrPrXml) > 0
            xml_ += '<w:rPr>' + wrPrXml + '</w:rPr>'
        ok
        xml_ += '<w:t>' + wordXmlEsc(tbText) + '</w:t>'
        xml_ += '</w:r>'
        xml_ += '</w:p>'
        xml_ += '</w:txbxContent>'
        xml_ += '</v:textbox>'
        xml_ += '</v:rect>'
        xml_ += '</w:pict>'
        xml_ += '</w:r>'
        xml_ += '</w:p>'

        return xml_

    # =========================================================================
    # Private: Chart XML Generators
    # =========================================================================

    func generateChartDrawing item
        /*
            Produce the <w:drawing> … <c:chart r:id="…"/> inline XML that
            anchors the chart in the document body.
        */
        relId    = item[:relId]
        chartId  = item[:chartId]
        docPrId  = item[:docPrId]
        widthEmu = item[:width]
        heightEmu= item[:height]
        centered = item[:centered]

        xml_ = '<w:p>'
        if centered
            xml_ += '<w:pPr><w:jc w:val="center"/></w:pPr>'
        ok
        xml_ += '<w:r><w:drawing>'
        xml_ += '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        xml_ += '<wp:extent cx="' + widthEmu + '" cy="' + heightEmu + '"/>'
        xml_ += '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        xml_ += '<wp:docPr id="' + docPrId + '" name="Chart ' + chartId + '"/>'
        xml_ += '<wp:cNvGraphicFramePr>'
        xml_ += '<a:graphicFrameLocks'
        xml_ += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        xml_ += ' noChangeAspect="1"/>'
        xml_ += '</wp:cNvGraphicFramePr>'
        xml_ += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        xml_ += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        xml_ += '<c:chart'
        xml_ += ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        xml_ += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        xml_ += ' r:id="' + relId + '"/>'
        xml_ += '</a:graphicData>'
        xml_ += '</a:graphic>'
        xml_ += '</wp:inline>'
        xml_ += '</w:drawing></w:r></w:p>'
        return xml_

    func writeChart tempDir, chart
        /*  Write word/charts/chartN.xml for the given chart definition.  */
        sep = wordGetSep()
        xml = buildChartXml(chart)
        write(tempDir + "word" + sep + "charts" + sep + "chart" + chart[:chartId] + ".xml", xml)

    func writeChartRels tempDir, chart
        /*  Write word/charts/_rels/chartN.xml.rels  (minimal — no external deps).  */
        sep = wordGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '</Relationships>'
        write(tempDir + "word" + sep + "charts" + sep + "_rels" + sep +
              "chart" + chart[:chartId] + ".xml.rels", c)

    func buildChartXml chart
        /*
            Build the complete chartN.xml content for a chart definition.
            Dispatches to the appropriate series-builder based on chart type.
        */
        chartType   = chart[:type]
        chartTitle  = chart[:title]
        categories  = chart[:categories]
        seriesList  = chart[:series]
        options     = chart[:options]

        # Default 8-colour palette (matches the Office "Colorful" theme)
        aPalette = ["4472C4","ED7D31","A9D18E","FFC000","5A9BD5","70AD47","255E91","C00000"]

        # Options defaults
        showLegend     = true
        legendPos      = "r"
        showDataLabels = false
        grouping       = "clustered"
        bSmooth        = false
        yAxisMin       = NULL   # number — value axis lower bound (NULL = auto)
        yAxisMax       = NULL   # number — value axis upper bound (NULL = auto)
        xAxisMin       = NULL   # number — X axis lower bound (scatter/bubble only)
        xAxisMax       = NULL   # number — X axis upper bound (scatter/bubble only)
        yAxisFormat    = ""     # number format string, e.g. "0.0%", "$#,##0", "0.00"
        xAxisFormat    = ""     # X axis number format (scatter/bubble)
        yAxisTitle     = ""     # value axis label (all chart types)
        xAxisTitle     = ""     # category/X axis label
        showGridlines  = true   # major gridlines on value axis

        showDataTable     = false   # show data table below chart
        dataTableShowKeys = true    # show legend colour keys in data table

        if isList(options)
            if options[:showLegend]        != NULL  showLegend        = options[:showLegend]        ok
            if options[:legendPos]         != NULL  legendPos         = options[:legendPos]         ok
            if options[:showDataLabels]    != NULL  showDataLabels    = options[:showDataLabels]    ok
            if options[:grouping]          != NULL  grouping          = options[:grouping]          ok
            if options[:smooth]            != NULL  bSmooth           = options[:smooth]            ok
            if isList(options[:colors])             aPalette          = options[:colors]            ok
            if options[:yAxisMin]          != NULL  yAxisMin          = options[:yAxisMin]          ok
            if options[:yAxisMax]          != NULL  yAxisMax          = options[:yAxisMax]          ok
            if options[:xAxisMin]          != NULL  xAxisMin          = options[:xAxisMin]          ok
            if options[:xAxisMax]          != NULL  xAxisMax          = options[:xAxisMax]          ok
            if options[:yAxisFormat]       != NULL  yAxisFormat       = options[:yAxisFormat]       ok
            if options[:xAxisFormat]       != NULL  xAxisFormat       = options[:xAxisFormat]       ok
            if options[:yAxisTitle]        != NULL  yAxisTitle        = options[:yAxisTitle]        ok
            if options[:xAxisTitle]        != NULL  xAxisTitle        = options[:xAxisTitle]        ok
            if options[:showGridlines]     != NULL  showGridlines     = options[:showGridlines]     ok
            if options[:showDataTable]     != NULL  showDataTable     = options[:showDataTable]     ok
            if options[:dataTableShowKeys] != NULL  dataTableShowKeys = options[:dataTableShowKeys] ok
        ok

        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<c:chartSpace'
        c += ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        c += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        c += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        c += '<c:roundedCorners val="0"/>'
        c += '<c:chart>'

        # Title
        if len(chartTitle) > 0
            c += '<c:title>'
            c += '<c:tx><c:rich>'
            c += '<a:bodyPr/><a:lstStyle/>'
            c += '<a:p><a:r>'
            c += '<a:rPr lang="en-US" b="1" dirty="0"/>'
            c += '<a:t>' + wordXmlEsc(chartTitle) + '</a:t>'
            c += '</a:r></a:p>'
            c += '</c:rich></c:tx>'
            c += '<c:overlay val="0"/>'
            c += '</c:title>'
            c += '<c:autoTitleDeleted val="0"/>'
        else
            c += '<c:autoTitleDeleted val="1"/>'
        ok

        c += '<c:plotArea><c:layout/>'

        # Chart-type body
        switch chartType
            on "column"
                c += buildBarChartXml(seriesList, categories, "col",  grouping, showDataLabels, aPalette)
            on "bar"
                c += buildBarChartXml(seriesList, categories, "bar",  grouping, showDataLabels, aPalette)
            on "line"
                c += buildLineChartXml(seriesList, categories, bSmooth, showDataLabels, aPalette)
            on "area"
                c += buildAreaChartXml(seriesList, categories, grouping, showDataLabels, aPalette)
            on "pie"
                c += buildPieChartXml(seriesList, categories, showDataLabels, aPalette, false)
            on "doughnut"
                c += buildPieChartXml(seriesList, categories, showDataLabels, aPalette, true)
            on "scatter"
                c += buildScatterChartXml(seriesList, options, showDataLabels, aPalette)
            on "bubble"
                c += buildBubbleChartXml(seriesList, options, showDataLabels, aPalette)
            other
                c += buildBarChartXml(seriesList, categories, "col", grouping, showDataLabels, aPalette)
        off

        # Axes — pie/doughnut need none; scatter/bubble need two valAx; others need catAx+valAx
        if chartType != "pie" and chartType != "doughnut"

            if chartType = "scatter" or chartType = "bubble"
                # Both axes are numeric — two <c:valAx>.
                # X axis number format: use xAxisFormat if provided, else General
                xFmt = "General"
                if len(xAxisFormat) > 0  xFmt = xAxisFormat  ok
                yFmt = "General"
                if len(yAxisFormat) > 0  yFmt = yAxisFormat  ok

                # X axis (bottom, valAx id=1)
                c += '<c:valAx>'
                c += '<c:axId val="1"/>'
                c += '<c:scaling><c:orientation val="minMax"/>'
                if xAxisMax != NULL  c += '<c:max val="' + xAxisMax + '"/>'  ok
                if xAxisMin != NULL  c += '<c:min val="' + xAxisMin + '"/>'  ok
                c += '</c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="b"/>'
                c += '<c:numFmt formatCode="' + wordXmlEsc(xFmt) + '" sourceLinked="0"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(xAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(xAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="2"/>'
                c += '<c:crossBetween val="midCat"/>'
                c += '</c:valAx>'

                # Y axis (left, valAx id=2)
                c += '<c:valAx>'
                c += '<c:axId val="2"/>'
                c += '<c:scaling><c:orientation val="minMax"/>'
                if yAxisMax != NULL  c += '<c:max val="' + yAxisMax + '"/>'  ok
                if yAxisMin != NULL  c += '<c:min val="' + yAxisMin + '"/>'  ok
                c += '</c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="l"/>'
                if showGridlines  c += '<c:majorGridlines/>'  ok
                c += '<c:numFmt formatCode="' + wordXmlEsc(yFmt) + '" sourceLinked="0"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(yAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(yAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="1"/>'
                c += '</c:valAx>'

            elseif chartType = "bar"
                # Horizontal bar: categories on left (l), values on bottom (b)
                yFmt = "General"
                if len(yAxisFormat) > 0  yFmt = yAxisFormat  ok

                c += '<c:catAx>'
                c += '<c:axId val="1"/>'
                c += '<c:scaling><c:orientation val="minMax"/></c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="l"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(xAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(xAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="2"/>'
                c += '</c:catAx>'
                c += '<c:valAx>'
                c += '<c:axId val="2"/>'
                c += '<c:scaling><c:orientation val="minMax"/>'
                if yAxisMax != NULL  c += '<c:max val="' + yAxisMax + '"/>'  ok
                if yAxisMin != NULL  c += '<c:min val="' + yAxisMin + '"/>'  ok
                c += '</c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="b"/>'
                if showGridlines  c += '<c:majorGridlines/>'  ok
                c += '<c:numFmt formatCode="' + wordXmlEsc(yFmt) + '" sourceLinked="0"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(yAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(yAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="1"/>'
                c += '<c:crossBetween val="between"/>'
                c += '</c:valAx>'

            else
                # Column / line / area: categories on bottom (b), values on left (l)
                yFmt = "General"
                if len(yAxisFormat) > 0  yFmt = yAxisFormat  ok

                c += '<c:catAx>'
                c += '<c:axId val="1"/>'
                c += '<c:scaling><c:orientation val="minMax"/></c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="b"/>'
                c += '<c:numFmt formatCode="General" sourceLinked="0"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(xAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(xAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="2"/>'
                c += '<c:auto val="1"/>'
                c += '<c:lblAlign val="ctr"/>'
                c += '<c:lblOffset val="100"/>'
                c += '<c:noMultiLvlLbl val="0"/>'
                c += '</c:catAx>'
                c += '<c:valAx>'
                c += '<c:axId val="2"/>'
                c += '<c:scaling><c:orientation val="minMax"/>'
                if yAxisMax != NULL  c += '<c:max val="' + yAxisMax + '"/>'  ok
                if yAxisMin != NULL  c += '<c:min val="' + yAxisMin + '"/>'  ok
                c += '</c:scaling>'
                c += '<c:delete val="0"/><c:axPos val="l"/>'
                if showGridlines  c += '<c:majorGridlines/>'  ok
                c += '<c:numFmt formatCode="' + wordXmlEsc(yFmt) + '" sourceLinked="0"/>'
                c += '<c:tickLblPos val="nextTo"/>'
                if len(yAxisTitle) > 0
                    c += '<c:title><c:tx><c:rich>'
                    c += '<a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/>'
                    c += '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>'
                    c += '<a:t>' + wordXmlEsc(yAxisTitle) + '</a:t>'
                    c += '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
                ok
                c += '<c:crossAx val="1"/>'
                c += '<c:crossBetween val="between"/>'
                c += '</c:valAx>'
            ok
        ok

        c += '</c:plotArea>'

        # Data table (only for category-based charts; not scatter/bubble/pie/doughnut)
        if showDataTable
            if chartType != "pie" and chartType != "doughnut" and
               chartType != "scatter" and chartType != "bubble"
                showKeysVal = "0"
                if dataTableShowKeys  showKeysVal = "1"  ok
                c += '<c:dTable>'
                c += '<c:showHorzBorder val="1"/>'
                c += '<c:showVertBorder val="1"/>'
                c += '<c:showOutline val="1"/>'
                c += '<c:showKeys val="' + showKeysVal + '"/>'
                c += '</c:dTable>'
            ok
        ok

        # Legend
        if showLegend
            c += '<c:legend>'
            c += '<c:legendPos val="' + legendPos + '"/>'
            c += '<c:overlay val="0"/>'
            c += '</c:legend>'
        ok

        c += '<c:plotVisOnly val="1"/>'
        c += '</c:chart>'

        # Chart space background (white fill, subtle border)
        c += '<c:spPr>'
        c += '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
        c += '<a:ln>'
        c += '<a:solidFill><a:schemeClr val="tx1">'
        c += '<a:lumMod val="15000"/><a:lumOff val="85000"/>'
        c += '</a:schemeClr></a:solidFill>'
        c += '</a:ln>'
        c += '</c:spPr>'

        c += '</c:chartSpace>'
        return c

    func buildSeriesCore serIdx, serName, categories, values, colorHex, showDataLabels
        /*
            Build the shared inner XML common to every series type:
            <c:idx>, <c:order>, series title, optional colour, optional data
            labels, category cache, and value cache.
        */
        c = '<c:idx val="' + serIdx + '"/>'
        c += '<c:order val="' + serIdx + '"/>'

        # Series name
        c += '<c:tx><c:strRef><c:f/><c:strCache>'
        c += '<c:ptCount val="1"/>'
        c += '<c:pt idx="0"><c:v>' + wordXmlEsc(serName) + '</c:v></c:pt>'
        c += '</c:strCache></c:strRef></c:tx>'

        # Series colour
        if len(colorHex) > 0
            c += '<c:spPr>'
            c += '<a:solidFill><a:srgbClr val="' + colorHex + '"/></a:solidFill>'
            c += '<a:ln><a:solidFill>'
            c += '<a:srgbClr val="' + colorHex + '"/>'
            c += '</a:solidFill></a:ln>'
            c += '</c:spPr>'
        ok

        # Data labels
        if showDataLabels
            c += '<c:dLbls>'
            c += '<c:numFmt formatCode="General" sourceLinked="1"/>'
            c += '<c:spPr/>'
            c += '<c:showLegendKey val="0"/><c:showVal val="1"/>'
            c += '<c:showCatName val="0"/><c:showSerName val="0"/>'
            c += '<c:showPercent val="0"/><c:showBubbleSize val="0"/>'
            c += '</c:dLbls>'
        ok

        # Category labels (shared string cache — no spreadsheet needed)
        c += '<c:cat><c:strRef><c:f/><c:strCache>'
        categoriesLen = len(categories)
        c += '<c:ptCount val="' + categoriesLen + '"/>'
        for i = 1 to categoriesLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>'
            c += wordXmlEsc("" + categories[i])
            c += '</c:v></c:pt>'
        next
        c += '</c:strCache></c:strRef></c:cat>'

        # Numeric values (inline numCache — no spreadsheet needed)
        c += '<c:val><c:numRef><c:f/><c:numCache>'
        c += '<c:formatCode>General</c:formatCode>'
        valuesLen = len(values)
        c += '<c:ptCount val="' + valuesLen + '"/>'
        for i = 1 to valuesLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>' + values[i] + '</c:v></c:pt>'
        next
        c += '</c:numCache></c:numRef></c:val>'
        return c

    func buildBarChartXml seriesList, categories, barDir, grouping, showDataLabels, aPalette
        /*  bar/column chart XML.  barDir = "bar" (horiz) | "col" (vert).  */
        ooGrouping = "clustered"
        if lower(grouping) = "stacked"         ooGrouping = "stacked"         ok
        if lower(grouping) = "percentstacked"  ooGrouping = "percentStacked"  ok
        if lower(grouping) = "percent"         ooGrouping = "percentStacked"  ok

        c = '<c:barChart>'
        c += '<c:barDir val="' + barDir + '"/>'
        c += '<c:grouping val="' + ooGrouping + '"/>'
        c += '<c:varyColors val="0"/>'

        seriesListLen = len(seriesList)
        for i = 1 to seriesListLen
            ser      = seriesList[i]
            serName  = ser[:name]
            if serName  = NULL  serName = "Series " + i  ok
            colorHex = ser[:color]
            if colorHex = NULL  colorHex = aPalette[((i-1) % len(aPalette)) + 1]  ok
            c += '<c:ser>'
            c += buildSeriesCore(i-1, serName, categories, ser[:values], colorHex, showDataLabels)
            c += '</c:ser>'
        next

        c += '<c:axId val="1"/><c:axId val="2"/>'
        c += '</c:barChart>'
        return c

    func buildLineChartXml seriesList, categories, bSmooth, showDataLabels, aPalette
        /*
            Line chart XML.
        */
        c = '<c:lineChart>'
        c += '<c:grouping val="standard"/>'
        c += '<c:varyColors val="0"/>'

        seriesListLen = len(seriesList)
        for i = 1 to seriesListLen
            ser      = seriesList[i]
            serName  = ser[:name]
            if serName  = NULL  serName = "Series " + i  ok
            colorHex = ser[:color]
            if colorHex = NULL  colorHex = aPalette[((i-1) % len(aPalette)) + 1]  ok
            c += '<c:ser>'
            c += buildSeriesCore(i-1, serName, categories, ser[:values], colorHex, showDataLabels)
            # Per-series marker (CT_Marker) — hide the dot marker on data points
            c += '<c:marker><c:symbol val="none"/></c:marker>'
            if bSmooth  c += '<c:smooth val="1"/>'  else  c += '<c:smooth val="0"/>'  ok
            c += '</c:ser>'
        next

        # <c:marker> at chart level is CT_Boolean (val="0" = no markers)
        c += '<c:marker val="0"/>'
        c += '<c:axId val="1"/><c:axId val="2"/>'
        c += '</c:lineChart>'
        return c

    func buildScatterSeriesCore serIdx, serName, xValues, yValues, colorHex, showDataLabels, markerStyle, markerSizeHpt
        /*
            Core XML for a single scatter or lines-only scatter series.
            xValues and yValues are parallel numeric lists.
            markerStyle : OOXML symbol string e.g. "circle", "square", "none"
            markerSizeHpt: marker size in half-points (e.g. 14 = 7pt)
        */
        c = '<c:idx val="' + serIdx + '"/>'
        c += '<c:order val="' + serIdx + '"/>'

        # Series name
        c += '<c:tx><c:strRef><c:f/><c:strCache>'
        c += '<c:ptCount val="1"/>'
        c += '<c:pt idx="0"><c:v>' + wordXmlEsc(serName) + '</c:v></c:pt>'
        c += '</c:strCache></c:strRef></c:tx>'

        # Solid colour for the line/marker
        if len(colorHex) > 0
            c += '<c:spPr>'
            c += '<a:solidFill><a:srgbClr val="' + colorHex + '"/></a:solidFill>'
            c += '<a:ln><a:solidFill>'
            c += '<a:srgbClr val="' + colorHex + '"/>'
            c += '</a:solidFill></a:ln>'
            c += '</c:spPr>'
        ok

        # Data labels
        if showDataLabels
            c += '<c:dLbls>'
            c += '<c:numFmt formatCode="General" sourceLinked="1"/>'
            c += '<c:spPr/>'
            c += '<c:showLegendKey val="0"/><c:showVal val="1"/>'
            c += '<c:showCatName val="0"/><c:showSerName val="0"/>'
            c += '<c:showPercent val="0"/><c:showBubbleSize val="0"/>'
            c += '</c:dLbls>'
        ok

        # Per-series marker (CT_Marker — inside <c:ser>)
        c += '<c:marker>'
        c += '<c:symbol val="' + markerStyle + '"/>'
        if markerStyle != "none"
            c += '<c:size val="' + markerSizeHpt + '"/>'
            if len(colorHex) > 0
                c += '<c:spPr>'
                c += '<a:solidFill><a:srgbClr val="' + colorHex + '"/></a:solidFill>'
                c += '<a:ln><a:solidFill>'
                c += '<a:srgbClr val="' + colorHex + '"/>'
                c += '</a:solidFill></a:ln>'
                c += '</c:spPr>'
            ok
        ok
        c += '</c:marker>'

        # X values
        c += '<c:xVal><c:numRef><c:f/><c:numCache>'
        c += '<c:formatCode>General</c:formatCode>'
        xValuesLen = len(xValues)
        c += '<c:ptCount val="' + xValuesLen + '"/>'
        for i = 1 to xValuesLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>' + xValues[i] + '</c:v></c:pt>'
        next
        c += '</c:numCache></c:numRef></c:xVal>'

        # Y values
        c += '<c:yVal><c:numRef><c:f/><c:numCache>'
        c += '<c:formatCode>General</c:formatCode>'
        yValuesLen = len(yValues)
        c += '<c:ptCount val="' + yValuesLen + '"/>'
        for i = 1 to yValuesLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>' + yValues[i] + '</c:v></c:pt>'
        next
        c += '</c:numCache></c:numRef></c:yVal>'

        return c

    func buildScatterChartXml seriesList, options, showDataLabels, aPalette
        /*
            Scatter chart XML.
            Supports markers-only (default), lines-only, and lines+markers.
            options:
              :markerStyle  "circle"|"square"|"diamond"|"triangle"|"x"|"star"|"dot"|"dash"|"none"
              :markerSize   integer (half-points, default 7 → val="7")
              :lines        true/false — connect points with lines (default false)
              :smooth       true/false — smooth curves (default false)
        */
        markerStyle = "circle"
        markerSize  = 7
        bLines      = false
        bSmooth     = false

        if isList(options)
            if options[:markerStyle] != NULL  markerStyle = options[:markerStyle]  ok
            if options[:markerSize]  != NULL  markerSize  = options[:markerSize]   ok
            if options[:lines]       != NULL  bLines      = options[:lines]        ok
            if options[:smooth]      != NULL  bSmooth     = options[:smooth]       ok
        ok

        # Pick OOXML scatter style
        if bLines and markerStyle != "none"
            scatterStyle = "lineMarker"
        elseif bLines
            scatterStyle = "line"
        else
            scatterStyle = "marker"
        ok

        c = '<c:scatterChart>'
        c += '<c:scatterStyle val="' + scatterStyle + '"/>'
        c += '<c:varyColors val="0"/>'

        seriesListLen = len(seriesList)
        for i = 1 to seriesListLen
            ser      = seriesList[i]
            serName  = ser[:name]
            if serName  = NULL  serName = "Series " + i  ok
            colorHex = ser[:color]
            if colorHex = NULL  colorHex = aPalette[((i-1) % len(aPalette)) + 1]  ok
            xVals    = ser[:xValues]
            yVals    = ser[:yValues]

            c += '<c:ser>'
            c += buildScatterSeriesCore(i-1, serName, xVals, yVals,
                                        colorHex, showDataLabels,
                                        markerStyle, markerSize)
            smoothVal = "0"
            if bSmooth  smoothVal = "1"  ok
            if bLines  c += '<c:smooth val="' + smoothVal + '"/>'  ok
            c += '</c:ser>'
        next

        c += '<c:axId val="1"/><c:axId val="2"/>'
        c += '</c:scatterChart>'
        return c

    func buildBubbleChartXml seriesList, options, showDataLabels, aPalette
        /*
            Bubble chart XML.
            Series items must supply :xValues, :yValues, and :sizes.
            options:
              :bubble3D   true/false — 3-D shading on bubbles (default false)
              :markerStyle / :markerSize / :lines / :smooth — ignored (bubbles have
              their own shape, these options only apply to scatter)
        */
        bBubble3D = false
        if isList(options)
            if options[:bubble3D] != NULL  bBubble3D = options[:bubble3D]  ok
        ok

        c = '<c:bubbleChart>'
        c += '<c:varyColors val="0"/>'

        seriesListLen = len(seriesList)
        for i = 1 to seriesListLen
            ser      = seriesList[i]
            serName  = ser[:name]
            if serName  = NULL  serName = "Series " + i  ok
            colorHex = ser[:color]
            if colorHex = NULL  colorHex = aPalette[((i-1) % len(aPalette)) + 1]  ok
            xVals    = ser[:xValues]
            yVals    = ser[:yValues]
            sizeVals = ser[:sizes]

            c += '<c:ser>'
            c += '<c:idx val="' + (i-1) + '"/>'
            c += '<c:order val="' + (i-1) + '"/>'

            # Series name
            c += '<c:tx><c:strRef><c:f/><c:strCache>'
            c += '<c:ptCount val="1"/>'
            c += '<c:pt idx="0"><c:v>' + wordXmlEsc(serName) + '</c:v></c:pt>'
            c += '</c:strCache></c:strRef></c:tx>'

            # Series colour (fill + outline)
            if len(colorHex) > 0
                c += '<c:spPr>'
                c += '<a:solidFill><a:srgbClr val="' + colorHex + '">'
                c += '<a:alpha val="70000"/>'   # 70% opacity so overlapping bubbles show through
                c += '</a:srgbClr></a:solidFill>'
                c += '<a:ln w="12700">'
                c += '<a:solidFill><a:srgbClr val="' + colorHex + '"/></a:solidFill>'
                c += '</a:ln>'
                c += '</c:spPr>'
            ok

            # Data labels (show bubble size)
            if showDataLabels
                c += '<c:dLbls>'
                c += '<c:numFmt formatCode="General" sourceLinked="1"/>'
                c += '<c:spPr/>'
                c += '<c:showLegendKey val="0"/><c:showVal val="0"/>'
                c += '<c:showCatName val="0"/><c:showSerName val="0"/>'
                c += '<c:showPercent val="0"/><c:showBubbleSize val="1"/>'
                c += '</c:dLbls>'
            ok

            # X values
            c += '<c:xVal><c:numRef><c:f/><c:numCache>'
            c += '<c:formatCode>General</c:formatCode>'
            xValsLen = len(xVals)
            c += '<c:ptCount val="' + xValsLen + '"/>'
            for j = 1 to xValsLen
                c += '<c:pt idx="' + (j-1) + '"><c:v>' + xVals[j] + '</c:v></c:pt>'
            next
            c += '</c:numCache></c:numRef></c:xVal>'

            # Y values
            c += '<c:yVal><c:numRef><c:f/><c:numCache>'
            c += '<c:formatCode>General</c:formatCode>'
            yValsLen = len(yVals)
            c += '<c:ptCount val="' + yValsLen + '"/>'
            for j = 1 to yValsLen
                c += '<c:pt idx="' + (j-1) + '"><c:v>' + yVals[j] + '</c:v></c:pt>'
            next
            c += '</c:numCache></c:numRef></c:yVal>'

            # Bubble sizes
            c += '<c:bubbleSize><c:numRef><c:f/><c:numCache>'
            c += '<c:formatCode>General</c:formatCode>'
            sizeValsLen = len(sizeVals)
            c += '<c:ptCount val="' + sizeValsLen + '"/>'
            for j = 1 to sizeValsLen
                c += '<c:pt idx="' + (j-1) + '"><c:v>' + sizeVals[j] + '</c:v></c:pt>'
            next
            c += '</c:numCache></c:numRef></c:bubbleSize>'

            bubble3DVal = "0"
            if bBubble3D  bubble3DVal = "1"  ok
            c += '<c:bubble3D val="' + bubble3DVal + '"/>'
            c += '</c:ser>'
        next

        c += '<c:axId val="1"/><c:axId val="2"/>'
        c += '</c:bubbleChart>'
        return c

    func buildAreaChartXml seriesList, categories, grouping, showDataLabels, aPalette
        /*  Area chart XML.  */
        ooGrouping = "standard"
        if lower(grouping) = "stacked"         ooGrouping = "stacked"         ok
        if lower(grouping) = "percentstacked"  ooGrouping = "percentStacked"  ok
        if lower(grouping) = "percent"         ooGrouping = "percentStacked"  ok

        c = '<c:areaChart>'
        c += '<c:grouping val="' + ooGrouping + '"/>'
        c += '<c:varyColors val="0"/>'

        seriesListLen = len(seriesList)
        for i = 1 to seriesListLen
            ser      = seriesList[i]
            serName  = ser[:name]
            if serName  = NULL  serName = "Series " + i  ok
            colorHex = ser[:color]
            if colorHex = NULL  colorHex = aPalette[((i-1) % len(aPalette)) + 1]  ok
            c += '<c:ser>'
            c += buildSeriesCore(i-1, serName, categories, ser[:values], colorHex, showDataLabels)
            c += '</c:ser>'
        next

        c += '<c:axId val="1"/><c:axId val="2"/>'
        c += '</c:areaChart>'
        return c

    func buildPieChartXml seriesList, categories, showDataLabels, aPalette, bDoughnut
        /*
            Pie / Doughnut chart XML.
            Uses only the first series; each category becomes a slice.
            bDoughnut = true emits <c:doughnutChart> with holeSize.

            NOTE: OOXML pie/doughnut charts only render the first series.
            Extra series are ignored with a printed warning.
        */
        chartKind = "Pie"
        if bDoughnut  chartKind = "Doughnut"  ok

        if len(seriesList) > 1
            ? "RingWordLib WARNING: " + chartKind + " charts support only one series. " +
              string(len(seriesList)) + " series supplied; only the first will be used. " +
              "Use a column or bar chart if you need multiple series."
        ok

        tagOpen  = "<c:pieChart>"
        tagClose = "</c:pieChart>"
        if bDoughnut
            tagOpen  = "<c:doughnutChart>"
            tagClose = "</c:doughnutChart>"
        ok

        c = tagOpen
        c += '<c:varyColors val="1"/>'

        ser      = seriesList[1]
        serName  = ser[:name]
        if serName = NULL  serName = "Series 1"  ok
        serVals  = ser[:values]

        c += '<c:ser>'
        c += '<c:idx val="0"/><c:order val="0"/>'
        c += '<c:tx><c:strRef><c:f/><c:strCache>'
        c += '<c:ptCount val="1"/>'
        c += '<c:pt idx="0"><c:v>' + wordXmlEsc(serName) + '</c:v></c:pt>'
        c += '</c:strCache></c:strRef></c:tx>'

        # Per-slice colour overrides
        categoriesLen = len(categories)
        for i = 1 to categoriesLen
            colorHex = aPalette[((i-1) % len(aPalette)) + 1]
            c += '<c:dPt><c:idx val="' + (i-1) + '"/>'
            c += '<c:spPr><a:solidFill>'
            c += '<a:srgbClr val="' + colorHex + '"/>'
            c += '</a:solidFill></c:spPr></c:dPt>'
        next

        # Data labels: show category name + percentage for pie/doughnut
        if showDataLabels
            c += '<c:dLbls>'
            c += '<c:numFmt formatCode="0%" sourceLinked="0"/>'
            c += '<c:spPr/>'
            c += '<c:showLegendKey val="0"/><c:showVal val="0"/>'
            c += '<c:showCatName val="1"/><c:showSerName val="0"/>'
            c += '<c:showPercent val="1"/><c:showBubbleSize val="0"/>'
            c += '<c:separator> — </c:separator>'
            c += '</c:dLbls>'
        ok

        # Category strings
        c += '<c:cat><c:strRef><c:f/><c:strCache>'
        c += '<c:ptCount val="' + categoriesLen + '"/>'
        for i = 1 to categoriesLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>'
            c += wordXmlEsc("" + categories[i])
            c += '</c:v></c:pt>'
        next
        c += '</c:strCache></c:strRef></c:cat>'

        # Values
        c += '<c:val><c:numRef><c:f/><c:numCache>'
        c += '<c:formatCode>General</c:formatCode>'
        serValsLen = len(serVals)
        c += '<c:ptCount val="' + serValsLen + '"/>'
        for i = 1 to serValsLen
            c += '<c:pt idx="' + (i-1) + '"><c:v>' + serVals[i] + '</c:v></c:pt>'
        next
        c += '</c:numCache></c:numRef></c:val>'
        c += '</c:ser>'

        c += '<c:firstSliceAng val="0"/>'
        if bDoughnut  c += '<c:holeSize val="50"/>'  ok
        c += tagClose
        return c

    # =========================================================================
    # createZip — packs everything into the final DOCX
    # =========================================================================

    func createZip tempDir, filename
        sep = wordGetSep()
        
        filesList = []
        
        filesList + ["[Content_Types].xml", read(tempDir + "[Content_Types].xml")]
        filesList + ["_rels/.rels", read(tempDir + "_rels" + sep + ".rels")]
        filesList + ["docProps/core.xml", read(tempDir + "docProps" + sep + "core.xml")]
        filesList + ["docProps/app.xml", read(tempDir + "docProps" + sep + "app.xml")]
        if len(aCustomProps) > 0
            filesList + ["docProps/custom.xml", read(tempDir + "docProps" + sep + "custom.xml")]
        ok
        filesList + ["word/document.xml", read(tempDir + "word" + sep + "document.xml")]
        filesList + ["word/_rels/document.xml.rels", read(tempDir + "word" + sep + "_rels" + sep + "document.xml.rels")]
        filesList + ["word/styles.xml", read(tempDir + "word" + sep + "styles.xml")]
        filesList + ["word/settings.xml", read(tempDir + "word" + sep + "settings.xml")]
        filesList + ["word/fontTable.xml", read(tempDir + "word" + sep + "fontTable.xml")]
        
        if len(aNumbering) > 0
            filesList + ["word/numbering.xml", read(tempDir + "word" + sep + "numbering.xml")]
        ok
        
        # Add header if exists (header text, text watermark, or image watermark)
        if len(cHeaderText) > 0 or bWatermark or bImgWatermark
            filesList + ["word/header1.xml", read(tempDir + "word" + sep + "header1.xml")]
        ok
        
        # Add first-page header/footer 
        if bFirstPageDifferent
            filesList + ["word/header_fp.xml", read(tempDir + "word" + sep + "header_fp.xml")]
            filesList + ["word/footer_fp.xml", read(tempDir + "word" + sep + "footer_fp.xml")]
        ok
        
        # Add even-page header/footer
        if bEvenAndOddHeaders
            filesList + ["word/header_even.xml", read(tempDir + "word" + sep + "header_even.xml")]
            filesList + ["word/footer_even.xml", read(tempDir + "word" + sep + "footer_even.xml")]
        ok
        
        # Add header relationships file (needed when image watermark is active)
        if bImgWatermark
            filesList + ["word/_rels/header1.xml.rels", read(tempDir + "word" + sep + "_rels" + sep + "header1.xml.rels")]
            # Add the watermark image binary itself into the ZIP
            filesList + ["word/media/" + cImgWatermarkFile, read(tempDir + "word" + sep + "media" + sep + cImgWatermarkFile)]
        ok
        
        # Add footer if exists
        if len(cFooterText) > 0 or bShowPageNumbers
            filesList + ["word/footer1.xml", read(tempDir + "word" + sep + "footer1.xml")]
        ok
        
        # Add footnotes if any
        if len(aFootnotes) > 0
            filesList + ["word/footnotes.xml", read(tempDir + "word" + sep + "footnotes.xml")]
        ok
        
        # Add endnotes if any
        if len(aEndnotes) > 0
            filesList + ["word/endnotes.xml", read(tempDir + "word" + sep + "endnotes.xml")]
        ok

        # Add comments if any 
        if len(aComments) > 0
            filesList + ["word/comments.xml", read(tempDir + "word" + sep + "comments.xml")]
        ok

        # Add theme if active 
        if len(cThemeName) > 0
            filesList + ["word/theme/theme1.xml", read(tempDir + "word" + sep + "theme" + sep + "theme1.xml")]
        ok
        
        # Add images
        for img in aImages
            filesList + ["word/media/" + img[:filename], img[:data]]
        next

        # Add chart XML files
        for ch in aCharts
            cid = ch[:chartId]
            filesList + ["word/charts/chart" + cid + ".xml",
                         read(tempDir + "word" + sep + "charts" + sep + "chart" + cid + ".xml")]
            filesList + ["word/charts/_rels/chart" + cid + ".xml.rels",
                         read(tempDir + "word" + sep + "charts" + sep + "_rels" + sep + "chart" + cid + ".xml.rels")]
        next
        
        return wordZipCreateFile(filename, filesList)
