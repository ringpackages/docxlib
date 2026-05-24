/*
    DOCXLib  
*/

# ============================================================================
# WordReader class
# ============================================================================

class WordReader

    cFilePath           # Original .docx path
    cTempDir            # Extracted temp directory path
    aBlocks             # Parsed content blocks
    aRelationships      # [:id,:type,:target] from document.xml.rels
    cTitle
    cAuthor
    cSubject
    cKeywords
    cDescription
    cHeaderText
    cFooterText
    cEvenPageHeaderText
    cEvenPageFooterText
    cFirstPageHeaderText
    cFirstPageFooterText
    bSrcEvenAndOddHeaders
    bSrcFirstPageDifferent
    cDefaultFont
    nDefaultSize
    nPageWidth
    nPageHeight
    nDocSpacingAfter        # Source document default para space-after (twips)
    nDocSpacingLine         # Source document default line height (twips)
    cSrcStylesXml           # Source styles.xml content for round-trip
    cSrcThemeXml            # Source theme1.xml content for round-trip
    nHeaderMargin           # Source document header margin (twips)
    nFooterMargin           # Source document footer margin (twips)
    nMarginTop
    nMarginBottom
    nMarginLeft
    nMarginRight
    aSourceThemeColors
    aListStyles         # numbering definitions
    aFootnotes
    aEndnotes
    nSrcColumns
    nSrcColumnSpace
    bSrcPageBorder
    cSrcPageBorderStyle
    cSrcPageBorderColor
    nSrcPageBorderSize
    nSrcPageBorderSpace
    aSrcPageBorderSides   # Which sides: list of "top","left","bottom","right"
    cSrcPageBgColor
    aComments
    cSrcOrientation
    bSrcDocumentRTL
    cDocXml
    cStylesXml
    aCharStyles         # character style definitions
    bSrcPageNumbers
    cSrcPageNumAlign
    cSrcPageNumLocation
    bSrcHasTOC          # document contains a Table of Contents
    cSrcTOCTitle        # TOC heading title if found
    cSrcRawHeaderXml    # accumulated raw header XML for watermark detection
    bSrcWatermark       # true if a text watermark was detected
    cSrcWatermarkText   # watermark string
    cSrcWatermarkColor  # hex color
    bSrcImgWatermark    # true if an image watermark was detected
    cSrcImgWatermarkPath # path to the source image watermark file
    nSrcImgWatermarkOpacity  # image watermark opacity 0-100
    nSrcImgWatermarkWidthCm  # display width in cm
    nSrcImgWatermarkHeightCm # display height in cm
    nSrcWatermarkOpacity  # 0-100
    nSrcWatermarkRotation # degrees
    cSrcWatermarkFont   # font family
    nSrcWatermarkSize   # pt
    aMergeFields        # list of unique merge field names found in doc
    aCustomProps        # list of [:name, :type, :value] from docProps/custom.xml
    aNumInstances       # list of [:numId, :abstractNumId, :startAt0] from numbering.xml

    func init filePath
        cFilePath = filePath
        cTempDir      = ""
        aNumInstances = []
        aBlocks   = []
        aRelationships = []
        cTitle    = ""
        cAuthor   = ""
        cSubject  = ""
        cKeywords = ""
        cDescription = ""
        cHeaderText  = ""
        cFooterText  = ""
        cEvenPageHeaderText   = ""
        cEvenPageFooterText   = ""
        cFirstPageHeaderText  = ""
        cFirstPageFooterText  = ""
        bSrcEvenAndOddHeaders  = false
        bSrcFirstPageDifferent = false
        cDefaultFont = "Calibri"
        nDefaultSize = 11
        nDocSpacingAfter = -1   # -1 = not found in source
        nDocSpacingLine  = -1
        cSrcStylesXml = ""
        cSrcThemeXml  = ""
        nPageWidth   = 12240
        nPageHeight  = 15840
        nMarginTop    = 1440
        nMarginBottom = 1440
        nMarginLeft   = 1800
        nMarginRight  = 1800
        nHeaderMargin = 720
        nFooterMargin = 720
        aSourceThemeColors = []
        aListStyles = []
        aFootnotes  = []
        aEndnotes   = []
        nSrcColumns      = 1
        nSrcColumnSpace  = 720
        bSrcPageBorder       = false
        cSrcPageBorderStyle  = "single"
        cSrcPageBorderColor  = "000000"
        nSrcPageBorderSize   = 24
        nSrcPageBorderSpace  = 24
        aSrcPageBorderSides  = []
        cSrcPageBgColor  = ""
        aComments        = []
        cSrcOrientation  = "portrait"
        bSrcDocumentRTL  = false
        aCharStyles      = []
        bSrcPageNumbers  = false
        cSrcPageNumAlign = "center"
        cSrcPageNumLocation = "footer"
        bSrcHasTOC          = false
        cSrcTOCTitle        = "Table of Contents"
        cSrcRawHeaderXml    = ""
        bSrcWatermark       = false
        cSrcWatermarkText   = ""
        cSrcWatermarkColor  = "C0C0C0"
        bSrcImgWatermark    = false
        cSrcImgWatermarkPath = ""
        nSrcImgWatermarkOpacity  = 50
        nSrcImgWatermarkWidthCm  = 15
        nSrcImgWatermarkHeightCm = 15
        nSrcWatermarkOpacity  = 50
        nSrcWatermarkRotation = -45
        cSrcWatermarkFont   = "Arial"
        nSrcWatermarkSize   = 72
        aMergeFields        = []
        aCustomProps        = []
        cDocXml    = ""
        cStylesXml = ""
        loadDocx()
        return self

    func loadDocx
        sep = wordGetSep()

        # Reset mutable state so calling loadDocx() twice is safe
        # (init() already calls loadDocx internally; explicit call must not double data)
        aBlocks        = []
        aRelationships = []
        aNumInstances  = []
        aListStyles    = []
        aFootnotes     = []
        aEndnotes      = []
        aComments      = []
        aMergeFields   = []
        aCustomProps   = []
        aCharStyles    = []
        aSourceThemeColors = []
        cDocXml    = ""
        cStylesXml = ""

        # Clean up any previous temp dir before creating a new one
        if len(cTempDir) > 0
            cDir = cTempDir
            cLast = substr(cDir, len(cDir), 1)
            if cLast = "/" or cLast = "\"
                cDir = left(cDir, len(cDir) - 1)
            ok
            if len(cDir) > 0 and DirExists(cDir)
                OSDeleteFolder(cDir)
            ok
        ok

        # Extract the .docx (zip) to a temp directory
        cTempDir = "wrdr_tmp_" + string(clock()) + sep
        wordMakeDir(cTempDir)
        wrExtractZip(cFilePath, cTempDir)

        # Read document.xml
        docPath = cTempDir + "word" + sep + "document.xml"
        if !fexists(docPath)  return  ok
        cDocXml = read(docPath)

        # Read styles.xml
        styPath = cTempDir + "word" + sep + "styles.xml"
        if fexists(styPath)
            cStylesXml    = read(styPath)
            cSrcStylesXml = cStylesXml   # keep verbatim for round-trip
        ok

        # Read relationships
        relsPath = cTempDir + "word" + sep + "_rels" + sep + "document.xml.rels"
        if fexists(relsPath)
            relsXml = read(relsPath)
            rPos = 1
            while true
                rs = wrFindFrom(relsXml, "<Relationship ", rPos)
                if rs = 0  break  ok
                re = wrFindFrom(relsXml, "/>", rs)
                if re = 0  re = wrFindFrom(relsXml, "</Relationship>", rs)  ok
                if re = 0  break  ok
                relEl  = substr(relsXml, rs, re - rs + 2)
                relId  = wrAttr(relEl, "Id")
                relTypeFull = wrAttr(relEl, "Type")
                relTarget = wrAttr(relEl, "Target")
                # simplify type
                relType = "other"
                if wrFindFrom(relTypeFull, "/image",       1) > 0  relType = "image"      ok
                if wrFindFrom(relTypeFull, "/hyperlink",   1) > 0  relType = "hyperlink"  ok
                if wrFindFrom(relTypeFull, "/header",      1) > 0  relType = "header"     ok
                if wrFindFrom(relTypeFull, "/footer",      1) > 0  relType = "footer"     ok
                if wrFindFrom(relTypeFull, "/footnotes",   1) > 0  relType = "footnotes"  ok
                if wrFindFrom(relTypeFull, "/endnotes",    1) > 0  relType = "endnotes"   ok
                if wrFindFrom(relTypeFull, "/chart",       1) > 0  relType = "chart"      ok
                if wrFindFrom(relTypeFull, "/numbering",   1) > 0  relType = "numbering"  ok
                if relId != ""
                    aRelationships + [:id=relId, :type=relType, :target=relTarget]
                ok
                rPos = re + 1
            end
        ok

        # Read core properties (title, author, etc.)
        corePath = cTempDir + "docProps" + sep + "core.xml"
        if fexists(corePath)
            coreXml = read(corePath)
            t = ""
            ts = wrFindFrom(coreXml, "<dc:title>", 1)
            if ts > 0
                te = wrFindFrom(coreXml, "</dc:title>", ts)
                if te > 0  t = substr(coreXml, ts + 10, te - ts - 10)  ok
            ok
            cTitle = wrXmlUnescape(t)
            a = ""
            as_ = wrFindFrom(coreXml, "<dc:creator>", 1)
            if as_ > 0
                ae = wrFindFrom(coreXml, "</dc:creator>", as_)
                if ae > 0  a = substr(coreXml, as_ + 12, ae - as_ - 12)  ok
            ok
            cAuthor = wrXmlUnescape(a)
        ok

        # Read custom document properties (docProps/custom.xml)
        customPath = cTempDir + "docProps" + sep + "custom.xml"
        if fexists(customPath)
            customXml = read(customPath)
            cpPos2 = 1
            while true
                propS = wrFindFrom(customXml, "<vt:", cpPos2)
                if propS = 0
                    # Also try without namespace
                    propS = wrFindFrom(customXml, "<property ", cpPos2)
                    if propS = 0  break  ok
                ok
                # Find enclosing <property ...> element
                propTagS = cpPos2
                while propTagS > 0 and propTagS < propS
                    nextPT = wrFindFrom(customXml, "<property ", propTagS)
                    if nextPT = 0  break  ok
                    if nextPT > propS  break  ok
                    propTagS = nextPT + 1
                end
                propTagS = wrFindFrom(customXml, "<property ", cpPos2)
                if propTagS = 0  break  ok
                propTagE = wrFindCloseTag(customXml, "property", propTagS)
                if propTagE = 0  propTagE = wrFindFrom(customXml, "</property>", propTagS)  ok
                if propTagE = 0  break  ok
                propEl = substr(customXml, propTagS, propTagE - propTagS + 11)
                propName = wrAttr(propEl, "name")
                # Find value: <vt:lpwstr>, <vt:i4>, <vt:r8>, <vt:bool>, <vt:filetime>
                propType = "string"; propValue = ""
                for vtTag in ["vt:lpwstr","vt:lpstr","vt:i4","vt:i8","vt:r8",
                               "vt:bool","vt:filetime","vt:decimal","vt:int"]
                    vtS = wrFindFrom(propEl, "<" + vtTag + ">", 1)
                    if vtS > 0
                        vtE = wrFindFrom(propEl, "</" + vtTag + ">", vtS)
                        if vtE > 0
                            propValue = wrXmlUnescape(substr(propEl, vtS + len(vtTag) + 2, vtE - vtS - len(vtTag) - 2))
                            if vtTag = "vt:i4" or vtTag = "vt:i8" or vtTag = "vt:int"
                                propType = "integer"
                            elseif vtTag = "vt:r8" or vtTag = "vt:decimal"
                                propType = "decimal"
                            elseif vtTag = "vt:bool"
                                propType = "boolean"
                            elseif vtTag = "vt:filetime"
                                propType = "datetime"
                            else
                                propType = "string"
                            ok
                            break
                        ok
                    ok
                next
                if len(propName) > 0
                    aCustomProps + [:name=propName, :type=propType, :value=propValue]
                ok
                cpPos2 = propTagE + 1
            end
        ok

        # Parse default font + size from styles.xml
        if len(cStylesXml) > 0
            # docDefaults -> rPrDefault -> rPr
            ddS = wrFindFrom(cStylesXml, "<w:docDefaults>", 1)
            if ddS = 0  ddS = wrFindFrom(cStylesXml, "<w:docDefaults ", 1)  ok
            if ddS > 0
                ddE = wrFindCloseTag(cStylesXml, "w:docDefaults", ddS)
                if ddE > 0
                    ddXml = substr(cStylesXml, ddS, ddE - ddS)
                    fnS = wrFindFrom(ddXml, "<w:rFonts ", 1)
                    if fnS > 0
                        fnEl = substr(ddXml, fnS, 120)
                        fv = wrAttr(fnEl, "w:ascii")
                        if len(fv) > 0  cDefaultFont = fv  ok
                    ok
                    szS = wrFindFrom(ddXml, "<w:sz ", 1)
                    if szS > 0
                        szEl = substr(ddXml, szS, 60)
                        sv = wrAttr(szEl, "w:val")
                        if len(sv) > 0  nDefaultSize = floor(number(sv) / 2)  ok
                    ok
                    # Read pPrDefault paragraph spacing
                    pprdS = wrFindFrom(ddXml, "<w:pPrDefault>", 1)
                    if pprdS = 0  pprdS = wrFindFrom(ddXml, "<w:pPrDefault ", 1)  ok
                    if pprdS > 0
                        pprdE = wrFindCloseTag(ddXml, "w:pPrDefault", pprdS)
                        if pprdE > 0
                            pprdXml = substr(ddXml, pprdS, pprdE - pprdS)
                            spDS = wrFindFrom(pprdXml, "<w:spacing ", 1)
                            if spDS > 0
                                spDEl = substr(pprdXml, spDS, 150)
                                spDAv = wrAttr(spDEl, "w:after")
                                spDLv = wrAttr(spDEl, "w:line")
                                if len(spDAv) > 0  nDocSpacingAfter = number(spDAv)  ok
                                if len(spDLv) > 0  nDocSpacingLine  = number(spDLv)   ok
                            ok
                            # Detect document-level RTL from pPrDefault <w:bidi/>
                            if wrFindFrom(pprdXml, "<w:bidi/>", 1) > 0 or
                               wrFindFrom(pprdXml, "<w:bidi ",  1) > 0
                                bSrcDocumentRTL = true
                            ok
                        ok
                    ok
                ok
            ok
        ok

        # Page size + margins from final sectPr
        bodyS = wrFindFrom(cDocXml, "<w:body>", 1)
        if bodyS = 0  bodyS = wrFindFrom(cDocXml, "<w:body ", 1)  ok
        if bodyS > 0
            bodyE = wrFindCloseTag(cDocXml, "w:body", bodyS)
            if bodyE > 0
                bodyXml = substr(cDocXml, bodyS, bodyE - bodyS)
                # Find the last sectPr
                lastSpS = 1
                lastSectPos = 0
                while true
                    spX = wrFindFrom(bodyXml, "<w:sectPr", lastSpS)
                    if spX = 0  break  ok
                    lastSectPos = spX
                    lastSpS = spX + 1
                end
                if lastSectPos > 0
                    lastSectXml = substr(bodyXml, lastSectPos, 2000)
                    # Page size
                    pgszS = wrFindFrom(lastSectXml, "<w:pgSz ", 1)
                    if pgszS > 0
                        pgszEl = substr(lastSectXml, pgszS, 150)
                        wv = wrAttr(pgszEl, "w:w")
                        hv = wrAttr(pgszEl, "w:h")
                        ov = wrAttr(pgszEl, "w:orient")
                        if len(wv) > 0  nPageWidth  = number(wv)  ok
                        if len(hv) > 0  nPageHeight = number(hv)  ok
                        if ov = "landscape"  cSrcOrientation = "landscape"  ok
                    ok
                    # Margins
                    pgMrS = wrFindFrom(lastSectXml, "<w:pgMar ", 1)
                    if pgMrS > 0
                        pgMrEl = substr(lastSectXml, pgMrS, 200)
                        tv  = wrAttr(pgMrEl, "w:top")
                        bv  = wrAttr(pgMrEl, "w:bottom")
                        lv  = wrAttr(pgMrEl, "w:left")
                        rv  = wrAttr(pgMrEl, "w:right")
                        if len(tv) > 0  nMarginTop    = number(tv)  ok
                        if len(bv) > 0  nMarginBottom = number(bv)  ok
                        if len(lv) > 0  nMarginLeft   = number(lv)  ok
                        if len(rv) > 0  nMarginRight  = number(rv)  ok
                        hv2 = wrAttr(pgMrEl, "w:header")
                        fv2 = wrAttr(pgMrEl, "w:footer")
                        if len(hv2) > 0  nHeaderMargin = number(hv2)  ok
                        if len(fv2) > 0  nFooterMargin = number(fv2)  ok
                    ok
                    # Columns
                    colS = wrFindFrom(lastSectXml, "<w:cols ", 1)
                    if colS > 0
                        colEl = substr(lastSectXml, colS, 120)
                        cnv = wrAttr(colEl, "w:num")
                        csv = wrAttr(colEl, "w:space")
                        if len(cnv) > 0  nSrcColumns     = number(cnv)  ok
                        if len(csv) > 0  nSrcColumnSpace = number(csv)  ok
                    ok
                    # Page border
                    pgBdrS = wrFindFrom(lastSectXml, "<w:pgBorders", 1)
                    if pgBdrS > 0
                        pgBdrE = wrFindCloseTag(lastSectXml, "w:pgBorders", pgBdrS)
                        if pgBdrE > 0
                            bdrXml = substr(lastSectXml, pgBdrS, pgBdrE - pgBdrS)
                            # Detect which sides are present and read attrs from first found side
                            aSrcPageBorderSides = []
                            bdrAttrRead = false
                            for bdrSideName in ["top", "left", "bottom", "right"]
                                bdrSideTag = "<w:" + bdrSideName + " "
                                bdrSideS2 = wrFindFrom(bdrXml, bdrSideTag, 1)
                                if bdrSideS2 > 0
                                    aSrcPageBorderSides + bdrSideName
                                    if !bdrAttrRead
                                        bdrEl = substr(bdrXml, bdrSideS2, 120)
                                        bSrcPageBorder = true
                                        sv2 = wrAttr(bdrEl, "w:val")
                                        cv2 = wrAttr(bdrEl, "w:color")
                                        szv = wrAttr(bdrEl, "w:sz")
                                        spv = wrAttr(bdrEl, "w:space")
                                        if len(sv2) > 0  cSrcPageBorderStyle = sv2          ok
                                        if len(cv2) > 0  cSrcPageBorderColor = cv2          ok
                                        if len(szv) > 0  nSrcPageBorderSize  = number(szv)  ok
                                        if len(spv) > 0  nSrcPageBorderSpace = number(spv)  ok
                                        bdrAttrRead = true
                                    ok
                                ok
                            next
                        ok
                    ok
                ok
            ok
        ok

        # Footnotes and endnotes
        fnPath = cTempDir + "word" + sep + "footnotes.xml"
        enPath = cTempDir + "word" + sep + "endnotes.xml"
        if fexists(fnPath)  aFootnotes = wrParseNotes(read(fnPath), "w:footnote")  ok
        if fexists(enPath)  aEndnotes  = wrParseNotes(read(enPath), "w:endnote")   ok

        # Comments
        cmPath = cTempDir + "word" + sep + "comments.xml"
        if fexists(cmPath)  aComments = wrParseComments(read(cmPath))  ok

        # Page background color
        cSrcPageBgColor = wrParsePageBgColor(cDocXml)

        # Character styles
        if len(cStylesXml) > 0
            aCharStyles = wrParseCharStyles(cStylesXml)
        ok

        # Document RTL - check settings.xml
        # Read theme1.xml for round-trip fidelity
        themePath = cTempDir + "word" + sep + "theme" + sep + "theme1.xml"
        if fexists(themePath)  cSrcThemeXml = read(themePath)  ok

        settPath = cTempDir + "word" + sep + "settings.xml"
        if fexists(settPath)
            settXml = read(settPath)
            if wrFindFrom(settXml, "<w:bidi/>", 1) > 0 or
               wrFindFrom(settXml, "<w:bidi ", 1) > 0
                bSrcDocumentRTL = true
            ok
        ok

        # Headers and footers
        wrLoadHeadersFooters(self)

        # Page number field detection
        sep2 = wordGetSep()
        hdrRaw = ""; ftrRaw = ""
        hdrPath2 = cTempDir + "word" + sep2 + "header1.xml"
        ftrPath2 = cTempDir + "word" + sep2 + "footer1.xml"
        if fexists(hdrPath2)  hdrRaw = read(hdrPath2)  ok
        if fexists(ftrPath2)  ftrRaw = read(ftrPath2)  ok
        if len(hdrRaw) + len(ftrRaw) > 0
            pgNum = wrDetectPageNumbers(hdrRaw, ftrRaw)
            if pgNum[:found] = true
                bSrcPageNumbers     = true
                cSrcPageNumAlign    = pgNum[:align]
                cSrcPageNumLocation = pgNum[:location]
            ok
        ok

        # Watermark detection from header XML
        wmResult = wrParseWatermark(cSrcRawHeaderXml)
        if wmResult[:found] = true
            bSrcWatermark        = true
            cSrcWatermarkText    = wmResult[:text]
            cSrcWatermarkColor   = wmResult[:color]
            nSrcWatermarkOpacity = wmResult[:opacity]
            nSrcWatermarkRotation= wmResult[:rotation]
            cSrcWatermarkFont    = wmResult[:font]
            nSrcWatermarkSize    = wmResult[:size]
        ok

        # Image watermark detection from header XML + rels
        imgWmResult = wrParseImageWatermark(cSrcRawHeaderXml)
        if imgWmResult[:found] = true
            imgRelId = imgWmResult[:relId]
            imgWmPath = ""
            hdrRelsPath = cTempDir + "word" + sep + "_rels" + sep + "header1.xml.rels"
            if fexists(hdrRelsPath)
                hdrRelsXml = read(hdrRelsPath)
                relS = wrFindFrom(hdrRelsXml, "Id=" + char(34) + imgRelId + char(34), 1)
                if relS > 0
                    relEl = substr(hdrRelsXml, relS, 200)
                    relTarget = wrAttr(relEl, "Target")
                    if len(relTarget) > 0
                        imgWmPath = cTempDir + "word" + sep + relTarget
                    ok
                ok
            ok
            if len(imgWmPath) > 0 and fexists(imgWmPath)
                bSrcImgWatermark = true
                cSrcImgWatermarkPath = imgWmPath
                nSrcImgWatermarkOpacity = imgWmResult[:opacity]
                # Convert pt to cm (1cm = 28.35pt)
                nSrcImgWatermarkWidthCm  = imgWmResult[:widthPt]  / 28.35
                nSrcImgWatermarkHeightCm = imgWmResult[:heightPt] / 28.35
            ok
        ok

        # TOC detection
        tocResult = wrDetectTOC(cDocXml)
        if tocResult[:found] = true
            bSrcHasTOC   = true
            cSrcTOCTitle = tocResult[:title]
        ok

        # Theme colors
        themePath = cTempDir + "word" + sep + "theme" + sep + "theme1.xml"
        if fexists(themePath)
            themeXml = read(themePath)
            # Look for dk1/lt1/accent1/accent2 in lumMod
            dk2S = wrFindFrom(themeXml, "dk2", 1)
            if dk2S > 0
                # Get the srgbClr value near dk2
                srgbS = wrFindFrom(themeXml, "<a:srgbClr ", dk2S)
                if srgbS > 0 and srgbS < dk2S + 500
                    srgbEl = substr(themeXml, srgbS, 80)
                    hexv = wrAttr(srgbEl, "val")
                    if len(hexv) = 6
                        aSourceThemeColors = ["_custom_", hexv, hexv]
                    ok
                ok
            ok
        ok

        # Parse numbering for list detection
        numPath = cTempDir + "word" + sep + "numbering.xml"
        if fexists(numPath)
            numXml = read(numPath)
            nPos = 1
            while true
                abstrS = wrFindFrom(numXml, "<w:abstractNum ", nPos)
                if abstrS = 0  break  ok
                abstrE = wrFindCloseTag(numXml, "w:abstractNum", abstrS)
                if abstrE = 0  break  ok
                abstrXml = substr(numXml, abstrS, abstrE - abstrS)
                aId = wrAttr(abstrXml, "w:abstractNumId")
                # Check lvl 0 numFmt
                lvlS = wrFindFrom(abstrXml, "<w:lvl ", 1)
                if lvlS > 0
                    lvlE = wrFindCloseTag(abstrXml, "w:lvl", lvlS)
                    if lvlE > 0
                        lvlXml = substr(abstrXml, lvlS, lvlE - lvlS)
                        fmtS = wrFindFrom(lvlXml, "<w:numFmt ", 1)
                        isBullet = false
                        if fmtS > 0
                            fmtEl = substr(lvlXml, fmtS, 80)
                            fmtV = wrAttr(fmtEl, "w:val")
                            if fmtV = "bullet"  isBullet = true  ok
                        ok
                        if len(aId) > 0
                            aListStyles + [:id = number(aId), :isBullet = isBullet]
                        ok
                    ok
                ok
                nPos = abstrE
            end
            # Parse <w:num> elements for numId->abstractNumId mapping + startOverride
            n2Pos = 1
            while true
                numElS2 = wrFindFrom(numXml, "<w:num ", n2Pos)
                if numElS2 = 0  break  ok
                numElE2 = wrFindCloseTag(numXml, "w:num", numElS2)
                if numElE2 = 0  break  ok
                numElXml2  = substr(numXml, numElS2, numElE2 - numElS2)
                numHdrEl2  = substr(numXml, numElS2, 60)
                numIdStr2  = wrAttr(numHdrEl2, "w:numId")
                # abstractNumId reference
                absRefS2 = wrFindFrom(numElXml2, "<w:abstractNumId ", 1)
                absNumId2 = 0
                if absRefS2 > 0
                    absRefEl2 = substr(numElXml2, absRefS2, 60)
                    absIdStr2 = wrAttr(absRefEl2, "w:val")
                    if len(absIdStr2) > 0  absNumId2 = number(absIdStr2)  ok
                ok
                # Level 0 startOverride
                startAt0_2 = -1
                loScan2 = 1
                while true
                    loS2 = wrFindFrom(numElXml2, "<w:lvlOverride ", loScan2)
                    if loS2 = 0  break  ok
                    loE2 = wrFindCloseTag(numElXml2, "w:lvlOverride", loS2)
                    if loE2 = 0  break  ok
                    loXml2  = substr(numElXml2, loS2, loE2 - loS2)
                    loHdr2  = substr(numElXml2, loS2, 80)
                    ilvlStr2 = wrAttr(loHdr2, "w:ilvl")
                    if ilvlStr2 = "0"
                        soS2 = wrFindFrom(loXml2, "<w:startOverride ", 1)
                        if soS2 > 0
                            soEl2 = substr(loXml2, soS2, 60)
                            soV2  = wrAttr(soEl2, "w:val")
                            if len(soV2) > 0  startAt0_2 = number(soV2)  ok
                        ok
                    ok
                    loScan2 = loE2
                end
                if len(numIdStr2) > 0
                    aNumInstances + [:numId=number(numIdStr2), :abstractNumId=absNumId2, :startAt0=startAt0_2]
                ok
                n2Pos = numElE2
            end
        ok

        # Parse body
        bodyS = wrFindFrom(cDocXml, "<w:body>", 1)
        if bodyS = 0  bodyS = wrFindFrom(cDocXml, "<w:body ", 1)  ok
        if bodyS > 0
            bodyE = wrFindCloseTag(cDocXml, "w:body", bodyS)
            if bodyE > 0
                bodyXml = substr(cDocXml, bodyS, bodyE - bodyS)
                parseBody(bodyXml)
            ok
        ok

        # Resolve footnote/endnote text (second pass — avoids depth-scope issue)
        for block in aBlocks
            bT = block[:type]
            if bT = "footnote"
                nId = block[:noteId]
                for fn in aFootnotes
                    if fn[:id] = nId
                        block[:noteText] = fn[:text]
                        block[:noteRuns] = fn[:runs]
                    ok
                next
            ok
            if bT = "endnote"
                nId = block[:noteId]
                for en in aEndnotes
                    if en[:id] = nId
                        block[:noteText] = en[:text]
                        block[:noteRuns] = en[:runs]
                    ok
                next
            ok
        next

    func parseBody bodyXml
        # wrParseAnchoredImages still called by getFloatingImageWrapTypes() query
        # but is NOT pre-added to aBlocks here; parseParagraph handles them in order

        # Text boxes are now preserved as rawparagraph blocks
        # at their original positions (wrParseTextBoxes pre-insertion removed)

        # Use wrSplitBodyElements for O(n) body parsing (avoids O(n^2) search)
        bodyElements = wrSplitBodyElements(bodyXml)
        for elem in bodyElements
            elemTag = elem[:tag]
            elemXml = elem[:xml]

            if elemTag = "w:tbl"
                tblBlock = parseTable(elemXml)
                aBlocks + tblBlock
            elseif elemTag = "w:sdt"
                sdtInfo = wrParseSdtBlock(elemXml)
                sdtType = sdtInfo[:type]
                if sdtType = "checkbox" or sdtType = "dropdown" or sdtType = "text"
                    sdtBlock = []
                    sdtBlock[:type]        = "formfield"
                    sdtBlock[:fieldType]   = sdtType
                    sdtBlock[:label]       = sdtInfo[:label]
                    sdtBlock[:checked]     = sdtInfo[:checked]
                    sdtBlock[:choices]     = sdtInfo[:choices]
                    sdtBlock[:default]     = sdtInfo[:default]
                    sdtBlock[:placeholder] = sdtInfo[:placeholder]
                    sdtBlock[:bookmark]    = ""
                    aBlocks + sdtBlock
                ok
            else
                # w:p paragraph
                pXml = elemXml
                # Handle self-closing <w:p .../> (empty para, no content)
                if wrIsSelfClosingTag(pXml, 1)
                    loop
                ok
                # Intercept TOC/SEQ field paragraphs + special-style paragraphs
                isTOCPara = false
                # TableOfFigures / Caption style paragraphs must be preserved verbatim
                pStyleCheck = substr(pXml, 1, 300)
                pStyleIsSpecial = (wrFindFrom(pStyleCheck, "TableOfFigures", 1) > 0) or
                                  (wrFindFrom(pStyleCheck, "Caption", 1) > 0 and wrFindFrom(pStyleCheck, "pStyle", 1) > 0)
                if pStyleIsSpecial
                    isTOCPara = true
                    rawBlock0 = [:type="rawparagraph", :rawXml=pXml, :text="", :bookmark=""]
                    aBlocks + rawBlock0
                ok
                if !isTOCPara
                    if wrFindFrom(pXml, "instrText", 1) > 0 and wrFindFrom(pXml, "MERGEFIELD", 1) = 0
                        itS3 = wrFindFrom(pXml, "<w:instrText", 1)
                        if itS3 > 0
                            itE3 = wrFindFrom(pXml, "</w:instrText>", itS3)
                            if itE3 > 0
                                itTxt3 = substr(pXml, itS3, itE3 - itS3)
                                isTocMain = (wrFindFrom(itTxt3, " TOC ", 1) > 0) and
                                            (wrFindFrom(itTxt3, char(92)+"o", 1) > 0 or
                                             wrFindFrom(itTxt3, char(92)+"h", 1) > 0 or
                                             wrFindFrom(itTxt3, char(92)+"u", 1) > 0 or
                                             wrFindFrom(itTxt3, char(92)+"z", 1) > 0)
                                isTocCaption = (wrFindFrom(itTxt3, " TOC ", 1) > 0) and
                                               (wrFindFrom(itTxt3, char(92)+"c", 1) > 0)
                                isSeqField = wrFindFrom(itTxt3, " SEQ ", 1) > 0
                                if isTocMain
                                    isTOCPara = true
                                    tocBlock3 = [:type="toc", :title=cSrcTOCTitle, :bookmark=""]
                                    aBlocks + tocBlock3
                                elseif isTocCaption or isSeqField
                                    isTOCPara = true
                                    rawBlock3 = [:type="rawparagraph", :rawXml=pXml, :text="", :bookmark=""]
                                    aBlocks + rawBlock3
                                ok
                            ok
                        ok
                    ok
                ok

                # Intercept MERGEFIELD paragraphs
                isMergePara = false
                isFieldPara2 = false
                if wrFindFrom(pXml, "MERGEFIELD", 1) > 0 and wrFindFrom(pXml, "instrText", 1) > 0
                    mfFields = wrParseMergeFields(pXml)
                    if len(mfFields) > 0
                        isMergePara = true
                        mfRuns = []
                        mfText = ""
                        mfScanPos = 1
                        mfInField = false
                        mfInDisplay = false
                        while true
                            nT1  = wrFindFrom(pXml, "<w:t>",           mfScanPos)
                            nT2  = wrFindFrom(pXml, "<w:t ",           mfScanPos)
                            nIT  = wrFindFrom(pXml, "<w:instrText",    mfScanPos)
                            nFC  = wrFindFrom(pXml, "<w:fldChar ",     mfScanPos)
                            nT   = wrMinPos(nT1, nT2)
                            nMin = nT
                            if nIT > 0
                                if nMin = 0 or nIT < nMin  nMin = nIT  ok
                            ok
                            if nFC > 0
                                if nMin = 0 or nFC < nMin  nMin = nFC  ok
                            ok
                            if nMin = 0  break  ok
                            if nFC = nMin
                                fcEl = substr(pXml, nFC, 100)
                                fcType = wrAttr(fcEl, "w:fldCharType")
                                if fcType = "begin"
                                    mfInField   = true
                                    mfInDisplay = false
                                elseif fcType = "separate"
                                    mfInField   = false
                                    mfInDisplay = true
                                elseif fcType = "end"
                                    mfInField   = false
                                    mfInDisplay = false
                                ok
                                fcE = wrFindFrom(pXml, ">", nFC)
                                if fcE = 0  break  ok
                                mfScanPos = fcE + 1
                            elseif nIT = nMin
                                itTagE   = wrFindFrom(pXml, ">", nIT)
                                itEndTag = wrFindFrom(pXml, "</w:instrText>", nIT)
                                if itTagE > 0 and itEndTag > 0
                                    itContent = trim(substr(pXml, itTagE+1, itEndTag-itTagE-1))
                                    mfIdx2 = wrFindFrom(itContent, "MERGEFIELD", 1)
                                    if mfIdx2 > 0
                                        mfNameRaw = trim(substr(itContent, mfIdx2 + 10))
                                        spPos = wrFindFrom(mfNameRaw, " ", 1)
                                        if spPos > 0  mfNameRaw = substr(mfNameRaw, 1, spPos-1)  ok
                                        if mfInField and len(mfNameRaw) > 0
                                            if len(mfText) > 0
                                                mfRuns + mfText
                                                mfText = ""
                                            ok
                                            mfRuns + [:field=mfNameRaw]
                                        ok
                                    ok
                                ok
                                mfScanPos = itEndTag + 14
                            else
                                # w:t text
                                tTagE = wrFindFrom(pXml, ">", nT)
                                tEndTag = wrFindFrom(pXml, "</w:t>", tTagE)
                                if tTagE > 0 and tEndTag > 0
                                    tContent = substr(pXml, tTagE+1, tEndTag-tTagE-1)
                                    if !mfInDisplay
                                        mfText += tContent
                                    ok
                                ok
                                mfScanPos = tEndTag + 6
                            ok
                        end
                        if len(mfText) > 0  mfRuns + mfText  ok
                        mfBlock = []
                        mfBlock[:type]      = "mergefield"
                        mfBlock[:runs]      = mfRuns
                        mfBlock[:bookmark]  = ""
                        aBlocks + mfBlock
                    ok
                ok
                # Intercept REF/PAGEREF cross-reference field paragraphs
                isRefPara = false
                if !isMergePara and !isTOCPara
                    if wrFindFrom(pXml, "instrText", 1) > 0
                        itRefS = wrFindFrom(pXml, "<w:instrText", 1)
                        if itRefS > 0
                            itRefE = wrFindFrom(pXml, "</w:instrText>", itRefS)
                            if itRefE > 0
                                itRefTxt = substr(pXml, itRefS, itRefE - itRefS)
                                if wrFindFrom(itRefTxt, " REF ", 1) > 0 or
                                   wrFindFrom(itRefTxt, " PAGEREF ", 1) > 0
                                    isRefPara = true
                                    refBlock2 = [:type="rawparagraph",
                                                  :rawXml=pXml, :text="", :bookmark=""]
                                    aBlocks + refBlock2
                                ok
                            ok
                        ok
                    ok
                ok
                if !isMergePara and !isFieldPara2 and !isTOCPara and !isRefPara
                    pBlock = parseParagraph(pXml)
                    aBlocks + pBlock
                ok
            ok
        next

    func parseParagraph pXml
        block = []

        # --- Paragraph properties ---
        pPrXml = ""
        pPrS = wrFindFrom(pXml, "<w:pPr>", 1)
        if pPrS = 0  pPrS = wrFindFrom(pXml, "<w:pPr ", 1)  ok
        if pPrS > 0
            pPrE = wrFindCloseTag(pXml, "w:pPr", pPrS)
            if pPrE > 0
                pPrXml = substr(pXml, pPrS, pPrE - pPrS)
            ok
        ok

        # Style name
        styleName = ""
        if len(pPrXml) > 0
            pStS = wrFindFrom(pPrXml, "<w:pStyle ", 1)
            if pStS > 0
                pStEl = substr(pPrXml, pStS, 80)
                styleName = wrAttr(pStEl, "w:val")
            ok
        ok

        # Heading level
        headingLevel = 0
        if len(styleName) > 0
            if lower(styleName) = "heading1" or styleName = "Heading1"  headingLevel = 1  ok
            if lower(styleName) = "heading2" or styleName = "Heading2"  headingLevel = 2  ok
            if lower(styleName) = "heading3" or styleName = "Heading3"  headingLevel = 3  ok
            if lower(styleName) = "heading4" or styleName = "Heading4"  headingLevel = 4  ok
            if lower(styleName) = "heading5" or styleName = "Heading5"  headingLevel = 5  ok
            if lower(styleName) = "heading6" or styleName = "Heading6"  headingLevel = 6  ok
            sn = lower(styleName)
            if sn = "heading 1" or sn = "title"     headingLevel = 1  ok
            if sn = "heading 2" or sn = "subtitle"  headingLevel = 2  ok
            if sn = "heading 3"                     headingLevel = 3  ok
            if sn = "heading 4"                     headingLevel = 4  ok
            if sn = "heading 5"                     headingLevel = 5  ok
            if sn = "heading 6"                     headingLevel = 6  ok
            # Numeric IDs
            if headingLevel = 0
                if styleName = "1"  headingLevel = 1  ok
                if styleName = "2"  headingLevel = 2  ok
                if styleName = "3"  headingLevel = 3  ok
            ok
        ok

        # Caption style
        isCaption = false
        if lower(styleName) = "caption"  isCaption = true  ok

        # Alignment
        alignVal = ""
        if len(pPrXml) > 0
            jcS = wrFindFrom(pPrXml, "<w:jc ", 1)
            if jcS > 0
                jcEl = substr(pPrXml, jcS, 80)
                alignVal = wrAttr(jcEl, "w:val")
            ok
        ok

        # Space before/after + line spacing
        spaceBefore   = 0
        spaceAfter    = 0
        lineSpacing   = 0.0
        lineSpacingPt = 0.0
        lineRule      = ""
        if len(pPrXml) > 0
            spS = wrFindFrom(pPrXml, "<w:spacing ", 1)
            if spS > 0
                spEl = substr(pPrXml, spS, 180)
                sbv  = wrAttr(spEl, "w:before")
                sav  = wrAttr(spEl, "w:after")
                slv  = wrAttr(spEl, "w:line")
                srv  = wrAttr(spEl, "w:lineRule")
                if len(sbv) > 0  spaceBefore = number(sbv)  ok
                if len(sav) > 0  spaceAfter  = number(sav)   ok
                if len(slv) > 0
                    lineRule = srv
                    if lineRule = "auto" or lineRule = ""
                        lineSpacing = number(slv) / 240.0
                    else
                        lineSpacingPt = number(slv) / 20.0
                    ok
                ok
            ok
        ok

        # Paragraph background shading
        paraBgColor = ""
        if len(pPrXml) > 0
            pshdS = wrFindFrom(pPrXml, "<w:shd ", 1)
            if pshdS > 0
                pshdEl = substr(pPrXml, pshdS, 100)
                pfill = wrAttr(pshdEl, "w:fill")
                if len(pfill) > 0 and pfill != "auto"  paraBgColor = pfill  ok
            ok
        ok

        # Indent, RTL, hline, blockquote, paragraph border
        isHorizLine      = false
        isBlockQuote     = false
        isRTL            = false
        paraHasBorder    = false
        paraBdrStyle = "single"
        paraBdrColor = "000000"
        paraBdrSize  = 6
        paraBdrSpace = 4
        paraBdrSides = []
        keepNext         = false
        keepLines        = false
        pageBreakBefore  = false
        contextualSpacing = false
        outlineLvl       = -1
        indentLeft       = 0
        indentRight      = 0
        indentFirstLine  = 0
        widowControl     = true    # Word default is ON; false means explicitly disabled
        noHyphenate      = false   
        if len(pPrXml) > 0
            # RTL
            if wrFindFrom(pPrXml, "<w:bidi/>", 1) > 0 or
               wrFindFrom(pPrXml, "<w:bidi ",  1) > 0
                isRTL = true
            ok
            # Keep-with-next / keep-lines / page-break-before
            if wrFindFrom(pPrXml, "<w:keepNext/>", 1) > 0 or
               wrFindFrom(pPrXml, "<w:keepNext ",  1) > 0
                keepNext = true
            ok
            if wrFindFrom(pPrXml, "<w:keepLines/>", 1) > 0 or
               wrFindFrom(pPrXml, "<w:keepLines ", 1) > 0
                keepLines = true
            ok
            if wrFindFrom(pPrXml, "<w:pageBreakBefore/>", 1) > 0 or
               wrFindFrom(pPrXml, "<w:pageBreakBefore ", 1) > 0
                pageBreakBefore = true
            ok
            # widow/orphan control — only detect explicit disable
            wctS = wrFindFrom(pPrXml, "<w:widowControl ", 1)
            if wctS > 0
                wctEl = substr(pPrXml, wctS, 80)
                wctV  = wrAttr(wctEl, "w:val")
                if wctV = "0" or wctV = "false"  widowControl = false  ok
            ok
            # suppress auto-hyphenation
            if wrFindFrom(pPrXml, "<w:suppressAutoHyphens/>", 1) > 0 or
               wrFindFrom(pPrXml, "<w:suppressAutoHyphens ",  1) > 0
                noHyphenate = true
            ok
            # Indent
            indS = wrFindFrom(pPrXml, "<w:ind ", 1)
            if indS > 0
                indEl = substr(pPrXml, indS, 160)
                ilv = wrAttr(indEl, "w:left")
                irv = wrAttr(indEl, "w:right")
                ifl = wrAttr(indEl, "w:firstLine")
                ihg = wrAttr(indEl, "w:hanging")
                if len(ilv) > 0  indentLeft      = number(ilv)   ok
                if len(irv) > 0  indentRight     = number(irv)   ok
                if len(ifl) > 0  indentFirstLine = number(ifl)   ok
                if len(ihg) > 0  indentFirstLine = -number(ihg)  ok
            ok
            # Outline level (for TOC depth / navigation pane)
            olvS = wrFindFrom(pPrXml, "<w:outlineLvl ", 1)
            if olvS > 0
                olvEl = substr(pPrXml, olvS, 80)
                olvV  = wrAttr(olvEl, "w:val")
                if len(olvV) > 0  outlineLvl = number(olvV)  ok
            ok
            # Contextual spacing
            if wrFindFrom(pPrXml, "<w:contextualSpacing", 1) > 0
                contextualSpacing = true
            ok
            # Paragraph border
            pBdrData = wrParsePBdr(pPrXml)
            if pBdrData[:found] = true
                # Existing heuristics: blockquote or hline
                if wrFindFrom(pPrXml, "<w:left ", 1) > 0 and indentLeft >= 200
                    isBlockQuote = true
                else
                    if wrFindFrom(pPrXml, "<w:bottom ", 1) > 0 and
                       wrFindFrom(pPrXml, "<w:left ",   1) = 0 and
                       wrFindFrom(pXml, "<w:t>", 1) = 0 and
                       wrFindFrom(pXml, "<w:t ", 1) = 0
                        isHorizLine = true
                    ok
                ok
                # Always save border data for replay
                paraHasBorder     = true
                paraBdrStyle  = pBdrData[:style]
                paraBdrColor  = pBdrData[:color]
                paraBdrSize   = pBdrData[:size]
                paraBdrSpace  = pBdrData[:space]
                paraBdrSides  = pBdrData[:sides]
            ok
        ok

        # List item detection
        isListItem  = false
        listNumId   = 0
        listIlvl    = 0
        listIsBullet = false
        listStartAt  = -1
        if len(pPrXml) > 0
            numPrS = wrFindFrom(pPrXml, "<w:numPr>", 1)
            if numPrS = 0  numPrS = wrFindFrom(pPrXml, "<w:numPr ", 1)  ok
            if numPrS > 0
                numPrE = wrFindCloseTag(pPrXml, "w:numPr", numPrS)
                if numPrE > 0
                    numPrXml = substr(pPrXml, numPrS, numPrE - numPrS)
                    ilvlS = wrFindFrom(numPrXml, "<w:ilvl ", 1)
                    numIdS = wrFindFrom(numPrXml, "<w:numId ", 1)
                    if ilvlS > 0
                        ilvlEl = substr(numPrXml, ilvlS, 60)
                        lv = wrAttr(ilvlEl, "w:val")
                        if len(lv) > 0  listIlvl = number(lv)  ok
                    ok
                    if numIdS > 0
                        numIdEl = substr(numPrXml, numIdS, 60)
                        nv = wrAttr(numIdEl, "w:val")
                        if len(nv) > 0
                            listNumId = number(nv)
                            if listNumId > 0  isListItem = true  ok
                        ok
                    ok
                    # Determine bullet vs numbered from aListStyles
                    # Resolve numId -> abstractNumId first, then look up isBullet
                    listAbstractId = -1
                    for ni2 in aNumInstances
                        if ni2[:numId] = listNumId
                            listAbstractId = ni2[:abstractNumId]
                        ok
                    next
                    for ls in aListStyles
                        if listAbstractId >= 0
                            if ls[:id] = listAbstractId
                                listIsBullet = ls[:isBullet]
                            ok
                        else
                            if ls[:id] = listNumId
                                listIsBullet = ls[:isBullet]
                            ok
                        ok
                    next
                    # Resolve startOverride from aNumInstances for numbered lists
                    if isListItem and !listIsBullet
                        for ni in aNumInstances
                            if ni[:numId] = listNumId
                                if ni[:startAt0] > 0
                                    listStartAt = ni[:startAt0]
                                ok
                            ok
                        next
                    ok
                ok
            ok
        ok

        # Bookmark
        bookmarkName = ""
        bmS = wrFindFrom(pXml, "<w:bookmarkStart ", 1)
        if bmS > 0
            bmEl  = substr(pXml, bmS, 120)
            bmName = wrAttr(bmEl, "w:name")
            if bmName != "_GoBack" and len(bmName) > 0
                bookmarkName = bmName
                # If the paragraph also has text content (inline bookmark),
                # preserve it verbatim so bookmark and text stay in one paragraph.
                if wrFindFrom(pXml, "<w:t>", 1) > 0 or wrFindFrom(pXml, "<w:t ", 1) > 0
                    block[:type]    = "rawparagraph"
                    block[:rawXml]  = pXml
                    block[:text]    = ""
                    block[:bookmark] = bookmarkName
                    return block
                ok
            ok
        ok

        # Footnote / endnote reference
        hasNote  = false
        noteType = ""
        noteId   = 0
        fnRefS = wrFindFrom(pXml, "<w:footnoteReference ", 1)
        enRefS = wrFindFrom(pXml, "<w:endnoteReference ", 1)
        if fnRefS > 0
            fnEl = substr(pXml, fnRefS, 80)
            idStr = wrAttr(fnEl, "w:id")
            if len(idStr) > 0 and number(idStr) > 0
                hasNote  = true
                noteType = "footnote"
                noteId   = number(idStr)
            ok
        elseif enRefS > 0
            enEl = substr(pXml, enRefS, 80)
            idStr = wrAttr(enEl, "w:id")
            if len(idStr) > 0 and number(idStr) > 0
                hasNote  = true
                noteType = "endnote"
                noteId   = number(idStr)
            ok
        ok

        # Section break
        sectBreakType     = ""
        sectHdrText       = ""
        sectFtrText       = ""
        sectHdrEvenText   = ""
        sectFtrEvenText   = ""
        sectNumCols       = 1
        sectColSpaceTwips = 720
        sectPrS = wrFindFrom(pPrXml, "<w:sectPr", 1)
        if sectPrS > 0
            sectPrE = wrFindCloseTag(pPrXml, "w:sectPr", sectPrS)
            if sectPrE > 0
                sectPrXml = substr(pPrXml, sectPrS, sectPrE - sectPrS)
                pgszS2 = wrFindFrom(sectPrXml, "<w:pgSz ", 1)
                if pgszS2 > 0
                    pgszEl2 = substr(sectPrXml, pgszS2, 120)
                    ov2 = wrAttr(pgszEl2, "w:orient")
                    if ov2 = "landscape"
                        block[:type]     = "landscapestart"
                        block[:bookmark] = bookmarkName
                        return block
                    ok
                ok
                typeS = wrFindFrom(sectPrXml, "<w:type ", 1)
                if typeS > 0
                    typeEl = substr(sectPrXml, typeS, 80)
                    sectBreakType = wrAttr(typeEl, "w:val")
                else
                    sectBreakType = "nextPage"
                ok
                # Extract per-section column count
                sectNumCols   = 1
                sectColSpaceTwips = 720
                colsS6 = wrFindFrom(sectPrXml, "<w:cols ", 1)
                if colsS6 > 0
                    colsEl6 = substr(sectPrXml, colsS6, 120)
                    ncv6 = wrAttr(colsEl6, "w:num")
                    spv6 = wrAttr(colsEl6, "w:space")
                    if len(ncv6) > 0  sectNumCols = number(ncv6)  ok
                    if len(spv6) > 0  sectColSpaceTwips = number(spv6)  ok
                ok
                # Extract per-section header/footer text (v3k)
                sep5 = wordGetSep()
                sectHdrText     = ""
                sectFtrText     = ""
                sectHdrEvenText = ""
                sectFtrEvenText = ""
                hrPos5 = 1
                while true
                    hrS5 = wrFindFrom(sectPrXml, "<w:headerReference ", hrPos5)
                    if hrS5 = 0  break  ok
                    hrEl5   = substr(sectPrXml, hrS5, 200)
                    hrType5 = wrAttr(hrEl5, "w:type")
                    hrId5   = wrAttr(hrEl5, "r:id")
                    hdrTarget5 = ""
                    for rel5 in aRelationships
                        if rel5[:id] = hrId5  hdrTarget5 = rel5[:target]  ok
                    next
                    if len(hdrTarget5) > 0
                        hdrFP5 = cTempDir + "word" + sep5 + hdrTarget5
                        if fexists(hdrFP5)
                            hdrXml5 = read(hdrFP5)
                            hdrTxt5 = ""
                            htP5    = 1
                            while true
                                tS5a = wrFindFrom(hdrXml5, "<w:t>", htP5)
                                tS5b = wrFindFrom(hdrXml5, "<w:t ", htP5)
                                tS5  = wrMinPos(tS5a, tS5b)
                                if tS5 = 0  break  ok
                                tTg5 = wrFindFrom(hdrXml5, ">", tS5)
                                tEd5 = wrFindFrom(hdrXml5, "</w:t>", tTg5)
                                if tTg5 > 0 and tEd5 > 0
                                    hdrTxt5 += wrXmlUnescape(substr(hdrXml5, tTg5+1, tEd5-tTg5-1))
                                    htP5 = tEd5 + 6
                                else
                                    break
                                ok
                            end
                            if hrType5 = "default"
                                sectHdrText = hdrTxt5
                            ok
                            if hrType5 = "even"
                                sectHdrEvenText = hdrTxt5
                            ok
                        ok
                    ok
                    hrPos5 = hrS5 + 1
                end
                frPos5 = 1
                while true
                    frS5 = wrFindFrom(sectPrXml, "<w:footerReference ", frPos5)
                    if frS5 = 0  break  ok
                    frEl5   = substr(sectPrXml, frS5, 200)
                    frType5 = wrAttr(frEl5, "w:type")
                    frId5   = wrAttr(frEl5, "r:id")
                    ftrTarget5 = ""
                    for rel5 in aRelationships
                        if rel5[:id] = frId5  ftrTarget5 = rel5[:target]  ok
                    next
                    if len(ftrTarget5) > 0
                        ftrFP5 = cTempDir + "word" + sep5 + ftrTarget5
                        if fexists(ftrFP5)
                            ftrXml5 = read(ftrFP5)
                            ftrTxt5 = ""
                            ftP5    = 1
                            while true
                                ftS5a = wrFindFrom(ftrXml5, "<w:t>", ftP5)
                                ftS5b = wrFindFrom(ftrXml5, "<w:t ", ftP5)
                                ftS5  = wrMinPos(ftS5a, ftS5b)
                                if ftS5 = 0  break  ok
                                ftTg5 = wrFindFrom(ftrXml5, ">", ftS5)
                                ftEd5 = wrFindFrom(ftrXml5, "</w:t>", ftTg5)
                                if ftTg5 > 0 and ftEd5 > 0
                                    ftrTxt5 += wrXmlUnescape(substr(ftrXml5, ftTg5+1, ftEd5-ftTg5-1))
                                    ftP5 = ftEd5 + 6
                                else
                                    break
                                ok
                            end
                            if frType5 = "default"
                                sectFtrText = ftrTxt5
                            ok
                            if frType5 = "even"
                                sectFtrEvenText = ftrTxt5
                            ok
                        ok
                    ok
                    frPos5 = frS5 + 1
                end
            ok
        ok

        # Column break
        if wrFindFrom(pXml, "<w:br ", 1) > 0
            brPos = 1
            while true
                brS2 = wrFindFrom(pXml, "<w:br ", brPos)
                if brS2 = 0  break  ok
                brEl2 = substr(pXml, brS2, 80)
                brType2 = wrAttr(brEl2, "w:type")
                if brType2 = "column"
                    block[:type]     = "columnbreak"
                    block[:bookmark] = bookmarkName
                    return block
                ok
                brPos = brS2 + 1
            end
        ok

        # Page break paragraph
        if wrFindFrom(pXml, "<w:lastRenderedPageBreak/>", 1) = 0
            if wrFindFrom(pXml, "<w:br ", 1) > 0
                brPos3 = 1
                while true
                    brS3 = wrFindFrom(pXml, "<w:br ", brPos3)
                    if brS3 = 0  break  ok
                    brEl3 = substr(pXml, brS3, 80)
                    brType3 = wrAttr(brEl3, "w:type")
                    if brType3 = "page"
                        block[:type]     = "pagebreak"
                        block[:bookmark] = bookmarkName
                        return block
                    ok
                    brPos3 = brS3 + 1
                end
            ok
        ok

        # Hyperlink
        hyperlinkUrl = ""
        hlS = wrFindFrom(pXml, "<w:hyperlink ", 1)
        if hlS > 0
            hlEl  = substr(pXml, hlS, 120)
            hlRid = wrAttr(hlEl, "r:id")
            if len(hlRid) > 0
                for rel in aRelationships
                    if rel[:id] = hlRid
                        hyperlinkUrl = rel[:target]
                    ok
                next
            ok
        ok

        # Drawing / Image / Chart
        isImage     = false
        isChart     = false
        isShape     = false
        shapeData   = []
        chartRelId2 = ""
        chartWidthCm  = 0.0
        chartHeightCm = 0.0
        isFloating  = false
        imageRelId  = ""
        imgAltText  = ""
        imgWidthCm  = 0.0
        imgHeightCm = 0.0
        imgWrapType = "wrapSquare"
        imgWrapSide = "bothSides"
        imgDistT    = 0.0
        imgDistB    = 0.0
        imgDistL    = 0.0
        imgDistR    = 0.0
        imgPosX     = 0.0
        imgPosY     = 0.0
        imgRunBold = false; imgRunItalic = false; imgRunUnder = false
        imgRunColor = ""; imgRunSize = 0
        drawS = wrFindFrom(pXml, "<w:drawing>", 1)
        if drawS = 0  drawS = wrFindFrom(pXml, "<w:drawing ", 1)  ok
        if drawS > 0
            isImage = true
            # Extract run-level formatting from the <w:r> that contains the drawing
            imgRunStart = 0
            imgRs1 = wrFindFrom(pXml, "<w:r>", 1)
            imgRs2 = wrFindFrom(pXml, "<w:r ", 1)
            imgRs  = wrMinPos(imgRs1, imgRs2)
            while imgRs > 0 and imgRs < drawS
                imgRunStart = imgRs
                imgRs = wrMinPos(wrFindFrom(pXml, "<w:r>", imgRs+1), wrFindFrom(pXml, "<w:r ", imgRs+1))
            end
            if imgRunStart > 0
                imgRunEnd = wrFindCloseTag(pXml, "w:r", imgRunStart)
                if imgRunEnd > 0
                    imgRXml = substr(pXml, imgRunStart, imgRunEnd - imgRunStart)
                    imgRPrS = wrFindFrom(imgRXml, "<w:rPr>", 1)
                    if imgRPrS = 0  imgRPrS = wrFindFrom(imgRXml, "<w:rPr ", 1)  ok
                    if imgRPrS > 0
                        imgRPrE = wrFindCloseTag(imgRXml, "w:rPr", imgRPrS)
                        if imgRPrE > 0
                            imgRPrXml = substr(imgRXml, imgRPrS, imgRPrE - imgRPrS)
                            if wrFindFrom(imgRPrXml, "<w:b/>", 1) > 0 or wrFindFrom(imgRPrXml, "<w:b ", 1) > 0
                                imgRunBold = true
                            ok
                            if wrFindFrom(imgRPrXml, "<w:i/>", 1) > 0 or wrFindFrom(imgRPrXml, "<w:i ", 1) > 0
                                imgRunItalic = true
                            ok
                            if wrFindFrom(imgRPrXml, "<w:u ", 1) > 0
                                imgRunUnder = true
                            ok
                            imgColS = wrFindFrom(imgRPrXml, "<w:color ", 1)
                            if imgColS > 0
                                imgColEl = substr(imgRPrXml, imgColS, 80)
                                imgColV  = wrAttr(imgColEl, "w:val")
                                if len(imgColV) > 0 and imgColV != "auto"  imgRunColor = imgColV  ok
                            ok
                            imgSzS = wrFindFrom(imgRPrXml, "<w:sz ", 1)
                            if imgSzS > 0
                                imgSzEl = substr(imgRPrXml, imgSzS, 80)
                                imgSzV  = wrAttr(imgSzEl, "w:val")
                                if len(imgSzV) > 0  imgRunSize = number(imgSzV)  ok
                            ok
                        ok
                    ok
                ok
            ok
            # Detect anchor vs inline
            anchorS = wrFindFrom(pXml, "<wp:anchor ", drawS)
            if anchorS > 0
                isFloating = true
                anchorE = wrFindCloseTag(pXml, "wp:anchor", anchorS)
                if anchorE = 0  anchorE = len(pXml)  ok
                anchorXml = substr(pXml, anchorS, anchorE - anchorS)
                # Size from wp:extent
                extS2 = wrFindFrom(anchorXml, "<wp:extent ", 1)
                if extS2 > 0
                    extEl2 = substr(anchorXml, extS2, 100)
                    cxv2 = wrAttr(extEl2, "cx")
                    cyv2 = wrAttr(extEl2, "cy")
                    if len(cxv2) > 0  imgWidthCm  = number(cxv2) / 360000.0  ok
                    if len(cyv2) > 0  imgHeightCm = number(cyv2) / 360000.0  ok
                ok
                # Position
                phS = wrFindFrom(anchorXml, "<wp:positionH ", 1)
                if phS > 0
                    poS2 = wrFindFrom(anchorXml, "<wp:posOffset>", phS)
                    if poS2 > 0
                        poE2 = wrFindFrom(anchorXml, "</wp:posOffset>", poS2)
                        if poE2 > 0
                            posStr = substr(anchorXml, poS2+14, poE2-poS2-14)
                            if isNumber(number(posStr))  imgPosX = number(posStr) / 360000.0  ok
                        ok
                    ok
                ok
                pvS = wrFindFrom(anchorXml, "<wp:positionV ", 1)
                if pvS > 0
                    poS3 = wrFindFrom(anchorXml, "<wp:posOffset>", pvS)
                    if poS3 > 0
                        poE3 = wrFindFrom(anchorXml, "</wp:posOffset>", poS3)
                        if poE3 > 0
                            posStr2 = substr(anchorXml, poS3+14, poE3-poS3-14)
                            if isNumber(number(posStr2))  imgPosY = number(posStr2) / 360000.0  ok
                        ok
                    ok
                ok
                # Wrap type and side
                for wrapTag in ["wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough",
                                 "wp:wrapTopAndBottom", "wp:wrapNone"]
                    wS3 = wrFindFrom(anchorXml, "<" + wrapTag, 1)
                    if wS3 > 0
                        imgWrapType = wrapTag
                        wEl3 = substr(anchorXml, wS3, 120)
                        ws3v = wrAttr(wEl3, "wrapText")
                        if len(ws3v) > 0  imgWrapSide = ws3v  ok
                        break
                    ok
                next
                # Distance from text (on the anchor element itself)
                anchorEl = substr(anchorXml, 1, 200)
                dT = wrAttr(anchorEl, "distT")
                dB = wrAttr(anchorEl, "distB")
                dL = wrAttr(anchorEl, "distL")
                dR = wrAttr(anchorEl, "distR")
                if len(dT) > 0  imgDistT = number(dT) / 360000.0  ok
                if len(dB) > 0  imgDistB = number(dB) / 360000.0  ok
                if len(dL) > 0  imgDistL = number(dL) / 360000.0  ok
                if len(dR) > 0  imgDistR = number(dR) / 360000.0  ok
            else
                # Inline image - get size from wp:extent
                extS = wrFindFrom(pXml, "<wp:extent ", drawS)
                if extS > 0
                    extEl = substr(pXml, extS, 100)
                    cxv = wrAttr(extEl, "cx")
                    cyv = wrAttr(extEl, "cy")
                    if len(cxv) > 0  imgWidthCm  = number(cxv) / 360000.0  ok
                    if len(cyv) > 0  imgHeightCm = number(cyv) / 360000.0  ok
                ok
            ok
            blipS = wrFindFrom(pXml, "<a:blip ", drawS)
            if blipS > 0
                blipEl = substr(pXml, blipS, 120)
                imageRelId = wrAttr(blipEl, "r:embed")
            ok
            # F2 (v3n): parse a:srcRect crop
            imgCropL = 0.0; imgCropR = 0.0; imgCropT = 0.0; imgCropB = 0.0
            srcRectS = wrFindFrom(pXml, "<a:srcRect ", drawS)
            if srcRectS > 0
                srcREl = substr(pXml, srcRectS, 200)
                scL = wrAttr(srcREl, "l"); scR = wrAttr(srcREl, "r")
                scT = wrAttr(srcREl, "t"); scB = wrAttr(srcREl, "b")
                if len(scL) > 0  imgCropL = number(scL) / 1000.0  ok
                if len(scR) > 0  imgCropR = number(scR) / 1000.0  ok
                if len(scT) > 0  imgCropT = number(scT) / 1000.0  ok
                if len(scB) > 0  imgCropB = number(scB) / 1000.0  ok
            ok
            # Alt text: read descr attribute from wp:docPr
            imgAltText = ""
            docPrS = wrFindFrom(pXml, "<wp:docPr ", drawS)
            if docPrS > 0
                docPrEl = substr(pXml, docPrS, 300)
                imgAltText = wrAttr(docPrEl, "descr")
                if imgAltText = NULL  imgAltText = ""  ok
                # Fall back to title if no descr
                if len(imgAltText) = 0
                    imgAltText = wrAttr(docPrEl, "title")
                    if imgAltText = NULL  imgAltText = ""  ok
                ok
            ok
            # Chart detection: <c:chart r:id="..."/> inside drawing
            cchS = wrFindFrom(pXml, "<c:chart ", drawS)
            if cchS > 0
                cchEl = substr(pXml, cchS, 400)  # wide enough to include xmlns+r:id
                chartRelId2 = wrAttr(cchEl, "r:id")
                if len(chartRelId2) > 0
                    isChart  = true
                    isImage  = false   # not an image
                    # Size from extent (already extracted into imgWidthCm/imgHeightCm)
                    chartWidthCm  = imgWidthCm
                    chartHeightCm = imgHeightCm
                ok
            ok
        ok

        # DrawingML shape detection (mc:AlternateContent with wps:wsp)
        # Must run even when isImage=true, because DrawingML shapes
        # use <w:drawing> which the image detector picks up first.
        if !isChart
            if (wrFindFrom(pXml, "<mc:AlternateContent", 1) > 0) and (wrFindFrom(pXml, "wps:wsp", 1) > 0)
                shapeData = wrParseDrawingShape(pXml)
                if shapeData[:found] = true
                    isShape = true
                    isImage = false   # shape takes priority over image
                ok
            ok
        ok

        # Comment reference
        commentId     = -1
        commentText   = ""
        commentAuthor = ""
        cmRS = wrFindFrom(pXml, "<w:commentRangeStart ", 1)
        if cmRS = 0  cmRS = wrFindFrom(pXml, "<w:commentReference ", 1)  ok
        if cmRS > 0
            cmEl  = substr(pXml, cmRS, 80)
            cmIdS = wrAttr(cmEl, "w:id")
            if len(cmIdS) > 0
                commentId = number(cmIdS)
                for cm in aComments
                    if cm[:id] = commentId
                        commentText   = cm[:text]
                        commentAuthor = cm[:author]
                    ok
                next
            ok
        ok

        # Tab stops
        aTabStops    = []
        aTabSegments = []
        if len(pPrXml) > 0
            tabsS = wrFindFrom(pPrXml, "<w:tabs>", 1)
            if tabsS = 0  tabsS = wrFindFrom(pPrXml, "<w:tabs ", 1)  ok
            if tabsS > 0
                tabsE = wrFindFrom(pPrXml, "</w:tabs>", tabsS)
                if tabsE > 0
                    tabsXml = substr(pPrXml, tabsS, tabsE - tabsS)
                    tPos2 = 1
                    while true
                        tabS2 = wrFindFrom(tabsXml, "<w:tab ", tPos2)
                        if tabS2 = 0  break  ok
                        tabEl2 = substr(tabsXml, tabS2, 120)
                        tabPos2   = wrAttr(tabEl2, "w:pos")
                        tabAlign2 = wrAttr(tabEl2, "w:val")
                        tabLead2  = wrAttr(tabEl2, "w:leader")
                        if len(tabAlign2) = 0  tabAlign2 = "left"  ok
                        if len(tabLead2)  = 0  tabLead2  = "none"  ok
                        if len(tabPos2) > 0
                            aTabStops + [:pos = number(tabPos2), :align = tabAlign2, :leader = tabLead2]
                        ok
                        tPos2 = tabS2 + 1
                    end
                ok
            ok
        ok

        # --- Tracked changes: strip <w:del>, unwrap <w:ins> ---
        pXml = wrStripTrackedChanges(pXml)

        # --- Runs ---
        aRuns    = []
        fullText = ""
        runPos   = 1

        while true
            rS1 = wrFindFrom(pXml, "<w:r>", runPos)
            rS2 = wrFindFrom(pXml, "<w:r ", runPos)
            rS  = wrMinPos(rS1, rS2)
            if rS = 0  break  ok
            rEnd = wrFindCloseTag(pXml, "w:r", rS)
            if rEnd = 0  break  ok
            rXml = substr(pXml, rS, rEnd - rS)

            tText = ""
            tS1 = wrFindFrom(rXml, "<w:t>", 1)
            tS2 = wrFindFrom(rXml, "<w:t ", 1)
            tS  = wrMinPos(tS1, tS2)
            if tS > 0
                tClose = wrFindFrom(rXml, ">", tS)
                if tClose > 0
                    tE = wrFindFrom(rXml, "</w:t>", tClose)
                    if tE > 0
                        tText = wrXmlUnescape(substr(rXml, tClose + 1, tE - tClose - 1))
                    ok
                ok
            ok

            # Tab character - preserve position relative to text
            tabPos2 = wrFindFrom(rXml, "<w:tab/>", 1)
            if tabPos2 = 0  tabPos2 = wrFindFrom(rXml, "<w:tab />", 1)  ok
            if tabPos2 > 0
                tS1tmp = wrFindFrom(rXml, "<w:t>", 1)
                tS2tmp = wrFindFrom(rXml, "<w:t ", 1)
                tStmp  = wrMinPos(tS1tmp, tS2tmp)
                if tStmp > 0 and tabPos2 < tStmp
                    tText = char(9) + tText   # tab before text
                else
                    tText = tText + char(9)   # tab after text
                ok
            ok

            # Run properties
            rBold = false; rItalic = false; rUnder = false; rStrike = false
            rDStrike = false; rCaps = false; rSmallCaps = false; rVanish = false
            rColor = ""; rSize = 0; rFont = ""; rHighlight = ""
            rSuper = false; rSub = false; rLang = ""; rStyleName = ""
            rBdrStyle = ""; rBdrColor = ""; rBdrSize = 0
            rRTL = false
            rPrS = wrFindFrom(rXml, "<w:rPr>", 1)
            if rPrS = 0  rPrS = wrFindFrom(rXml, "<w:rPr ", 1)  ok
            if rPrS > 0
                rPrE = wrFindCloseTag(rXml, "w:rPr", rPrS)
                if rPrE > 0
                    rPrXml = substr(rXml, rPrS, rPrE - rPrS)
                    if wrFindFrom(rPrXml, "<w:b/>",  1) > 0  rBold   = true  ok
                    if wrFindFrom(rPrXml, "<w:b ",   1) > 0  rBold   = true  ok
                    if wrFindFrom(rPrXml, "<w:i/>",  1) > 0  rItalic = true  ok
                    if wrFindFrom(rPrXml, "<w:i ",   1) > 0  rItalic = true  ok
                    if wrFindFrom(rPrXml, "<w:u ",   1) > 0  rUnder  = true  ok
                    if wrFindFrom(rPrXml, "<w:strike",  1) > 0  rStrike   = true  ok
                    if wrFindFrom(rPrXml, "<w:dstrike", 1) > 0  rDStrike  = true  ok
                    if wrFindFrom(rPrXml, "<w:caps",    1) > 0  rCaps     = true  ok
                    if wrFindFrom(rPrXml, "<w:smallCaps",1)> 0  rSmallCaps= true  ok
                    if wrFindFrom(rPrXml, "<w:vanish",  1) > 0  rVanish   = true  ok
                    if wrFindFrom(rPrXml, "<w:rtl/>",  1) > 0 or wrFindFrom(rPrXml, "<w:rtl ", 1) > 0  rRTL = true  ok
                    colS2 = wrFindFrom(rPrXml, "<w:color ", 1)
                    if colS2 > 0
                        colEl2 = substr(rPrXml, colS2, 80)
                        cv2 = wrAttr(colEl2, "w:val")
                        if len(cv2) > 0 and cv2 != "auto"  rColor = cv2  ok
                    ok
                    szS2 = wrFindFrom(rPrXml, "<w:sz ", 1)
                    if szS2 > 0
                        szEl2 = substr(rPrXml, szS2, 60)
                        sv2 = wrAttr(szEl2, "w:val")
                        if len(sv2) > 0  rSize = floor(number(sv2) / 2)  ok
                    ok
                    fnS2 = wrFindFrom(rPrXml, "<w:rFonts ", 1)
                    if fnS2 > 0
                        fnEl2 = substr(rPrXml, fnS2, 120)
                        fv2 = wrAttr(fnEl2, "w:ascii")
                        if len(fv2) > 0  rFont = fv2  ok
                    ok
                    hlS2 = wrFindFrom(rPrXml, "<w:highlight ", 1)
                    if hlS2 > 0
                        hlEl2 = substr(rPrXml, hlS2, 60)
                        rHighlight = wrAttr(hlEl2, "w:val")
                    ok
                    vaS2 = wrFindFrom(rPrXml, "<w:vertAlign ", 1)
                    if vaS2 > 0
                        vaEl2 = substr(rPrXml, vaS2, 80)
                        vav2 = wrAttr(vaEl2, "w:val")
                        if vav2 = "superscript"  rSuper = true  ok
                        if vav2 = "subscript"    rSub   = true  ok
                    ok
                    # Language tag
                    rLang = ""
                    lgS2 = wrFindFrom(rPrXml, "<w:lang ", 1)
                    if lgS2 > 0
                        lgEl2 = substr(rPrXml, lgS2, 80)
                        rLang = wrAttr(lgEl2, "w:val")
                    ok
                    # Run border (w:bdr)
                    rBdrStyle = ""; rBdrColor = ""; rBdrSize = 0
                    bdrS2 = wrFindFrom(rPrXml, "<w:bdr ", 1)
                    if bdrS2 > 0
                        bdrEl2 = substr(rPrXml, bdrS2, 120)
                        bdrSV = wrAttr(bdrEl2, "w:val")
                        bdrCV = wrAttr(bdrEl2, "w:color")
                        bdrSZ = wrAttr(bdrEl2, "w:sz")
                        if len(bdrSV) > 0 and bdrSV != "none"
                            rBdrStyle = bdrSV
                            if len(bdrCV) > 0 and bdrCV != "auto"  rBdrColor = bdrCV  ok
                            if len(bdrSZ) > 0  rBdrSize = number(bdrSZ)  ok
                        ok
                    ok
                    # Character style (rStyle) - resolve to formatting
                    rStyleName = ""
                    rsS = wrFindFrom(rPrXml, "<w:rStyle ", 1)
                    if rsS > 0
                        rsEl = substr(rPrXml, rsS, 80)
                        rStyleName = wrAttr(rsEl, "w:val")
                        # Resolve style to formatting (unless it is a Word built-in ref style)
                        builtins = ["FootnoteReference","EndnoteReference","CommentReference","Hyperlink","PlaceholderText"]
                        isBuiltin = false
                        for bi in builtins
                            if rStyleName = bi  isBuiltin = true  ok
                        next
                        if !isBuiltin and len(rStyleName) > 0
                            cs = wrResolveCharStyle(rStyleName, aCharStyles)
                            if isList(cs) and len(cs) > 0
                                if cs[:bold]      = true  rBold     = true  ok
                                if cs[:italic]    = true  rItalic   = true  ok
                                if cs[:underline] = true  rUnder    = true  ok
                                if cs[:strike]    = true  rStrike   = true  ok
                                if cs[:dstrike]   = true  rDStrike  = true  ok
                                if cs[:caps]      = true  rCaps     = true  ok
                                if cs[:smallCaps] = true  rSmallCaps= true  ok
                                if cs[:vanish]    = true  rVanish   = true  ok
                                if len(cs[:color]) > 0 and len(rColor) = 0  rColor = cs[:color]  ok
                                if cs[:size] > 0   and rSize  = 0  rSize  = cs[:size]   ok
                                if len(cs[:font]) > 0 and len(rFont) = 0   rFont  = cs[:font]   ok
                            ok
                        ok
                    ok
                ok
            ok

            if len(tText) > 0
                run = [:text = tText, :bold = rBold, :italic = rItalic]
                run[:underline]    = rUnder
                run[:strike]       = rStrike
                run[:dstrike]      = rDStrike
                run[:caps]         = rCaps
                run[:smallCaps]    = rSmallCaps
                run[:vanish]       = rVanish
                run[:color]        = rColor
                run[:size]         = rSize
                run[:font]         = rFont
                run[:highlight]    = rHighlight
                run[:superscript]  = rSuper
                run[:subscript]    = rSub
                run[:lang]         = rLang
                run[:styleName]    = rStyleName
                run[:borderStyle]  = rBdrStyle
                run[:borderColor]  = rBdrColor
                run[:borderSize]   = rBdrSize
                run[:rtl]          = rRTL
                aRuns + run
                # Vanished runs are hidden — exclude from plain text
                if !rVanish
                    fullText += tText
                ok
            ok
            runPos = rEnd
        end

        # Tab segments - use wrFindFrom since Ring 1.27 lacks split()
        if len(aTabStops) > 0
            tabSep = char(9)
            tabPos = 1
            tabLen = len(fullText)
            while tabPos <= tabLen + 1  # +1 to catch trailing tab
                tabP2 = wrFindFrom(fullText, tabSep, tabPos)
                if tabP2 = 0
                    # No more tabs - add remaining text (may be empty)
                    seg5 = ""
                    if tabPos <= tabLen
                        seg5 = substr(fullText, tabPos, tabLen - tabPos + 1)
                    ok
                    # Only add empty segment if it follows a tab (trailing tab case)
                    if len(seg5) > 0 or len(aTabSegments) > 0
                        aTabSegments + seg5
                    ok
                    exit
                ok
                aTabSegments + substr(fullText, tabPos, tabP2 - tabPos)
                tabPos = tabP2 + 1
            end
        ok

        # --- Inline SDT (content control) detection ---
        # Check for inline <w:sdt> inside this paragraph
        if (wrFindFrom(pXml, "<w:sdt>", 1) > 0 or
            wrFindFrom(pXml, "<w:sdt ", 1) > 0)
            inSdtS = wrFindFrom(pXml, "<w:sdt>", 1)
            if inSdtS = 0  inSdtS = wrFindFrom(pXml, "<w:sdt ", 1)  ok
            if inSdtS > 0
                inSdtE = wrFindCloseTag(pXml, "w:sdt", inSdtS)
                if inSdtE > 0
                    inSdtXml = substr(pXml, inSdtS, inSdtE - inSdtS)
                    inSdtInfo = wrParseSdtBlock(inSdtXml)
                    inSdtType = inSdtInfo[:type]
                    if inSdtType = "checkbox" or inSdtType = "dropdown"
                       or inSdtType = "text"
                        # For checkbox: label is text OUTSIDE the SDT in this para
                        inSdtLabel = inSdtInfo[:label]
                        if inSdtType = "checkbox"
                            # Collect text from runs outside the SDT
                            outsideText = ""
                            # Text before SDT
                            preXml = substr(pXml, 1, inSdtS - 1)
                            tPrePos = 1
                            while true
                                tpS1 = wrFindFrom(preXml, "<w:t>", tPrePos)
                                tpS2 = wrFindFrom(preXml, "<w:t ", tPrePos)
                                tpS  = wrMinPos(tpS1, tpS2)
                                if tpS = 0  break  ok
                                tpGt = wrFindFrom(preXml, ">", tpS)
                                tpE  = wrFindFrom(preXml, "</w:t>", tpGt)
                                if tpGt > 0 and tpE > 0
                                    outsideText += wrXmlUnescape(substr(preXml, tpGt+1, tpE-tpGt-1))
                                ok
                                tPrePos = tpE + 1
                            end
                            # Text after SDT
                            postXml = substr(pXml, inSdtE)
                            tPostPos = 1
                            while true
                                tpS1 = wrFindFrom(postXml, "<w:t>", tPostPos)
                                tpS2 = wrFindFrom(postXml, "<w:t ", tPostPos)
                                tpS  = wrMinPos(tpS1, tpS2)
                                if tpS = 0  break  ok
                                tpGt = wrFindFrom(postXml, ">", tpS)
                                tpE  = wrFindFrom(postXml, "</w:t>", tpGt)
                                if tpGt > 0 and tpE > 0
                                    outsideText += wrXmlUnescape(substr(postXml, tpGt+1, tpE-tpGt-1))
                                ok
                                tPostPos = tpE + 1
                            end
                            outsideText = trim(outsideText)
                            if len(outsideText) > 0  inSdtLabel = outsideText  ok
                        ok
                        block[:type]        = "formfield"
                        block[:fieldType]   = inSdtType
                        block[:label]       = inSdtLabel
                        block[:checked]     = inSdtInfo[:checked]
                        block[:choices]     = inSdtInfo[:choices]
                        block[:default]     = inSdtInfo[:default]
                        block[:placeholder] = inSdtInfo[:placeholder]
                        block[:bookmark]    = bookmarkName
                        return block
                    ok
                ok
            ok
        ok

        # --- Block assembly ---
        if len(sectBreakType) > 0
            block[:type]               = "sectionbreak"
            block[:breakType]          = sectBreakType
            block[:bookmark]           = bookmarkName
            block[:sectHeader]         = sectHdrText
            block[:sectFooter]         = sectFtrText
            block[:sectEvenHeader]     = sectHdrEvenText
            block[:sectEvenFooter]     = sectFtrEvenText
            block[:numColumns]         = sectNumCols
            block[:columnSpaceTwips]   = sectColSpaceTwips
            return block
        ok

        # Quick DrawingML shape check to avoid falsely returning "empty"
        # before full shape detection runs below
        hasMcShape = (wrFindFrom(pXml, "<mc:AlternateContent", 1) > 0) and (wrFindFrom(pXml, "wps:wsp", 1) > 0)
        # Detect VML textbox paragraphs (<w:pict> with <w:txbxContent> or <v:shapetype>)
        # These must be preserved verbatim as floating elements
        hasPictTxbx = (wrFindFrom(pXml, "<w:pict>", 1) > 0 or wrFindFrom(pXml, "<w:pict ", 1) > 0) and
                      (wrFindFrom(pXml, "<w:txbxContent>", 1) > 0 or
                       wrFindFrom(pXml, "<v:shapetype ", 1) > 0 or
                       wrFindFrom(pXml, "<v:shape ", 1) > 0 or
                       wrFindFrom(pXml, "<v:rect ", 1) > 0)
        if hasPictTxbx
            block[:type]   = "rawparagraph"
            block[:rawXml] = pXml
            block[:text]   = ""
            block[:bookmark] = bookmarkName
            return block
        ok
        if len(fullText) = 0 and !isImage and !isChart and !isHorizLine and !isListItem and
           !hasNote and len(sectBreakType) = 0 and !hasMcShape
            block[:type] = "empty"
            return block
        ok

        if isShape
            block[:type]           = "shape"
            block[:shapeType]      = shapeData[:shapeType]
            block[:widthCm]        = shapeData[:widthCm]
            block[:heightCm]       = shapeData[:heightCm]
            block[:fillColor]      = shapeData[:fillColor]
            block[:noFill]         = shapeData[:noFill]
            block[:lineColor]      = shapeData[:lineColor]
            block[:lineWidthPt]    = shapeData[:lineWidthPt]
            block[:noBorder]       = shapeData[:noBorder]
            block[:text]           = shapeData[:text]
            block[:textColor]      = shapeData[:textColor]
            block[:textBold]       = shapeData[:textBold]
            block[:textSize]       = shapeData[:textSize]
            block[:align]          = shapeData[:align]
            block[:bookmark]       = bookmarkName
            return block
        ok

        if isChart
            chartData2 = wrParseChartData(chartRelId2, aRelationships, cTempDir)
            block[:type]            = "chart"
            block[:chartType]       = chartData2[:type]
            block[:chartScatterStyle] = chartData2[:scatterStyle]
            block[:chartBubble3D]     = chartData2[:bubble3D]
            block[:chartTitle]      = chartData2[:title]
            block[:chartData]       = chartData2[:series]
            block[:chartGrouping]   = chartData2[:grouping]
            block[:chartBarDir]     = chartData2[:barDir]
            block[:chartSmooth]     = chartData2[:smooth]
            block[:chartLegendPos]  = chartData2[:legendPos]
            block[:chartShowLabels] = chartData2[:showDataLabels]
            block[:chartSerColors]  = chartData2[:serColors]
            block[:widthCm]         = chartWidthCm
            block[:heightCm]        = chartHeightCm
            block[:bookmark]        = bookmarkName
            return block
        ok

        if isImage
            imgPath = resolveRelPath(imageRelId)
            block[:type]     = "image"
            block[:path]     = imgPath
            block[:widthCm]  = imgWidthCm
            block[:heightCm] = imgHeightCm
            block[:relId]    = imageRelId
            block[:altText]  = imgAltText
            block[:bookmark] = bookmarkName
            block[:align]      = alignVal
            block[:runBold]    = imgRunBold
            block[:runItalic]  = imgRunItalic
            block[:runUnder]   = imgRunUnder
            block[:runColor]   = imgRunColor
            block[:runSize]    = imgRunSize
            block[:floating]   = isFloating
            block[:wrapType] = imgWrapType
            block[:wrapSide] = imgWrapSide
            block[:distT]    = imgDistT
            block[:distB]    = imgDistB
            block[:distL]    = imgDistL
            block[:distR]    = imgDistR
            block[:x]        = imgPosX
            block[:y]        = imgPosY
            block[:cropL]    = imgCropL
            block[:cropR]    = imgCropR
            block[:cropT]    = imgCropT
            block[:cropB]    = imgCropB
            return block
        ok

        if isHorizLine
            block[:type]     = "hline"
            block[:bookmark] = bookmarkName
            return block
        ok

        if headingLevel > 0
            block[:type]     = "heading"
            block[:level]    = headingLevel
            block[:text]     = fullText
            block[:runs]     = aRuns
            # Numbered heading: check if numId present in pPr
            headingIsNumbered = false
            headingNumId      = 1
            headingIlvl       = 0
            if len(pPrXml) > 0
                hnumPrS = wrFindFrom(pPrXml, "<w:numPr>", 1)
                if hnumPrS = 0  hnumPrS = wrFindFrom(pPrXml, "<w:numPr ", 1)  ok
                if hnumPrS > 0
                    hnumPrE = wrFindCloseTag(pPrXml, "w:numPr", hnumPrS)
                    if hnumPrE > 0
                        hnumPrXml = substr(pPrXml, hnumPrS, hnumPrE - hnumPrS)
                        hnidS = wrFindFrom(hnumPrXml, "<w:numId ", 1)
                        if hnidS > 0
                            hnidEl = substr(hnumPrXml, hnidS, 60)
                            hnidV  = wrAttr(hnidEl, "w:val")
                            if len(hnidV) > 0 and hnidV != "0"
                                headingIsNumbered = true
                                headingNumId = number(hnidV)
                                # Extract ilvl
                                hilvlS = wrFindFrom(hnumPrXml, "<w:ilvl ", 1)
                                if hilvlS > 0
                                    hilvlEl = substr(hnumPrXml, hilvlS, 60)
                                    hilvlV  = wrAttr(hilvlEl, "w:val")
                                    if len(hilvlV) > 0
                                        headingIlvl = number(hilvlV)
                                    ok
                                ok
                            ok
                        ok
                    ok
                ok
            ok
            block[:numbered] = headingIsNumbered
            block[:numId]    = headingNumId
            block[:ilvl]     = headingIlvl
            block[:bookmark] = bookmarkName
            return block
        ok

        # Detect writer-generated captions by inline formatting
        # (italic + small font <=9pt + no bullet)
        if !isCaption and !isListItem and !isBlockQuote and len(fullText) > 0
            if isList(aRuns) and len(aRuns) > 0
                allItalicSmall = true
                for capRun in aRuns
                    if capRun[:italic] != true  allItalicSmall = false  ok
                    if capRun[:size] > 18 or capRun[:size] = 0  allItalicSmall = false  ok
                next
                if allItalicSmall  isCaption = true  ok
            ok
        ok

        if isCaption
            block[:type]     = "caption"
            block[:text]     = fullText
            block[:runs]     = aRuns
            block[:bookmark] = bookmarkName
            return block
        ok

        if len(aTabStops) > 0
            block[:type]      = "tabbed"
            block[:text]      = fullText
            block[:runs]      = aRuns
            block[:tabStops]  = aTabStops
            block[:segments]  = aTabSegments
            block[:align]     = alignVal
            block[:bookmark]  = bookmarkName
            return block
        ok

        if isListItem
            block[:type]         = "listitem"
            block[:text]         = fullText
            block[:runs]         = aRuns
            block[:listNumId]    = listNumId
            block[:listIlvl]     = listIlvl
            block[:listIsBullet] = listIsBullet
            block[:listStartAt]  = listStartAt
            block[:align]        = alignVal
            block[:bookmark]     = bookmarkName
            return block
        ok

        if len(hyperlinkUrl) > 0
            block[:type]     = "hyperlink"
            block[:text]     = fullText
            block[:url]      = hyperlinkUrl
            block[:runs]     = aRuns
            block[:bookmark] = bookmarkName
            return block
        ok

        if isBlockQuote
            block[:type]       = "blockquote"
            block[:text]       = fullText
            block[:runs]       = aRuns
            block[:bgColor]    = paraBgColor
            block[:bookmark]   = bookmarkName
            # Store border + formatting data for faithful round-trip
            block[:hasBorder]  = paraHasBorder
            block[:borderStyle] = paraBdrStyle
            block[:borderColor] = paraBdrColor
            block[:borderSize]  = paraBdrSize
            block[:borderSpace] = paraBdrSpace
            block[:borderSides] = paraBdrSides
            block[:align]       = alignVal
            block[:indentLeft]  = indentLeft
            block[:indentRight] = indentRight
            block[:spaceBefore] = spaceBefore
            block[:spaceAfter]  = spaceAfter
            return block
        ok

        if hasNote
            block[:type]     = noteType
            block[:text]     = fullText
            block[:noteText] = ""
            block[:noteId]   = noteId
            block[:runs]     = aRuns
            block[:bookmark] = bookmarkName
            return block
        ok

        # Standard paragraph (or RTL paragraph)
        pType = "paragraph"
        if isRTL  pType = "rtlparagraph"  ok
        block[:type]            = pType
        block[:text]            = fullText
        block[:runs]            = aRuns
        block[:style]           = styleName
        # TOCHeading paragraphs: mark as tocheading so toWriter can skip
        # (addTOC() will add its own heading with the correct title)
        if styleName = "TOCHeading"
            # Preserve TOCHeading paragraphs verbatim (List of Figures,
            # List of Tables, etc.). The main TOC title is handled by addTOC().
            block[:type]   = "rawparagraph"
            block[:rawXml] = pXml
        ok
        block[:align]           = alignVal
        block[:spaceBefore]     = spaceBefore
        block[:spaceAfter]      = spaceAfter
        block[:lineSpacing]     = lineSpacing
        block[:lineSpacingPt]   = lineSpacingPt
        block[:lineRule]        = lineRule
        block[:bgColor]         = paraBgColor
        block[:indent]          = indentLeft
        block[:indentRight]     = indentRight
        block[:indentFirstLine] = indentFirstLine
        block[:keepNext]        = keepNext
        block[:keepLines]       = keepLines
        block[:pageBreakBefore] = pageBreakBefore
        block[:outlineLvl]      = outlineLvl
        block[:contextualSpacing] = contextualSpacing
        block[:widowControl]    = widowControl
        block[:noHyphenate]     = noHyphenate
        block[:bookmark]        = bookmarkName
        if commentId >= 0
            block[:commentId]     = commentId
            block[:commentText]   = commentText
            block[:commentAuthor] = commentAuthor
        ok
        block[:hasBorder]       = paraHasBorder
        block[:borderStyle]     = paraBdrStyle
        block[:borderColor]     = paraBdrColor
        block[:borderSize]      = paraBdrSize
        block[:borderSpace]     = paraBdrSpace
        block[:borderSides]     = paraBdrSides
        return block

    func parseTable tblXml
        block = []
        block[:type] = "table"

        # Column widths from tblGrid
        aColWidths = []
        gridS = wrFindFrom(tblXml, "<w:tblGrid>", 1)
        if gridS = 0  gridS = wrFindFrom(tblXml, "<w:tblGrid ", 1)  ok
        gridE = wrFindFrom(tblXml, "</w:tblGrid>", 1)
        if gridS > 0 and gridE > 0
            gridXml = substr(tblXml, gridS, gridE - gridS + 12)
            gPos = 1
            while true
                gcS = wrFindFrom(gridXml, "<w:gridCol ", gPos)
                if gcS = 0  break  ok
                gcEl = substr(gridXml, gcS, 100)
                wKey = "w:w=" + char(34)
                wPos = wrFindFrom(gcEl, wKey, 1)
                if wPos > 0
                    wValStart = wPos + len(wKey)
                    wValEnd   = wrFindFrom(gcEl, char(34), wValStart)
                    if wValEnd > 0
                        wv = substr(gcEl, wValStart, wValEnd - wValStart)
                        if len(wv) > 0
                            aColWidths + (number(wv) / 567.0)
                        ok
                    ok
                ok
                gPos = gcS + 1
            end
        ok

        # Table border style
        tblBorderStyle = "single"
        tblBorderColor = "auto"
        tblBordersMap  = []
        tblBgFill      = ""
        tblStyleName   = ""
        tblWidthTwips  = 0
        tblWidthType   = "auto"
        tblAlignStr    = ""
        tblIndentTwips = 0
        # vars (defaults; overwritten inside tblPr block)
        tblWidthTwips  = 0
        tblWidthType   = "auto"
        tblAlignStr    = ""
        tblIndentTwips = 0
        tblPrS = wrFindFrom(tblXml, "<w:tblPr>", 1)
        if tblPrS = 0  tblPrS = wrFindFrom(tblXml, "<w:tblPr ", 1)  ok
        if tblPrS > 0
            tblPrE = wrFindCloseTag(tblXml, "w:tblPr", tblPrS)
            if tblPrE > 0
                tblPrXml = substr(tblXml, tblPrS, tblPrE - tblPrS)
                # Table style name
                tsS = wrFindFrom(tblPrXml, "<w:tblStyle ", 1)
                if tsS > 0
                    tsEl = substr(tblPrXml, tsS, 80)
                    tblStyleName = wrAttr(tsEl, "w:val")
                ok
                # Table width (w:tblW)
                tblWidthTwips = 0
                tblWidthType  = "auto"
                twS = wrFindFrom(tblPrXml, "<w:tblW ", 1)
                if twS > 0
                    twEl = substr(tblPrXml, twS, 80)
                    twW  = wrAttr(twEl, "w:w")
                    twT  = wrAttr(twEl, "w:type")
                    if len(twT) > 0  tblWidthType = twT  ok
                    if len(twW) > 0 and twT != "auto"
                        tblWidthTwips = number(twW)
                    ok
                ok
                # Table alignment (w:jc inside tblPr)
                tblAlignStr = ""
                jcS2 = wrFindFrom(tblPrXml, "<w:jc ", 1)
                if jcS2 > 0
                    jcEl2 = substr(tblPrXml, jcS2, 60)
                    tblAlignStr = wrAttr(jcEl2, "w:val")
                ok
                # Table indent (w:tblInd)
                tblIndentTwips = 0
                tiS = wrFindFrom(tblPrXml, "<w:tblInd ", 1)
                if tiS > 0
                    tiEl = substr(tblPrXml, tiS, 80)
                    tiW  = wrAttr(tiEl, "w:w")
                    if len(tiW) > 0  tblIndentTwips = number(tiW)  ok
                ok
                # Table width, alignment, indent (v3k)
                tblWidthTwips  = 0
                tblWidthType   = "auto"
                tblAlignStr    = ""
                tblIndentTwips = 0
                twS1 = wrFindFrom(tblPrXml, "<w:tblW ", 1)
                if twS1 > 0
                    twEl1 = substr(tblPrXml, twS1, 100)
                    twW1  = wrAttr(twEl1, "w:w")
                    twT1  = wrAttr(twEl1, "w:type")
                    if twT1 != "auto" and twT1 != "nil" and len(twW1) > 0
                        tblWidthTwips = number(twW1)
                        if len(twT1) > 0  tblWidthType = twT1  ok
                    ok
                ok
                tjcS1 = wrFindFrom(tblPrXml, "<w:jc ", 1)
                if tjcS1 > 0
                    tjcEl1 = substr(tblPrXml, tjcS1, 60)
                    tblAlignStr = wrAttr(tjcEl1, "w:val")
                ok
                tiS1 = wrFindFrom(tblPrXml, "<w:tblInd ", 1)
                if tiS1 > 0
                    tiEl1 = substr(tblPrXml, tiS1, 100)
                    tiW1  = wrAttr(tiEl1, "w:w")
                    if len(tiW1) > 0  tblIndentTwips = number(tiW1)  ok
                ok
                tblBdrS = wrFindFrom(tblPrXml, "<w:tblBorders>", 1)
                if tblBdrS > 0
                    tblBdrE = wrFindCloseTag(tblPrXml, "w:tblBorders", tblBdrS)
                    if tblBdrE > 0
                        tblBdrXml = substr(tblPrXml, tblBdrS, tblBdrE - tblBdrS)
                        # Parse each border side independently
                        aSides = ["top", "left", "bottom", "right", "insideH", "insideV"]
                        tblBordersMap = []
                        for bdrSideName in aSides
                            bdrSideS = wrFindFrom(tblBdrXml, "<w:" + bdrSideName + " ", 1)
                            if bdrSideS > 0
                                bdrSideEl = substr(tblBdrXml, bdrSideS, 150)
                                bdrSideVal  = wrAttr(bdrSideEl, "w:val")
                                bdrSideSz   = wrAttr(bdrSideEl, "w:sz")
                                bdrSideColor= wrAttr(bdrSideEl, "w:color")
                                tblBordersMap + [:side=bdrSideName, :style=bdrSideVal, :sz=bdrSideSz, :color=bdrSideColor]
                            ok
                        next
                        # Set legacy single-style fields from top (or insideH fallback)
                        for bdrEntry in tblBordersMap
                            if bdrEntry[:side] = "top"
                                if len(bdrEntry[:style]) > 0  tblBorderStyle = bdrEntry[:style]  ok
                                if len(bdrEntry[:color]) > 0 and bdrEntry[:color] != "auto"  tblBorderColor = bdrEntry[:color]  ok
                            ok
                        next
                        if tblBorderStyle = "single" and len(tblBordersMap) > 0
                            for bdrEntry in tblBordersMap
                                if bdrEntry[:side] = "insideH" and len(bdrEntry[:style]) > 0
                                    tblBorderStyle = bdrEntry[:style]
                                ok
                            next
                        ok
                    ok
                ok
                # Table-level background shading (tblPr > w:shd)
                tblShdS = wrFindFrom(tblPrXml, "<w:shd ", 1)
                if tblShdS > 0
                    tblShdEl = substr(tblPrXml, tblShdS, 150)
                    tblShdFillV = wrAttr(tblShdEl, "w:fill")
                    if len(tblShdFillV) > 0 and tblShdFillV != "auto"
                        tblBgFill = tblShdFillV
                    ok
                ok
                # tblLook flags (banding row/col control)
                tblLookFirstRow = -1
                tblLookLastRow  = -1
                tblLookFirstCol = -1
                tblLookLastCol  = -1
                tblLookNoHBand  = -1
                tblLookNoVBand  = -1
                tlS1 = wrFindFrom(tblPrXml, "<w:tblLook ", 1)
                if tlS1 > 0
                    tlEl1 = substr(tblPrXml, tlS1, 300)
                    tlFR = wrAttr(tlEl1, "w:firstRow")
                    tlLR = wrAttr(tlEl1, "w:lastRow")
                    tlFC = wrAttr(tlEl1, "w:firstColumn")
                    tlLC = wrAttr(tlEl1, "w:lastColumn")
                    tlNH = wrAttr(tlEl1, "w:noHBand")
                    tlNV = wrAttr(tlEl1, "w:noVBand")
                    if len(tlFR) > 0  tblLookFirstRow = number(tlFR)  ok
                    if len(tlLR) > 0  tblLookLastRow  = number(tlLR)  ok
                    if len(tlFC) > 0  tblLookFirstCol = number(tlFC)  ok
                    if len(tlLC) > 0  tblLookLastCol  = number(tlLC)  ok
                    if len(tlNH) > 0  tblLookNoHBand  = number(tlNH)  ok
                    if len(tlNV) > 0  tblLookNoVBand  = number(tlNV)  ok
                ok
            ok
        ok

        # Rows
        aRows       = []
        aRowHeights = []   # per-row height in twips (0=auto)
        aRowHRules  = []   # per-row hRule: "atLeast","exact","auto"
        aRowCantSplit = []  # per-row cantSplit boolean
        aRowIsHeader  = []  # per-row tblHeader boolean
        aRowBgColors  = []  # per-row background color (empty string = none)
        trPos = 1
        while true
            trS1 = wrFindFrom(tblXml, "<w:tr>", trPos)
            trS2 = wrFindFrom(tblXml, "<w:tr ", trPos)
            trS  = wrMinPos(trS1, trS2)
            if trS = 0  break  ok
            trE = wrFindCloseTag(tblXml, "w:tr", trS)
            if trE = 0  break  ok
            trXml = substr(tblXml, trS, trE - trS)
            row   = parseTableRow(trXml)
            aRows + row
            rowProps = wrParseRowProps(trXml)
            aRowHeights  + rowProps[:height]
            aRowHRules   + rowProps[:hRule]
            aRowCantSplit + rowProps[:cantSplit]
            aRowIsHeader  + rowProps[:isHeaderRow]
            aRowBgColors  + rowProps[:rowBgColor]
            trPos = trE
        end

        # Detect header bg from first row
        headerBg = ""
        if len(aRows) > 0
            firstRow = aRows[1]
            if len(firstRow) > 0
                fc = firstRow[1]
                if isList(fc)
                    hbg = fc[:bgColor]
                    if hbg = NULL  hbg = ""  ok
                    if len(hbg) > 0  headerBg = hbg  ok
                ok
            ok
        ok

        # Detect even-row bg from row 3
        evenRowBg = ""
        if len(aRows) >= 3
            secRow = aRows[3]
            if len(secRow) > 0
                sc = secRow[1]
                if isList(sc)
                    erbg = sc[:bgColor]
                    if erbg = NULL  erbg = ""  ok
                    if len(erbg) > 0 and erbg != headerBg  evenRowBg = erbg  ok
                ok
            ok
        ok

        block[:rows]         = aRows
        block[:colWidths]    = aColWidths
        block[:tblStyle]       = tblStyleName
        block[:tblWidthTwips]  = tblWidthTwips
        block[:tblWidthType]   = tblWidthType
        block[:tblAlign]       = tblAlignStr
        block[:tblIndentTwips] = tblIndentTwips
        block[:hasHeader]    = true
        block[:headerBg]     = headerBg
        block[:evenRowBg]    = evenRowBg
        block[:borderStyle]  = tblBorderStyle
        block[:borderColor]  = tblBorderColor
        block[:bordersMap]   = tblBordersMap
        block[:tblBgFill]    = tblBgFill
        block[:rowHeights]   = aRowHeights
        block[:rowHRules]    = aRowHRules
        block[:rowCantSplit] = aRowCantSplit
        block[:rowIsHeader]  = aRowIsHeader
        block[:rowBgColors]  = aRowBgColors
        # tblLook banding flags
        block[:tblLookFirstRow] = tblLookFirstRow
        block[:tblLookLastRow]  = tblLookLastRow
        block[:tblLookFirstCol] = tblLookFirstCol
        block[:tblLookLastCol]  = tblLookLastCol
        block[:tblLookNoHBand]  = tblLookNoHBand
        block[:tblLookNoVBand]  = tblLookNoVBand
        return block

    func parseTableRow trXml
        aRow = []
        tcPos = 1
        while true
            tcS1 = wrFindFrom(trXml, "<w:tc>", tcPos)
            tcS2 = wrFindFrom(trXml, "<w:tc ", tcPos)
            tcS  = wrMinPos(tcS1, tcS2)
            if tcS = 0  break  ok
            tcEnd = wrFindCloseTag(trXml, "w:tc", tcS)
            if tcEnd = 0  break  ok
            tcXml = substr(trXml, tcS, tcEnd - tcS)

            cellBg     = ""
            spanVal    = 1
            rowSpanVal = 1
            cellVAlign = ""
            isMergeContinue = false

            tcPrS = wrFindFrom(tcXml, "<w:tcPr>", 1)
            if tcPrS = 0  tcPrS = wrFindFrom(tcXml, "<w:tcPr ", 1)  ok
            if tcPrS > 0
                tcPrE = wrFindCloseTag(tcXml, "w:tcPr", tcPrS)
                if tcPrE > 0
                    tcPrXml = substr(tcXml, tcPrS, tcPrE - tcPrS)
                    # Shading
                    shdS = wrFindFrom(tcPrXml, "<w:shd ", 1)
                    if shdS > 0
                        shdEl = substr(tcPrXml, shdS, 100)
                        fillV = wrAttr(shdEl, "w:fill")
                        if len(fillV) > 0 and fillV != "auto"  cellBg = fillV  ok
                    ok
                    # Colspan
                    gsS = wrFindFrom(tcPrXml, "<w:gridSpan ", 1)
                    if gsS > 0
                        gsEl = substr(tcPrXml, gsS, 60)
                        gsV  = wrAttr(gsEl, "w:val")
                        if len(gsV) > 0  spanVal = number(gsV)  ok
                    ok
                    # Vertical align
                    vaS = wrFindFrom(tcPrXml, "<w:vAlign ", 1)
                    if vaS > 0
                        vaEl = substr(tcPrXml, vaS, 60)
                        cellVAlign = wrAttr(vaEl, "w:val")
                    ok
                    # Vertical merge
                    vmS = wrFindFrom(tcPrXml, "<w:vMerge", 1)
                    if vmS > 0
                        vmEl = substr(tcPrXml, vmS, 80)
                        restartQ1 = "w:val=" + char(34) + "restart" + char(34)
                        restartQ2 = "w:val=" + char(39) + "restart" + char(39)
                        if wrFindFrom(vmEl, restartQ1, 1) = 0 and
                           wrFindFrom(vmEl, restartQ2, 1) = 0
                            isMergeContinue = true
                        else
                            rowSpanVal = 2
                        ok
                    ok
                ok
            ok

            # Per-cell borders from <w:tcBorders>
            cellBdrStyle = ""
            cellBdrColor = ""
            cellBdrSides = []
            cellTextDir  = ""
            if tcPrS > 0
                tcPrE2 = wrFindCloseTag(tcXml, "w:tcPr", tcPrS)
                if tcPrE2 > 0
                    tcPrXml2 = substr(tcXml, tcPrS, tcPrE2 - tcPrS)
                    bdrInfo = wrParseTcBorders(tcPrXml2)
                    if bdrInfo[:found] = true
                        cellBdrStyle = bdrInfo[:style]
                        cellBdrColor = bdrInfo[:color]
                        cellBdrSides = bdrInfo[:sides]
                    ok
                    # text direction
                    tdS = wrFindFrom(tcPrXml2, "<w:textDirection ", 1)
                    if tdS > 0
                        tdEl = substr(tcPrXml2, tdS, 80)
                        cellTextDir = wrAttr(tdEl, "w:val")
                    ok
                ok
            ok

            # Cell padding from <w:tcMar>
            cellPadTop    = -1
            cellPadBottom = -1
            cellPadLeft   = -1
            cellPadRight  = -1
            if tcPrS > 0
                tcmS = wrFindFrom(tcXml, "<w:tcMar>", tcPrS)
                if tcmS = 0  tcmS = wrFindFrom(tcXml, "<w:tcMar ", tcPrS)  ok
                if tcmS > 0
                    tcmE = wrFindCloseTag(tcXml, "w:tcMar", tcmS)
                    if tcmE > 0
                        tcmXml = substr(tcXml, tcmS, tcmE - tcmS)
                        for side in ["top","bottom","left","right"]
                            sS = wrFindFrom(tcmXml, "<w:" + side + " ", 1)
                            if sS > 0
                                sEl = substr(tcmXml, sS, 80)
                                sv = wrAttr(sEl, "w:w")
                                if len(sv) > 0
                                    if side = "top"     cellPadTop    = number(sv)  ok
                                    if side = "bottom"  cellPadBottom = number(sv)  ok
                                    if side = "left"    cellPadLeft   = number(sv)  ok
                                    if side = "right"   cellPadRight  = number(sv)  ok
                                ok
                            ok
                        next
                    ok
                ok
            ok

            if isMergeContinue
                mcBlk = [:text="", :bgColor="", :colspan=1, :rowspan=0,
                         :align="", :vAlign="", :runs=[], :isMergeContinue=true,
                         :hlUrl="", :hlText="", :imgPath="", :imgW=0.0, :imgH=0.0]
                aRow + mcBlk
                tcPos = tcEnd
                loop
            ok

            # Collect text and runs from all paragraphs in cell
            cellRuns    = []
            cellText    = ""
            firstPAlign = ""
            cpPos = 1
            while true
                cpS1 = wrFindFrom(tcXml, "<w:p>", cpPos)
                cpS2 = wrFindFrom(tcXml, "<w:p ", cpPos)
                cpS  = wrMinPos(cpS1, cpS2)
                if cpS = 0  break  ok
                # Handle self-closing <w:p .../> in table cell
                if wrIsSelfClosingTag(tcXml, cpS)
                    cpPos = wrSelfClosingEnd(tcXml, cpS)
                    loop
                ok
                # Skip paragraphs inside nested tables
                preCtx2 = substr(tcXml, 1, cpS - 1)
                tblO2 = 0  tblC2 = 0  scanP2 = 1
                while true
                    to2p = wrFindFrom(preCtx2, "<w:tbl>", scanP2)
                    if to2p = 0  break  ok
                    tblO2++  scanP2 = to2p + 1
                end
                scanP2 = 1
                while true
                    tc2p = wrFindFrom(preCtx2, "</w:tbl>", scanP2)
                    if tc2p = 0  break  ok
                    tblC2++  scanP2 = tc2p + 1
                end
                if tblO2 > tblC2
                    nxtTblE2 = wrFindFrom(tcXml, "</w:tbl>", cpS)
                    if nxtTblE2 > 0
                        cpPos = nxtTblE2 + 8
                    else
                        cpPos = cpS + 1
                    ok
                    loop
                ok
                cpE = wrFindCloseTag(tcXml, "w:p", cpS)
                if cpE = 0  break  ok
                cpXml = substr(tcXml, cpS, cpE - cpS)

                # Alignment + pPr rPr highlight (paragraph mark formatting)
                cellPprHighlight = ""
                # Parse pPr rPr highlight every cell paragraph (not just first)
                cpPrHlS = wrFindFrom(cpXml, "<w:pPr>", 1)
                if cpPrHlS = 0  cpPrHlS = wrFindFrom(cpXml, "<w:pPr ", 1)  ok
                if cpPrHlS > 0
                    cpPrHlE = wrFindCloseTag(cpXml, "w:pPr", cpPrHlS)
                    if cpPrHlE > 0
                        cpPrXmlH = substr(cpXml, cpPrHlS, cpPrHlE - cpPrHlS)
                        cpRprHlS = wrFindFrom(cpPrXmlH, "<w:rPr>", 1)
                        if cpRprHlS > 0
                            cpRprHlE = wrFindCloseTag(cpPrXmlH, "w:rPr", cpRprHlS)
                            if cpRprHlE > 0
                                cpRprHlXml = substr(cpPrXmlH, cpRprHlS, cpRprHlE - cpRprHlS)
                                cpHlTagS = wrFindFrom(cpRprHlXml, "<w:highlight ", 1)
                                if cpHlTagS > 0
                                    cpHlTagEl = substr(cpRprHlXml, cpHlTagS, 80)
                                    cellPprHighlight = wrAttr(cpHlTagEl, "w:val")
                                ok
                            ok
                        ok
                    ok
                ok
                if len(firstPAlign) = 0
                    pPrS2 = wrFindFrom(cpXml, "<w:pPr>", 1)
                    if pPrS2 = 0  pPrS2 = wrFindFrom(cpXml, "<w:pPr ", 1)  ok
                    if pPrS2 > 0
                        pPrE2 = wrFindCloseTag(cpXml, "w:pPr", pPrS2)
                        if pPrE2 > 0
                            pPrXml2 = substr(cpXml, pPrS2, pPrE2 - pPrS2)
                            jcS2 = wrFindFrom(pPrXml2, "<w:jc ", 1)
                            if jcS2 > 0
                                jcEl2 = substr(pPrXml2, jcS2, 80)
                                firstPAlign = wrAttr(jcEl2, "w:val")
                            ok
                        ok
                    ok
                ok

                # Separator between paragraphs
                if len(cellText) > 0
                    cellText += " "
                    if len(cellRuns) > 0
                        crSep = [:text=" ",:bold=false,:italic=false,:underline=false,
                                 :strike=false,:color="",:size=0,:font="",:highlight=""]
                        cellRuns + crSep
                    ok
                ok

                # Runs in this paragraph
                rPos2 = 1
                while true
                    rS1b = wrFindFrom(cpXml, "<w:r>", rPos2)
                    rS2b = wrFindFrom(cpXml, "<w:r ", rPos2)
                    rSb  = wrMinPos(rS1b, rS2b)
                    if rSb = 0  break  ok
                    rEb  = wrFindCloseTag(cpXml, "w:r", rSb)
                    if rEb = 0  break  ok
                    rXmlb = substr(cpXml, rSb, rEb - rSb)

                    tTextb = ""
                    tS1b = wrFindFrom(rXmlb, "<w:t>", 1)
                    tS2b = wrFindFrom(rXmlb, "<w:t ", 1)
                    tSb  = wrMinPos(tS1b, tS2b)
                    if tSb > 0
                        tClb = wrFindFrom(rXmlb, ">", tSb)
                        if tClb > 0
                            tEb = wrFindFrom(rXmlb, "</w:t>", tClb)
                            if tEb > 0
                                tTextb = wrXmlUnescape(substr(rXmlb, tClb + 1, tEb - tClb - 1))
                            ok
                        ok
                    ok

                    rBoldb=false; rItalicb=false; rUnderb=false; rStrikeb=false
                    rColorb=""; rSizeb=0; rFontb=""; rHighlightb=""
                    rRTLb = false
                    rPrSb = wrFindFrom(rXmlb, "<w:rPr>", 1)
                    if rPrSb = 0  rPrSb = wrFindFrom(rXmlb, "<w:rPr ", 1)  ok
                    if rPrSb > 0
                        rPrEb = wrFindCloseTag(rXmlb, "w:rPr", rPrSb)
                        if rPrEb > 0
                            rPrXmlb = substr(rXmlb, rPrSb, rPrEb - rPrSb)
                            if wrFindFrom(rPrXmlb, "<w:b/>",  1) > 0  rBoldb = true    ok
                            if wrFindFrom(rPrXmlb, "<w:b ",   1) > 0  rBoldb = true    ok
                            if wrFindFrom(rPrXmlb, "<w:i/>",  1) > 0  rItalicb = true  ok
                            if wrFindFrom(rPrXmlb, "<w:i ",   1) > 0  rItalicb = true  ok
                            if wrFindFrom(rPrXmlb, "<w:u ",   1) > 0  rUnderb  = true  ok
                            if wrFindFrom(rPrXmlb, "<w:strike", 1) > 0  rStrikeb = true  ok
                            colSb = wrFindFrom(rPrXmlb, "<w:color ", 1)
                            if colSb > 0
                                colElb = substr(rPrXmlb, colSb, 80)
                                cv = wrAttr(colElb, "w:val")
                                if len(cv) > 0 and cv != "auto"  rColorb = cv  ok
                            ok
                            szSb = wrFindFrom(rPrXmlb, "<w:sz ", 1)
                            if szSb > 0
                                szElb = substr(rPrXmlb, szSb, 60)
                                sv = wrAttr(szElb, "w:val")
                                if len(sv) > 0  rSizeb = floor(number(sv) / 2)  ok
                            ok
                            fnSb = wrFindFrom(rPrXmlb, "<w:rFonts ", 1)
                            if fnSb > 0
                                fnElb = substr(rPrXmlb, fnSb, 120)
                                fv = wrAttr(fnElb, "w:ascii")
                                if len(fv) > 0  rFontb = fv  ok
                            ok
                            hlSb = wrFindFrom(rPrXmlb, "<w:highlight ", 1)
                            if hlSb > 0
                                hlElb = substr(rPrXmlb, hlSb, 60)
                                rHighlightb = wrAttr(hlElb, "w:val")
                            ok
                            # RTL run direction in cell
                            if wrFindFrom(rPrXmlb, "<w:rtl/>", 1) > 0 or wrFindFrom(rPrXmlb, "<w:rtl ", 1) > 0
                                rRTLb = true
                            ok
                        ok
                    ok

                    # Check for line break run (<w:br/>)
                    brS3 = wrFindFrom(rXmlb, "<w:br/>", 1)
                    if brS3 = 0  brS3 = wrFindFrom(rXmlb, "<w:br ", 1)  ok
                    if brS3 > 0 and len(tTextb) = 0
                        # Pure line break run - store as special marker
                        brBlk = [:text="", :bold=false, :italic=false,
                                 :underline=false, :strike=false, :color="",
                                 :size=0, :font="", :highlight="", :rtl=false,
                                 :isLineBreak=true]
                        cellRuns + brBlk
                    ok
                    if len(tTextb) > 0
                        cellText += tTextb
                        effectiveHlb = rHighlightb
                        if len(effectiveHlb) = 0 and len(cellPprHighlight) > 0
                            effectiveHlb = cellPprHighlight
                        ok
                        crBlk = [:text=tTextb,:bold=rBoldb,:italic=rItalicb,
                                 :underline=rUnderb,:strike=rStrikeb,:color=rColorb,
                                 :size=rSizeb,:font=rFontb,:highlight=effectiveHlb,
                                 :rtl=rRTLb]
                        cellRuns + crBlk
                    ok
                    rPos2 = rEb
                end
                cpPos = cpE
            end

            # Parse each cell paragraph as a full block for rich round-trip.
            # Only DIRECT paragraphs (not inside nested <w:tbl>) are included.
            cellParas = []
            cpPos3 = 1
            while true
                cp3S1 = wrFindFrom(tcXml, "<w:p>", cpPos3)
                cp3S2 = wrFindFrom(tcXml, "<w:p ", cpPos3)
                cp3S  = wrMinPos(cp3S1, cp3S2)
                if cp3S = 0  break  ok
                # Count nested table opens/closes before this position
                preCtx3 = substr(tcXml, 1, cp3S - 1)
                tblO3 = 0  tblScan3 = 1
                while true
                    to3 = wrFindFrom(preCtx3, "<w:tbl>", tblScan3)
                    if to3 = 0  break  ok
                    tblO3++  tblScan3 = to3 + 1
                end
                tblC3 = 0  tblScan3 = 1
                while true
                    tc3 = wrFindFrom(preCtx3, "</w:tbl>", tblScan3)
                    if tc3 = 0  break  ok
                    tblC3++  tblScan3 = tc3 + 1
                end
                if tblO3 > tblC3
                    # Inside nested table - skip to after next </w:tbl>
                    nxtTblE3 = wrFindFrom(tcXml, "</w:tbl>", cp3S)
                    if nxtTblE3 > 0  cpPos3 = nxtTblE3 + 8
                    else  cpPos3 = cp3S + 1  ok
                    loop
                ok
                if wrIsSelfClosingTag(tcXml, cp3S)
                    cpPos3 = wrSelfClosingEnd(tcXml, cp3S)
                    loop
                ok
                cp3E = wrFindCloseTag(tcXml, "w:p", cp3S)
                if cp3E = 0  break  ok
                cp3Xml = substr(tcXml, cp3S, cp3E - cp3S)
                cp3Block = parseParagraph(cp3Xml)
                cellParas + cp3Block
                cpPos3 = cp3E
            end

            # Cell hyperlink
            cellHlUrl  = ""
            cellHlText = ""
            hlCS = wrFindFrom(tcXml, "<w:hyperlink ", 1)
            if hlCS > 0
                hlCEl  = substr(tcXml, hlCS, 120)
                hlCRid = wrAttr(hlCEl, "r:id")
                if len(hlCRid) > 0
                    for rel in aRelationships
                        if rel[:id] = hlCRid
                            cellHlUrl  = rel[:target]
                            cellHlText = cellText
                        ok
                    next
                ok
            ok

            # Cell image
            cellImgRelId = ""
            cellImgPath  = ""
            cellImgW     = 3.0
            cellImgH     = 2.0
            drawCS = wrFindFrom(tcXml, "<w:drawing>", 1)
            if drawCS = 0  drawCS = wrFindFrom(tcXml, "<w:drawing ", 1)  ok
            if drawCS > 0
                extCS = wrFindFrom(tcXml, "<wp:extent ", drawCS)
                if extCS > 0
                    extCEl = substr(tcXml, extCS, 100)
                    cxvc = wrAttr(extCEl, "cx")
                    cyvc = wrAttr(extCEl, "cy")
                    if len(cxvc) > 0  cellImgW = number(cxvc) / 360000.0  ok
                    if len(cyvc) > 0  cellImgH = number(cyvc) / 360000.0  ok
                ok
                blipCS = wrFindFrom(tcXml, "<a:blip ", drawCS)
                if blipCS > 0
                    blipCEl  = substr(tcXml, blipCS, 120)
                    cellImgRelId = wrAttr(blipCEl, "r:embed")
                    if len(cellImgRelId) > 0
                        cellImgPath = resolveRelPath(cellImgRelId)
                    ok
                ok
            ok

            # Nested table inside cell
            cellNestedTable = NULL
            ntblS = wrFindFrom(tcXml, "<w:tbl>", 1)
            if ntblS = 0  ntblS = wrFindFrom(tcXml, "<w:tbl ", 1)  ok
            if ntblS > 0
                ntblE = wrFindCloseTag(tcXml, "w:tbl", ntblS)
                if ntblE > 0
                    ntblXml = substr(tcXml, ntblS, ntblE - ntblS)
                    cellNestedTable = parseTable(ntblXml)
                ok
            ok

            cell = [:text          = cellText,
                    :bgColor      = cellBg,
                    :colspan      = spanVal,
                    :rowspan      = rowSpanVal,
                    :align        = firstPAlign,
                    :vAlign       = cellVAlign,
                    :runs         = cellRuns,
                    :hlUrl        = cellHlUrl,
                    :hlText       = cellHlText,
                    :imgPath      = cellImgPath,
                    :imgW         = cellImgW,
                    :imgH         = cellImgH,
                    :padTop       = cellPadTop,
                    :padBottom    = cellPadBottom,
                    :padLeft      = cellPadLeft,
                    :padRight     = cellPadRight,
                    :borderStyle  = cellBdrStyle,
                    :borderColor  = cellBdrColor,
                    :borderSides  = cellBdrSides,
                    :textDir      = cellTextDir,
                    :nestedTable  = cellNestedTable,
                    :isMergeContinue = false,
                    :cellParas = cellParas]
            aRow + cell
            tcPos = tcEnd
        end
        return aRow

    func resolveRelPath relId
        sep = wordGetSep()
        for rel in aRelationships
            if rel[:id] = relId
                target = rel[:target]
                if left(target, 6) = "media/"
                    return cTempDir + "word" + sep + target
                ok
                if left(target, 9) = "../media/"
                    return cTempDir + "word" + sep + substr(target, 4)
                ok
                return cTempDir + "word" + sep + target
            ok
        next
        return ""

    # =========================================================================
    # Summary and listing
    # =========================================================================

    func summary
        nPara      = 0
        nHead      = 0
        nTbl       = 0
        nImg       = 0
        nPB        = 0
        nList      = 0
        nLink      = 0
        nHLine     = 0
        nBQ        = 0
        nFN        = 0
        nEN        = 0
        nSB        = 0
        nChart     = 0
        nTextBox   = 0
        nFloatImg  = 0
        nFormField   = 0
        nBorderedP   = 0
        nNestedTbl   = 0
        nMergeField  = 0
        nNumberedHdg = 0
        nCap       = 0
        nTabbed    = 0
        nLandscape = 0
        nColBreak  = 0
        nRTL          = 0
        nCommented    = 0
        nKeepNext     = 0
        nPBBefore     = 0
        nTrackedKept  = 0
        for block in aBlocks
            bType = block[:type]
            if bType = "paragraph"    nPara++  ok
            if bType = "rtlparagraph" nPara++  nRTL++  ok
            if bType = "heading"      nHead++  ok
            if bType = "table"        nTbl++   ok
            if bType = "image"        nImg++   ok
            if bType = "pagebreak"    nPB++    ok
            if bType = "listitem"     nList++  ok
            if bType = "hyperlink"    nLink++  ok
            if bType = "hline"        nHLine++ ok
            if bType = "blockquote"   nBQ++    ok
            if bType = "footnote"     nFN++    ok
            if bType = "endnote"      nEN++    ok
            if bType = "sectionbreak" nSB++    ok
            if bType = "chart"        nChart++   ok
            if bType = "textbox"      nTextBox++ ok
            if bType = "formfield"    nFormField++ ok
            if bType = "mergefield"   nMergeField++ ok
            if bType = "heading"
                nb = block[:numbered]
                if nb = true  nNumberedHdg++  ok
            ok
            if bType = "paragraph" or bType = "rtlparagraph"
                hb = block[:hasBorder]
                if hb = true  nBorderedP++  ok
            ok
            if bType = "table"
                rows2 = block[:rows]
                if isList(rows2)
                    for row2 in rows2
                        if isList(row2)
                            for cell2 in row2
                                if isList(cell2)
                                    nt2 = cell2[:nestedTable]
                                    if isList(nt2) and len(nt2) > 0  nNestedTbl++  ok
                                ok
                            next
                        ok
                    next
                ok
            ok
            if bType = "image"
                isFloat = block[:floating]
                if isFloat = true  nFloatImg++  ok
            ok
            if bType = "caption"      nCap++   ok
            if bType = "tabbed"       nTabbed++ ok
            if bType = "landscapestart" nLandscape++ ok
            if bType = "columnbreak"  nColBreak++   ok
            if bType = "paragraph"
                cmt = block[:commentText]
                if cmt != NULL and len(cmt) > 0  nCommented++  ok
                if block[:keepNext] = true        nKeepNext++   ok
                if block[:pageBreakBefore] = true nPBBefore++   ok
            ok
            if bType = "rtlparagraph"
                cmt = block[:commentText]
                if cmt != NULL and len(cmt) > 0  nCommented++  ok
                if block[:keepNext] = true        nKeepNext++   ok
                if block[:pageBreakBefore] = true nPBBefore++   ok
            ok
        next
        bgLine = ""
        if len(cSrcPageBgColor) > 0  bgLine = "  PageBg : #" + cSrcPageBgColor  ok
        wmLabel = ""
        if bSrcWatermark = true  wmLabel = " [" + cSrcWatermarkText + "]"  ok
        lines = [
            "=== WordReader Summary ===",
            "  File   : " + cFilePath,
            "  Title  : " + cTitle,
            "  Author : " + cAuthor,
            "  Font   : " + cDefaultFont + " " + string(nDefaultSize) + "pt",
            "  Header : " + cHeaderText,
            "  Footer : " + cFooterText,
            "  Columns: " + string(nSrcColumns),
            "  Blocks : " + string(len(aBlocks)),
            "    Headings     : " + string(nHead),
            "    Paragraphs   : " + string(nPara),
            "      RTL paras  : " + string(nRTL),
            "      Commented  : " + string(nCommented),
            "      KeepNext   : " + string(nKeepNext),
            "      PgBrkBefore: " + string(nPBBefore),
            "    Tables       : " + string(nTbl),
            "    List items   : " + string(nList),
            "    Hyperlinks   : " + string(nLink),
            "    Images       : " + string(nImg),
            "    Charts       : " + string(nChart),
            "    Text boxes   : " + string(nTextBox),
            "    Floating imgs: " + string(nFloatImg),
            "  DocRTL  : " + string(bSrcDocumentRTL),
            "  PageNums: " + string(bSrcPageNumbers) + " loc=" + cSrcPageNumLocation + " align=" + cSrcPageNumAlign,
            "  CharStyles: " + string(len(aCharStyles)),
            "  Form fields  : " + string(nFormField),
            "  Merge fields : " + string(nMergeField) + " blocks, " + string(len(aMergeFields)) + " unique names",
            "  Numbered hdgs : " + string(nNumberedHdg),
            "  Bordered paras: " + string(nBorderedP),
            "  Nested tables : " + string(nNestedTbl),
            "  Has TOC       : " + string(bSrcHasTOC),
            "  Has Watermark : " + string(bSrcWatermark) + wmLabel,
            "    Captions     : " + string(nCap),
            "    Tabbed paras : " + string(nTabbed),
            "    Landscape sec: " + string(nLandscape),
            "    Column breaks: " + string(nColBreak),
            "    Footnotes    : " + string(nFN),
            "    Endnotes     : " + string(nEN),
            "    Section brks : " + string(nSB),
            "    Horiz lines  : " + string(nHLine),
            "    Blockquotes  : " + string(nBQ),
            "    Page breaks  : " + string(nPB),
            "    Comments     : " + string(len(aComments)),
            "=========================="
        ]
        if len(bgLine) > 0  add(lines, bgLine)  ok
        result = ""
        for line in lines
            result += line + char(10)
        next
        return result

    func listBlocks
        bLen = len(aBlocks)
        for i = 1 to bLen
            block = aBlocks[i]
            bType = block[:type]
            desc  = ""
            if bType = "paragraph" or bType = "rtlparagraph"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 55  t = substr(t, 1, 55) + "..."  ok
                prefix = ""
                if bType = "rtlparagraph"  prefix = "[RTL] "  ok
                cmt = block[:commentText]
                suffix = ""
                if cmt != NULL and len(cmt) > 0
                    if len(cmt) > 25  cmt = left(cmt, 25) + "..."  ok
                    suffix = " [comment: " + cmt + "]"
                ok
                hb = block[:hasBorder]
                if hb = true
                    bs = block[:borderStyle]
                    if bs = NULL  bs = "single"  ok
                    desc = prefix + "{border:" + bs + "} " + t + suffix
                else
                    desc = prefix + t + suffix
                ok
            elseif bType = "heading"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 55  t = substr(t, 1, 55) + "..."  ok
                hnum = block[:numbered]
                numMark = ""
                if hnum = true  numMark = "# "  ok
                desc = "[H" + string(block[:level]) + "] " + numMark + t
            elseif bType = "table"
                rows = block[:rows]
                nr = 0
                nc = 0
                if isList(rows)  nr = len(rows)  ok
                if nr > 0 and isList(rows[1])  nc = len(rows[1])  ok
                desc = "[Table " + string(nr) + "r x " + string(nc) + "c]"
            elseif bType = "image"
                p = block[:path]
                if p = NULL  p = ""  ok
                if len(p) > 40  p = "..." + substr(p, len(p)-37)  ok
                prefix2 = "[Image"
                isFloat2 = block[:floating]
                if isFloat2 = true  prefix2 = "[FloatImg"  ok
                wt2 = block[:wrapType]
                if wt2 = NULL  wt2 = ""  ok
                wrapNote = ""
                if isFloat2 = true and len(wt2) > 0  wrapNote = " wrap=" + wt2  ok
                desc = prefix2 + " " + string(block[:widthCm]) + "x" + string(block[:heightCm]) + "cm" + wrapNote + "] " + p
            elseif bType = "hyperlink"
                t = block[:text]
                if t = NULL  t = ""  ok
                u = block[:url]
                if u = NULL  u = ""  ok
                desc = "[Link] " + t + " -> " + u
            elseif bType = "listitem"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 50  t = substr(t, 1, 50) + "..."  ok
                bul = "•"
                if block[:listIsBullet] = false  bul = string(block[:listIlvl]+1) + "."  ok
                desc = "[List " + bul + "] " + t
            elseif bType = "pagebreak"
                desc = "[--- Page Break ---]"
            elseif bType = "hline"
                desc = "[Horizontal Line]"
            elseif bType = "blockquote"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 50  t = substr(t, 1, 50) + "..."  ok
                desc = "[Blockquote] " + t
            elseif bType = "footnote"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 40  t = substr(t, 1, 40) + "..."  ok
                desc = "[Footnote] " + t
            elseif bType = "endnote"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 40  t = substr(t, 1, 40) + "..."  ok
                desc = "[Endnote] " + t
            elseif bType = "sectionbreak"
                sbNC3 = block[:numColumns]
                sbCS3 = block[:columnSpaceTwips]
                if sbNC3 = NULL or !isNumber(sbNC3)  sbNC3 = 1    ok
                if sbCS3 = NULL or !isNumber(sbCS3)  sbCS3 = 720  ok
                desc = "[Section Break: " + block[:breakType]
                if sbNC3 > 1
                    desc += " cols=" + string(sbNC3) + " gap=" + string(sbCS3/567.0) + "cm"
                ok
                desc += "]"
            elseif bType = "chart"
                chartT2 = block[:chartTitle]
                if chartT2 = NULL  chartT2 = ""  ok
                chartS2 = block[:chartData]
                nSer2 = 0
                if isList(chartS2)  nSer2 = len(chartS2)  ok
                desc = "[Chart: " + block[:chartType] + " " + string(block[:widthCm]) + "x" + string(block[:heightCm]) + "cm"
                if len(chartT2) > 0  desc += " title=" + chartT2  ok
                desc += " series=" + string(nSer2) + "]"
            elseif bType = "textbox"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 45  t = left(t, 45) + "..."  ok
                tbw = block[:width]; tbh = block[:height]
                if !isNumber(tbw)  tbw = 0  ok
                if !isNumber(tbh)  tbh = 0  ok
                desc = "[TextBox " + string(tbw) + "x" + string(tbh) + "cm] " + t
            elseif bType = "mergefield"
                flds = block[:fields]
                if !isList(flds)  flds = []  ok
                fStr = ""
                for mf in flds  fStr += "<<" + mf + ">> "  next
                if len(fStr) > 0  fStr = substr(fStr, 1, len(fStr)-1)  ok
                desc = "[MergeField] " + fStr
            elseif bType = "formfield"
                ft = block[:fieldType]
                if ft = NULL  ft = "unknown"  ok
                lbl = block[:label]
                if lbl = NULL  lbl = ""  ok
                if len(lbl) > 40  lbl = left(lbl, 40) + "..."  ok
                desc = "[FormField:" + ft + "] " + lbl
            elseif bType = "caption"
                t = block[:text]
                if t = NULL  t = ""  ok
                desc = "[Caption] " + t
            elseif bType = "tabbed"
                t = block[:text]
                if t = NULL  t = ""  ok
                if len(t) > 50  t = substr(t, 1, 50) + "..."  ok
                nts = 0
                if isList(block[:tabStops])  nts = len(block[:tabStops])  ok
                desc = "[Tabbed tabStops=" + string(nts) + "] " + t
            elseif bType = "landscapestart"
                desc = "[--- Landscape Start ---]"
            elseif bType = "landscapeend"
                desc = "[--- Landscape End ---]"
            elseif bType = "columnbreak"
                desc = "[Column Break]"
            else
                desc = "[" + bType + "]"
            ok
            ? string(i) + ". " + desc
        next

    # =========================================================================
    # Query methods
    # =========================================================================

    func getBlocks             return aBlocks          ok
    func getFilePath           return cFilePath         ok
    func getTitle              return cTitle            ok
    func getAuthor             return cAuthor           ok
    func getDefaultFont        return cDefaultFont      ok
    func getDefaultFontSize    return nDefaultSize      ok
    func getPageWidthCm        return nPageWidth / 567.0   ok
    func getPageHeightCm       return nPageHeight / 567.0  ok
    func getMarginTopCm        return nMarginTop / 567.0   ok
    func getMarginBottomCm     return nMarginBottom / 567.0  ok
    func getMarginLeftCm       return nMarginLeft / 567.0   ok
    func getMarginRightCm      return nMarginRight / 567.0  ok
    func getOrientation        return cSrcOrientation   ok
    func getHeaderText         return cHeaderText       ok
    func getFooterText         return cFooterText       ok
    func getEvenPageHeaderText return cEvenPageHeaderText   ok
    func getEvenPageFooterText return cEvenPageFooterText   ok
    func getFirstPageHeaderText return cFirstPageHeaderText ok
    func getFirstPageFooterText return cFirstPageFooterText ok
    func hasEvenAndOddHeaders  return bSrcEvenAndOddHeaders  ok
    func hasFirstPageDifferent return bSrcFirstPageDifferent ok
    func hasPageBorder         return bSrcPageBorder    ok
    func getPageBorderStyle    return cSrcPageBorderStyle ok
    func getPageBorderColor    return cSrcPageBorderColor ok
    func getPageBorderSizePt   return nSrcPageBorderSize / 8.0  ok
    func getPageBgColor        return cSrcPageBgColor   ok
    func hasPageBackground     return len(cSrcPageBgColor) > 0  ok
    func getColumnCount        return nSrcColumns       ok
    func getColumnSpaceCm      return nSrcColumnSpace / 567.0  ok
    func getComments           return aComments         ok
    func getFootnotes
        result = []
        for block in aBlocks
            if block[:type] = "footnote"  result + block  ok
        next
        return result
    func getEndnotes
        result = []
        for block in aBlocks
            if block[:type] = "endnote"  result + block  ok
        next
        return result
    func getHeadings
        result = []
        for block in aBlocks
            if block[:type] = "heading"  result + block  ok
        next
        return result
    func getParagraphs
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"  result + block  ok
        next
        return result
    func getRTLParagraphs
        result = []
        for block in aBlocks
            if block[:type] = "rtlparagraph"  result + block  ok
        next
        return result
    func getCommentedParagraphs
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                cmt = block[:commentText]
                if cmt != NULL and len(cmt) > 0
                    result + block
                ok
            ok
        next
        return result

    func getKeepNextParagraphs
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                if block[:keepNext] = true  result + block  ok
            ok
        next
        return result

    func getPageBreakBeforeParagraphs
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                if block[:pageBreakBefore] = true  result + block  ok
            ok
        next
        return result

    func getCellsWithPadding
        # Returns all cells (from all tables) that have custom padding
        result = []
        for block in aBlocks
            if block[:type] = "table"
                rows = block[:rows]
                if isList(rows)
                    for row in rows
                        if isList(row)
                            for cell in row
                                if isList(cell)
                                    hasPad = false
                                    pt = cell[:padTop]
                                    pb = cell[:padBottom]
                                    pl = cell[:padLeft]
                                    pr = cell[:padRight]
                                    if isNumber(pt) and pt >= 0  hasPad = true  ok
                                    if isNumber(pb) and pb >= 0  hasPad = true  ok
                                    if isNumber(pl) and pl >= 0  hasPad = true  ok
                                    if isNumber(pr) and pr >= 0  hasPad = true  ok
                                    if hasPad  result + cell  ok
                                ok
                            next
                        ok
                    next
                ok
            ok
        next
        return result

    func getTextBoxes
        result = []
        for block in aBlocks
            if block[:type] = "textbox"  result + block  ok
        next
        return result

    func getFloatingImages
        result = []
        for block in aBlocks
            if block[:type] = "image"
                isFloat = block[:floating]
                if isFloat = true  result + block  ok
            ok
        next
        return result

    func getCharStyles
        return aCharStyles

    func isDocumentRTL
        return bSrcDocumentRTL

    func hasPageNumbers
        return bSrcPageNumbers

    func getPageNumberAlign
        return cSrcPageNumAlign

    func getPageNumberLocation
        return cSrcPageNumLocation

    func getCharStylesForRun runBlock
        # Given a run block, return the resolved character style name if any
        sn = runBlock[:styleName]
        if sn = NULL or len(sn) = 0  return []  ok
        return wrResolveCharStyle(sn, aCharStyles)

    func getFormFields
        # Returns all form field blocks (checkbox, dropdown, text)
        result = []
        for block in aBlocks
            if block[:type] = "formfield"  result + block  ok
        next
        return result

    func getCheckboxes
        result = []
        for block in aBlocks
            if block[:type] = "formfield"
                if block[:fieldType] = "checkbox"  result + block  ok
            ok
        next
        return result

    func getDropdowns
        result = []
        for block in aBlocks
            if block[:type] = "formfield"
                if block[:fieldType] = "dropdown"  result + block  ok
            ok
        next
        return result

    func getTextInputs
        result = []
        for block in aBlocks
            if block[:type] = "formfield"
                if block[:fieldType] = "text"  result + block  ok
            ok
        next
        return result

    func getBorderedParagraphs
        # Returns all paragraphs/rtlparagraphs that have a visible border
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                hb = block[:hasBorder]
                if hb = true  result + block  ok
            ok
        next
        return result

    func getCellsWithBorders
        # Returns all table cells that have per-cell border style set
        result = []
        for block in aBlocks
            if block[:type] = "table"
                rows = block[:rows]
                if isList(rows)
                    for row in rows
                        if isList(row)
                            for cell in row
                                if isList(cell)
                                    bs = cell[:borderStyle]
                                    if bs != NULL and len(bs) > 0  result + cell  ok
                                ok
                            next
                        ok
                    next
                ok
            ok
        next
        return result

    func getNestedTables
        # Returns all nested table objects found inside table cells
        result = []
        for block in aBlocks
            if block[:type] = "table"
                rows = block[:rows]
                if isList(rows)
                    for row in rows
                        if isList(row)
                            for cell in row
                                if isList(cell)
                                    nt = cell[:nestedTable]
                                    if isList(nt) and len(nt) > 0
                                        result + nt
                                    ok
                                ok
                            next
                        ok
                    next
                ok
            ok
        next
        return result

    func hasTOC
        return bSrcHasTOC

    func getTOCTitle
        return cSrcTOCTitle

    func hasWatermark
        return bSrcWatermark

    func getWatermarkText
        return cSrcWatermarkText

    func getWatermarkOptions
        return [:color=cSrcWatermarkColor, :opacity=nSrcWatermarkOpacity,
                :rotation=nSrcWatermarkRotation, :font=cSrcWatermarkFont,
                :size=nSrcWatermarkSize]

    func getCustomProperties
        /*  Returns list of [:name, :type, :value] from docProps/custom.xml  */
        return aCustomProps

    func getCustomProperty name
        /*  Returns value of a named custom property, or NULL if not found  */
        lname = lower(name)
        for cp in aCustomProps
            if lower(cp[:name]) = lname  return cp[:value]  ok
        next
        return NULL

    func getHiddenRuns
        /*  Returns list of [:blockIndex, :text] for runs with vanish=true  */
        result = []
        aBlocksLen = len(aBlocks)
        for i = 1 to aBlocksLen
            block = aBlocks[i]
            runs = block[:runs]
            if !isList(runs)  loop  ok
            for run in runs
                if !isList(run)  loop  ok
                if run[:vanish] = true
                    result + [:blockIndex=i, :text=run[:text], :blockType=block[:type]]
                ok
            next
        next
        return result

    func getCapsRuns
        /*  Returns list of [:text, :type] for runs with caps=true or smallCaps=true  */
        result = []
        for block in aBlocks
            runs = block[:runs]
            if !isList(runs)  loop  ok
            for run in runs
                if !isList(run)  loop  ok
                if run[:caps] = true
                    result + [:text=run[:text], :type="caps"]
                elseif run[:smallCaps] = true
                    result + [:text=run[:text], :type="smallCaps"]
                ok
            next
        next
        return result

    func getDStrikeRuns
        /*  Returns list of text strings for runs with double strikethrough  */
        result = []
        for block in aBlocks
            runs = block[:runs]
            if !isList(runs)  loop  ok
            for run in runs
                if !isList(run)  loop  ok
                if run[:dstrike] = true  result + run[:text]  ok
            next
        next
        return result

    func getChartData
        /*  Returns list of full chart blocks including :chartData series info  */
        result = []
        for block in aBlocks
            if block[:type] = "chart"  result + block  ok
        next
        return result

    func getMergeFields
        # Returns all mergefield blocks
        result = []
        for block in aBlocks
            if block[:type] = "mergefield"  result + block  ok
        next
        return result

    func getMergeFieldNames
        # Returns list of unique merge field names
        return aMergeFields

    func getNumberedHeadings
        result = []
        for block in aBlocks
            if block[:type] = "heading"
                nb = block[:numbered]
                if nb = true  result + block  ok
            ok
        next
        return result

    func getFloatingImageWrapTypes
        # Returns list of [:path, :wrapType, :wrapSide] for all floating images
        result = []
        for block in aBlocks
            if block[:type] = "image"
                isF = block[:floating]
                if isF = true
                    wt = block[:wrapType]
                    ws = block[:wrapSide]
                    if wt = NULL  wt = ""  ok
                    if ws = NULL  ws = ""  ok
                    result + [:path=block[:path], :wrapType=wt, :wrapSide=ws,
                              :distT=block[:distT], :distB=block[:distB],
                              :distL=block[:distL], :distR=block[:distR]]
                ok
            ok
        next
        return result

    func getTableRowProperties
        # Returns list of row-property summaries per table block
        result = []
        for block in aBlocks
            if block[:type] = "table"
                rh = block[:rowHeights]
                rc = block[:rowCantSplit]
                ri = block[:rowIsHeader]
                rb = block[:rowBgColors]
                if isList(rh)
                    result + [:rowHeights=rh, :rowCantSplit=rc, :rowIsHeader=ri, :rowBgColors=rb]
                ok
            ok
        next
        return result

    func getNamedStyleParagraphs
        /*  Returns all paragraph blocks that carry a named paragraph style
            (e.g. "Quote", "Caption", "ListParagraph", "IntenseEmphasis") */
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                sn = block[:style]
                if sn != NULL and len(sn) > 0  result + block  ok
            ok
        next
        return result

    func getOutlinedParagraphs
        /*  Returns paragraphs with w:outlineLvl set (used in TOC / nav pane).
            Each item is [:block, :level]. */
        result = []
        for block in aBlocks
            bt = block[:type]
            if bt = "paragraph" or bt = "rtlparagraph"
                olvl = block[:outlineLvl]
                if isNumber(olvl) and olvl >= 0
                    result + [:block=block, :level=olvl]
                ok
            ok
        next
        return result

    func getRunBorderRuns
        /*  Returns list of [:text, :borderStyle, :borderColor, :borderSize]
            for all runs that carry a w:bdr run border. */
        result = []
        for block in aBlocks
            runs = block[:runs]
            if !isList(runs)  loop  ok
            for run in runs
                if !isList(run)  loop  ok
                bs = run[:borderStyle]
                if bs != NULL and len(bs) > 0
                    result + [:text=run[:text], :borderStyle=bs,
                               :borderColor=run[:borderColor],
                               :borderSize=run[:borderSize]]
                ok
            next
        next
        return result

    func getTableStyle blockIndex
        /*  Returns the tblStyle name for the table block at blockIndex (1-based),
            or empty string if not set / not a table. */
        if blockIndex < 1 or blockIndex > len(aBlocks)  return ""  ok
        block = aBlocks[blockIndex]
        if block[:type] != "table"  return ""  ok
        sn = block[:tblStyle]
        if sn = NULL  return ""  ok
        return sn
    func getTables
        result = []
        for block in aBlocks
            if block[:type] = "table"  result + block  ok
        next
        return result
    func getImages
        result = []
        for block in aBlocks
            if block[:type] = "image"  result + block  ok
        next
        return result

    # =========================================================================
    # Block access helpers
    # =========================================================================

    func blockCount
        return len(aBlocks)

    func getBlock idx
        if idx >= 1 and idx <= len(aBlocks)
            return aBlocks[idx]
        ok
        return []

    # =========================================================================
    # Table query helpers
    # =========================================================================

    func getTableBlockIndex n
        /*  Return the aBlocks index of the nth table (1-based).
            Returns 0 if not found.  */
        count = 0
        for i = 1 to len(aBlocks)
            if aBlocks[i][:type] = "table"
                count++
                if count = n  return i  ok
            ok
        next
        return 0

    func getTableAt n
        /*  Return the nth table block object (1-based).
            Returns [] if not found.  */
        count = 0
        for block in aBlocks
            if block[:type] = "table"
                count++
                if count = n  return block  ok
            ok
        next
        return []

    # =========================================================================
    # Table mutation methods
    # =========================================================================

    func replaceInBlock idx, oldStr, newStr
        /*  Replace all occurrences of oldStr with newStr in the block's
            :text and in each run's :text.  */
        if idx < 1 or idx > len(aBlocks)  return self  ok
        t = aBlocks[idx][:text]
        if t != NULL and len(t) > 0
            aBlocks[idx][:text] = substr(t, oldStr, newStr)
        ok
        runs = aBlocks[idx][:runs]
        if isList(runs)
            for ri = 1 to len(runs)
                rt = runs[ri][:text]
                if rt != NULL and len(rt) > 0
                    runs[ri][:text] = substr(rt, oldStr, newStr)
                ok
            next
            aBlocks[idx][:runs] = runs
        ok
        return self

    # =========================================================================
    # Block insertion
    # =========================================================================

    func insertAfter idx, block
        /*  Insert block after position idx in aBlocks.
            All subsequent indices shift by 1.  */
        if idx < 0                  idx = 0          ok
        if idx > len(aBlocks)       idx = len(aBlocks)  ok
        # Build new list with the inserted block
        newBlocks = []
        for i = 1 to idx
            newBlocks + aBlocks[i]
        next
        newBlocks + block
        for i = idx + 1 to len(aBlocks)
            newBlocks + aBlocks[i]
        next
        aBlocks = newBlocks
        return self

    func findBlocks searchStr
        /*  Return a list of 1-based block indices whose :text contains
            searchStr (case-sensitive). Also searches table cell text.  */
        result = []
        if searchStr = NULL or len(searchStr) = 0  return result  ok
        for i = 1 to len(aBlocks)
            block = aBlocks[i]
            bType = block[:type]
            if bType = "table"
                rows = block[:rows]
                if isList(rows)
                    found = false
                    for row in rows
                        if found  break  ok
                        if isList(row)
                            for cell in row
                                if isList(cell)
                                    ct = cell[:text]
                                    if ct != NULL and substr(ct, searchStr) > 0
                                        found = true
                                        break
                                    ok
                                ok
                            next
                        ok
                    next
                    if found  result + i  ok
                ok
            else
                t = block[:text]
                if t != NULL and substr(t, searchStr) > 0
                    result + i
                ok
            ok
        next
        return result

    # =========================================================================
    # Table cell read
    # =========================================================================

    func getCell blockIndex, rowNum, colNum
        /*  Return the text of cell [rowNum, colNum] (1-based)
            from the table block at aBlocks[blockIndex].
            Returns "" if not found.  */
        if blockIndex < 1 or blockIndex > len(aBlocks)  return ""  ok
        if aBlocks[blockIndex][:type] != "table"        return ""  ok
        rows = aBlocks[blockIndex][:rows]
        if !isList(rows)                                return ""  ok
        if rowNum < 1 or rowNum > len(rows)             return ""  ok
        row = rows[rowNum]
        if !isList(row)                                 return ""  ok
        if colNum < 1 or colNum > len(row)              return ""  ok
        cell = row[colNum]
        if !isList(cell)                                return ""  ok
        t = cell[:text]
        if t = NULL  return ""  ok
        return t

    # =========================================================================
    # Convenience: set cell by table number (not block index)
    # =========================================================================

    func setTableCell tableN, rowNum, colNum, newText
        /*  Like setCell but takes the nth table (1-based) instead of
            a raw block index.  */
        idx = getTableBlockIndex(tableN)
        if idx = 0  return self  ok
        return setCell(idx, rowNum, colNum, newText)

    # =========================================================================
    # Global find/replace across all blocks
    # =========================================================================

    func replaceText oldStr, newStr
        /*  Replace all occurrences of oldStr with newStr in every block:
            paragraph/heading/etc. :text, run :text fields, and table
            cell text and cell run text.  */
        if oldStr = NULL or len(oldStr) = 0  return self  ok
        if newStr = NULL                      newStr = ""  ok
        for i = 1 to len(aBlocks)
            bType = aBlocks[i][:type]
            if bType = "table"
                rows = aBlocks[i][:rows]
                if isList(rows)
                    for ri = 1 to len(rows)
                        if isList(rows[ri])
                            for ci = 1 to len(rows[ri])
                                if isList(rows[ri][ci])
                                    ct = rows[ri][ci][:text]
                                    if ct != NULL and substr(ct, oldStr) > 0
                                        rows[ri][ci][:text] = substr(ct, oldStr, newStr)
                                        rows[ri][ci][:runs] = []
                                    ok
                                ok
                            next
                        ok
                    next
                    aBlocks[i][:rows] = rows
                ok
            else
                t = aBlocks[i][:text]
                if t != NULL and substr(t, oldStr) > 0
                    aBlocks[i][:text] = substr(t, oldStr, newStr)
                ok
                runs = aBlocks[i][:runs]
                if isList(runs)
                    for ri = 1 to len(runs)
                        rt = runs[ri][:text]
                        if rt != NULL and substr(rt, oldStr) > 0
                            runs[ri][:text] = substr(rt, oldStr, newStr)
                        ok
                    next
                    aBlocks[i][:runs] = runs
                ok
            ok
        next
        return self

    # =========================================================================
    # Block-level text replace
    # =========================================================================

    func setBlockText idx, newText
        /*  Replace the entire :text of block idx and clear its runs,
            so round-trip writes newText as a plain paragraph.  */
        if newText = NULL  newText = ""  ok
        if idx < 1 or idx > len(aBlocks)  return self  ok
        aBlocks[idx][:text] = newText
        aBlocks[idx][:runs] = []
        return self

    # =========================================================================
    # Append a plain table (rows of strings, colWidths in cm)
    # =========================================================================

    func setHeadingText n, newText
        /*  Replace the text of the nth heading block (1-based).  */
        if newText = NULL  newText = ""  ok
        count = 0
        for i = 1 to len(aBlocks)
            if aBlocks[i][:type] = "heading"
                count++
                if count = n
                    aBlocks[i][:text] = newText
                    # Also clear runs so round-trip uses plain text
                    aBlocks[i][:runs] = []
                    return self
                ok
            ok
        next
        return self

    # =========================================================================
    # Table cell mutation
    # =========================================================================

    func setCell blockIndex, rowNum, colNum, newText
        /*  Replace the text of cell [rowNum, colNum] (both 1-based)
            in the table at aBlocks[blockIndex].  */
        if newText = NULL  newText = ""  ok
        if blockIndex < 1 or blockIndex > len(aBlocks)  return self  ok
        if aBlocks[blockIndex][:type] != "table"        return self  ok
        rows = aBlocks[blockIndex][:rows]
        if !isList(rows)                                return self  ok
        if rowNum < 1 or rowNum > len(rows)             return self  ok
        row = rows[rowNum]
        if !isList(row)                                 return self  ok
        if colNum < 1 or colNum > len(row)              return self  ok
        row[colNum][:text] = newText
        row[colNum][:runs] = []
        rows[rowNum] = row
        aBlocks[blockIndex][:rows] = rows
        return self

    # =========================================================================

    func getImagesWithAlt
        /*  Returns list of image blocks that have a non-empty altText.
            Each item: block with :altText, :widthCm, :heightCm, :path  */
        result = []
        for block in aBlocks
            if block[:type] = "image"
                alt = block[:altText]
                if alt != NULL and len(alt) > 0
                    result + block
                ok
            ok
        next
        return result

    func getTableLayouts
        /*  Returns list of layout info for every table block.
            Each item: [:tblStyle, :tblAlign, :tblWidthCm, :tblWidthType, :tblIndentCm, :numRows, :numCols]  */
        result = []
        for block in aBlocks
            if block[:type] = "table"
                twTwips = block[:tblWidthTwips]
                twType  = block[:tblWidthType]
                twCm    = 0.0
                if isNumber(twTwips) and twTwips > 0
                    twCm = twTwips / 567.0
                ok
                tiTwips = block[:tblIndentTwips]
                tiCm    = 0.0
                if isNumber(tiTwips) and tiTwips > 0
                    tiCm = tiTwips / 567.0
                ok
                nR = 0
                nC = 0
                rws = block[:rows]
                if isList(rws)
                    nR = len(rws)
                    if nR > 0 and isList(rws[1])
                        nC = len(rws[1])
                    ok
                ok
                entry = [
                    :tblStyle     = block[:tblStyle],
                    :tblAlign     = block[:tblAlign],
                    :tblWidthCm   = twCm,
                    :tblWidthType = twType,
                    :tblIndentCm  = tiCm,
                    :numRows      = nR,
                    :numCols      = nC
                ]
                result + entry
            ok
        next
        return result

    func getSectionLayouts
        /*  Returns info for every sectionbreak block.
            Each item: [:breakType, :numColumns, :columnSpaceCm,
                        :sectHeader, :sectFooter]  */
        result = []
        for block in aBlocks
            if block[:type] = "sectionbreak"
                nc = block[:numColumns]
                cs = block[:columnSpaceTwips]
                if nc = NULL or !isNumber(nc)  nc = 1    ok
                if cs = NULL or !isNumber(cs)  cs = 720  ok
                sh = block[:sectHeader]
                sf = block[:sectFooter]
                if sh = NULL  sh = ""  ok
                if sf = NULL  sf = ""  ok
                entry = [
                    :breakType      = block[:breakType],
                    :numColumns     = nc,
                    :columnSpaceCm  = cs / 567.0,
                    :sectHeader     = sh,
                    :sectFooter     = sf
                ]
                result + entry
            ok
        next
        return result

    func getRunLanguages
        /*  F2 (v3l): Returns all runs that have an explicit language tag.
            Each item: [:lang, :text]  */
        result = []
        for block in aBlocks
            runs = block[:runs]
            if isList(runs)
                for run in runs
                    if isList(run)
                        lg = run[:lang]
                        if lg != NULL and len(lg) > 0
                            result + [:lang=lg, :text=run[:text]]
                        ok
                    ok
                next
            ok
        next
        return result

    func getCharStyleRuns
        /*  Returns all runs that carry a named character style.
            Each item: [:styleName, :text]  */
        result = []
        for block in aBlocks
            runs = block[:runs]
            if isList(runs)
                for run in runs
                    if isList(run)
                        sn = run[:styleName]
                        if sn != NULL and len(sn) > 0
                            result + [:styleName=sn, :text=run[:text]]
                        ok
                    ok
                next
            ok
        next
        return result

    func getFields
        /*  Returns all field blocks (DATE, AUTHOR, TITLE, etc.) found in the document.
            Each item: [:fieldType, :fieldResult]  */
        result = []
        for block in aBlocks
            if block[:type] = "field"  result + block  ok
        next
        return result

    func getListRestartPoints
        /*  Returns list items where the numbered list explicitly restarts.
            Each item: [:text, :level, :startAt, :numId]  */
        result = []
        for block in aBlocks
            if block[:type] = "listitem"
                sa = block[:listStartAt]
                if sa != NULL and isNumber(sa) and sa != 1
                    entry = [:text=block[:text], :level=block[:listIlvl],
                             :startAt=sa, :numId=block[:listNumId]]
                    result + entry
                ok
            ok
        next
        return result

    func getSectionHeaders
        /*  Returns per-section header/footer text found in mid-document sectPr elements.
            Each item: [:headerText, :footerText, :evenHeaderText, :evenFooterText, :breakType]  */
        result = []
        for block in aBlocks
            if block[:type] = "sectionbreak"
                sh = block[:sectHeader]
                sf = block[:sectFooter]
                seh = block[:sectEvenHeader]
                sef = block[:sectEvenFooter]
                if sh = NULL  sh = ""  ok
                if sf = NULL  sf = ""  ok
                if seh = NULL  seh = ""  ok
                if sef = NULL  sef = ""  ok
                if len(sh) > 0 or len(sf) > 0 or len(seh) > 0 or len(sef) > 0
                    entry = [:headerText=sh, :footerText=sf,
                             :evenHeaderText=seh, :evenFooterText=sef,
                             :breakType=block[:breakType]]
                    result + entry
                ok
            ok
        next
        return result

    func getHyperlinks
        result = []
        for block in aBlocks
            if block[:type] = "hyperlink"  result + block  ok
        next
        return result
    func getListItems
        result = []
        for block in aBlocks
            if block[:type] = "listitem"  result + block  ok
        next
        return result
    func getCaptions
        result = []
        for block in aBlocks
            if block[:type] = "caption"  result + block  ok
        next
        return result
    func getTabbedParagraphs
        result = []
        for block in aBlocks
            if block[:type] = "tabbed"  result + block  ok
        next
        return result
    func getSectionBreaks
        result = []
        for block in aBlocks
            if block[:type] = "sectionbreak"  result + block  ok
        next
        return result
    func getCharts
        result = []
        for block in aBlocks
            if block[:type] = "chart"  result + block  ok
        next
        return result
    func getBookmarks
        result = []
        aBlocksLen = len(aBlocks)
        for i = 1 to aBlocksLen
            block = aBlocks[i]
            bm = block[:bookmark]
            if bm != NULL and len(bm) > 0
                result + [:name=bm, :blockIndex=i]
            ok
        next
        return result
    func getLandscapeSections
        result = []
        inLand = false
        startIdx = 0
        aBlocksLen = len(aBlocks)
        for i = 1 to aBlocksLen
            bt = aBlocks[i][:type]
            if bt = "landscapestart"
                inLand = true
                startIdx = i
            ok
            if bt = "landscapeend" and inLand
                result + [:startIndex=startIdx, :endIndex=i]
                inLand = false
            ok
        next
        return result

    # =========================================================================
    # toWriter — convert parsed blocks back to a WordWriter
    # =========================================================================

    func toWriter
        writer = new WordWriter()
        writer.setTitle(cTitle)
        writer.setAuthor(cAuthor)
        writer.setDefaultFont(cDefaultFont, nDefaultSize)

        # Document RTL
        if bSrcDocumentRTL
            writer.setDocumentRTL(true)
        ok

        # Watermark round-trip
        if bSrcWatermark = true
            wmOpts = [:color=cSrcWatermarkColor, :opacity=nSrcWatermarkOpacity,
                      :rotation=nSrcWatermarkRotation, :font=cSrcWatermarkFont,
                      :size=nSrcWatermarkSize]
            writer.setWatermark(cSrcWatermarkText, wmOpts)
        ok
        # Image watermark round-trip
        if bSrcImgWatermark = true and len(cSrcImgWatermarkPath) > 0
            if fexists(cSrcImgWatermarkPath)
                imgWmOpts = [:opacity=nSrcImgWatermarkOpacity,
                             :width=nSrcImgWatermarkWidthCm,
                             :height=nSrcImgWatermarkHeightCm]
                writer.setImageWatermark(cSrcImgWatermarkPath, imgWmOpts)
            ok
        ok

        # Page setup - pass exact page size (in twips) from source document
        # Set orientation FIRST (it swaps dims), then override with actual source dims
        if cSrcOrientation = "landscape"
            writer.setOrientation("landscape")
        ok
        writer.setCustomPageSize(nPageWidth/567.0, nPageHeight/567.0)
        # Pass source document default paragraph spacing so output matches
        if nDocSpacingAfter >= 0 and nDocSpacingLine > 0
            writer.setDocDefaultSpacing(nDocSpacingAfter, nDocSpacingLine)
        ok
        # Pass source styles.xml and theme1.xml for round-trip fidelity
        if len(cSrcStylesXml) > 0  writer.useSourceStyles(cSrcStylesXml)  ok
        if len(cSrcThemeXml)  > 0  writer.useSourceTheme(cSrcThemeXml)    ok
        writer.setMargins(nMarginTop/567.0, nMarginBottom/567.0,
                          nMarginLeft/567.0, nMarginRight/567.0)
        writer.setHeaderFooterMargins(nHeaderMargin, nFooterMargin)

        # Header / footer
        if len(cHeaderText) > 0
            writer.setHeader(cHeaderText)
        ok
        if len(cFooterText) > 0
            writer.setFooter(cFooterText)
        ok
        if bSrcEvenAndOddHeaders
            writer.setEvenAndOddHeaders(true)
            if len(cEvenPageHeaderText) > 0  writer.setEvenPageHeader(cEvenPageHeaderText)  ok
            if len(cEvenPageFooterText) > 0  writer.setEvenPageFooter(cEvenPageFooterText)  ok
        ok
        if bSrcFirstPageDifferent
            writer.setFirstPageDifferent(true)
            if len(cFirstPageHeaderText) > 0  writer.setFirstPageHeader(cFirstPageHeaderText)  ok
            if len(cFirstPageFooterText) > 0  writer.setFirstPageFooter(cFirstPageFooterText)  ok
        ok

        # Page border
        if bSrcPageBorder
            # nSrcPageBorderSize is in eighths-of-point; setPageBorder expects points
            pgBdrOpts = [:style=cSrcPageBorderStyle,
                         :color=cSrcPageBorderColor,
                         :size=nSrcPageBorderSize/8.0,
                         :space=nSrcPageBorderSpace]
            # Pass only the sides that were present in the source
            if isList(aSrcPageBorderSides) and len(aSrcPageBorderSides) > 0
                pgBdrOpts[:sides] = aSrcPageBorderSides
            ok
            writer.setPageBorder(pgBdrOpts)
        ok

        # Page background
        if len(cSrcPageBgColor) > 0
            writer.setPageBackground(cSrcPageBgColor)
        ok

        # Page numbers
        if bSrcPageNumbers
            writer.showPageNumbers(cSrcPageNumAlign)
        ok

        # Theme
        if isList(aSourceThemeColors) and len(aSourceThemeColors) >= 2
            if aSourceThemeColors[1] = "_theme_"
                writer.setTheme(aSourceThemeColors[2])
            elseif aSourceThemeColors[1] = "_custom_" and len(aSourceThemeColors) >= 3
                writer.setThemeColors(aSourceThemeColors[2], aSourceThemeColors[3])
            ok
        ok

        # Columns
        if nSrcColumns >= 2
            writer.setColumns(nSrcColumns, nSrcColumnSpace / 567.0)
        ok

        # Replay blocks (index-based so list grouping can look-ahead)
        # Pre-pass — build header injection map from mid-doc sectionbreak blocks
        aHdrInjections5 = []
        prevSectEnd5 = 0
        prePassN5 = len(aBlocks)
        for hpi5 = 1 to prePassN5
            if aBlocks[hpi5][:type] = "sectionbreak"
                sh5pre  = aBlocks[hpi5][:sectHeader]
                sf5pre  = aBlocks[hpi5][:sectFooter]
                seh5pre = aBlocks[hpi5][:sectEvenHeader]
                sef5pre = aBlocks[hpi5][:sectEvenFooter]
                if sh5pre  = NULL  sh5pre  = ""  ok
                if sf5pre  = NULL  sf5pre  = ""  ok
                if seh5pre = NULL  seh5pre = ""  ok
                if sef5pre = NULL  sef5pre = ""  ok
                if len(sh5pre) > 0 or len(sf5pre) > 0 or len(seh5pre) > 0 or len(sef5pre) > 0
                    injAt5 = prevSectEnd5 + 1
                    aHdrInjections5 + [:atBlock=injAt5, :hdr=sh5pre, :ftr=sf5pre, :evenHdr=seh5pre, :evenFtr=sef5pre]
                ok
                prevSectEnd5 = hpi5
            ok
        next
        nBlocks = len(aBlocks)
        blkIdx  = 0
        while blkIdx < nBlocks
            blkIdx++
            block = aBlocks[blkIdx]
            bType = block[:type]
            if bType = NULL  bType = ""  ok
            # inject setHeader/setFooter when this block starts a new section
            aHdrInjections5Len = len(aHdrInjections5)
            for hii5 = 1 to aHdrInjections5Len
                if aHdrInjections5[hii5][:atBlock] = blkIdx
                    h5tx  = aHdrInjections5[hii5][:hdr]
                    f5tx  = aHdrInjections5[hii5][:ftr]
                    eh5tx = aHdrInjections5[hii5][:evenHdr]
                    ef5tx = aHdrInjections5[hii5][:evenFtr]
                    if len(h5tx)  > 0  writer.setHeader(h5tx)      ok
                    if len(f5tx)  > 0  writer.setFooter(f5tx)      ok
                    if len(eh5tx) > 0  writer.setEvenPageHeader(eh5tx)  ok
                    if len(ef5tx) > 0  writer.setEvenPageFooter(ef5tx)  ok
                ok
            next

            if bType = "empty"
                # Empty paragraph - reproduce as blank line for layout fidelity
                writer.addParagraph("", [])

            elseif bType = "heading"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                runs = block[:runs]
                hLvl = block[:level]
                if hLvl = NULL  hLvl = 1  ok
                hPOpts = []
                if isList(runs) and len(runs) > 1
                    wrRuns = []
                    for run in runs
                        rOpts2 = []
                        if run[:bold]      = true  rOpts2[:bold]      = true  ok
                        if run[:italic]    = true  rOpts2[:italic]    = true  ok
                        if run[:underline] = true  rOpts2[:underline] = true  ok
                        if len(run[:color]) > 0    rOpts2[:color]     = run[:color]  ok
                        if len(run[:font])  > 0    rOpts2[:font]      = run[:font]   ok
                        if run[:size] > 0          rOpts2[:size]      = run[:size]   ok
                        if run[:rtl] = true  rOpts2[:rtl] = true  ok
                        wPair = [run[:text], rOpts2]
                        wrRuns + Ref(wPair)
                    next
                    if block[:numbered] = true
                        writer.addNumberedHeading(block[:text], hLvl, block[:numId])
                    else
                        writer.addHeading(block[:text], hLvl)
                    ok
                else
                    if block[:numbered] = true
                        writer.addNumberedHeading(block[:text], hLvl, block[:numId])
                    else
                        writer.addHeading(block[:text], hLvl)
                    ok
                ok

            elseif bType = "paragraph" or bType = "rtlparagraph"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                runs = block[:runs]
                pOpts = []
                al   = block[:align]
                sb   = block[:spaceBefore]
                sa   = block[:spaceAfter]
                pbg  = block[:bgColor]
                pin  = block[:indent]
                pir  = block[:indentRight]
                pifl = block[:indentFirstLine]
                pls  = block[:lineSpacing]
                plsp = block[:lineSpacingPt]
                plr  = block[:lineRule]
                pkn  = block[:keepNext]
                pkl  = block[:keepLines]
                ppbb = block[:pageBreakBefore]
                pstyle = block[:style]
                polvl  = block[:outlineLvl]
                pctxsp = block[:contextualSpacing]
                if al  != NULL and len(al) > 0   pOpts[:align]       = al   ok
                if sb  != NULL and sb  > 0        pOpts[:spaceBefore] = sb   ok
                if sa  != NULL and sa  > 0        pOpts[:spaceAfter]  = sa   ok
                if pbg != NULL and len(pbg) > 0   pOpts[:bgColor]     = pbg  ok
                if pin != NULL and pin > 0        pOpts[:indent]      = pin  ok
                if pir != NULL and pir > 0        pOpts[:indentRight] = pir  ok
                if pifl != NULL and pifl > 0      pOpts[:indentFirstLine] = pifl   ok
                if pifl != NULL and pifl < 0      pOpts[:indentHanging]   = -pifl  ok
                # Named paragraph style (non-empty, not a heading variant)
                if pstyle != NULL and len(pstyle) > 0
                    sln = lower(pstyle)
                    isHdgStyle = (sln = "heading1" or sln = "heading2" or
                                  sln = "heading3" or sln = "heading4" or
                                  sln = "heading5" or sln = "heading6" or
                                  sln = "heading 1" or sln = "heading 2" or
                                  sln = "heading 3" or sln = "heading 4" or
                                  sln = "heading 5" or sln = "heading 6" or
                                  sln = "normal" or sln = "default")
                    if !isHdgStyle
                        pOpts[:style] = pstyle
                    ok
                ok
                # Outline level
                if isNumber(polvl) and polvl >= 0 and polvl <= 8
                    pOpts[:outlineLvl] = polvl
                ok
                # Contextual spacing
                if pctxsp = true  pOpts[:contextualSpacing] = true  ok
                if isNumber(pls) and pls > 0 and pls != 1.0
                    pOpts[:lineSpacing] = pls
                ok
                if isNumber(plsp) and plsp > 0 and len(plr) > 0 and plr != "auto"
                    pOpts[:lineSpacingPt] = plsp
                    pOpts[:lineRule]      = plr
                ok
                if pkn  = true  pOpts[:keepNext]  = true  ok
                if pkl  = true  pOpts[:keepLines] = true  ok
                # widow/orphan control
                pwc = block[:widowControl]
                if pwc = false  pOpts[:widowControl] = false  ok
                # suppress auto-hyphenation
                pnh = block[:noHyphenate]
                if pnh = true  pOpts[:noHyphenate] = true  ok
                hb = block[:hasBorder]
                if hb = true
                    pOpts[:_hasBorder]  = true
                    pOpts[:borderStyle] = block[:borderStyle]
                    pOpts[:borderColor] = block[:borderColor]
                    pOpts[:borderSize]  = block[:borderSize]
                    pOpts[:borderSpace] = block[:borderSpace]
                    bsides = block[:borderSides]
                    if isList(bsides) and len(bsides) > 0
                        pOpts[:sides] = bsides
                    ok
                ok
                if bType = "rtlparagraph"
                    pOpts[:rtl] = true
                    if al = NULL or len(al) = 0  pOpts[:align] = "right"  ok
                ok
                cmt   = block[:commentText]
                cAuth = block[:commentAuthor]
                hasComment = cmt != NULL and len(cmt) > 0
                # page-break-before: insert explicit page break before this paragraph
                if ppbb = true  writer.addPageBreak()  ok
                if isList(runs) and len(runs) > 1
                    wrRuns = []
                    for run in runs
                        if !isList(run)  loop  ok
                        rOpts2 = []
                        if run[:bold]      = true  rOpts2[:bold]      = true  ok
                        if run[:italic]    = true  rOpts2[:italic]    = true  ok
                        if run[:underline] = true  rOpts2[:underline] = true  ok
                        if run[:strike]    = true  rOpts2[:strike]    = true  ok
                        if run[:dstrike]   = true  rOpts2[:dstrike]   = true  ok
                        if run[:caps]      = true  rOpts2[:caps]      = true  ok
                        if run[:smallCaps] = true  rOpts2[:smallCaps] = true  ok
                        if len(run[:color])     > 0  rOpts2[:color]        = run[:color]     ok
                        if len(run[:font])      > 0  rOpts2[:font]         = run[:font]       ok
                        if len(run[:highlight]) > 0  rOpts2[:highlight]    = run[:highlight]  ok
                        if run[:size] > 0            rOpts2[:size]         = run[:size]       ok
                        if run[:superscript] = true  rOpts2[:superscript]  = true             ok
                        if run[:subscript]   = true  rOpts2[:subscript]    = true             ok
                        # run language
                        rLng2 = run[:lang]
                        if rLng2 != NULL and len(rLng2) > 0  rOpts2[:lang] = rLng2  ok
                        # named character style
                        rSN2 = run[:styleName]
                        if rSN2 != NULL and len(rSN2) > 0  rOpts2[:charStyle] = rSN2  ok
                        rbdrS2 = run[:borderStyle]
                        if rbdrS2 != NULL and len(rbdrS2) > 0
                            rOpts2[:runBorderStyle] = rbdrS2
                            rOpts2[:runBorderColor] = run[:borderColor]
                            rOpts2[:runBorderSize]  = run[:borderSize]
                        ok
                        wPair = [run[:text], rOpts2]
                        wrRuns + Ref(wPair)
                    next
                    if hasComment
                        writer.addCommentedParagraph(block[:text], cmt, pOpts)
                    else
                        writer.addRichParagraph(wrRuns, pOpts)
                    ok
                else
                    txt = block[:text]
                    if txt = NULL  txt = ""  ok
                    if isList(runs) and len(runs) = 1 and isList(runs[1])
                        run = runs[1]
                        if run[:bold]      = true  pOpts[:bold]      = true  ok
                        if run[:italic]    = true  pOpts[:italic]    = true  ok
                        if run[:underline] = true  pOpts[:underline] = true  ok
                        if run[:strike]    = true  pOpts[:strike]    = true  ok
                        if run[:dstrike]   = true  pOpts[:dstrike]   = true  ok
                        if run[:caps]      = true  pOpts[:caps]      = true  ok
                        if run[:smallCaps] = true  pOpts[:smallCaps] = true  ok
                        if len(run[:color])  > 0   pOpts[:color]  = run[:color]  ok
                        if len(run[:font])   > 0   pOpts[:font]   = run[:font]   ok
                        if run[:size] > 0          pOpts[:size]   = run[:size]   ok
                        # F2 (v3l): run language
                        rLng3 = run[:lang]
                        if rLng3 != NULL and len(rLng3) > 0  pOpts[:lang] = rLng3  ok
                        # named character style
                        rSN3 = run[:styleName]
                        if rSN3 != NULL and len(rSN3) > 0  pOpts[:charStyle] = rSN3  ok
                        rbdrS3 = run[:borderStyle]
                        if rbdrS3 != NULL and len(rbdrS3) > 0
                            pOpts[:runBorderStyle] = rbdrS3
                            pOpts[:runBorderColor] = run[:borderColor]
                            pOpts[:runBorderSize]  = run[:borderSize]
                        ok
                        if run[:rtl] = true  pOpts[:rtl] = true  ok
                    ok
                    if hasComment
                        writer.addComment(txt, cmt, cAuth)
                    else
                        writer.addParagraph(txt, pOpts)
                    ok
                ok

            elseif bType = "table"
                rows     = block[:rows]
                colWids  = block[:colWidths]
                hdrBg    = block[:headerBg]
                evenBg   = block[:evenRowBg]
                brdStyle = block[:borderStyle]
                brdColor = block[:borderColor]
                if brdStyle = NULL  brdStyle = "single"  ok
                if brdColor = NULL  brdColor = "auto"    ok
                tableData = []
                isFirstRow = true
                rowIdx2 = 0
                rowBgList = block[:rowBgColors]
                for row in rows
                    if !isList(row)  loop  ok
                    rowIdx2++
                    rowBg2 = ""
                    if isList(rowBgList) and rowIdx2 <= len(rowBgList)
                        rowBg2 = rowBgList[rowIdx2]
                        if rowBg2 = NULL  rowBg2 = ""  ok
                    ok
                    rowData  = []
                    isHeader = isFirstRow
                    isFirstRow = false
                    for cell in row
                        if isList(cell)
                            # Skip merge-continue cells
                            if cell[:isMergeContinue] = true
                                rowData + wordMergeCell()
                                loop
                            ok
                            cText  = cell[:text]
                            cBg    = cell[:bgColor]
                            cSpan  = cell[:colspan]
                            cRSpan = cell[:rowspan]
                            cAlign = cell[:align]
                            cVAl   = cell[:vAlign]
                            cRuns  = cell[:runs]
                            cHlUrl = cell[:hlUrl]
                            cImgPth = cell[:imgPath]
                            if cText  = NULL  cText  = ""  ok
                            if cBg    = NULL  cBg    = ""  ok
                            if cSpan  = NULL  cSpan  = 1   ok
                            if cRSpan = NULL  cRSpan = 1   ok
                            if cAlign = NULL  cAlign = ""  ok
                            if cVAl   = NULL  cVAl   = ""  ok
                            if cHlUrl = NULL  cHlUrl = ""  ok
                            if cImgPth = NULL  cImgPth = ""  ok

                            # Build cell options
                            cOpts = []
                            # Cell bg: use cell-level if set, else row-level bg
                            effectiveBg = cBg
                            if len(effectiveBg) = 0 and len(rowBg2) > 0
                                effectiveBg = rowBg2
                            ok
                            if len(effectiveBg) > 0  cOpts[:bgColor]      = effectiveBg  ok
                            if cSpan       > 1  cOpts[:colspan]       = cSpan   ok
                            if cRSpan      > 1  cOpts[:rowspan]       = cRSpan  ok
                            if len(cAlign) > 0  cOpts[:align]         = cAlign  ok
                            if len(cVAl)   > 0  cOpts[:verticalAlign] = cVAl    ok
                            # Do not force bold for header row; let source run formatting carry through
                            # if isHeader  cOpts[:bold] = true  ok

                            # Decide cell creation strategy:
                            # - Image cells: empty wordCell + addCellImage
                            # - Hyperlink cells: wordCell with text + addCellHyperlink
                            # - Multi-run rich cells: mergeCell with run list
                            # - Plain text cells: wordCell (simplest, most reliable)
                            hasImg = len(cImgPth) > 0 and fexists(cImgPth)
                            hasHl  = len(cHlUrl) > 0
                            hasRichRuns = isList(cRuns) and len(cRuns) > 1
                            # Check for rich multi-paragraph cell content
                            cParas = cell[:cellParas]
                            hasRichParas = false
                            if isList(cParas) and len(cParas) > 1
                                hasRichParas = true
                            elseif isList(cParas) and len(cParas) = 1
                                ptype1 = cParas[1][:type]
                                if ptype1 = "listitem" or ptype1 = "blockquote"
                                   or ptype1 = "caption"
                                    hasRichParas = true
                                ok
                            ok

                            if hasRichParas
                                # Multi-paragraph rich cell: reproduce each paragraph
                                wc = new WordCell()
                                if len(cBg) > 0  wc.cBgColor = cBg  ok
                                if len(cAlign) > 0  wc.setAlign(cAlign)  ok
                                if len(cVAl) > 0    wc.setVerticalAlign(cVAl)  ok
                                if cSpan > 1        wc.nColSpan = cSpan  ok
                                bFirstPara = true
                                listRunsQ = []
                                listIsBulletQ = true
                                for cpBlk in cParas
                                    cpType = cpBlk[:type]
                                    cpText = cpBlk[:text]
                                    if cpText = NULL  cpText = ""  ok
                                    cpRuns = cpBlk[:runs]
                                    # Flush list queue if current para is not listitem
                                    if cpType != "listitem" and len(listRunsQ) > 0
                                        listIdQ = writer.registerCellList(listIsBulletQ)
                                        if listIsBulletQ
                                            wc.addCellBulletList(listRunsQ, listIdQ)
                                        else
                                            wc.addCellNumberedList(listRunsQ, listIdQ)
                                        ok
                                        listRunsQ = []
                                    ok
                                    if cpType = "listitem"
                                        listRunsQ + cpText
                                        cpBulletFlag = cpBlk[:isBullet]
                                        if cpBulletFlag != NULL  listIsBulletQ = cpBulletFlag = true
                                        else  listIsBulletQ = true  ok
                                    elseif cpType = "blockquote"
                                        wc.addCellBlockQuote(cpText)
                                    elseif cpType = "caption"
                                        wc.addCellCaption(cpText)
                                    elseif cpType = "hyperlink"
                                        hlUrl3 = cpBlk[:url]
                                        if hlUrl3 = NULL  hlUrl3 = ""  ok
                                        if len(hlUrl3) > 0
                                            hlRelId3 = writer.registerHyperlink(hlUrl3)
                                            wc.addCellHyperlink(cpText, hlRelId3)
                                        else
                                            if !bFirstPara  wc.addLineBreak()  ok
                                            wc.addRun(cpText, [])
                                        ok
                                    elseif cpType = "empty"
                                        # empty paragraph in cell - skip
                                    else
                                        # Normal paragraph: add runs with formatting
                                        if !bFirstPara  wc.addLineBreak()  ok
                                        if isList(cpRuns) and len(cpRuns) > 0
                                            for cpRun in cpRuns
                                                rOptsCP = []
                                                if cpRun[:bold]      = true  rOptsCP[:bold]      = true  ok
                                                if cpRun[:italic]    = true  rOptsCP[:italic]    = true  ok
                                                if cpRun[:underline] = true  rOptsCP[:underline] = true  ok
                                                if len(cpRun[:color] + "")    > 0  rOptsCP[:color]    = cpRun[:color]    ok
                                                if len(cpRun[:font]  + "")    > 0  rOptsCP[:font]     = cpRun[:font]     ok
                                                if cpRun[:size]               > 0  rOptsCP[:size]     = cpRun[:size]     ok
                                                wc.addRun(cpRun[:text], rOptsCP)
                                            next
                                        elseif len(cpText) > 0
                                            wc.addRun(cpText, [])
                                        ok
                                        bFirstPara = false
                                    ok
                                next
                                # Flush any remaining list items
                                if len(listRunsQ) > 0
                                    listIdQ2 = writer.registerCellList(listIsBulletQ)
                                    if listIsBulletQ
                                        wc.addCellBulletList(listRunsQ, listIdQ2)
                                    else
                                        wc.addCellNumberedList(listRunsQ, listIdQ2)
                                    ok
                                ok
                            elseif hasImg
                                # Create empty cell then attach image
                                wc = wordCell("", cOpts)
                                cIW = cell[:imgW]
                                cIH = cell[:imgH]
                                if !isNumber(cIW) or cIW <= 0  cIW = 3.0  ok
                                if !isNumber(cIH) or cIH <= 0  cIH = 2.0  ok
                                imgReg = writer.registerImage(cImgPth)
                                wc.addCellImage(cImgPth, cIW, cIH)
                                wc.setCellImageRel(1, imgReg[:relId], imgReg[:imageId])
                            elseif hasHl
                                # Plain cell + hyperlink attachment
                                wc = wordCell(cText, cOpts)
                                hlRelId = writer.registerHyperlink(cHlUrl)
                                hlTxt   = cell[:hlText]
                                if hlTxt = NULL or len(hlTxt) = 0  hlTxt = cText  ok
                                wc.addCellHyperlink(hlTxt, hlRelId)
                            elseif hasRichRuns
                                # Multiple runs with mixed formatting
                                wcRunList = []
                                for run in cRuns
                                    if !isList(run)  loop  ok
                                    rOpts2 = []
                                    if run[:bold]      = true  rOpts2[:bold]      = true  ok
                                    if run[:italic]    = true  rOpts2[:italic]    = true  ok
                                    if run[:underline] = true  rOpts2[:underline] = true  ok
                                    if run[:strike]    = true  rOpts2[:strike]    = true  ok
                                    if len(run[:color])      > 0  rOpts2[:color]     = run[:color]     ok
                                    if len(run[:font])       > 0  rOpts2[:font]      = run[:font]       ok
                                    if run[:size] > 0             rOpts2[:size]      = run[:size]       ok
                                    if len(run[:highlight]) > 0   rOpts2[:highlight] = run[:highlight]  ok
                                    if run[:rtl] = true            rOpts2[:rtl]       = true            ok
                                    if run[:isLineBreak] = true
                                        # Line break: add as line break run
                                        lbEntry = [:text="", :lineBreak=true]
                                        wcRunList + Ref(lbEntry)
                                    else
                                        # Build proper associative list for mergeCell
                                        wcEntry = [:text=run[:text]]
                                        if rOpts2[:bold]      = true  wcEntry[:bold]      = true  ok
                                        if rOpts2[:italic]    = true  wcEntry[:italic]    = true  ok
                                        if rOpts2[:underline] = true  wcEntry[:underline] = true  ok
                                        if rOpts2[:strike]    = true  wcEntry[:strike]    = true  ok
                                        if rOpts2[:color]     != NULL  if len(rOpts2[:color])     > 0  wcEntry[:color]     = rOpts2[:color]     ok  ok
                                        if rOpts2[:font]      != NULL  if len(rOpts2[:font])      > 0  wcEntry[:font]      = rOpts2[:font]      ok  ok
                                        if rOpts2[:size]              > 0  wcEntry[:size]      = rOpts2[:size]      ok
                                        if rOpts2[:highlight] != NULL  if len(rOpts2[:highlight]) > 0  wcEntry[:highlight] = rOpts2[:highlight] ok  ok
                                        if rOpts2[:rtl]       = true  wcEntry[:rtl]       = true  ok
                                        wcRunList + Ref(wcEntry)
                                    ok
                                next
                                wc = mergeCell(wcRunList, cOpts)
                            else
                                # Plain or single-run cell — use wordCell directly
                                # Apply single-run formatting to cOpts if present
                                if isList(cRuns) and len(cRuns) = 1
                                    run = cRuns[1]
                                    if run[:bold]      = true  cOpts[:bold]      = true  ok
                                    if run[:italic]    = true  cOpts[:italic]    = true  ok
                                    if run[:underline] = true  cOpts[:underline] = true  ok
                                    if len(run[:color])     > 0  cOpts[:color]     = run[:color]     ok
                                    if len(run[:font])      > 0  cOpts[:font]      = run[:font]       ok
                                    if run[:size]           > 0  cOpts[:size]      = run[:size]       ok
                                    if len(run[:highlight]) > 0  cOpts[:highlight] = run[:highlight]  ok
                                    if run[:rtl] = true  cOpts[:rtl] = true  ok
                                ok
                                wc = wordCell(cText, cOpts)
                            ok
                            # Apply cell padding if parsed
                            if !hasImg
                                cPadT = cell[:padTop]
                                cPadB = cell[:padBottom]
                                cPadL = cell[:padLeft]
                                cPadR = cell[:padRight]
                                hasPad = false
                                if isNumber(cPadT) and cPadT >= 0  hasPad = true  ok
                                if isNumber(cPadB) and cPadB >= 0  hasPad = true  ok
                                if isNumber(cPadL) and cPadL >= 0  hasPad = true  ok
                                if isNumber(cPadR) and cPadR >= 0  hasPad = true  ok
                                if hasPad
                                    ptCm = 0.0; pbCm = 0.0; plCm = 0.0; prCm = 0.0
                                    if isNumber(cPadT) and cPadT >= 0  ptCm = cPadT / 567.0  ok
                                    if isNumber(cPadB) and cPadB >= 0  pbCm = cPadB / 567.0  ok
                                    if isNumber(cPadL) and cPadL >= 0  plCm = cPadL / 567.0  ok
                                    if isNumber(cPadR) and cPadR >= 0  prCm = cPadR / 567.0  ok
                                    wc.setPadding(ptCm, pbCm, plCm, prCm)
                                ok
                            ok
                            # Per-cell border
                            cbStyle = cell[:borderStyle]
                            cbColor = cell[:borderColor]
                            cbSides = cell[:borderSides]
                            cbTextDir = cell[:textDir]
                            if isList(cbSides) and len(cbSides) > 0
                                # per-side border replay
                                for sideEntry in cbSides
                                    wc.setBorderSide(sideEntry[:side], sideEntry[:style],
                                                     sideEntry[:color], sideEntry[:size])
                                next
                            elseif cbStyle != NULL and len(cbStyle) > 0
                                if cbColor = NULL  cbColor = "auto"  ok
                                wc.setBorder(cbStyle, cbColor)
                            ok
                            # cell text direction
                            if cbTextDir != NULL and len(cbTextDir) > 0 and cbTextDir != "lrTb"
                                wc.setTextDir(cbTextDir)
                            ok
                            # Nested table inside cell
                            ntObj2 = cell[:nestedTable]
                            if isList(ntObj2)
                                ntRows2 = ntObj2[:rows]
                                if isList(ntRows2) and len(ntRows2) > 0
                                    ntData2 = []
                                    for nt_r2 in ntRows2
                                        ntRow2 = []
                                        if isList(nt_r2)
                                            for nt_c2 in nt_r2
                                                if isList(nt_c2)
                                                    isMC2 = nt_c2[:isMergeContinue]
                                                    if isMC2 = true
                                                        ntRow2 + wordCell("", [])
                                                    else
                                                        # Build nested cell with formatting from cellParas or runs
                                                        nt_txt2  = nt_c2[:text]
                                                        nt_isBold2 = false
                                                        nt_clr2    = ""
                                                        nt_aln2    = nt_c2[:align]
                                                        if nt_aln2 = NULL  nt_aln2 = ""  ok
                                                        # Try cellParas first for richer formatting
                                                        ntCPs2 = nt_c2[:cellParas]
                                                        if isList(ntCPs2) and len(ntCPs2) > 0
                                                            ntCP2 = ntCPs2[1]
                                                            if isList(ntCP2)
                                                                cpAln2 = ntCP2[:align]
                                                                if cpAln2 != NULL and len(cpAln2) > 0
                                                                    nt_aln2 = cpAln2
                                                                ok
                                                                cpRuns2 = ntCP2[:runs]
                                                                if isList(cpRuns2) and len(cpRuns2) > 0
                                                                    cpr2 = cpRuns2[1]
                                                                    if isList(cpr2)
                                                                        if cpr2[:bold] = true  nt_isBold2 = true  ok
                                                                        cpClr2 = cpr2[:color]
                                                                        if cpClr2 != NULL and len(cpClr2) > 0  nt_clr2 = cpClr2  ok
                                                                    ok
                                                                ok
                                                            ok
                                                        ok
                                                        # Fall back to runs
                                                        if !nt_isBold2 and nt_clr2 = ""
                                                            ntRuns2 = nt_c2[:runs]
                                                            if isList(ntRuns2) and len(ntRuns2) > 0
                                                                nr2 = ntRuns2[1]
                                                                if isList(nr2)
                                                                    if nr2[:bold] = true  nt_isBold2 = true  ok
                                                                    ntClr2 = nr2[:color]
                                                                    if ntClr2 != NULL and len(ntClr2) > 0 and ntClr2 != "auto"  nt_clr2 = ntClr2  ok
                                                                ok
                                                            ok
                                                        ok
                                                        # Build wordCell with formatting options
                                                        nt_rOpts2 = []
                                                        if nt_isBold2  nt_rOpts2[:bold] = true  ok
                                                        if len(nt_clr2) > 0  nt_rOpts2[:color] = nt_clr2  ok
                                                        wcnt2 = wordCell(nt_txt2, nt_rOpts2)
                                                        bg2 = nt_c2[:bgColor]
                                                        if bg2 != NULL and len(bg2) > 0  wcnt2.cBgColor = bg2  ok
                                                        if len(nt_aln2) > 0  wcnt2.cAlign = nt_aln2  ok
                                                        ntRow2 + wcnt2
                                                    ok
                                                ok
                                            next
                                        ok
                                        ntData2 + ntRow2
                                    next
                                    if len(ntData2) > 0
                                        # Detect header row from nested table properties
                                        ntHdrBg2 = ntObj2[:headerBg]
                                        if ntHdrBg2 = NULL  ntHdrBg2 = ""  ok
                                        ntHasHdr2 = len(ntHdrBg2) > 0
                                        ntOpts2 = [:borderStyle="single", :headerRow=ntHasHdr2]
                                        if ntHasHdr2  ntOpts2[:headerBgColor] = ntHdrBg2  ok
                                        # Also pass nested table style
                                        ntStyle2 = ntObj2[:tblStyle]
                                        if ntStyle2 != NULL and len(ntStyle2) > 0
                                            ntOpts2[:tblStyle] = ntStyle2
                                        ok
                                        ntXml2 = writer.generateNestedTable(ntData2, ntOpts2)
                                        wc.addCellTable(ntXml2)
                                    ok
                                ok
                            ok
                            rowData + wc
                        else
                            rowData + ("" + cell)
                        ok
                    next
                    tableData + rowData
                next
                # Only treat first row as header if it has a distinct background.
                # Using headerRow=true unconditionally adds bold+center to the
                # first row even when the source had no such formatting.
                bHasRealHeader = len(hdrBg) > 0
                tOpts = [:headerRow=bHasRealHeader, :borderStyle=brdStyle]
                if brdColor != "auto"   tOpts[:borderColor]    = brdColor  ok
                brdMap = block[:bordersMap]
                if isList(brdMap) and len(brdMap) > 0  tOpts[:bordersMap] = brdMap  ok
                tblBg2 = block[:tblBgFill]
                if tblBg2 != NULL and len(tblBg2) > 0  tOpts[:tblBgFill] = tblBg2  ok
                if len(hdrBg) > 0       tOpts[:headerBgColor]  = hdrBg     ok
                if len(evenBg) > 0      tOpts[:evenRowBgColor] = evenBg    ok
                if len(colWids) > 0     tOpts[:colWidths]      = colWids   ok
                # Table style name
                tblSN = block[:tblStyle]
                if tblSN != NULL and len(tblSN) > 0  tOpts[:tblStyle] = tblSN  ok
                # Table width / alignment / indent round-trip
                twTwips = block[:tblWidthTwips]
                twType  = block[:tblWidthType]
                if isNumber(twTwips) and twTwips > 0 and len(twType) > 0 and twType != "auto"
                    tOpts[:tblWidthTwips] = twTwips
                    tOpts[:tblWidthType]  = twType
                ok
                taln = block[:tblAlign]
                if taln != NULL and len(taln) > 0  tOpts[:tblAlign] = taln  ok
                tiTwips = block[:tblIndentTwips]
                if isNumber(tiTwips) and tiTwips > 0  tOpts[:tblIndentTwips] = tiTwips  ok
                # tblLook banding flags
                tlFR2 = block[:tblLookFirstRow]
                tlLR2 = block[:tblLookLastRow]
                tlFC2 = block[:tblLookFirstCol]
                tlLC2 = block[:tblLookLastCol]
                tlNH2 = block[:tblLookNoHBand]
                tlNV2 = block[:tblLookNoVBand]
                if isNumber(tlFR2) and tlFR2 >= 0
                    tblLookMap = []
                    tblLookMap[:firstRow] = tlFR2
                    tblLookMap[:lastRow]  = tlLR2
                    tblLookMap[:firstCol] = tlFC2
                    tblLookMap[:lastCol]  = tlLC2
                    tblLookMap[:noHBand]  = tlNH2
                    tblLookMap[:noVBand]  = tlNV2
                    tOpts[:tblLook] = tblLookMap
                ok
                # Row properties
                rHeights = block[:rowHeights]
                rHRules  = block[:rowHRules]
                rIsHdr   = block[:rowIsHeader]
                if isList(rHeights) and len(rHeights) > 0
                    rHeightsLen = len(rHeights)
                    h1 = rHeights[1]
                    if isNumber(h1) and h1 > 0  tOpts[:headerHeight] = h1  ok
                    # F2 (v3m): pass full per-row height list (cm)
                    hasNonZeroRow = false
                    for rhi2 = 2 to rHeightsLen
                        rh2 = rHeights[rhi2]
                        if isNumber(rh2) and rh2 > 0  hasNonZeroRow = true  ok
                    next
                    if hasNonZeroRow
                        rowHCmList = []
                        for rhi3 = 1 to rHeightsLen
                            rhTwips3 = rHeights[rhi3]
                            if isNumber(rhTwips3) and rhTwips3 > 0
                                rowHCmList + (rhTwips3 / 567.0)
                            else
                                rowHCmList + 0
                            ok
                        next
                        tOpts[:rowHeights] = rowHCmList
                        # hRule from first non-auto row
                        if isList(rHRules) and len(rHRules) > 0
                            for rhr in rHRules
                                if rhr != NULL and len(rhr) > 0 and rhr != "auto"
                                    tOpts[:rowHRule] = rhr
                                    break
                                ok
                            next
                        ok
                    ok
                ok
                if isList(rIsHdr)
                    rIsHdrLen = len(rIsHdr)
                    for rhi = 1 to rIsHdrLen
                        if rIsHdr[rhi] = true  tOpts[:repeatHeader] = true  ok
                    next
                ok
                if len(tableData) > 0
                    writer.addTable(tableData, tOpts)
                ok


            elseif bType = "image_handled_below"
                # image handling moved to end of block loop (handles both inline + floating)

            elseif bType = "hyperlink"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                writer.addHyperlink(block[:text], block[:url])

            elseif bType = "listitem"
                # Accumulate list group: collect this and following list items
                # with the same numId into one addBulletList / addNumberedList call
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                isBulletGroup = block[:listIsBullet]
                curNumId      = block[:listNumId]
                listItems     = []
                # Add current item
                txt = block[:text]
                if txt = NULL  txt = ""  ok
                lvl = block[:listIlvl]
                if lvl = NULL  lvl = 0  ok
                listItems + [txt, lvl]
                # Look ahead for consecutive list items with same numId
                while blkIdx < nBlocks
                    nxtBlk = aBlocks[blkIdx + 1]
                    if nxtBlk[:type] != "listitem"  break  ok
                    if nxtBlk[:listNumId] != curNumId  break  ok
                    blkIdx++
                    ntxt = nxtBlk[:text]
                    if ntxt = NULL  ntxt = ""  ok
                    nlvl = nxtBlk[:listIlvl]
                    if nlvl = NULL  nlvl = 0  ok
                    listItems + [ntxt, nlvl]
                end
                if isBulletGroup = true
                    writer.addNestedBulletList(listItems)
                else
                    sa2 = block[:listStartAt]
                    if isNumber(sa2) and sa2 > 1
                        writer.addNestedNumberedListAt(listItems, sa2)
                    else
                        writer.addNestedNumberedList(listItems)
                    ok
                ok

            elseif bType = "hline"
                writer.addHorizontalLine()

            elseif bType = "pagebreak"
                writer.addPageBreak()

            elseif bType = "blockquote"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                txt = block[:text]
                if txt = NULL  txt = ""  ok
                bqHasBorder = block[:hasBorder]
                if bqHasBorder = true
                    # Bordered paragraph: reproduce exact border/formatting
                    bqOpts = [:_hasBorder=true]
                    bqStyle = block[:borderStyle]
                    bqColor = block[:borderColor]
                    bqSize  = block[:borderSize]
                    bqSpace = block[:borderSpace]
                    bqSides = block[:borderSides]
                    bqAlign = block[:align]
                    bqIndL  = block[:indentLeft]
                    bqIndR  = block[:indentRight]
                    bqSpB   = block[:spaceBefore]
                    bqSpA   = block[:spaceAfter]
                    bqBg    = block[:bgColor]
                    if bqStyle != NULL and len(bqStyle) > 0  bqOpts[:borderStyle] = bqStyle  ok
                    if bqColor != NULL and len(bqColor) > 0  bqOpts[:borderColor] = bqColor  ok
                    if isNumber(bqSize) and bqSize > 0       bqOpts[:borderSize]  = bqSize   ok
                    if isNumber(bqSpace) and bqSpace > 0     bqOpts[:borderSpace] = bqSpace  ok
                    if isList(bqSides) and len(bqSides) > 0  bqOpts[:sides]       = bqSides  ok
                    if bqAlign != NULL and len(bqAlign) > 0  bqOpts[:align]       = bqAlign  ok
                    if isNumber(bqIndL) and bqIndL > 0       bqOpts[:indent]      = bqIndL   ok
                    if isNumber(bqIndR) and bqIndR > 0       bqOpts[:indentRight] = bqIndR   ok
                    if isNumber(bqSpB)  and bqSpB  > 0       bqOpts[:spaceBefore] = bqSpB    ok
                    if isNumber(bqSpA)  and bqSpA  > 0       bqOpts[:spaceAfter]  = bqSpA    ok
                    if bqBg != NULL and len(bqBg) > 0        bqOpts[:bgColor]     = bqBg     ok
                    # Replay runs if present
                    bqRuns = block[:runs]
                    if isList(bqRuns) and len(bqRuns) > 1
                        bqRunList = []
                        for bqRun in bqRuns
                            if !isList(bqRun)  loop  ok
                            bqROpts = []
                            if bqRun[:bold]      = true  bqROpts[:bold]      = true  ok
                            if bqRun[:italic]    = true  bqROpts[:italic]    = true  ok
                            if bqRun[:underline] = true  bqROpts[:underline] = true  ok
                            if len(bqRun[:color]+"") > 0  bqROpts[:color] = bqRun[:color]  ok
                            if len(bqRun[:font] +"") > 0  bqROpts[:font]  = bqRun[:font]   ok
                            if bqRun[:size]      > 0     bqROpts[:size]  = bqRun[:size]    ok
                            bqRunPair = [:text=bqRun[:text], :bold=bqROpts[:bold],
                                          :italic=bqROpts[:italic], :color=bqROpts[:color],
                                          :underline=bqROpts[:underline]]
                            bqRunList + bqRunPair
                        next
                        writer.addBorderedParagraph(bqRunList, bqOpts)
                    else
                        writer.addBorderedParagraph(txt, bqOpts)
                    ok
                else
                    writer.addBlockQuote(txt)
                ok

            elseif bType = "footnote"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                txt = block[:text]
                nt  = block[:noteText]
                nr  = block[:noteRuns]
                if txt = NULL  txt = ""  ok
                if nt  = NULL  nt  = ""  ok
                # replay rich note runs if available
                if isList(nr) and len(nr) > 1
                    noteRunList = []
                    for nRun in nr
                        nROpts = []
                        if nRun[:bold]      = true  nROpts[:bold]      = true  ok
                        if nRun[:italic]    = true  nROpts[:italic]    = true  ok
                        if nRun[:underline] = true  nROpts[:underline] = true  ok
                        if len(nRun[:color]) > 0    nROpts[:color]     = nRun[:color]  ok
                        if len(nRun[:font])  > 0    nROpts[:font]      = nRun[:font]   ok
                        if nRun[:size] > 0          nROpts[:size]      = nRun[:size]   ok
                        noteRunList + [nRun[:text], nROpts]
                    next
                    writer.addFootnote(txt, noteRunList, [])
                else
                    writer.addFootnote(txt, nt, [])
                ok

            elseif bType = "endnote"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                txt = block[:text]
                nt  = block[:noteText]
                nr  = block[:noteRuns]
                if txt = NULL  txt = ""  ok
                if nt  = NULL  nt  = ""  ok
                # replay rich note runs if available
                if isList(nr) and len(nr) > 1
                    noteRunList = []
                    for nRun in nr
                        nROpts = []
                        if nRun[:bold]      = true  nROpts[:bold]      = true  ok
                        if nRun[:italic]    = true  nROpts[:italic]    = true  ok
                        if nRun[:underline] = true  nROpts[:underline] = true  ok
                        if len(nRun[:color]) > 0    nROpts[:color]     = nRun[:color]  ok
                        if len(nRun[:font])  > 0    nROpts[:font]      = nRun[:font]   ok
                        if nRun[:size] > 0          nROpts[:size]      = nRun[:size]   ok
                        noteRunList + [nRun[:text], nROpts]
                    next
                    writer.addEndnote(txt, noteRunList, [])
                else
                    writer.addEndnote(txt, nt, [])
                ok

            elseif bType = "sectionbreak"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                bt2 = block[:breakType]
                if bt2 = NULL  bt2 = "nextPage"  ok
                # Look ahead: if a landscapestart follows, use addLandscapeStart()
                # so the writer correctly inserts portrait sectPr before landscape
                nextIsLandscape = false
                for sbLook = blkIdx+1 to len(aBlocks)
                    sbLookType = aBlocks[sbLook][:type]
                    if sbLookType = "landscapestart"
                        nextIsLandscape = true
                        exit
                    elseif sbLookType = "sectionbreak" or sbLookType = "pagebreak"
                        exit
                    ok
                next
                if nextIsLandscape
                    writer.addLandscapeStart()
                else
                    # pass column count if > 1
                    sbNC2 = block[:numColumns]
                    sbCS2 = block[:columnSpaceTwips]
                    if isNumber(sbNC2) and sbNC2 > 1
                        sbOpts2 = []
                        sbOpts2[:numColumns] = sbNC2
                        if isNumber(sbCS2) and sbCS2 > 0
                            sbOpts2[:columnSpaceCm] = sbCS2 / 567.0
                        ok
                        writer.addSectionBreak(bt2, sbOpts2)
                    else
                        writer.addSectionBreak(bt2, [])
                    ok
                ok

            elseif bType = "caption"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                txt = block[:text]
                if txt = NULL  txt = ""  ok
                writer.addCaption(txt)

            elseif bType = "tabbed"
                bm = block[:bookmark]
                if bm != NULL and len(bm) > 0  writer.addBookmark(bm)  ok
                segs = block[:segments]
                tabs = block[:tabStops]
                al2  = block[:align]
                tOpts2 = []
                if isList(tabs) and len(tabs) > 0  tOpts2[:tabStops] = tabs  ok
                if al2 != NULL and len(al2) > 0    tOpts2[:align]    = al2   ok
                if isList(segs) and len(segs) > 0
                    writer.addTabbedParagraph(segs, tOpts2)
                else
                    txt = block[:text]
                    if txt = NULL  txt = ""  ok
                    writer.addParagraph(txt, tOpts2)
                ok

            elseif bType = "landscapestart"
                # landscapestart block = sectPr{landscape} in source
                # = end of landscape section -> call addLandscapeEnd()
                writer.addLandscapeEnd()

            elseif bType = "landscapeend"
                writer.addLandscapeEnd()

            elseif bType = "columnbreak"
                writer.addColumnBreak()

            elseif bType = "shape"
                # Round-trip DrawingML shape using addShape()
                shpType  = block[:shapeType]
                shpW     = block[:widthCm]
                shpH     = block[:heightCm]
                shpFill  = block[:fillColor]
                shpNoFill= block[:noFill]
                shpLnClr = block[:lineColor]
                shpLnW   = block[:lineWidthPt]
                shpNoBdr = block[:noBorder]
                shpTxt   = block[:text]
                shpTxtClr= block[:textColor]
                shpTxtB  = block[:textBold]
                shpTxtSz = block[:textSize]
                shpAlign = block[:align]
                if shpType  = NULL   shpType  = "rect"  ok
                if shpW     = NULL   shpW     = 5.0     ok
                if shpH     = NULL   shpH     = 3.0     ok
                if shpFill  = NULL   shpFill  = "4472C4" ok
                if shpNoFill= NULL   shpNoFill= false    ok
                if shpLnClr = NULL   shpLnClr = "2E4E7E" ok
                if shpLnW   = NULL   shpLnW   = 1.0     ok
                if shpNoBdr = NULL   shpNoBdr = false    ok
                if shpTxt   = NULL   shpTxt   = ""      ok
                if shpTxtClr= NULL   shpTxtClr= "FFFFFF" ok
                if shpTxtB  = NULL   shpTxtB  = true    ok
                if shpTxtSz = NULL   shpTxtSz = 11      ok
                if shpAlign = NULL   shpAlign = "center" ok
                # Map prstGeom shape type to writer API type
                apiShpType = shpType
                if shpType = "roundRect"   apiShpType = "rect"  ok
                if shpType = "rtTriangle"  apiShpType = "triangle"  ok
                if shpType = "triangle"    apiShpType = "triangle"  ok
                shpOpts = []
                shpOpts[:width]     = shpW
                shpOpts[:height]    = shpH
                shpOpts[:align]     = shpAlign
                if shpNoFill = true
                    shpOpts[:noFill] = true
                else
                    shpOpts[:fillColor] = shpFill
                ok
                if shpNoBdr = true
                    shpOpts[:noBorder] = true
                else
                    shpOpts[:lineColor]  = shpLnClr
                    shpOpts[:lineWidth]  = shpLnW
                ok
                if shpType = "roundRect"  shpOpts[:rounded] = true  ok
                if len(shpTxt) > 0
                    shpOpts[:text]      = shpTxt
                    shpOpts[:textColor] = shpTxtClr
                    shpOpts[:textBold]  = shpTxtB
                    shpOpts[:textSize]  = shpTxtSz
                ok
                writer.addShape(apiShpType, shpOpts)

            elseif bType = "chart"
                # Round-trip chart using native addChart()
                ctRaw  = block[:chartType]
                if ctRaw = NULL  ctRaw = "barChart"  ok
                cw = block[:widthCm]
                ch = block[:heightCm]
                if !isNumber(cw) or cw <= 0  cw = 14  ok
                if !isNumber(ch) or ch <= 0  ch = 9   ok
                ctTitle       = block[:chartTitle]
                ctSeries      = block[:chartData]
                ctGrouping    = block[:chartGrouping]
                ctBarDir      = block[:chartBarDir]
                ctSmooth      = block[:chartSmooth]
                ctLegendPos   = block[:chartLegendPos]
                ctShowLabels  = block[:chartShowLabels]
                ctSerColors   = block[:chartSerColors]
                if ctTitle      = NULL  ctTitle      = ""     ok
                if ctSeries     = NULL  ctSeries     = []     ok
                if ctGrouping   = NULL  ctGrouping   = ""     ok
                if ctBarDir     = NULL  ctBarDir     = ""     ok
                if ctSmooth     = NULL  ctSmooth     = false  ok
                if ctLegendPos  = NULL  ctLegendPos  = "r"    ok
                if ctShowLabels = NULL  ctShowLabels = false  ok
                if !isList(ctSerColors)  ctSerColors = []  ok
                if len(ctLegendPos) = 0  ctLegendPos = "r"  ok

                # Map chartType to addChart type string
                # barChart with barDir=col -> column; barDir=bar -> bar
                ctApiType = "column"
                if ctRaw = "barChart"
                    if ctBarDir = "bar"  ctApiType = "bar"
                    else  ctApiType = "column"  ok
                elseif ctRaw = "lineChart"     ctApiType = "line"
                elseif ctRaw = "areaChart"     ctApiType = "area"
                elseif ctRaw = "pieChart"      ctApiType = "pie"
                elseif ctRaw = "doughnutChart" ctApiType = "doughnut"
                elseif ctRaw = "scatterChart"  ctApiType = "scatter"
                elseif ctRaw = "bubbleChart"   ctApiType = "bubble"
                else  ctApiType = "column"  ok
                ctScatterStyle = block[:chartScatterStyle]
                ctBubble3D     = block[:chartBubble3D]
                if ctScatterStyle = NULL  ctScatterStyle = "marker"  ok
                if ctBubble3D     = NULL  ctBubble3D     = false     ok

                # Build categories from first series (all share same cats)
                ctCats = []
                if isList(ctSeries) and len(ctSeries) > 0
                    firstSer = ctSeries[1]
                    if isList(firstSer)
                        cats2 = firstSer[:categories]
                        if isList(cats2)  ctCats = cats2  ok
                    ok
                ok

                # Build series list for addChart
                ctSerList = []
                serIdx2 = 0
                if isList(ctSeries)
                    for cser2 in ctSeries
                        serIdx2++
                        if isList(cser2)
                            sEntry = [:name=cser2[:name], :values=cser2[:values]]
                            # For scatter/bubble: include xValues, yValues, sizes
                            cXVals = cser2[:xValues]
                            cYVals = cser2[:yValues]
                            cSizes = cser2[:sizes]
                            cMrkSt = cser2[:markerStyle]
                            cMrkSz = cser2[:markerSize]
                            if isList(cXVals) and len(cXVals) > 0
                                sEntry[:xValues] = cXVals
                            ok
                            if isList(cYVals) and len(cYVals) > 0
                                sEntry[:yValues] = cYVals
                            ok
                            if isList(cSizes) and len(cSizes) > 0
                                sEntry[:sizes] = cSizes
                            ok
                            if cMrkSt != NULL and len(cMrkSt) > 0
                                sEntry[:markerStyle] = cMrkSt
                            ok
                            if isNumber(cMrkSz) and cMrkSz > 0
                                sEntry[:markerSize] = cMrkSz
                            ok
                            if serIdx2 <= len(ctSerColors)
                                sc2 = ctSerColors[serIdx2]
                                if sc2 != NULL and len(sc2) > 0
                                    sEntry[:color] = sc2
                                ok
                            ok
                            ctSerList + sEntry
                        ok
                    next
                ok

                # Build chart options
                ctOpts = [:widthCm=cw, :heightCm=ch,
                          :centered=true, :legendPos=ctLegendPos,
                          :showDataLabels=ctShowLabels]
                if len(ctGrouping) > 0    ctOpts[:grouping]      = ctGrouping     ok
                if ctSmooth = true        ctOpts[:smooth]        = true           ok
                if len(ctScatterStyle) > 0 and ctScatterStyle != "marker"
                    ctOpts[:scatterStyle] = ctScatterStyle
                ok
                if ctBubble3D = true      ctOpts[:bubble3D]      = true           ok

                writer.addChart(ctApiType, ctTitle, ctCats, ctSerList, ctOpts)

            elseif bType = "rawparagraph"
                # Preserve REF/PAGEREF field paragraphs verbatim
                rawXml2 = block[:rawXml]
                if rawXml2 != NULL and len(rawXml2) > 0
                    writer.addRawParagraph(rawXml2)
                ok

            elseif bType = "toc"
                tocTitle = block[:title]
                if tocTitle = NULL  tocTitle = "Table of Contents"  ok
                writer.addTableOfContents(tocTitle)

            elseif bType = "formfield"
                ft  = block[:fieldType]
                lbl = block[:label]
                if ft = NULL   ft  = ""  ok
                if lbl = NULL  lbl = ""  ok
                if ft = "checkbox"
                    chk = block[:checked]
                    if chk = NULL  chk = false  ok
                    writer.addCheckbox(lbl, chk)
                elseif ft = "dropdown"
                    chs = block[:choices]
                    dflt = block[:default]
                    if !isList(chs)  chs = []  ok
                    if dflt = NULL  dflt = ""  ok
                    if len(chs) = 0  chs = [dflt]  ok
                    # Pass "" label: preceding paragraph already has the label
                    writer.addDropdown("", chs, dflt)
                elseif ft = "text"
                    dflt2 = block[:default]
                    ph2   = block[:placeholder]
                    if dflt2 = NULL  dflt2 = ""  ok
                    if ph2   = NULL  ph2   = ""  ok
                    # Pass "" label: preceding paragraph already has the label
                    writer.addTextInput("", dflt2, ph2)
                else
                    writer.addParagraph("[Form field: " + ft + "] " + lbl, [:italic=true, :color="888888"])
                ok

            elseif bType = "mergefield"
                # Use stored :runs (mixed literal strings + [:field="Name"] dicts)
                mfRuns3 = block[:runs]
                if !isList(mfRuns3) or len(mfRuns3) = 0
                    # Fallback: use field names only
                    mfRuns3 = []
                    mfFlds3 = block[:fields]
                    if !isList(mfFlds3)  mfFlds3 = ["Field"]  ok
                    for mff3 in mfFlds3
                        add(mfRuns3, [:field = mff3])
                    next
                ok
                writer.addMergeFieldParagraph(mfRuns3, [])

            elseif bType = "textbox"
                # Replay text box using writer.addTextBox
                tbTxt = block[:text]
                tbX   = block[:x]
                tbY   = block[:y]
                tbW   = block[:width]
                tbH   = block[:height]
                if tbTxt = NULL  tbTxt = ""  ok
                if !isNumber(tbX) or tbX <= 0  tbX = 2.0  ok
                if !isNumber(tbY) or tbY <= 0  tbY = 5.0  ok
                if !isNumber(tbW) or tbW <= 0  tbW = 6.0  ok
                if !isNumber(tbH) or tbH <= 0  tbH = 3.0  ok
                writer.addTextBox(tbTxt, [:x=tbX, :y=tbY, :width=tbW, :height=tbH])

            elseif bType = "image"
                bm3 = block[:bookmark]
                if bm3 != NULL and len(bm3) > 0  writer.addBookmark(bm3)  ok
                imgPath3 = block[:path]
                if imgPath3 != NULL and len(imgPath3) > 0 and fexists(imgPath3)
                    w3 = block[:widthCm]
                    h3 = block[:heightCm]
                    if !isNumber(w3) or w3 <= 0  w3 = 10   ok
                    if !isNumber(h3) or h3 <= 0  h3 = 7.5  ok
                    alt3 = block[:altText]
                    if alt3 = NULL  alt3 = ""  ok
                    # F2 (v3n): image crop round-trip
                    iOpts3 = []
                    if len(alt3) > 0  iOpts3[:altText] = alt3  ok
                    al3 = block[:align]
                    if al3 != NULL and len(al3) > 0  iOpts3[:align] = al3  ok
                    if block[:runBold]   = true  iOpts3[:bold]      = true  ok
                    if block[:runItalic] = true  iOpts3[:italic]    = true  ok
                    if block[:runUnder]  = true  iOpts3[:underline] = true  ok
                    irColor = block[:runColor]
                    irSize  = block[:runSize]
                    if irColor != NULL and len(irColor) > 0  iOpts3[:color] = irColor  ok
                    if isNumber(irSize) and irSize > 0  iOpts3[:size] = irSize  ok
                    icL = block[:cropL]; icR = block[:cropR]
                    icT = block[:cropT]; icB = block[:cropB]
                    if isNumber(icL) and icL > 0  iOpts3[:cropLeft]   = icL  ok
                    if isNumber(icR) and icR > 0  iOpts3[:cropRight]  = icR  ok
                    if isNumber(icT) and icT > 0  iOpts3[:cropTop]    = icT  ok
                    if isNumber(icB) and icB > 0  iOpts3[:cropBottom] = icB  ok
                    isF3 = block[:floating]
                    if isF3 = true
                        # Replay as floating image with original wrap/position
                        fOpts3 = []
                        fOpts3[:wrapType] = block[:wrapType]
                        fOpts3[:wrapSide] = block[:wrapSide]
                        fOpts3[:distCm]   = block[:distL]
                        fOpts3[:posX]     = block[:x]
                        fOpts3[:posY]     = block[:y]
                        if len(alt3) > 0  fOpts3[:altText] = alt3  ok
                        if isNumber(icL) and icL > 0  fOpts3[:cropLeft]   = icL  ok
                        if isNumber(icR) and icR > 0  fOpts3[:cropRight]  = icR  ok
                        if isNumber(icT) and icT > 0  fOpts3[:cropTop]    = icT  ok
                        if isNumber(icB) and icB > 0  fOpts3[:cropBottom] = icB  ok
                        writer.addFloatingImage(imgPath3, w3, h3, fOpts3)
                    else
                        writer.addImageWithOptions(imgPath3, w3, h3, iOpts3)
                    ok
                ok

            elseif bType = "field"
                # F2 round-trip: replay auto-updating document fields
                ft3 = block[:fieldType]
                fr3 = block[:fieldResult]
                if ft3 = NULL  ft3 = ""  ok
                if fr3 = NULL  fr3 = ""  ok
                if len(ft3) > 0
                    writer.addField(ft3, fr3, [])
                ok
            ok
        end

        return writer

    func save outputPath
        wrSaveWriter(toWriter(), outputPath)
        cleanup()
        return self

    func cleanup
        /*  Delete the temporary extraction directory used during parsing.
            Called automatically after save(). Can also be called manually
            when you are done querying and no longer need image file access.
            Uses OSDeleteFolder() + DirExists() from stdlibcore.ring.
            IMPORTANT: fexists() is files-only on Windows and always returns
            false for directories, so we must use DirExists() here.
            Both OSDeleteFolder and DirExists expect a bare path WITHOUT a
            trailing separator, so we strip it first.  */
        if len(cTempDir) > 0
            cDir = cTempDir
            cLast = substr(cDir, len(cDir), 1)
            if cLast = "/" or cLast = "\"
                cDir = left(cDir, len(cDir) - 1)
            ok
            if len(cDir) > 0 and DirExists(cDir)
                OSDeleteFolder(cDir)
            ok
            cTempDir = ""
        ok
        return self

