/*
    DOCXLib  
*/

# ============================================================================
# Global helper functions
# ============================================================================

func wrFindFrom str, sub, fromPos
    /*  Find sub in str starting at fromPos (1-based). Returns position or 0.  */
    if fromPos < 1  fromPos = 1  ok
    sLen = len(str)
    subLen = len(sub)
    if subLen = 0 or sLen = 0  return 0  ok
    pos = fromPos
    while pos <= sLen - subLen + 1
        if substr(str, pos, subLen) = sub
            return pos
        ok
        pos++
    end
    return 0

func wrAttr xml, attr
    /*  Extract the value of attr="..." or attr='...' from an XML element string.  */
    key1 = attr + '="'
    key2 = attr + "='"
    p = wrFindFrom(xml, key1, 1)
    if p > 0
        vs = p + len(key1)
        ve = wrFindFrom(xml, '"', vs)
        if ve > 0  return substr(xml, vs, ve - vs)  ok
    ok
    p = wrFindFrom(xml, key2, 1)
    if p > 0
        vs = p + len(key2)
        ve = wrFindFrom(xml, "'", vs)
        if ve > 0  return substr(xml, vs, ve - vs)  ok
    ok
    return ""

func wrFindCloseTag xml, tagName, startPos
    /*
        Return the position AFTER the matching close tag for the element
        whose open tag begins at startPos. Handles nesting.
        depth starts at 1 and pos skips past the opening tag to avoid
        counting it again.
    */
    openStr1 = "<" + tagName + ">"
    openStr2 = "<" + tagName + " "
    closeStr = "</" + tagName + ">"
    depth = 1
    pos   = startPos + 1
    while true
        nextOpen1 = wrFindFrom(xml, openStr1, pos)
        nextOpen2 = wrFindFrom(xml, openStr2, pos)
        if nextOpen1 = 0
            nextOpen = nextOpen2
        elseif nextOpen2 = 0
            nextOpen = nextOpen1
        elseif nextOpen1 < nextOpen2
            nextOpen = nextOpen1
        else
            nextOpen = nextOpen2
        ok
        nextClose = wrFindFrom(xml, closeStr, pos)
        if nextClose = 0  return 0  ok
        if nextOpen > 0 and nextOpen < nextClose
            depth++
            pos = nextOpen + 1
        else
            depth--
            if depth = 0
                return nextClose + len(closeStr)
            ok
            pos = nextClose + len(closeStr)
        ok
    end

func wrXmlUnescape s
    s = substr(s, "&amp;",  "&")
    s = substr(s, "&lt;",   "<")
    s = substr(s, "&gt;",   ">")
    s = substr(s, "&quot;", '"')
    s = substr(s, "&apos;", "'")
    return s

func wrMinPos a, b
    if a = 0  return b  ok
    if b = 0  return a  ok
    if a < b  return a  ok
    return b


func wrReadU16 data, pos
    b1 = ascii(substr(data, pos,   1))
    b2 = ascii(substr(data, pos+1, 1))
    return b1 + b2 * 256

func wrReadU32 data, pos
    b1 = ascii(substr(data, pos,   1))
    b2 = ascii(substr(data, pos+1, 1))
    b3 = ascii(substr(data, pos+2, 1))
    b4 = ascii(substr(data, pos+3, 1))
    return b1 + b2*256 + b3*65536 + b4*16777216

func wrExtractZip wriez_zipPath, wriez_destDir
    /*
        Extract all entries from a ZIP file into destDir.
        Handles STORE (method=0) via pure-Ring code.
        Handles DEFLATE (method=8) via pure-Ring wrInflateData().
        No external tools or OS calls needed - works everywhere.
    */
    wriez_data    = read(wriez_zipPath)
    wriez_dataLen = len(wriez_data)
    wriez_sep     = wordGetSep()
    if right(wriez_destDir, 1) != wriez_sep and right(wriez_destDir, 1) != "/"
        wriez_destDir += wriez_sep
    ok

    wriez_pos = 1
    while wriez_pos <= wriez_dataLen - 4
        wriez_b1 = ascii(substr(wriez_data, wriez_pos,   1))
        wriez_b2 = ascii(substr(wriez_data, wriez_pos+1, 1))
        wriez_b3 = ascii(substr(wriez_data, wriez_pos+2, 1))
        wriez_b4 = ascii(substr(wriez_data, wriez_pos+3, 1))
        if wriez_b1 = 80 and wriez_b2 = 75 and wriez_b3 = 3 and wriez_b4 = 4
            wriez_method   = wrReadU16(wriez_data, wriez_pos + 8)
            wriez_compSize = wrReadU32(wriez_data, wriez_pos + 18)
            wriez_fnLen    = wrReadU16(wriez_data, wriez_pos + 26)
            wriez_extraLen = wrReadU16(wriez_data, wriez_pos + 28)
            wriez_entryName = substr(wriez_data, wriez_pos + 30, wriez_fnLen)
            wriez_dataStart = wriez_pos + 30 + wriez_fnLen + wriez_extraLen

            if wriez_method = 0 or wriez_method = 8
                if wriez_method = 0
                    wriez_fileData = substr(wriez_data, wriez_dataStart, wriez_compSize)
                else
                    wriez_fileData = wrInflateData(substr(wriez_data, wriez_dataStart, wriez_compSize))
                ok

                if len(wriez_fileData) > 0
                    wriez_outPath = wriez_destDir + wriez_entryName
                    if wriez_sep = char(92)
                        wriez_outPath = substr(wriez_outPath, "/", wriez_sep)
                    ok
                    wriez_lastSep = 0
                    for wriez_ci = len(wriez_outPath) to 1 step -1
                        if substr(wriez_outPath, wriez_ci, 1) = wriez_sep or substr(wriez_outPath, wriez_ci, 1) = "/"
                            wriez_lastSep = wriez_ci
                            break
                        ok
                    next
                    if wriez_lastSep > 0
                        wordMakeDir(substr(wriez_outPath, 1, wriez_lastSep - 1))
                    ok
                    write(wriez_outPath, wriez_fileData)
                ok
            ok
            wriez_pos = wriez_dataStart + wriez_compSize
        else
            wriez_pos++
        ok
    end

/*
    wrInflateData - Pure-Ring DEFLATE (RFC 1951) decompressor
    For Ring language (which uses dynamic/shared variable scope).
    ALL internal variables are prefixed wri_ to avoid caller conflicts.
    
    Usage:
        result = wrInflateData(compressedString)
        Returns decompressed string, or "" on error.
*/

func wrInflateData wri_cData
    wri_nDataLen = len(wri_cData)
    # State list: [bitbuf, bitcnt, bytepos, output]
    wri_st = [0, 0, 1, ""]
    wri_bFinal = false
    while not wri_bFinal
        # Read BFINAL (1 bit)
        wrInfNeed(wri_st, wri_cData, wri_nDataLen, 1)
        wri_bFinal = (wrInfPeek(wri_st, 1) = 1)
        wrInfDrop(wri_st, 1)
        # Read BTYPE (2 bits)
        wrInfNeed(wri_st, wri_cData, wri_nDataLen, 2)
        wri_nType = wrInfPeek(wri_st, 2)
        wrInfDrop(wri_st, 2)
        if wri_nType = 0
            wrInfStored(wri_st, wri_cData, wri_nDataLen)
        elseif wri_nType = 1
            wrInfFixed(wri_st, wri_cData, wri_nDataLen)
        elseif wri_nType = 2
            wrInfDynamic(wri_st, wri_cData, wri_nDataLen)
        else
            return ""
        ok
        if len(wri_st[4]) > 67108864  return ""  ok
    end
    return wri_st[4]

func wrInfNeed wri_st, wri_cData, wri_nDataLen, wri_n
    while wri_st[2] < wri_n
        if wri_st[3] > wri_nDataLen  return  ok
        wri_st[1] = wri_st[1] | (ascii(substr(wri_cData, wri_st[3], 1)) << wri_st[2])
        wri_st[3]++
        wri_st[2] += 8
    end

func wrInfPeek wri_st, wri_n
    return wri_st[1] & ((1 << wri_n) - 1)

func wrInfDrop wri_st, wri_n
    wri_st[1] = wri_st[1] >> wri_n
    wri_st[2] -= wri_n

func wrInfBits wri_st, wri_cData, wri_nDataLen, wri_n
    wrInfNeed(wri_st, wri_cData, wri_nDataLen, wri_n)
    wri_v = wrInfPeek(wri_st, wri_n)
    wrInfDrop(wri_st, wri_n)
    return wri_v

func wrInfStored wri_st, wri_cData, wri_nDataLen
    wri_st[1] = 0
    wri_st[2] = 0
    if wri_st[3] + 3 > wri_nDataLen  return  ok
    wri_nLEN = ascii(substr(wri_cData, wri_st[3], 1)) | (ascii(substr(wri_cData, wri_st[3]+1, 1)) << 8)
    wri_st[3] += 4
    if wri_st[3] + wri_nLEN - 1 > wri_nDataLen  return  ok
    wri_st[4] += substr(wri_cData, wri_st[3], wri_nLEN)
    wri_st[3] += wri_nLEN

func wrBuildHuff wri_aLens, wri_nSyms
    wri_aCount = list(16)
    for wri_i = 1 to 16  wri_aCount[wri_i] = 0  next
    for wri_i = 1 to wri_nSyms
        if wri_aLens[wri_i] > 0  wri_aCount[wri_aLens[wri_i]]++  ok
    next
    wri_aNext = list(16)
    wri_code = 0
    wri_aNext[1] = 0
    for wri_i = 2 to 15
        wri_code = (wri_code + wri_aCount[wri_i-1]) << 1
        wri_aNext[wri_i] = wri_code
    next
    wri_aTable = []
    for wri_sym = 1 to wri_nSyms
        wri_l = wri_aLens[wri_sym]
        if wri_l > 0
            wri_entry = [wri_aNext[wri_l], wri_l, wri_sym - 1]
            wri_aTable + wri_entry
            wri_aNext[wri_l]++
        ok
    next
    return wri_aTable

func wrHuffDecode wri_st, wri_cData, wri_nDataLen, wri_aTable
    wri_nCode = 0
    wri_nBits = 0
    for wri_b = 1 to 15
        wrInfNeed(wri_st, wri_cData, wri_nDataLen, 1)
        wri_nCode = (wri_nCode << 1) | wrInfPeek(wri_st, 1)
        wrInfDrop(wri_st, 1)
        wri_nBits++
        for wri_i = 1 to len(wri_aTable)
            wri_e = wri_aTable[wri_i]
            if wri_e[2] = wri_nBits and wri_e[1] = wri_nCode
                return wri_e[3]
            ok
        next
    next
    return -1

func wrFixedLitLens
    wri_aLens = list(288)
    for wri_i = 1   to 144  wri_aLens[wri_i] = 8  next
    for wri_i = 145 to 256  wri_aLens[wri_i] = 9  next
    for wri_i = 257 to 280  wri_aLens[wri_i] = 7  next
    for wri_i = 281 to 288  wri_aLens[wri_i] = 8  next
    return wri_aLens

func wrFixedDistLens
    wri_aLens = list(32)
    for wri_i = 1 to 32  wri_aLens[wri_i] = 5  next
    return wri_aLens

func wrInfFixed wri_st, wri_cData, wri_nDataLen
    wri_aLitLens  = wrFixedLitLens()
    wri_aDistLens = wrFixedDistLens()
    wri_aLitTab   = wrBuildHuff(wri_aLitLens,  288)
    wri_aDistTab  = wrBuildHuff(wri_aDistLens, 32)
    wrInfCodes(wri_st, wri_cData, wri_nDataLen, wri_aLitTab, wri_aDistTab)

func wrInfCodes wri_st, wri_cData, wri_nDataLen, wri_aLitTab, wri_aDistTab
    wri_aLenBase  = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258]
    wri_aLenExtra = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0]
    wri_aDistBase = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577]
    wri_aDistExtra= [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13]
    while true
        wri_sym = wrHuffDecode(wri_st, wri_cData, wri_nDataLen, wri_aLitTab)
        if wri_sym < 0  return  ok
        if wri_sym < 256
            wri_st[4] += char(wri_sym)
        elseif wri_sym = 256
            return
        else
            wri_lenIdx = wri_sym - 256  # 1-based: sym=257->idx=1
            if wri_lenIdx < 1 or wri_lenIdx > 29  return  ok
            wri_nLen = wri_aLenBase[wri_lenIdx]
            wri_nEx  = wri_aLenExtra[wri_lenIdx]
            if wri_nEx > 0
                wri_nLen += wrInfBits(wri_st, wri_cData, wri_nDataLen, wri_nEx)
            ok
            wri_distSym = wrHuffDecode(wri_st, wri_cData, wri_nDataLen, wri_aDistTab)
            if wri_distSym < 0 or wri_distSym > 29  return  ok
            wri_nDist = wri_aDistBase[wri_distSym + 1]
            wri_nDex  = wri_aDistExtra[wri_distSym + 1]
            if wri_nDex > 0
                wri_nDist += wrInfBits(wri_st, wri_cData, wri_nDataLen, wri_nDex)
            ok
            wri_nOutLen = len(wri_st[4])
            wri_nStart  = wri_nOutLen - wri_nDist + 1
            for wri_k = 1 to wri_nLen
                wri_pos = wri_nStart + wri_k - 1
                if wri_pos < 1
                    wri_st[4] += char(0)
                else
                    wri_st[4] += substr(wri_st[4], wri_pos, 1)
                ok
            next
        ok
    end

func wrInfDynamic wri_st, wri_cData, wri_nDataLen
    wri_nHLIT  = wrInfBits(wri_st, wri_cData, wri_nDataLen, 5) + 257
    wri_nHDIST = wrInfBits(wri_st, wri_cData, wri_nDataLen, 5) + 1
    wri_nHCLEN = wrInfBits(wri_st, wri_cData, wri_nDataLen, 4) + 4
    wri_aOrder = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15]
    wri_aCLLens = list(19)
    for wri_i = 1 to 19  wri_aCLLens[wri_i] = 0  next
    for wri_i = 1 to wri_nHCLEN
        wri_aCLLens[wri_aOrder[wri_i] + 1] = wrInfBits(wri_st, wri_cData, wri_nDataLen, 3)
    next
    wri_aCLTab = wrBuildHuff(wri_aCLLens, 19)
    wri_nTotal = wri_nHLIT + wri_nHDIST
    wri_aAllLens = list(wri_nTotal)
    for wri_i = 1 to wri_nTotal  wri_aAllLens[wri_i] = 0  next
    wri_i = 1
    while wri_i <= wri_nTotal
        wri_sym = wrHuffDecode(wri_st, wri_cData, wri_nDataLen, wri_aCLTab)
        if wri_sym < 0  return  ok
        if wri_sym <= 15
            wri_aAllLens[wri_i] = wri_sym
            wri_i++
        elseif wri_sym = 16
            if wri_i <= 1  return  ok
            wri_prev = wri_aAllLens[wri_i-1]
            wri_rep = wrInfBits(wri_st, wri_cData, wri_nDataLen, 2) + 3
            for wri_r = 1 to wri_rep
                if wri_i > wri_nTotal  return  ok
                wri_aAllLens[wri_i] = wri_prev
                wri_i++
            next
        elseif wri_sym = 17
            wri_rep = wrInfBits(wri_st, wri_cData, wri_nDataLen, 3) + 3
            for wri_r = 1 to wri_rep
                if wri_i > wri_nTotal  return  ok
                wri_aAllLens[wri_i] = 0
                wri_i++
            next
        elseif wri_sym = 18
            wri_rep = wrInfBits(wri_st, wri_cData, wri_nDataLen, 7) + 11
            for wri_r = 1 to wri_rep
                if wri_i > wri_nTotal  return  ok
                wri_aAllLens[wri_i] = 0
                wri_i++
            next
        else
            return
        ok
    end
    wri_aLitLens  = list(wri_nHLIT)
    wri_aDistLens = list(wri_nHDIST)
    for wri_i = 1 to wri_nHLIT   wri_aLitLens[wri_i]  = wri_aAllLens[wri_i]             next
    for wri_i = 1 to wri_nHDIST  wri_aDistLens[wri_i] = wri_aAllLens[wri_nHLIT + wri_i] next
    wri_aLitTab  = wrBuildHuff(wri_aLitLens,  wri_nHLIT)
    wri_aDistTab = wrBuildHuff(wri_aDistLens, wri_nHDIST)
    wrInfCodes(wri_st, wri_cData, wri_nDataLen, wri_aLitTab, wri_aDistTab)

func wrIsSelfClosingTag wrisc_xml, wrisc_tagStart
    /*
        Return true if the opening tag at wrisc_tagStart is self-closing (ends with "/>").
        Correctly skips over attribute values enclosed in double-quotes.
    */
    wrisc_len    = len(wrisc_xml)
    wrisc_pos    = wrisc_tagStart + 1
    wrisc_inStr  = false
    while wrisc_pos <= wrisc_len
        wrisc_ch = substr(wrisc_xml, wrisc_pos, 1)
        if wrisc_inStr
            if wrisc_ch = char(34)  wrisc_inStr = false  ok
        else
            if wrisc_ch = char(34)
                wrisc_inStr = true
            elseif wrisc_ch = ">"
                if wrisc_pos > 1
                    return substr(wrisc_xml, wrisc_pos - 1, 1) = "/"
                ok
                return false
            ok
        ok
        wrisc_pos++
    end
    return false

func wrSelfClosingEnd wrisc_xml, wrisc_tagStart
    /*
        Return the position immediately after the ">" of a self-closing tag.
        Works for both self-closing (<w:p ... />) and regular opening tags.
    */
    wrisc_len   = len(wrisc_xml)
    wrisc_pos   = wrisc_tagStart + 1
    wrisc_inStr = false
    while wrisc_pos <= wrisc_len
        wrisc_ch = substr(wrisc_xml, wrisc_pos, 1)
        if wrisc_inStr
            if wrisc_ch = char(34)  wrisc_inStr = false  ok
        else
            if wrisc_ch = char(34)
                wrisc_inStr = true
            elseif wrisc_ch = ">"
                return wrisc_pos + 1
            ok
        ok
        wrisc_pos++
    end
    return 0

func wrStripTrackedChanges pXml
    /*
        Pre-process a paragraph XML string for tracked changes:
        - Remove <w:del>...</w:del> blocks entirely (deleted text is discarded)
        - Unwrap <w:ins>...</w:ins> blocks (keep their inner content as normal runs)
        This lets the existing run-parsing loop work unchanged.
    */
    # Remove del blocks
    pos = 1
    result = pXml
    while true
        ds = wrFindFrom(result, "<w:del ", pos)
        if ds = 0  ds = wrFindFrom(result, "<w:del>", pos)  ok
        if ds = 0  break  ok
        de = wrFindCloseTag(result, "w:del", ds)
        if de = 0  break  ok
        result = substr(result, 1, ds - 1) + substr(result, de)
        pos = ds
    end
    # Unwrap ins blocks (keep inner content, remove <w:ins ...> and </w:ins> tags)
    pos = 1
    while true
        is_ = wrFindFrom(result, "<w:ins ", pos)
        if is_ = 0  is_ = wrFindFrom(result, "<w:ins>", pos)  ok
        if is_ = 0  break  ok
        # Find end of opening tag
        igt = wrFindFrom(result, ">", is_)
        if igt = 0  break  ok
        ie = wrFindCloseTag(result, "w:ins", is_)
        if ie = 0  break  ok
        inner = substr(result, igt + 1, ie - igt - 1 - len("</w:ins>"))
        result = substr(result, 1, is_ - 1) + inner + substr(result, ie)
        pos = is_
    end
    return result


func wrParseCharStyles stylesXml
    /*
        Parse all character styles from styles.xml.
        Returns list of [:name, :bold, :italic, :color, :size, :font, :underline].
    */
    result = []
    pos = 1
    while true
        ss = wrFindFrom(stylesXml, "<w:style ", pos)
        if ss = 0  break  ok
        se = wrFindCloseTag(stylesXml, "w:style", ss)
        if se = 0  break  ok
        styleEl = substr(stylesXml, ss, se - ss)
        sType = wrAttr(styleEl, "w:type")
        if sType = "character"
            sId = wrAttr(styleEl, "w:styleId")
            sBold = false; sItalic = false; sUnder = false; sStrike = false
            sDStrike = false; sCaps = false; sSmallCaps = false; sVanish = false
            sColor = ""; sSize = 0; sFont = ""
            rPrS = wrFindFrom(styleEl, "<w:rPr>", 1)
            if rPrS = 0  rPrS = wrFindFrom(styleEl, "<w:rPr ", 1)  ok
            if rPrS > 0
                rPrE = wrFindCloseTag(styleEl, "w:rPr", rPrS)
                if rPrE > 0
                    rPrXml = substr(styleEl, rPrS, rPrE - rPrS)
                    if wrFindFrom(rPrXml, "<w:b/>", 1) > 0  sBold = true  ok
                    if wrFindFrom(rPrXml, "<w:b ",  1) > 0  sBold = true  ok
                    if wrFindFrom(rPrXml, "<w:i/>", 1) > 0  sItalic = true  ok
                    if wrFindFrom(rPrXml, "<w:i ",  1) > 0  sItalic = true  ok
                    if wrFindFrom(rPrXml, "<w:u ",  1) > 0  sUnder = true   ok
                    if wrFindFrom(rPrXml, "<w:strike", 1) > 0  sStrike   = true  ok
                    if wrFindFrom(rPrXml, "<w:dstrike",1) > 0  sDStrike  = true  ok
                    if wrFindFrom(rPrXml, "<w:caps",   1) > 0  sCaps     = true  ok
                    if wrFindFrom(rPrXml, "<w:smallCaps",1)> 0 sSmallCaps= true  ok
                    if wrFindFrom(rPrXml, "<w:vanish", 1) > 0  sVanish   = true  ok
                    cS = wrFindFrom(rPrXml, "<w:color ", 1)
                    if cS > 0
                        cEl = substr(rPrXml, cS, 80)
                        cv = wrAttr(cEl, "w:val")
                        if len(cv) > 0 and cv != "auto"  sColor = cv  ok
                    ok
                    szS = wrFindFrom(rPrXml, "<w:sz ", 1)
                    if szS > 0
                        szEl = substr(rPrXml, szS, 60)
                        sv = wrAttr(szEl, "w:val")
                        if len(sv) > 0  sSize = floor(number(sv) / 2)  ok
                    ok
                    fnS = wrFindFrom(rPrXml, "<w:rFonts ", 1)
                    if fnS > 0
                        fnEl = substr(rPrXml, fnS, 120)
                        fv = wrAttr(fnEl, "w:ascii")
                        if len(fv) > 0  sFont = fv  ok
                    ok
                ok
            ok
            if len(sId) > 0
                result + [:name=sId, :bold=sBold, :italic=sItalic,
                          :underline=sUnder, :strike=sStrike, :dstrike=sDStrike,
                          :caps=sCaps, :smallCaps=sSmallCaps, :vanish=sVanish,
                          :color=sColor, :size=sSize, :font=sFont]
            ok
        ok
        pos = se
    end
    return result

func wrResolveCharStyle styleName, aCharStyles
    /*
        Look up a character style by name and return its formatting as an assoc list.
        Returns [] if not found.
    */
    for cs in aCharStyles
        if cs[:name] = styleName
            return cs
        ok
    next
    return []

func wrDetectPageNumbers hdrXml, ftrXml
    /*
        Detect PAGE / NUMPAGES field codes in header and footer XML.
        Returns [:found, :location, :align].
    */
    result = [:found=false, :location="footer", :align="center"]
    # Check footer first (most common)
    if wrFindFrom(ftrXml, "PAGE", 1) > 0 and wrFindFrom(ftrXml, "instrText", 1) > 0
        result[:found] = true
        result[:location] = "footer"
        # Try to detect alignment
        jcS = wrFindFrom(ftrXml, "<w:jc ", 1)
        if jcS > 0
            jcEl = substr(ftrXml, jcS, 60)
            result[:align] = wrAttr(jcEl, "w:val")
        ok
        return result
    ok
    # Check header
    if wrFindFrom(hdrXml, "PAGE", 1) > 0 and wrFindFrom(hdrXml, "instrText", 1) > 0
        result[:found] = true
        result[:location] = "header"
        jcS = wrFindFrom(hdrXml, "<w:jc ", 1)
        if jcS > 0
            jcEl = substr(hdrXml, jcS, 60)
            result[:align] = wrAttr(jcEl, "w:val")
        ok
        return result
    ok
    return result

func wrParseAnchoredImages bodyXml, aRels, tempDir
    /*
        Find all <wp:anchor> floating images in bodyXml.
        Returns list of [:path, :widthCm, :heightCm, :x, :y].
    */
    result = []
    sep = wordGetSep()
    pos = 1
    while true
        as_ = wrFindFrom(bodyXml, "<wp:anchor", pos)
        if as_ = 0  break  ok
        ae = wrFindCloseTag(bodyXml, "wp:anchor", as_)
        if ae = 0  break  ok
        anchorXml = substr(bodyXml, as_, ae - as_)
        # Size
        extS = wrFindFrom(anchorXml, "<wp:extent ", 1)
        imgW = 10.0; imgH = 7.5
        if extS > 0
            extEl = substr(anchorXml, extS, 100)
            cxv = wrAttr(extEl, "cx")
            cyv = wrAttr(extEl, "cy")
            if len(cxv) > 0  imgW = number(cxv) / 360000.0  ok
            if len(cyv) > 0  imgH = number(cyv) / 360000.0  ok
        ok
        # Position
        posX = 0.0; posY = 0.0
        phS = wrFindFrom(anchorXml, "<wp:posOffset>", 1)
        if phS > 0
            # simplePos or posOffset
            smpS = wrFindFrom(anchorXml, "<wp:simplePos ", 1)
            if smpS > 0
                smpEl = substr(anchorXml, smpS, 100)
                xv = wrAttr(smpEl, "x")
                yv = wrAttr(smpEl, "y")
                if len(xv) > 0  posX = number(xv) / 360000.0  ok
                if len(yv) > 0  posY = number(yv) / 360000.0  ok
            ok
        ok
        # Wrap type
        wrapType = "square"
        wrapSide = "bothSides"
        distT2 = 0.0; distB2 = 0.0; distL2 = 0.0; distR2 = 0.0
        for wrapTag in ["wrapSquare","wrapTight","wrapThrough","wrapTopAndBottom","wrapNone"]
            wtS = wrFindFrom(anchorXml, "<wp:" + wrapTag, 1)
            if wtS > 0
                wrapType = wrapTag
                if wrapTag != "wrapTopAndBottom" and wrapTag != "wrapNone"
                    wtEl = substr(anchorXml, wtS, 120)
                    ws = wrAttr(wtEl, "wrapText")
                    if len(ws) > 0  wrapSide = ws  ok
                ok
                break
            ok
        next
        # Distance from text (EMU -> cm)
        anchorEl = substr(anchorXml, 1, 200)
        dT = wrAttr(anchorEl, "distT"); if len(dT) > 0  distT2 = number(dT)/360000.0  ok
        dB = wrAttr(anchorEl, "distB"); if len(dB) > 0  distB2 = number(dB)/360000.0  ok
        dL = wrAttr(anchorEl, "distL"); if len(dL) > 0  distL2 = number(dL)/360000.0  ok
        dR = wrAttr(anchorEl, "distR"); if len(dR) > 0  distR2 = number(dR)/360000.0  ok

        # Image relId
        blipS = wrFindFrom(anchorXml, "<a:blip ", 1)
        imgPath = ""
        if blipS > 0
            blipEl = substr(anchorXml, blipS, 120)
            relId = wrAttr(blipEl, "r:embed")
            if len(relId) > 0
                for rel in aRels
                    if rel[:id] = relId
                        target = rel[:target]
                        if left(target, 6) = "media/"
                            imgPath = tempDir + "word" + sep + target
                        elseif left(target, 9) = "../media/"
                            imgPath = tempDir + "word" + sep + substr(target, 4)
                        else
                            imgPath = tempDir + "word" + sep + target
                        ok
                    ok
                next
            ok
        ok
        if len(imgPath) > 0
            result + [:path=imgPath, :widthCm=imgW, :heightCm=imgH, :x=posX, :y=posY,
                   :wrapType=wrapType, :wrapSide=wrapSide,
                   :distT=distT2, :distB=distB2, :distL=distL2, :distR=distR2]
        ok
        pos = ae
    end
    return result

func wrParseTextBoxes bodyXml
    /*
        Extract text content from VML text boxes (<w:txbxContent>) in bodyXml.
        Returns list of [:text, :x, :y, :width, :height].
    */
    result = []
    pos = 1
    while true
        ts = wrFindFrom(bodyXml, "<w:txbxContent>", pos)
        if ts = 0  ts = wrFindFrom(bodyXml, "<w:txbxContent ", pos)  ok
        if ts = 0  break  ok
        te = wrFindCloseTag(bodyXml, "w:txbxContent", ts)
        if te = 0  break  ok
        tbXml = substr(bodyXml, ts, te - ts)
        # Extract all text
        txt = ""
        tPos = 1
        while true
            tts1 = wrFindFrom(tbXml, "<w:t>", tPos)
            tts2 = wrFindFrom(tbXml, "<w:t ", tPos)
            tts  = wrMinPos(tts1, tts2)
            if tts = 0  break  ok
            ttc = wrFindFrom(tbXml, ">", tts)
            tte = wrFindFrom(tbXml, "</w:t>", ttc)
            if ttc > 0 and tte > 0
                txt += wrXmlUnescape(substr(tbXml, ttc + 1, tte - ttc - 1))
            ok
            tPos = tte + 1
        end
        if len(txt) > 0
            # Try to get position from nearby shape element (search backwards)
            shapeStart = max(1, ts - 500)
            shapeXml = substr(bodyXml, shapeStart, ts - shapeStart)
            tbW = 6.0; tbH = 3.0; tbX = 2.0; tbY = 5.0
            # Look for style="position:absolute;..." in VML shape
            styleS = wrFindFrom(shapeXml, "width:", 1)
            if styleS > 0
                wEnd = wrFindFrom(shapeXml, "pt", styleS)
                if wEnd > 0
                    wStr = substr(shapeXml, styleS + 6, wEnd - styleS - 6)
                    wVal = number(wStr)
                    if wVal > 0  tbW = wVal * 2.54 / 72.0  ok
                ok
            ok
            result + [:text=txt, :x=tbX, :y=tbY, :width=tbW, :height=tbH]
        ok
        pos = te
    end
    return result


func wrParsePBdr pPrXml
    /*
        Parse <w:pBdr> from a paragraph pPr block.
        Returns [:found, :style, :color, :size, :space, :sides]
        sides is a list of "top","left","bottom","right" that have borders.
    */
    result = [:found=false, :style="single", :color="000000",
              :size=6, :space=4, :sides=[]]
    bdrS = wrFindFrom(pPrXml, "<w:pBdr>", 1)
    if bdrS = 0  bdrS = wrFindFrom(pPrXml, "<w:pBdr ", 1)  ok
    if bdrS = 0  return result  ok
    bdrE = wrFindCloseTag(pPrXml, "w:pBdr", bdrS)
    if bdrE = 0  return result  ok
    bdrXml = substr(pPrXml, bdrS, bdrE - bdrS)
    result[:found] = true
    sides = []
    for side in ["top", "left", "bottom", "right"]
        sS = wrFindFrom(bdrXml, "<w:" + side + " ", 1)
        if sS > 0
            sEl = substr(bdrXml, sS, 120)
            sv = wrAttr(sEl, "w:val")
            sc = wrAttr(sEl, "w:color")
            ssz = wrAttr(sEl, "w:sz")
            ssp = wrAttr(sEl, "w:space")
            if len(sv) > 0 and sv != "none" and sv != "nil"
                sides + side
                if result[:style] = "single" and len(sv) > 0  result[:style] = sv  ok
                if len(sc) > 0 and sc != "auto"  result[:color] = sc  ok
                if len(ssz) > 0  result[:size]  = number(ssz)  ok
                if len(ssp) > 0  result[:space] = number(ssp)  ok
            ok
        ok
    next
    result[:sides] = sides
    return result

func wrParseTcBorders tcPrXml
    /*
        Parse <w:tcBorders> from a table cell tcPr block.
        Returns [:found, :style, :color,
                 :sides = list of [:side, :style, :color, :size] per side].
        :style/:color are still set to the first non-none side found (legacy compat).
    */
    result = [:found=false, :style="single", :color="auto", :sides=[]]
    tbS = wrFindFrom(tcPrXml, "<w:tcBorders>", 1)
    if tbS = 0  tbS = wrFindFrom(tcPrXml, "<w:tcBorders ", 1)  ok
    if tbS = 0  return result  ok
    tbE = wrFindCloseTag(tcPrXml, "w:tcBorders", tbS)
    if tbE = 0  return result  ok
    tbXml = substr(tcPrXml, tbS, tbE - tbS)
    result[:found] = true
    foundFirst = false
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]
        sS = wrFindFrom(tbXml, "<w:" + side + " ", 1)
        if sS > 0
            sEl  = substr(tbXml, sS, 150)
            sv   = wrAttr(sEl, "w:val")
            sc   = wrAttr(sEl, "w:color")
            ssz  = wrAttr(sEl, "w:sz")
            if len(sv) = 0  sv = "single"  ok
            if len(sc) = 0  sc = "auto"    ok
            szN  = 4
            if len(ssz) > 0  szN = number(ssz)  ok
            result[:sides] + [:side=side, :style=sv, :color=sc, :size=szN]
            if !foundFirst and sv != "none" and sv != "nil"
                result[:style] = sv
                if sc != "auto"  result[:color] = sc  ok
                foundFirst = true
            ok
        ok
    next
    return result

func wrParseSdtBlock sdtXml
    /*
        Parse a <w:sdt> content control block.
        Returns [:type, :label, :checked, :choices, :default, :placeholder]
        type: "checkbox", "dropdown", "text", "unknown"
    */
    result = [:type="unknown", :label="", :checked=false,
              :choices=[], :default="", :placeholder=""]
    # Get sdtPr
    prS = wrFindFrom(sdtXml, "<w:sdtPr>", 1)
    if prS = 0  prS = wrFindFrom(sdtXml, "<w:sdtPr ", 1)  ok
    if prS > 0
        prE = wrFindCloseTag(sdtXml, "w:sdtPr", prS)
        if prE > 0
            prXml = substr(sdtXml, prS, prE - prS)
            # Checkbox detection (w14:checkbox)
            if wrFindFrom(prXml, "w14:checkbox", 1) > 0 or
               wrFindFrom(prXml, "w:checkbox",   1) > 0
                result[:type] = "checkbox"
                # Check if checked
                chkS = wrFindFrom(prXml, "w14:checked", 1)
                if chkS = 0  chkS = wrFindFrom(prXml, "w:checked", 1)  ok
                if chkS > 0
                    chkEl = substr(prXml, chkS, 80)
                    chkV  = wrAttr(chkEl, "w14:val")
                    if chkV = NULL or len(chkV) = 0  chkV = wrAttr(chkEl, "w:val")  ok
                    if chkV = "1" or chkV = "true"  result[:checked] = true  ok
                ok
            ok
            # Dropdown detection
            if wrFindFrom(prXml, "<w:dropDownList", 1) > 0
                result[:type] = "dropdown"
                choices = []
                cPos = 1
                while true
                    liS = wrFindFrom(prXml, "<w:listItem ", cPos)
                    if liS = 0  break  ok
                    liEl = substr(prXml, liS, 120)
                    dv = wrAttr(liEl, "w:displayText")
                    if len(dv) = 0  dv = wrAttr(liEl, "w:value")  ok
                    if len(dv) > 0  choices + dv  ok
                    cPos = liS + 1
                end
                result[:choices] = choices
            ok
            # Plain text detection (tag or alias)
            if result[:type] = "unknown"
                if wrFindFrom(prXml, "<w:text>",  1) > 0 or
                   wrFindFrom(prXml, "<w:text ",  1) > 0 or
                   wrFindFrom(prXml, "PlainText",  1) > 0
                    result[:type] = "text"
                ok
            ok
            # Get tag/alias as label hint
            tagS = wrFindFrom(prXml, "<w:tag ", 1)
            if tagS > 0
                tagEl = substr(prXml, tagS, 80)
                tv = wrAttr(tagEl, "w:val")
                if len(tv) > 0 and result[:type] = "unknown"
                    if tv = "toc" or tv = "TOC"  result[:type] = "toc"  ok
                ok
            ok
            # Placeholder text
            phS = wrFindFrom(prXml, "<w:placeholder>", 1)
            if phS = 0  phS = wrFindFrom(prXml, "<w:placeholder ", 1)  ok
            if phS > 0
                phE = wrFindCloseTag(sdtXml, "w:placeholder", phS)
                if phE > 0
                    phXml = substr(sdtXml, phS, phE - phS)
                    dtS = wrFindFrom(phXml, "<w:docPart ", 1)
                    if dtS > 0
                        dtEl = substr(phXml, dtS, 80)
                        result[:placeholder] = wrAttr(dtEl, "w:val")
                    ok
                ok
            ok
        ok
    ok
    # Get content text for default value / label
    cntS = wrFindFrom(sdtXml, "<w:sdtContent>", 1)
    if cntS = 0  cntS = wrFindFrom(sdtXml, "<w:sdtContent ", 1)  ok
    if cntS > 0
        cntE = wrFindCloseTag(sdtXml, "w:sdtContent", cntS)
        if cntE > 0
            cntXml = substr(sdtXml, cntS, cntE - cntS)
            txt = ""
            tPos = 1
            while true
                tS1 = wrFindFrom(cntXml, "<w:t>", tPos)
                tS2 = wrFindFrom(cntXml, "<w:t ", tPos)
                tS  = wrMinPos(tS1, tS2)
                if tS = 0  break  ok
                tCl = wrFindFrom(cntXml, ">", tS)
                tE  = wrFindFrom(cntXml, "</w:t>", tCl)
                if tCl > 0 and tE > 0
                    txt += wrXmlUnescape(substr(cntXml, tCl+1, tE-tCl-1))
                ok
                tPos = tE + 1
            end
            result[:default] = txt
            if len(result[:label]) = 0  result[:label] = txt  ok
        ok
    ok
    return result

func wrDetectTOC bodyXml
    /*
        Detect if bodyXml contains a Table of Contents field (TOC instrText).
        Returns [:found, :title]
    */
    result = [:found=false, :title="Table of Contents"]
    # Check for TOC instrText
    itS = wrFindFrom(bodyXml, "instrText", 1)
    while itS > 0
        itE = wrFindFrom(bodyXml, "</w:instrText>", itS)
        if itE > 0
            itTxt = substr(bodyXml, itS, itE - itS)
            if wrFindFrom(itTxt, " TOC ", 1) > 0 or wrFindFrom(itTxt, "TOC ", 1) = 1
                result[:found] = true
                return result
            ok
        ok
        itS = wrFindFrom(bodyXml, "instrText", itS+1)
    end
    # Also check for TOCHeading paragraph style
    if wrFindFrom(bodyXml, "TOCHeading", 1) > 0
        result[:found] = true
    ok
    return result


func wrParseWatermark headerXml
    /*
        Detect and extract a VML text watermark from header XML.
        Returns [:found, :text, :color, :opacity, :rotation, :font, :size]
    */
    result = [:found=false, :text="", :color="C0C0C0",
              :opacity=50, :rotation=-45, :font="Arial", :size=72]
    # Must contain a v:textpath (VML watermark)
    tpS = wrFindFrom(headerXml, "<v:textpath ", 1)
    if tpS = 0  return result  ok
    tpEl = substr(headerXml, tpS, 300)
    wText = wrAttr(tpEl, "string")
    if len(wText) = 0  return result  ok
    result[:found] = true
    result[:text]  = wrXmlUnescape(wText)
    # Shape for color/opacity/rotation
    shpS = wrFindFrom(headerXml, "<v:shape ", 1)
    if shpS > 0
        shpEl = substr(headerXml, shpS, 500)
        fc = wrAttr(shpEl, "fillcolor")
        if len(fc) > 1 and left(fc,1) = "#"  result[:color] = substr(fc, 2)  ok
        # Rotation from style
        stS = wrFindFrom(shpEl, "rotation:", 1)
        if stS > 0
            rotStr = ""
            i2 = stS + 9
            c2start = substr(shpEl, i2, 1)
            if c2start = "-"  rotStr = "-"  i2++  ok
            shpElLen = len(shpEl)
            while i2 <= shpElLen and i2 <= stS + 20
                c2 = substr(shpEl, i2, 1)
                if ascii(c2) >= ascii("0") and ascii(c2) <= ascii("9")  rotStr += c2
                else  break  ok
                i2++
            end
            if len(rotStr) > 0 and rotStr != "-"  result[:rotation] = number(rotStr)  ok
        ok
    ok
    # Opacity from v:fill
    fillS = wrFindFrom(headerXml, "<v:fill ", 1)
    if fillS > 0
        fillEl = substr(headerXml, fillS, 150)
        opStr = wrAttr(fillEl, "opacity")
        if len(opStr) > 0
            result[:opacity] = floor(number(opStr) * 100)
        ok
    ok
    # Font from v:textpath style (value may be &quot;FontName&quot; or plain FontName)
    stpS = wrFindFrom(tpEl, "font-family:", 1)
    if stpS > 0
        fontStr = ""
        afterColon = substr(tpEl, stpS + 12)
        if left(afterColon, 6) = "&quot;"
            # Font name enclosed in &quot; ... &quot;
            inner = substr(afterColon, 7)
            closeQ = wrFindFrom(inner, "&quot;", 1)
            if closeQ > 0
                fontStr = substr(inner, 1, closeQ - 1)
            ok
        else
            i3 = stpS + 12
            tpElLen = len(tpEl)
            while i3 <= tpElLen and i3 <= stpS + 60
                c3 = substr(tpEl, i3, 1)
                if c3 = ";" or c3 = char(34) or c3 = "'"  break  ok
                fontStr += c3
                i3++
            end
        ok
        if len(fontStr) > 0  result[:font] = fontStr  ok
    ok
    # Size from font-size
    szS = wrFindFrom(tpEl, "font-size:", 1)
    if szS > 0
        szStr = ""
        i4 = szS + 10
        tpElLen = len(tpEl)
        while i4 <= tpElLen and i4 <= szS + 10
            c4 = substr(tpEl, i4, 1)
            if ascii(c4) >= ascii("0") and ascii(c4) <= ascii("9")  szStr += c4
            else  break  ok
            i4++
        end
        if len(szStr) > 0  result[:size] = number(szStr)  ok
    ok
    return result

func wrParseRowProps trXml
    /*
        Parse <w:trPr> from a table row.
        Returns [:height, :hRule, :cantSplit, :isHeaderRow, :rowBgColor]
        height in twips (0 = auto), hRule: "atLeast"|"exact"|"auto"
        cantSplit: true if row should not split across pages
        isHeaderRow: true if w:tblHeader present
        rowBgColor: hex fill color from w:shd on trPr (empty string if none)
    */
    result = [:height=0, :hRule="auto", :cantSplit=false, :isHeaderRow=false, :rowBgColor=""]
    trPrS = wrFindFrom(trXml, "<w:trPr>", 1)
    if trPrS = 0  trPrS = wrFindFrom(trXml, "<w:trPr ", 1)  ok
    if trPrS = 0  return result  ok
    trPrE = wrFindCloseTag(trXml, "w:trPr", trPrS)
    if trPrE = 0  return result  ok
    trPrXml = substr(trXml, trPrS, trPrE - trPrS)
    # Row height
    rhS = wrFindFrom(trPrXml, "<w:trHeight ", 1)
    if rhS > 0
        rhEl = substr(trPrXml, rhS, 120)
        hv = wrAttr(rhEl, "w:val")
        hr = wrAttr(rhEl, "w:hRule")
        if len(hv) > 0  result[:height] = number(hv)  ok
        if len(hr) > 0  result[:hRule]  = hr           ok
    ok
    # Can't split across pages
    if wrFindFrom(trPrXml, "<w:cantSplit", 1) > 0
        result[:cantSplit] = true
    ok
    # Header row repeat
    if wrFindFrom(trPrXml, "<w:tblHeader", 1) > 0
        result[:isHeaderRow] = true
    ok
    # Row-level background shading
    rowShdS = wrFindFrom(trPrXml, "<w:shd ", 1)
    if rowShdS > 0
        rowShdEl = substr(trPrXml, rowShdS, 100)
        rowFill  = wrAttr(rowShdEl, "w:fill")
        if len(rowFill) > 0 and rowFill != "auto"  result[:rowBgColor] = rowFill  ok
    ok
    return result

func wrParseMergeFields pXml
    /*
        Scan paragraph XML for MERGEFIELD instrText.
        Returns list of merge field names found, or [] if none.
    */
    fields = []
    pos = 1
    while true
        itS = wrFindFrom(pXml, "<w:instrText", pos)
        if itS = 0  break  ok
        itCl = wrFindFrom(pXml, ">", itS)
        itE  = wrFindFrom(pXml, "</w:instrText>", itCl)
        if itCl = 0 or itE = 0  break  ok
        instrTxt = wrXmlUnescape(substr(pXml, itCl+1, itE-itCl-1))
        # Trim leading/trailing space
        while len(instrTxt) > 0 and left(instrTxt,1) = " "
            instrTxt = substr(instrTxt, 2)
        end
        # Check for MERGEFIELD keyword
        if left(instrTxt, 10) = "MERGEFIELD"
            fldPart = substr(instrTxt, 12)
            # field name ends at first space or end
            spPos = wrFindFrom(fldPart, " ", 1)
            if spPos > 0
                fldName = substr(fldPart, 1, spPos-1)
            else
                fldName = fldPart
            ok
            if len(fldName) > 0  fields + fldName  ok
        ok
        pos = itE + 1
    end
    return fields

func wrParseNotes xml, tagName
    /*
        Parse footnotes.xml or endnotes.xml.
        Returns list of [:id, :text, :runs] per note.
        :text  — concatenated plain text (backward compat)
        :runs  — list of [:text, :bold, :italic, :underline, :color, :font, :size]
                 per formatted run (rich footnote/endnote content)
        Skips separator and continuationSeparator entries.
    */
    result = []
    pos = 1
    openTag = "<" + tagName + " "
    while true
        s = wrFindFrom(xml, openTag, pos)
        if s = 0  break  ok
        e = wrFindCloseTag(xml, tagName, s)
        if e = 0  break  ok
        el    = substr(xml, s, e - s)
        idStr = wrAttr(el, "w:id")
        noteType = wrAttr(el, "w:type")
        # Skip separator entries
        if noteType = "separator" or noteType = "continuationSeparator"
            pos = e
            loop
        ok
        # Collect runs
        aRuns = []
        txt   = ""
        rPos  = 1
        while true
            rS1 = wrFindFrom(el, "<w:r>",  rPos)
            rS2 = wrFindFrom(el, "<w:r ",  rPos)
            rS  = wrMinPos(rS1, rS2)
            if rS = 0  break  ok
            rE  = wrFindCloseTag(el, "w:r", rS)
            if rE = 0  break  ok
            rXml = substr(el, rS, rE - rS)
            # Skip footnoteRef/endnoteRef runs
            if wrFindFrom(rXml, "<w:footnoteRef", 1) > 0 or
               wrFindFrom(rXml, "<w:endnoteRef",  1) > 0
                rPos = rE
                loop
            ok
            # Get text
            rTxt = ""
            tS1 = wrFindFrom(rXml, "<w:t>", 1)
            tS2 = wrFindFrom(rXml, "<w:t ", 1)
            tS  = wrMinPos(tS1, tS2)
            if tS > 0
                tC = wrFindFrom(rXml, ">", tS)
                tE = wrFindFrom(rXml, "</w:t>", tS)
                if tC > 0 and tE > 0
                    rTxt = wrXmlUnescape(substr(rXml, tC + 1, tE - tC - 1))
                ok
            ok
            # Get run formatting from rPr
            rBold  = false; rItal  = false; rUnder = false
            rColor = ""; rFont = ""; rSize = 0
            rPrS = wrFindFrom(rXml, "<w:rPr>", 1)
            if rPrS = 0  rPrS = wrFindFrom(rXml, "<w:rPr ", 1)  ok
            if rPrS > 0
                rPrE = wrFindCloseTag(rXml, "w:rPr", rPrS)
                if rPrE > 0
                    rPrXml = substr(rXml, rPrS, rPrE - rPrS)
                    if wrFindFrom(rPrXml, "<w:b/>", 1) > 0 or
                       wrFindFrom(rPrXml, "<w:b ",  1) > 0  rBold  = true  ok
                    if wrFindFrom(rPrXml, "<w:i/>", 1) > 0 or
                       wrFindFrom(rPrXml, "<w:i ",  1) > 0  rItal  = true  ok
                    if wrFindFrom(rPrXml, "<w:u ",  1) > 0  rUnder = true  ok
                    cS2 = wrFindFrom(rPrXml, "<w:color ", 1)
                    if cS2 > 0
                        cE2 = substr(rPrXml, cS2, 80)
                        cv2 = wrAttr(cE2, "w:val")
                        if cv2 != NULL and cv2 != "auto"  rColor = cv2  ok
                    ok
                    fS2 = wrFindFrom(rPrXml, "<w:rFonts ", 1)
                    if fS2 > 0
                        fE2 = substr(rPrXml, fS2, 120)
                        fv2 = wrAttr(fE2, "w:ascii")
                        if fv2 = NULL or len(fv2) = 0  fv2 = wrAttr(fE2, "w:hAnsi")  ok
                        if fv2 != NULL  rFont = fv2  ok
                    ok
                    szS2 = wrFindFrom(rPrXml, "<w:sz ", 1)
                    if szS2 > 0
                        szE2 = substr(rPrXml, szS2, 60)
                        szV2 = wrAttr(szE2, "w:val")
                        if len(szV2) > 0  rSize = number(szV2) / 2.0  ok
                    ok
                ok
            ok
            if len(rTxt) > 0
                aRuns + [:text=rTxt, :bold=rBold, :italic=rItal,
                         :underline=rUnder, :color=rColor,
                         :font=rFont, :size=rSize]
                txt += rTxt
            ok
            rPos = rE
        end
        if len(idStr) > 0
            result + [:id=number(idStr), :text=txt, :runs=aRuns]
        ok
        pos = e
    end
    return result

func wrParseComments xml
    /*  Parse comments.xml and return [:id,:text,:author] list.  */
    result = []
    pos = 1
    while true
        s = wrFindFrom(xml, "<w:comment ", pos)
        if s = 0  break  ok
        e = wrFindCloseTag(xml, "w:comment", s)
        if e = 0  break  ok
        el     = substr(xml, s, e - s)
        idStr  = wrAttr(el, "w:id")
        author = wrAttr(el, "w:author")
        txt = ""
        tPos = 1
        while true
            ts1 = wrFindFrom(el, "<w:t>", tPos)
            ts2 = wrFindFrom(el, "<w:t ", tPos)
            ts  = wrMinPos(ts1, ts2)
            if ts = 0  break  ok
            tc = wrFindFrom(el, ">", ts)
            te = wrFindFrom(el, "</w:t>", tc)
            if tc > 0 and te > 0
                txt += substr(el, tc + 1, te - tc - 1)
            ok
            tPos = te + 1
        end
        if len(idStr) > 0
            result + [:id = number(idStr), :text = txt, :author = author]
        ok
        pos = e
    end
    return result

func wrParsePageBgColor docXml
    /*  Extract page background color from <w:background w:color="...">.  */
    bgS = wrFindFrom(docXml, "<w:background ", 1)
    if bgS = 0  return ""  ok
    bgE = wrFindFrom(docXml, ">", bgS)
    if bgE = 0  return ""  ok
    bgEl = substr(docXml, bgS, bgE - bgS + 1)
    col  = wrAttr(bgEl, "w:color")
    if len(col) > 0 and col != "auto" and col != "FFFFFF" and col != "ffffff"
        return col
    ok
    return ""

func wrResolveChartType relId, aRels, tempDir
    /*  Kept for backward compatibility - returns chart type string.  */
    result = wrParseChartData(relId, aRels, tempDir)
    return result[:type]

func wrParseChartData relId, aRels, tempDir
    /*
        Read chart XML and return full chart data:
        [:type, :title, :series]
        Each series: [:name, :categories, :values]
        categories and values are flat lists of strings/numbers.
    */
    sep = wordGetSep()
    result = [:type="chart", :title="", :series=[]]
    chartXml = ""
    for rel in aRels
        if rel[:id] = relId
            target = rel[:target]
            chartPath = ""
            if left(target, 7) = "charts/"
                chartPath = tempDir + "word" + sep + target
            elseif left(target, 10) = "../charts/"
                chartPath = tempDir + "word" + sep + substr(target, 4)
            else
                chartPath = tempDir + "word" + sep + target
            ok
            if fexists(chartPath)
                chartXml = read(chartPath)
            ok
        ok
    next
    if len(chartXml) = 0  return result  ok

    # Chart type
    types = ["barChart","lineChart","pieChart","areaChart",
             "scatterChart","bubbleChart","doughnutChart",
             "radarChart","stockChart","surface3DChart","surfaceChart"]
    for t in types
        if wrFindFrom(chartXml, "<c:" + t, 1) > 0
            result[:type] = t
            break
        ok
    next

    # Chart title from <c:title><c:tx><c:rich>...<a:t>
    titleS = wrFindFrom(chartXml, "<c:title>", 1)
    if titleS = 0  titleS = wrFindFrom(chartXml, "<c:title ", 1)  ok
    if titleS > 0
        titleE = wrFindCloseTag(chartXml, "c:title", titleS)
        if titleE > 0
            titleXml = substr(chartXml, titleS, titleE - titleS)
            atS = wrFindFrom(titleXml, "<a:t>", 1)
            if atS = 0  atS = wrFindFrom(titleXml, "<a:t ", 1)  ok
            if atS > 0
                atCl = wrFindFrom(titleXml, ">", atS)
                atE  = wrFindFrom(titleXml, "</a:t>", atCl)
                if atCl > 0 and atE > 0
                    result[:title] = wrXmlUnescape(substr(titleXml, atCl+1, atE-atCl-1))
                ok
            ok
        ok
    ok

    # Series — scan all <c:ser> elements
    aSeries = []
    serPos = 1
    while true
        serS = wrFindFrom(chartXml, "<c:ser>", serPos)
        if serS = 0  serS = wrFindFrom(chartXml, "<c:ser ", serPos)  ok
        if serS = 0  break  ok
        serE = wrFindCloseTag(chartXml, "c:ser", serS)
        if serE = 0  break  ok
        serXml = substr(chartXml, serS, serE - serS)

        # Series name from <c:tx><c:strRef><c:f> or <c:tx><c:strRef><c:strCache><c:v>
        serName = ""
        txS = wrFindFrom(serXml, "<c:tx>", 1)
        if txS = 0  txS = wrFindFrom(serXml, "<c:tx ", 1)  ok
        if txS > 0
            txE = wrFindCloseTag(serXml, "c:tx", txS)
            if txE > 0
                txXml = substr(serXml, txS, txE - txS)
                # Try <c:v> (inline value)
                vS = wrFindFrom(txXml, "<c:v>", 1)
                if vS > 0
                    vE = wrFindFrom(txXml, "</c:v>", vS)
                    if vE > 0  serName = wrXmlUnescape(substr(txXml, vS+5, vE-vS-5))  ok
                ok
                # Try <a:t> (rich text in title)
                if len(serName) = 0
                    atS2 = wrFindFrom(txXml, "<a:t>", 1)
                    if atS2 = 0  atS2 = wrFindFrom(txXml, "<a:t ", 1)  ok
                    if atS2 > 0
                        atCl2 = wrFindFrom(txXml, ">", atS2)
                        atE2  = wrFindFrom(txXml, "</a:t>", atCl2)
                        if atCl2 > 0 and atE2 > 0
                            serName = wrXmlUnescape(substr(txXml, atCl2+1, atE2-atCl2-1))
                        ok
                    ok
                ok
            ok
        ok

        # Categories from <c:cat><c:strRef><c:strCache> or <c:numRef><c:numCache>
        aCats = []
        catS = wrFindFrom(serXml, "<c:cat>", 1)
        if catS = 0  catS = wrFindFrom(serXml, "<c:cat ", 1)  ok
        if catS > 0
            catE = wrFindCloseTag(serXml, "c:cat", catS)
            if catE > 0
                catXml = substr(serXml, catS, catE - catS)
                cpPos = 1
                while true
                    cvS = wrFindFrom(catXml, "<c:v>", cpPos)
                    if cvS = 0  break  ok
                    cvE = wrFindFrom(catXml, "</c:v>", cvS)
                    if cvE > 0
                        add(aCats, wrXmlUnescape(substr(catXml, cvS+5, cvE-cvS-5)))
                        cpPos = cvE + 1
                    else
                        cpPos = cvS + 1
                    ok
                end
            ok
        ok

        # Values from <c:val><c:numRef><c:numCache>
        aVals = []
        valS = wrFindFrom(serXml, "<c:val>", 1)
        if valS = 0  valS = wrFindFrom(serXml, "<c:val ", 1)  ok
        if valS > 0
            valE = wrFindCloseTag(serXml, "c:val", valS)
            if valE > 0
                valXml = substr(serXml, valS, valE - valS)
                vpPos = 1
                while true
                    vvS = wrFindFrom(valXml, "<c:v>", vpPos)
                    if vvS = 0  break  ok
                    vvE = wrFindFrom(valXml, "</c:v>", vvS)
                    if vvE > 0
                        rawV = wrXmlUnescape(substr(valXml, vvS+5, vvE-vvS-5))
                        add(aVals, rawV)
                        vpPos = vvE + 1
                    else
                        vpPos = vvS + 1
                    ok
                end
            ok
        ok

        ser = [:name=serName, :categories=aCats, :values=aVals]
        add(aSeries, ser)
        serPos = serE
    end
    result[:series] = aSeries
    return result

func wrExtractHeaderText hdrXml
    /*  Extract plain paragraph text from header/footer XML.  */
    result = ""
    pos = 1
    while true
        ps1 = wrFindFrom(hdrXml, "<w:p>", pos)
        ps2 = wrFindFrom(hdrXml, "<w:p ", pos)
        ps  = wrMinPos(ps1, ps2)
        if ps = 0  break  ok
        pe = wrFindCloseTag(hdrXml, "w:p", ps)
        if pe = 0  break  ok
        pXml = substr(hdrXml, ps, pe - ps)
        tPos = 1
        pTxt = ""
        while true
            ts1 = wrFindFrom(pXml, "<w:t>", tPos)
            ts2 = wrFindFrom(pXml, "<w:t ", tPos)
            ts  = wrMinPos(ts1, ts2)
            if ts = 0  break  ok
            tc = wrFindFrom(pXml, ">", ts)
            te = wrFindFrom(pXml, "</w:t>", tc)
            if tc > 0 and te > 0
                pTxt += wrXmlUnescape(substr(pXml, tc + 1, te - tc - 1))
            ok
            tPos = te + 1
        end
        if len(pTxt) > 0
            if len(result) > 0  result += " "  ok
            result += pTxt
        ok
        pos = pe
    end
    return result

func wrLoadHeadersFooters wr
    /*
        Populate all header/footer fields on a WordReader instance.
        Global function to avoid Ring private-method depth limit.
        wr must be a WordReader object.
    */
    sep = wordGetSep()
    docRelsPath = wr.cTempDir + "word" + sep + "_rels" + sep + "document.xml.rels"
    if !fexists(docRelsPath)  return  ok
    relsXml = read(docRelsPath)

    # Map relId -> filename for headers/footers
    sPos = 1
    lastSect = 0

    # Parse all headerReference and footerReference in the final sectPr
    finalSectXml = ""
    docBody = wr.cDocXml
    # Find last sectPr
    sp = 1
    while true
        spFound = wrFindFrom(docBody, "<w:sectPr", sp)
        if spFound = 0  break  ok
        finalSectXml = substr(docBody, spFound, 500)
        sp = spFound + 1
    end

    # Read each header/footer reference in final sectPr
    refPos = 1
    while true
        # headerReference
        hrS = wrFindFrom(finalSectXml, "<w:headerReference ", refPos)
        frS = wrFindFrom(finalSectXml, "<w:footerReference ", refPos)
        nextRef = wrMinPos(hrS, frS)
        if nextRef = 0  break  ok
        refEl  = substr(finalSectXml, nextRef, 150)
        refEnd = wrFindFrom(finalSectXml, "/>", nextRef)
        if refEnd = 0  refPos = nextRef + 1  loop  ok
        isHeader = (nextRef = hrS)
        refType  = wrAttr(refEl, "w:type")
        refRelId = wrAttr(refEl, "r:id")

        # Resolve relId to filename
        hdrFile = ""
        rPos = 1
        while true
            rs = wrFindFrom(relsXml, "<Relationship ", rPos)
            if rs = 0  break  ok
            re = wrFindFrom(relsXml, "/>", rs)
            if re = 0  re = wrFindFrom(relsXml, "</Relationship>", rs)  ok
            if re = 0  break  ok
            relEl  = substr(relsXml, rs, re - rs + 2)
            relId2 = wrAttr(relEl, "Id")
            if relId2 = refRelId
                target = wrAttr(relEl, "Target")
                # Detect header/footer by relationship Type (reliable across all docx)
                relType = wrAttr(relEl, "Type")
                isHdrRel = wrFindFrom(relType, "/header", 1) > 0
                isFtrRel = wrFindFrom(relType, "/footer", 1) > 0
                if isHdrRel or isFtrRel or left(target, 6) = "header" or left(target, 6) = "footer"
                    hdrFile = wr.cTempDir + "word" + sep + target
                ok
            ok
            rPos = re + 1
        end

        if len(hdrFile) > 0 and fexists(hdrFile)
            hdrXml = read(hdrFile)
            hdrTxt = wrExtractHeaderText(hdrXml)
            # Accumulate raw header XML for watermark detection
            wr.cSrcRawHeaderXml += hdrXml
            if isHeader
                if refType = "default" or refType = "odd"
                    if len(wr.cHeaderText) = 0  wr.cHeaderText = hdrTxt  ok
                elseif refType = "even"
                    wr.cEvenPageHeaderText = hdrTxt
                    wr.bSrcEvenAndOddHeaders = true
                elseif refType = "first"
                    wr.cFirstPageHeaderText = hdrTxt
                    wr.bSrcFirstPageDifferent = true
                ok
            else
                if refType = "default" or refType = "odd"
                    if len(wr.cFooterText) = 0  wr.cFooterText = hdrTxt  ok
                elseif refType = "even"
                    wr.cEvenPageFooterText = hdrTxt
                elseif refType = "first"
                    wr.cFirstPageFooterText = hdrTxt
                ok
            ok
        ok
        refPos = refEnd + 1
    end

    # Check titlePg in sectPr
    if wrFindFrom(finalSectXml, "<w:titlePg/>", 1) > 0 or
       wrFindFrom(finalSectXml, "<w:titlePg ", 1) > 0
        wr.bSrcFirstPageDifferent = true
    ok

func wrSaveWriter writer, outputPath
    writer.save(outputPath)

