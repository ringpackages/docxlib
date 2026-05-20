/*
    DOCXLib - Functions
*/

# ============================================================================
# Global Variables
# ============================================================================

# CRC32 lookup table for ZIP creation
aWordZipCRC32Table = []

# ============================================================================
# Quick Helper Functions
# ============================================================================

func quickWord filename, content
    /*
        Quick function to create a simple Word document
        content can be a string or list of strings (each becomes a paragraph)
    */
    doc = new WordWriter()
    
    if isString(content)
        doc.addParagraph(content, NULL)
    elseif isList(content)
        for item in content
            doc.addParagraph(item, NULL)
        next
    ok
    
    return doc.save(filename)

# ============================================================================
# Platform Helper Functions
# ============================================================================

func wordIsWindows
    return isWindows()

func wordGetSep
    if wordIsWindows()
        return "\"
    else
        return "/"
    ok

func wordXmlEsc str
    /*
        Escape special XML characters and strip any control characters that
        are illegal in XML 1.0 (anything below 0x20 except tab/LF/CR).
        This prevents corrupt documents from Ring char() values > 127 which
        ring resolves as byte-modulo values that may land in the control range.
    */
    str = "" + str
    # Strip illegal XML 1.0 control characters (U+0000-U+0008, U+000B-U+000C,
    # U+000E-U+001F) — keep tab (9), LF (10), CR (13)
    result = ""
    strLen = len(str)
    for ci = 1 to strLen
        ch = substr(str, ci, 1)
        cp = ascii(ch)
        if cp = 9 or cp = 10 or cp = 13 or cp >= 32
            result += ch
        ok
    next
    str = result
    str = substr(str, "&", "&amp;")
    str = substr(str, "<", "&lt;")
    str = substr(str, ">", "&gt;")
    str = substr(str, char(34), "&quot;")
    str = substr(str, char(39), "&apos;")
    return str

func wordMakeDir path
    if wordIsWindows()
        path = substr(path, "/", "\")
        system('mkdir "' + path + '" 2>nul')
    else
        system("mkdir -p '" + path + "'")
    ok

func wordGetImageExtension filepath
    # Find the last dot in the filepath
    dotPos = 0
    fpLen = len(filepath)
    # Forward loop, remember last dot position
    for i = 1 to fpLen
        if substr(filepath, i, 1) = "."
            dotPos = i
        ok
    next
    # Extract extension after the last dot
    if dotPos > 0 and dotPos < fpLen
        return substr(filepath, dotPos + 1, fpLen - dotPos)
    ok
    return "png"

func wordColorToHex color
    # Convert color name or hex to proper format
    color = upper(color)
    
    # Common color names
    colors = [
        :black = "000000",
        :white = "FFFFFF",
        :red = "FF0000",
        :green = "00FF00",
        :blue = "0000FF",
        :yellow = "FFFF00",
        :orange = "FFA500",
        :purple = "800080",
        :gray = "808080",
        :grey = "808080",
        :navy = "000080",
        :teal = "008080",
        :maroon = "800000"
    ]
    
    if colors[lower(color)] != NULL
        return colors[lower(color)]
    ok
    
    # Remove # if present
    if left(color, 1) = "#"
        color = substr(color, 2)
    ok
    
    return color

func wordCell text, options
    /*
        Quick helper to create a WordCell with text and optional formatting.
        Options can include all run properties (bold, italic, underline, strike,
        font, size, color, highlight) plus cell properties (align, bgColor, 
        verticalAlign, colspan).
    */
    cell = new WordCell()
    if text != NULL
        cell.addRun(text, options)
    ok
    if isList(options)
        if options[:align] != NULL
            cell.setAlign(options[:align])
        ok
        if options[:bgColor] != NULL or options[:bgcolor] != NULL
            bg = options[:bgColor]
            if bg = NULL bg = options[:bgcolor] ok
            cell.setBgColor(bg)
        ok
        if options[:verticalAlign] != NULL
            cell.setVerticalAlign(options[:verticalAlign])
        ok
        if options[:colspan] != NULL
            cell.setColSpan(options[:colspan])
        ok
        if options[:rowspan] != NULL
            cell.setRowSpan(options[:rowspan])
        ok
    ok
    return cell

func wordMergeCell
    /*
        Creates a continuation cell for vertical merge (row span).
        Place this in rows below the cell that has setRowSpan().
        The cell renders as empty with <w:vMerge/> (continue).
    */
    cell = new WordCell()
    cell.cMerge = "continue"
    return cell

func mergeCell runs, cellOptions
    /*
        Create a WordCell with multiple styled text runs for use
        inside mail-merge table templates.  {{FIELD}} tokens in run texts
        are filled automatically when mergeRecord() or mergeAll() is called.

        runs : a list of run definitions, each an associative list:
                 [:text = "Hello {{Name}}", :bold = true, :color = "C00000"]
               Supported per-run keys: :text, :bold, :italic, :underline,
                 :strike, :font, :size, :color, :highlight

        cellOptions : cell-level options (applied to the whole cell):
                 [:align, :bgColor, :verticalAlign, :colspan, :rowspan]
                 Pass [] or NULL for no cell-level formatting.

        Example — header cell with label + token on separate runs:
            mergeCell([
                [:text = "Client: ",  :bold = true],
                [:text = "{{Name}}", :bold = true, :color = "1F3864"]
            ], [:bgColor = "DEEAF1"])

        Example — status cell coloured at record-fill time via :bgColor token:
            mergeCell([[:text = "{{Status}}", :bold = true]], [:align = "center"])
    */
    cell = new WordCell()
    if !isList(runs)  runs = []  ok
    runsLen = len(runs)
    for ri = 1 to runsLen
        run = runs[ri]
        if !isList(run)  loop  ok
        txt = run[:text]
        if txt = NULL  txt = ""  ok
        cell.addRun("" + txt, run)
    next
    if isList(cellOptions)
        if cellOptions[:align]         != NULL  cell.setAlign(cellOptions[:align])               ok
        if cellOptions[:bgColor]       != NULL  cell.setBgColor(cellOptions[:bgColor])           ok
        if cellOptions[:bgcolor]       != NULL  cell.setBgColor(cellOptions[:bgcolor])           ok
        if cellOptions[:verticalAlign] != NULL  cell.setVerticalAlign(cellOptions[:verticalAlign]) ok
        if cellOptions[:colspan]       != NULL  cell.setColSpan(cellOptions[:colspan])           ok
        if cellOptions[:rowspan]       != NULL  cell.setRowSpan(cellOptions[:rowspan])           ok
    ok
    return cell

# ============================================================================
# ZIP Functions for Word-Compatible ZIP Creation
# ============================================================================

func wordZipInitCRC
    if len(aWordZipCRC32Table) > 0
        return
    ok
    
    for i = 0 to 255
        crc = i
        for j = 1 to 8
            if (crc & 1) = 1
                crc = (crc >> 1) ^ 0xEDB88320
            else
                crc = crc >> 1
            ok
        next
        aWordZipCRC32Table + crc
    next

func wordZipCRC32 data
    wordZipInitCRC()
    
    crc = 0xFFFFFFFF
    dataLen = len(data)
    for i = 1 to dataLen
        b = ascii(data[i])
        idx = ((crc ^ b) & 0xFF) + 1
        crc = (crc >> 8) ^ aWordZipCRC32Table[idx]
    next
    
    return crc ^ 0xFFFFFFFF

func wordZipWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF)

func wordZipDWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF) + char((value >> 16) & 0xFF) + char((value >> 24) & 0xFF)

func wordZipCreateFile filename, filesList
    output = ""
    centralDir = ""
    offset = 0
    
    filesCount = len(filesList)
    for i = 1 to filesCount
        entry = filesList[i]
        zipPath = entry[1]
        content = entry[2]
        
        crc32 = wordZipCRC32(content)
        uncompSize = len(content)
        compMethod = 0
        compSize = uncompSize
        fileData = content
        
        dosTime = 0x0000
        dosDate = 0x0021
        
        localHeader = ""
        localHeader += char(0x50) + char(0x4B) + char(0x03) + char(0x04)
        localHeader += wordZipWord(20)
        localHeader += wordZipWord(0)
        localHeader += wordZipWord(compMethod)
        localHeader += wordZipWord(dosTime)
        localHeader += wordZipWord(dosDate)
        localHeader += wordZipDWord(crc32)
        localHeader += wordZipDWord(compSize)
        localHeader += wordZipDWord(uncompSize)
        localHeader += wordZipWord(len(zipPath))
        localHeader += wordZipWord(0)
        localHeader += zipPath
        
        output += localHeader
        output += fileData
        
        centralEntry = ""
        centralEntry += char(0x50) + char(0x4B) + char(0x01) + char(0x02)
        centralEntry += wordZipWord(20)
        centralEntry += wordZipWord(20)
        centralEntry += wordZipWord(0)
        centralEntry += wordZipWord(compMethod)
        centralEntry += wordZipWord(dosTime)
        centralEntry += wordZipWord(dosDate)
        centralEntry += wordZipDWord(crc32)
        centralEntry += wordZipDWord(compSize)
        centralEntry += wordZipDWord(uncompSize)
        centralEntry += wordZipWord(len(zipPath))
        centralEntry += wordZipWord(0)
        centralEntry += wordZipWord(0)
        centralEntry += wordZipWord(0)
        centralEntry += wordZipWord(0)
        centralEntry += wordZipDWord(0)
        centralEntry += wordZipDWord(offset)
        centralEntry += zipPath
        
        centralDir += centralEntry
        offset += len(localHeader) + len(fileData)
    next
    
    eocd = ""
    eocd += char(0x50) + char(0x4B) + char(0x05) + char(0x06)
    eocd += wordZipWord(0)
    eocd += wordZipWord(0)
    eocd += wordZipWord(len(filesList))
    eocd += wordZipWord(len(filesList))
    eocd += wordZipDWord(len(centralDir))
    eocd += wordZipDWord(offset)
    eocd += wordZipWord(0)
    
    output += centralDir
    output += eocd
    
    write(filename, output)
    return fexists(filename)

