/*
    DOCXLIB
    =======

    Demonstrates 
      * addScatterChart()  - XY scatter (markers, lines+markers, lines only)
      * addBubbleChart()   - Bubble (X/Y + size third dimension)

*/

load "docxlib.ring"

? "================================================="
? "  DOCXLib Scatter & Bubble Charts"
? "================================================="
? ""

outFile = "demo_charts2.docx"

doc = new WordWriter()
doc.setTitle("DOCXLib - Scatter & Bubble Chart Demo")
doc.setAuthor("Mahmoud")
doc.setPageSize("a4")
doc.setTheme("Blue")
doc.setHeader("DOCXLib - Scatter & Bubble Charts")
doc.showPageNumbers("right")
doc.setMargins(2, 2, 2.2, 2.2)

# -- Introduction --------------------------------------------------------------
doc.addHeading("DOCXLib - Scatter & Bubble Charts", 1)
doc.addParagraph(
    "Unlike column/line/pie charts, scatter and bubble charts have no string " +
    "category axis - both axes are fully numeric. Series supply :xValues and " +
    ":yValues lists instead of :values, and bubble series additionally supply " +
    "a :sizes list to control each bubble's area.", NULL)
doc.addEmptyParagraph()

doc.addTable([
    ["addScatterChart()", "XY scatter - markers, lines+markers, or lines only"],
    ["addBubbleChart()",  "Bubble - XY position plus a third numeric dimension as bubble size"]
], [
    :headerRow      = false,
    :evenRowBgColor = "EEF3FA",
    :borderStyle    = "single",
    :colWidths      = [5, 12]
])
doc.addEmptyParagraph()

# =============================================================================
# CHART 1 - Markers-only scatter (default style)
# =============================================================================
doc.addHeading("1. Scatter Chart - Markers Only", 2)
doc.addParagraph(
    "The default scatter style shows only markers at each (X, Y) data point. " +
    "Three research groups are shown; each series has its own colour from the " +
    "automatic palette. Axis labels are set via :xAxisTitle and :yAxisTitle.", NULL)
doc.addEmptyParagraph()

doc.addScatterChart(
    "Study Hours vs Exam Score - Three Cohorts",
    [
        [:name = "Cohort A",
         :xValues = [2, 3, 4, 5, 6, 7, 8, 9, 10],
         :yValues = [45, 52, 60, 67, 73, 79, 84, 88, 93]],
        [:name = "Cohort B",
         :xValues = [1, 2, 4, 5, 7, 8, 9, 10, 11],
         :yValues = [38, 48, 55, 63, 70, 75, 81, 86, 90]],
        [:name = "Cohort C",
         :xValues = [3, 4, 5, 6, 7, 8, 9, 10, 12],
         :yValues = [50, 58, 65, 71, 76, 82, 87, 91, 96]]
    ],
    [
        :widthCm     = 14,
        :heightCm    = 9,
        :centered    = true,
        :legendPos   = "r",
        :markerStyle = "circle",
        :markerSize  = 7,
        :xAxisTitle  = "Daily Study Hours",
        :yAxisTitle  = "Exam Score (%)"
    ])
doc.addCaption("Figure 1 - Markers-only scatter, three cohorts, axis titles set")
doc.addEmptyParagraph()

# =============================================================================
# CHART 2 - Lines + markers scatter
# =============================================================================
doc.addHeading("2. Scatter Chart - Lines + Markers", 2)
doc.addParagraph(
    "Passing :lines = true connects data points with straight lines. " +
    "This variant is ideal for tracking paired measurements over time " +
    "when the X values are not evenly spaced. Two sensors are compared.", NULL)
doc.addEmptyParagraph()

doc.addScatterChart(
    "Temperature vs Pressure - Two Sensors",
    [
        [:name = "Sensor A",
         :xValues = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
         :yValues = [101.2, 100.8, 100.1, 99.6, 99.0, 98.3, 97.8, 97.1, 96.4, 95.8],
         :color   = "2E75B6"],
        [:name = "Sensor B",
         :xValues = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
         :yValues = [101.5, 101.0, 100.4, 99.9, 99.2, 98.7, 98.0, 97.4, 96.8, 96.1],
         :color   = "ED7D31"]
    ],
    [
        :widthCm     = 14,
        :heightCm    = 9,
        :centered    = true,
        :legendPos   = "r",
        :lines       = true,
        :smooth      = false,
        :markerStyle = "square",
        :markerSize  = 6,
        :xAxisTitle  = "Temperature (C)",
        :yAxisTitle  = "Pressure (kPa)"
    ])
doc.addCaption("Figure 2 - Lines + markers scatter, two sensor series")
doc.addEmptyParagraph()

# =============================================================================
# CHART 3 - Smooth lines + markers scatter
# =============================================================================
doc.addHeading("3. Scatter Chart - Smooth Curves + Markers", 2)
doc.addParagraph(
    "Adding :smooth = true renders bezier curves instead of straight " +
    "segments - useful for data that is expected to follow a continuous " +
    "function. Here two algorithm performance curves are compared.", NULL)
doc.addEmptyParagraph()

doc.addScatterChart(
    "Input Size vs Execution Time (ms)",
    [
        [:name = "Algorithm A  O(n log n)",
         :xValues = [100, 500, 1000, 2000, 5000, 10000],
         :yValues = [0.8, 5.2, 11.5, 24.8, 68.3, 145.2],
         :color   = "4472C4"],
        [:name = "Algorithm B  O(n^2)",
         :xValues = [100, 500, 1000, 2000, 5000, 10000],
         :yValues = [1.2, 28.5, 112.4, 448.1, 2802.5, 11210.0],
         :color   = "C00000"]
    ],
    [
        :widthCm     = 14,
        :heightCm    = 9,
        :centered    = true,
        :legendPos   = "b",
        :lines       = true,
        :smooth      = true,
        :markerStyle = "diamond",
        :markerSize  = 7,
        :xAxisTitle  = "Input Size (n)",
        :yAxisTitle  = "Time (ms)"
    ])
doc.addCaption("Figure 3 - Smooth lines + markers, algorithm complexity comparison")
doc.addEmptyParagraph()

# =============================================================================
# CHART 4 - Lines-only scatter (no markers)
# =============================================================================
doc.addHeading("4. Scatter Chart - Lines Only", 2)
doc.addParagraph(
    "Setting :markerStyle = 'none' with :lines = true gives a clean " +
    "continuous line without point markers - suitable for dense datasets " +
    "where markers would clutter the chart.", NULL)
doc.addEmptyParagraph()

aX = []
aY1 = []
aY2 = []
for i = 0 to 36
    aX  + i
    aY1 + (50 + 30 * sin(i * 10 * 3.14159 / 180))
    aY2 + (50 + 25 * cos(i * 10 * 3.14159 / 180))
next

doc.addScatterChart(
    "Sine vs Cosine Wave",
    [
        [:name = "sin(x)",  :xValues = aX, :yValues = aY1, :color = "2E75B6"],
        [:name = "cos(x)",  :xValues = aX, :yValues = aY2, :color = "70AD47"]
    ],
    [
        :widthCm     = 14,
        :heightCm    = 8,
        :centered    = true,
        :legendPos   = "r",
        :lines       = true,
        :smooth      = true,
        :markerStyle = "none",
        :xAxisTitle  = "Degrees / 10",
        :yAxisTitle  = "Amplitude"
    ])
doc.addCaption("Figure 4 - Lines-only scatter with 37 data points per series")
doc.addEmptyParagraph()

# =============================================================================
# CHART 5 - Bubble chart (basic)
# =============================================================================
doc.addHeading("5. Bubble Chart - Market Analysis", 2)
doc.addParagraph(
    "A bubble chart plots (X, Y) coordinates and encodes a third numeric " +
    "dimension as the bubble's area. Each bubble below represents a product: " +
    "X = market share, Y = growth rate, bubble size = annual revenue ($M). " +
    "Each product is its own series so it gets a distinct colour.", NULL)
doc.addEmptyParagraph()

aProducts = [
    [:name = "Ring",      :xValues = [35], :yValues = [22], :sizes = [450], :color = "4472C4"],
    [:name = "PWCT",      :xValues = [28], :yValues = [15], :sizes = [280], :color = "ED7D31"],
    [:name = "Supernova", :xValues = [18], :yValues = [31], :sizes = [160], :color = "70AD47"],
    [:name = "PWCT2",     :xValues = [12], :yValues = [28], :sizes = [110], :color = "FFC000"],
    [:name = "Other",     :xValues = [7],  :yValues = [8],  :sizes = [60],  :color = "5A9BD5"]
]

doc.addBubbleChart(
    "Product Portfolio - Market Share / Growth / Revenue",
    aProducts,
    [
        :widthCm     = 14,
        :heightCm    = 9,
        :centered    = true,
        :legendPos   = "r",
        :xAxisTitle  = "Market Share (%)",
        :yAxisTitle  = "YoY Growth (%)",
        :bubble3D    = false
    ])
doc.addCaption("Figure 5 - Bubble chart: X=market share, Y=growth, size=revenue ($M)")
doc.addEmptyParagraph()

# =============================================================================
# CHART 6 - Bubble chart with 3-D shading
# =============================================================================
doc.addHeading("6. Bubble Chart - Research Metrics (3-D shading)", 2)
doc.addParagraph(
    "Setting :bubble3D = true renders each bubble with a specular highlight, " +
    "giving a glossy sphere appearance. The data shows research projects: " +
    "X = team size, Y = months to completion, size = citation count.", NULL)
doc.addEmptyParagraph()

aProjects = [
    [:name = "Project Alpha",
     :xValues = [3, 5, 8, 12],
     :yValues = [6, 9, 14, 18],
     :sizes   = [45, 120, 280, 500],
     :color   = "1F3864"],
    [:name = "Project Beta",
     :xValues = [2, 4, 7, 10],
     :yValues = [4, 7, 11, 15],
     :sizes   = [30, 90, 200, 380],
     :color   = "C00000"]
]

doc.addBubbleChart(
    "Research Projects - Team / Duration / Impact",
    aProjects,
    [
        :widthCm     = 14,
        :heightCm    = 9,
        :centered    = true,
        :legendPos   = "r",
        :xAxisTitle  = "Team Size",
        :yAxisTitle  = "Months to Completion",
        :bubble3D    = true,
        :showDataLabels = true
    ])
doc.addCaption("Figure 6 - Bubble chart with 3-D shading and bubble-size labels")
doc.addEmptyParagraph()

# =============================================================================
# Series format summary
# =============================================================================
doc.addHeading("Scatter & Bubble Series Format Summary", 2)
doc.addTable([
    ["Key",       "Scatter", "Bubble", "Description"],
    [":name",     "Yes",     "Yes",    "Series label shown in legend"],
    [":xValues",  "Yes",     "Yes",    "List of X numeric values"],
    [":yValues",  "Yes",     "Yes",    "List of Y numeric values (same length as :xValues)"],
    [":sizes",    "No",      "Yes",    "Bubble size values (positive numbers; relative scale)"],
    [":color",    "Optional","Optional","Hex colour override; palette used if omitted"]
], [
    :headerRow      = true,
    :headerBgColor  = "1F3864",
    :evenRowBgColor = "EEF3FA",
    :borderStyle    = "single",
    :colWidths      = [3.5, 2, 2, 9.5]
])
doc.addEmptyParagraph()

doc.addHeading("Options Reference", 2)
doc.addTable([
    ["Option",         "Type",   "Default",   "Applies to",     "Description"],
    [":widthCm",       "Number", "14",        "Both",           "Chart width in centimetres"],
    [":heightCm",      "Number", "9",         "Both",           "Chart height in centimetres"],
    [":centered",      "Bool",   "true",      "Both",           "Centre on page"],
    [":showLegend",    "Bool",   "true",      "Both",           "Show legend"],
    [":legendPos",     "String", char(34)+"r"+char(34),      "Both",           "Legend position: r b t l"],
    [":showDataLabels","Bool",   "false",                    "Both",           "Show value labels on data points"],
    [":xAxisTitle",    "String", char(34)+char(34),          "Both",           "Numeric X axis label"],
    [":yAxisTitle",    "String", char(34)+char(34),          "Both",           "Numeric Y axis label"],
    [":markerStyle",   "String", char(34)+"circle"+char(34),"Scatter only",   "circle square diamond triangle x star dot dash none"],
    [":markerSize",    "Number", "7",         "Scatter only",   "Marker size (half-points)"],
    [":lines",         "Bool",   "false",     "Scatter only",   "Connect points with lines"],
    [":smooth",        "Bool",   "false",     "Scatter only",   "Smooth bezier lines"],
    [":bubble3D",      "Bool",   "false",     "Bubble only",    "3-D specular shading on bubbles"]
], [
    :headerRow      = true,
    :headerBgColor  = "1F3864",
    :evenRowBgColor = "EEF3FA",
    :borderStyle    = "single",
    :colWidths      = [2.8, 1.8, 2.2, 2.8, 7.4]
])

# -- Save ----------------------------------------------------------------------
? "Building document ..."
if doc.save(outFile)
    ? "Saved: " + outFile
    ? ""
    ? "Charts included:"
    ? "  1. Scatter - markers only          (3 cohorts, custom axis labels)"
    ? "  2. Scatter - lines + markers       (2 sensors, straight lines)"
    ? "  3. Scatter - smooth lines + markers(2 algorithms, bezier curves)"
    ? "  4. Scatter - lines only            (sine/cosine, 37 points each)"
    ? "  5. Bubble  - basic                 (5 products, flat shading)"
    ? "  6. Bubble  - 3-D shading           (2 projects, glossy bubbles)"
else
    ? "Error: save() failed."
ok
