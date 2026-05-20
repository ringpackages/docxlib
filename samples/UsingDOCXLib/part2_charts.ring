/*
	DOCXLib - Demo (Charts)
*/

load "docxlib.ring"

? "============================================="
? "  DOCXLib - Charts Demo"
? "============================================="
? ""

outFile = "demo_charts.docx"

# Shared data
aQuarters   = ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"]
aMonths     = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
aRegions    = ["North", "South", "East", "West", "Central"]
aProducts   = ["Ring", "PWCT", "Supernova", "PWCT2"]

# -- Build document ------------------------------------------------------------
doc = new WordWriter()
doc.setTitle("DOCXLib — Charts Demo")
doc.setAuthor("Mahmoud Fayed")
doc.setPageSize("a4")
doc.setTheme("Blue")
doc.setHeader("DOCXLib — Native OOXML Charts")
doc.showPageNumbers("right")
doc.setMargins(2, 2, 2.2, 2.2)

# -- Cover ---------------------------------------------------------------------
doc.addHeading("DOCXLib — Native OOXML Chart Support", 1)
doc.addParagraph(
    "This document demonstrates native chart types. " +
    "Charts are stored as pure OOXML XML inside the DOCX package ", NULL)
doc.addEmptyParagraph()

doc.addTable([
    ["addColumnChart()",  "Vertical bars — best for comparing values across categories"],
    ["addBarChart()",     "Horizontal bars — great for long category labels"           ],
    ["addLineChart()",    "Trend lines — ideal for time-series data"                   ],
    ["addAreaChart()",    "Filled trend areas — stacked variant shows part-to-whole"   ],
    ["addPieChart()",     "Slices — single-series, shows proportions"                  ],
    ["addDoughnutChart()","Hollow-centre pie — shows proportions with inner space"     ]
], [
    :headerRow      = false,
    :evenRowBgColor = "EEF3FA",
    :borderStyle    = "single",
    :colWidths      = [5.5, 11.5]
])
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 1 — Column chart
# -----------------------------------------------------------------------------
doc.addHeading("1. Column Chart — Quarterly Revenue vs Costs", 2)
doc.addParagraph(
    "A clustered column chart with three series and data labels enabled. " +
    "addColumnChart() is the most common chart type for category comparisons.", NULL)
doc.addEmptyParagraph()

doc.addColumnChart(
    "Quarterly Revenue vs Costs ($K)",
    aQuarters,
    [
        [:name = "Revenue",         :values = [320, 410, 380, 490]],
        [:name = "Operating Costs", :values = [210, 255, 230, 290]],
        [:name = "Net Profit",      :values = [110, 155, 150, 200]]
    ],
    [
        :widthCm        = 14,
        :heightCm       = 9,
        :centered       = true,
        :showDataLabels = true,
        :showLegend     = true,
        :legendPos      = "b"
    ])
doc.addCaption("Figure 1 — Clustered column chart, three series, data labels")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 2 — Line chart
# -----------------------------------------------------------------------------
doc.addHeading("2. Line Chart — Monthly Website Traffic", 2)
doc.addParagraph(
    "A smooth-curve line chart comparing two years of monthly visit data. " +
    "Pass :smooth = true in options for bezier curves instead of straight segments.", NULL)
doc.addEmptyParagraph()

doc.addLineChart(
    "Monthly Website Traffic (Unique Visits)",
    aMonths,
    [
        [:name = "2023", :values = [12000, 14500, 13200, 16800, 18000, 17400]],
        [:name = "2024", :values = [15000, 18200, 17500, 21000, 23500, 22800]]
    ],
    [
        :widthCm  = 14,
        :heightCm = 9,
        :centered = true,
        :smooth   = true,
        :legendPos = "r"
    ])
doc.addCaption("Figure 2 — Smooth line chart, year-over-year comparison")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 3 — Pie chart
# -----------------------------------------------------------------------------
doc.addHeading("3. Pie Chart — Product Revenue Share", 2)
doc.addParagraph(
    "A single-series pie chart. Each slice gets an automatic colour from the " +
    "default palette. Data labels show the category name and percentage.", NULL)
doc.addEmptyParagraph()

doc.addPieChart(
    "Product Revenue Share",
    aProducts,
    [
        [:name = "Revenue ($K)", :values = [450, 280, 160, 110]]
    ],
    [
        :widthCm        = 12,
        :heightCm       = 8,
        :centered       = true,
        :showDataLabels = true,
        :legendPos      = "r"
    ])
doc.addCaption("Figure 3 — Pie chart with per-slice colours and percent labels")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 4 — Doughnut chart
# -----------------------------------------------------------------------------
doc.addHeading("4. Doughnut Chart — Market Share by Region", 2)
doc.addParagraph(
    "The doughnut variant of the pie chart has a hollow centre (50% hole). " +
    "Custom colours are supplied via the :colors option.", NULL)
doc.addEmptyParagraph()

doc.addDoughnutChart(
    "Market Share by Region (%)",
    aRegions,
    [
        [:name = "Market Share", :values = [35, 22, 18, 15, 10]]
    ],
    [
        :widthCm        = 12,
        :heightCm       = 8,
        :centered       = true,
        :showDataLabels = true,
        :legendPos      = "r",
        :colors = ["1F3864", "2E75B6", "70AD47", "FFC000", "C00000"]
    ])
doc.addCaption("Figure 4 — Doughnut chart with custom colour palette")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 5 — Stacked area / column chart
# -----------------------------------------------------------------------------
doc.addHeading("5. Stacked Column Chart — Revenue by Product", 2)
doc.addParagraph(
    "A four-series stacked column chart produced by addColumnChart() with the " +
    ":grouping = 'stacked' option.  The total height of each bar equals the " +
    "combined revenue across all products for that quarter.", NULL)
doc.addEmptyParagraph()

doc.addColumnChart(
    "Stacked Revenue by Product ($K)",
    aQuarters,
    [
        [:name = "Ring",      :values = [120, 155, 140, 180]],
        [:name = "PWCT",      :values = [80,  95,  90,  110]],
        [:name = "Supernova", :values = [60,  75,  70,   90]],
        [:name = "PWCT2",     :values = [40,  55,  50,   70]]
    ],
    [
        :widthCm   = 14,
        :heightCm  = 9,
        :centered  = true,
        :grouping  = "stacked",
        :legendPos = "r"
    ])
doc.addCaption("Figure 5 — Stacked column chart, four product series")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# CHART 6 — Horizontal bar chart
# -----------------------------------------------------------------------------
doc.addHeading("6. Horizontal Bar Chart — Support Tickets by Region", 2)
doc.addParagraph(
    "Horizontal bars are produced by addBarChart(). This orientation works " +
    "well when category labels are long.  Two series show open vs resolved " +
    "tickets per region.", NULL)
doc.addEmptyParagraph()

doc.addBarChart(
    "Support Tickets by Region",
    aRegions,
    [
        [:name = "Open",     :values = [45, 32, 28, 51, 19]],
        [:name = "Resolved", :values = [120, 98, 110, 145, 67]]
    ],
    [
        :widthCm   = 14,
        :heightCm  = 9,
        :centered  = true,
        :legendPos = "b"
    ])
doc.addCaption("Figure 6 — Horizontal bar chart with two series")
doc.addEmptyParagraph()

# -----------------------------------------------------------------------------
# API summary
# -----------------------------------------------------------------------------
doc.addHeading("Chart API Summary", 2)
doc.addTable([
    ["Method",                                        "Chart type",      "Notes"                        ],
    ["addColumnChart(t, cats, series, opts)",         "Column (vertical bar)", "Most common; supports :grouping"],
    ["addBarChart(t, cats, series, opts)",            "Bar (horizontal)",      "Good for long category labels" ],
    ["addLineChart(t, cats, series, opts)",           "Line",                  ":smooth for bezier curves"     ],
    ["addAreaChart(t, cats, series, opts)",           "Area (filled line)",    "Stacked variant via :grouping" ],
    ["addPieChart(t, cats, series, opts)",            "Pie",                   "Uses first series only"        ],
    ["addDoughnutChart(t, cats, series, opts)",       "Doughnut",              "50% hollow centre"             ],
    ["addChart(type, t, cats, series, opts)",         "Any of the above",      "Generic entry point"           ]
], [
    :headerRow      = true,
    :headerBgColor  = "1F3864",
    :evenRowBgColor = "EEF3FA",
    :borderStyle    = "single",
    :colWidths      = [6.5, 4.5, 6]
])
doc.addEmptyParagraph()
doc.addBorderedParagraph(
    "Every series accepts :name (string), :values (list of numbers), and :color (hex string). " +
    "Common options: :widthCm, :heightCm, :centered, :showLegend, :legendPos, " +
    ":showDataLabels, :grouping, :smooth, :colors.",
    [:borderStyle = "single", :borderColor = "1E6BB5",
     :bgColor = "DEEAF1", :indent = 240])

# -- Save ----------------------------------------------------------------------
? "Building document ..."
if doc.save(outFile)
    ? "Saved: " + outFile
else
    ? "Error: save() failed."
ok
