"""
Builds the Clean and Green Lawn Care - Kill & Reseed pricing workbook.

Methodology mirrors an M&A valuation model:
- Assumptions tab holds every driver (lawn size, markup, labor rate, hours/phase).
- Inputs (Materials Catalog) holds every SKU with a $/1000 sqft formula.
- Each phase tab pulls from the catalog by name (named ranges) and reports
  Material Cost, Marked-up Material, Labor Cost, and Phase Total.
- Project Total consolidates all four phases.
- Summary is the customer-facing one-pager with a signature block.

Run:  python3 build_workbook.py
Output: Clean_and_Green_KillAndReseed_Pricing.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# ---------- styling ----------
GREEN_DARK = "1B4D1F"
GREEN_PRIMARY = "2E7D32"
GREEN_LIGHT = "A5D6A7"
LIMESTONE = "F4F4EF"
CHARCOAL = "2C2C2C"
YELLOW = "FFC107"

THIN = Side(border_style="thin", color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FONT_TITLE = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
FONT_H1 = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
FONT_H2 = Font(name="Calibri", size=12, bold=True, color=GREEN_DARK)
FONT_LABEL = Font(name="Calibri", size=11, bold=True)
FONT_BODY = Font(name="Calibri", size=11)
FONT_INPUT = Font(name="Calibri", size=11, bold=True, color="0000AA")
FONT_TOTAL = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
FONT_NOTE = Font(name="Calibri", size=9, italic=True, color="666666")

FILL_TITLE = PatternFill("solid", fgColor=GREEN_DARK)
FILL_H1 = PatternFill("solid", fgColor=GREEN_PRIMARY)
FILL_H2 = PatternFill("solid", fgColor=GREEN_LIGHT)
FILL_INPUT = PatternFill("solid", fgColor="FFF8E1")
FILL_TOTAL = PatternFill("solid", fgColor=GREEN_PRIMARY)
FILL_GRANDTOTAL = PatternFill("solid", fgColor=GREEN_DARK)
FILL_LIMESTONE = PatternFill("solid", fgColor=LIMESTONE)
FILL_SIG = PatternFill("solid", fgColor="FFFFFF")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title_bar(ws, text, row=1, span=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = FONT_TITLE
    c.fill = FILL_TITLE
    c.alignment = CENTER
    ws.row_dimensions[row].height = 32


def section_header(ws, text, row, span=8, level=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    if level == 1:
        c.font = FONT_H1
        c.fill = FILL_H1
    else:
        c.font = FONT_H2
        c.fill = FILL_H2
    c.alignment = LEFT
    ws.row_dimensions[row].height = 22


# ============================================================
# WORKBOOK
# ============================================================
wb = Workbook()

# =====================================================================
# Tab 1: SUMMARY (cover + signature) -- created first so it shows first
# =====================================================================
ws_sum = wb.active
ws_sum.title = "Summary"

# =====================================================================
# Tab 2: ASSUMPTIONS
# =====================================================================
ws_a = wb.create_sheet("Assumptions")
title_bar(ws_a, "Assumptions  |  Clean and Green Lawn Care", span=4)
set_col_widths(ws_a, [38, 16, 12, 50])

ws_a.cell(row=2, column=1, value="Driver").font = FONT_LABEL
ws_a.cell(row=2, column=2, value="Value").font = FONT_LABEL
ws_a.cell(row=2, column=3, value="Unit").font = FONT_LABEL
ws_a.cell(row=2, column=4, value="Notes").font = FONT_LABEL
for col in range(1, 5):
    ws_a.cell(row=2, column=col).fill = FILL_H2

# Driver rows -- (label, value, unit, note, name)
assumptions = [
    ("Lawn Size",                       5000,   "sq ft",   "Edit this cell to re-quote any lawn.",            "LawnSize"),
    ("Material Markup",                 0.30,   "%",       "30% markup applied to material cost.",            "Markup"),
    ("Labor Rate",                      75,     "$/hour",  "Blended hourly rate for technician labor.",       "LaborRate"),
    ("Sales Tax (materials)",           0.00,   "%",       "Set if you collect tax on materials.",            "SalesTax"),
    ("Trip / Mobilization Fee",         0,      "$",       "Optional flat fee per visit; set to 0 if rolled into labor.", "TripFee"),
    # Labor hour estimates per 1000 sq ft (empirical best-guesses; tune in field)
    ("Phase 1 Labor (per 1,000 sq ft)", 0.50,   "hours",   "Mix + spray glyphosate; 2 visits at 0.25 hr each.", "P1Hrs"),
    ("Phase 2 Labor (per 1,000 sq ft)", 1.50,   "hours",   "Scalp, dethatch, aerate, debris haul.",             "P2Hrs"),
    ("Phase 3 Labor (per 1,000 sq ft)", 0.75,   "hours",   "Starter fert, seed, Tenacity, peat top-dress.",     "P3Hrs"),
    ("Phase 4 Labor (per 1,000 sq ft)", 0.50,   "hours",   "Follow-up feeding + 2nd Tenacity at ~30 days.",     "P4Hrs"),
    # Per-job (not per 1000 sqft) labor add-ons
    ("Phase 1 Mobilization Hours",      0.50,   "hours",   "Drive, setup, breakdown for Phase 1.",              "P1Mob"),
    ("Phase 2 Mobilization Hours",      1.00,   "hours",   "Equipment pickup/return for aerator + dethatcher.", "P2Mob"),
    ("Phase 3 Mobilization Hours",      0.75,   "hours",   "Material loadout, customer walk-through.",          "P3Mob"),
    ("Phase 4 Mobilization Hours",      0.50,   "hours",   "Visit + walkthrough.",                              "P4Mob"),
]

for i, (label, val, unit, note, name) in enumerate(assumptions, start=3):
    ws_a.cell(row=i, column=1, value=label).font = FONT_LABEL
    c = ws_a.cell(row=i, column=2, value=val)
    c.font = FONT_INPUT
    c.fill = FILL_INPUT
    c.alignment = RIGHT
    if unit == "%":
        c.number_format = "0.0%"
    elif unit == "$" or unit == "$/hour":
        c.number_format = '"$"#,##0.00'
    elif unit == "sq ft":
        c.number_format = "#,##0"
    else:
        c.number_format = "0.00"
    ws_a.cell(row=i, column=3, value=unit).alignment = CENTER
    ws_a.cell(row=i, column=4, value=note).font = FONT_NOTE
    # Defined name
    dn = DefinedName(name, attr_text=f"Assumptions!${get_column_letter(2)}${i}")
    wb.defined_names[name] = dn

# =====================================================================
# Tab 3: MATERIALS CATALOG (Inputs)
# =====================================================================
ws_m = wb.create_sheet("Materials Catalog")
title_bar(ws_m, "Materials Catalog  |  Inputs (priced & broken down per 1,000 sq ft)", span=10)
set_col_widths(ws_m, [4, 30, 14, 14, 16, 12, 14, 16, 14, 36])

headers = ["#", "Product", "Brand", "Vendor", "Package Size", "Unit",
           "Package Price", "App Rate / 1,000 sq ft", "$ / 1,000 sq ft", "Notes"]
for i, h in enumerate(headers, 1):
    c = ws_m.cell(row=2, column=i, value=h)
    c.font = FONT_LABEL
    c.fill = FILL_H2
    c.alignment = CENTER
    c.border = BORDER

# Catalog rows -- (Product, Brand, Vendor, PackageSize, Unit, PackagePrice, AppRate, Notes, name)
# Cost / 1000 = PackagePrice / PackageSize * AppRate
# All prices are May 2026 best-available retail; verify before quoting.
catalog = [
    ("Glyphosate (Roundup Super Concentrate 50%)", "Roundup",   "Home Depot",      128,  "fl oz", 87.00, 6.00,
     "Phase 1 burn-down. Rate: 6 oz product per 1,000 sq ft (mixed in 1 gal water).", "Glyphosate"),
    ("Non-Ionic Surfactant",                      "Southern Ag", "Amazon",          16,   "fl oz", 11.00, 0.50,
     "0.5 oz / gal tank mix to improve glyphosate + Tenacity uptake.",               "Surfactant"),
    ("Tenacity Herbicide (Mesotrione)",           "Syngenta",   "Amazon",           8,   "fl oz", 63.00, 0.18,
     "At-seeding pre-emergent. Rate: 0.18 oz / 1,000 sq ft (5 mL).",                "Tenacity"),
    ("Valkyrie Tall Fescue Seed",                 "United Seeds","unitedseeds.com", 50,   "lb",    200.00,10.00,
     "Hero product. Rate: 10 lb / 1,000 sq ft for full renovation.",                 "Seed"),
    ("Starter Fertilizer 24-25-4",                "Scotts",     "Home Depot",      15,   "lb",    26.00, 3.00,
     "Apply at seeding. 1 bag covers 5,000 sq ft = 3 lb / 1,000.",                  "Starter"),
    ("Peat Moss (3.8 cu ft compressed bale)",     "Premier",    "Home Depot",      3.8,  "cu ft", 14.00, 1.00,
     "Light top-dress over seed for moisture retention. ~1 cu ft / 1,000 sq ft.",  "Peat"),
    ("Milorganite (organic slow-release N)",      "Milorganite","Home Depot",      32,   "lb",    24.97, 12.80,
     "Phase 4 follow-up feed at week 4-6. 1 bag = 2,500 sq ft.",                   "Milo"),
]

# Per-job materials (not per-1000 sqft -- one-time line items)
per_job = [
    ("Pump Sprayer 2-gal (HDX)",         "HDX",        "Home Depot", 1,  "ea",    17.47, 1.00,
     "One-time per job. Reusable across phases.",                                   "Sprayer"),
    ("Aerator Rental (4-hr)",            "Bluebird",   "Home Depot", 1,  "rental",80.00, 1.00,
     "Phase 2 only. 4-hour Home Depot tool rental window.",                         "Aerator"),
    ("Dethatcher Rental (4-hr)",         "Classen",    "Home Depot", 1,  "rental",75.00, 1.00,
     "Phase 2 only.",                                                               "Dethatcher"),
]

# Write per-1000-sqft catalog rows
row = 3
catalog_rows = {}
for i, (prod, brand, vendor, pack, unit, price, rate, note, name) in enumerate(catalog, start=1):
    ws_m.cell(row=row, column=1, value=i).alignment = CENTER
    ws_m.cell(row=row, column=2, value=prod).alignment = LEFT
    ws_m.cell(row=row, column=3, value=brand).alignment = CENTER
    ws_m.cell(row=row, column=4, value=vendor).alignment = CENTER
    c = ws_m.cell(row=row, column=5, value=pack); c.alignment = CENTER; c.number_format = "0.0"
    ws_m.cell(row=row, column=6, value=unit).alignment = CENTER
    c = ws_m.cell(row=row, column=7, value=price); c.font = FONT_INPUT; c.fill = FILL_INPUT
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_m.cell(row=row, column=8, value=rate); c.font = FONT_INPUT; c.fill = FILL_INPUT
    c.number_format = "0.00"; c.alignment = RIGHT
    # cost per 1000 sqft formula
    c = ws_m.cell(row=row, column=9,
                  value=f"=G{row}/E{row}*H{row}")
    c.number_format = '"$"#,##0.00'
    c.font = FONT_LABEL
    c.alignment = RIGHT
    ws_m.cell(row=row, column=10, value=note).font = FONT_NOTE
    for col in range(1, 11):
        ws_m.cell(row=row, column=col).border = BORDER
    catalog_rows[name] = row  # remember row for later refs
    # Named range pointing to $ / 1000 sqft cell
    dn = DefinedName(name, attr_text=f"'Materials Catalog'!${get_column_letter(9)}${row}")
    wb.defined_names[name] = dn
    row += 1

# Per-job catalog rows
section_header(ws_m, "Per-Job Items (one-time, not scaled per sq ft)", row, span=10, level=2)
row += 1
for i, h in enumerate(headers, 1):
    c = ws_m.cell(row=row, column=i, value=h)
    c.font = FONT_LABEL; c.fill = FILL_H2; c.alignment = CENTER; c.border = BORDER
row += 1
for i, (prod, brand, vendor, pack, unit, price, rate, note, name) in enumerate(per_job, start=1):
    ws_m.cell(row=row, column=1, value=i).alignment = CENTER
    ws_m.cell(row=row, column=2, value=prod).alignment = LEFT
    ws_m.cell(row=row, column=3, value=brand).alignment = CENTER
    ws_m.cell(row=row, column=4, value=vendor).alignment = CENTER
    c = ws_m.cell(row=row, column=5, value=pack); c.alignment = CENTER
    ws_m.cell(row=row, column=6, value=unit).alignment = CENTER
    c = ws_m.cell(row=row, column=7, value=price); c.font = FONT_INPUT; c.fill = FILL_INPUT
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_m.cell(row=row, column=8, value=rate); c.alignment = CENTER
    c = ws_m.cell(row=row, column=9, value=f"=G{row}*H{row}")
    c.number_format = '"$"#,##0.00'; c.font = FONT_LABEL; c.alignment = RIGHT
    ws_m.cell(row=row, column=10, value=note).font = FONT_NOTE
    for col in range(1, 11):
        ws_m.cell(row=row, column=col).border = BORDER
    # Named range = total per-job cost (price * qty)
    dn = DefinedName(name, attr_text=f"'Materials Catalog'!${get_column_letter(9)}${row}")
    wb.defined_names[name] = dn
    row += 1

# Footnote
row += 1
ws_m.cell(row=row, column=1,
          value="Prices reflect best-available May 2026 retail; verify before customer-facing quotes.")
ws_m.cell(row=row, column=1).font = FONT_NOTE
ws_m.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)


# =====================================================================
# Helper to build a Phase tab
# =====================================================================
def build_phase(name, title, materials, per_job_items, phase_hours_name, phase_mob_name, narrative):
    """
    materials:    list of (named_range, label, apps)  -- per-1000-sqft items
    per_job_items: list of (named_range, label, qty)  -- one-time items
    phase_hours_name: defined name for hr/1000sqft labor
    phase_mob_name:   defined name for fixed mobilization hours
    """
    ws = wb.create_sheet(name)
    title_bar(ws, title, span=7)
    set_col_widths(ws, [4, 36, 14, 14, 18, 18, 36])

    section_header(ws, "Scope of Work", 2, span=7, level=2)
    ws.cell(row=3, column=1, value=narrative).alignment = LEFT
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=7)
    ws.row_dimensions[3].height = 60

    section_header(ws, "Materials  (per 1,000 sq ft items)", 5, span=7, level=2)
    hdrs = ["#", "Item", "$ / 1,000 sq ft", "# Apps", "Cost (raw)", "Cost (w/ markup)", "Notes"]
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(row=6, column=i, value=h)
        c.font = FONT_LABEL; c.fill = FILL_H2; c.alignment = CENTER; c.border = BORDER

    r = 7
    mat_rows = []
    for i, (nm, label, apps) in enumerate(materials, 1):
        ws.cell(row=r, column=1, value=i).alignment = CENTER
        ws.cell(row=r, column=2, value=label).alignment = LEFT
        c = ws.cell(row=r, column=3, value=f"={nm}")
        c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
        c = ws.cell(row=r, column=4, value=apps)
        c.font = FONT_INPUT; c.fill = FILL_INPUT; c.alignment = CENTER
        # raw cost = $/1000 * apps * (LawnSize/1000)
        c = ws.cell(row=r, column=5,
                    value=f"=C{r}*D{r}*(LawnSize/1000)")
        c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
        # marked-up
        c = ws.cell(row=r, column=6, value=f"=E{r}*(1+Markup)")
        c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
        ws.cell(row=r, column=7, value="").font = FONT_NOTE
        for col in range(1, 8):
            ws.cell(row=r, column=col).border = BORDER
        mat_rows.append(r)
        r += 1

    # Per-job items
    if per_job_items:
        section_header(ws, "Materials  (per-job / one-time)", r, span=7, level=2)
        r += 1
        hdrs2 = ["#", "Item", "Cost each", "Qty", "Cost (raw)", "Cost (w/ markup)", "Notes"]
        for i, h in enumerate(hdrs2, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.font = FONT_LABEL; c.fill = FILL_H2; c.alignment = CENTER; c.border = BORDER
        r += 1
        for i, (nm, label, qty) in enumerate(per_job_items, 1):
            ws.cell(row=r, column=1, value=i).alignment = CENTER
            ws.cell(row=r, column=2, value=label).alignment = LEFT
            c = ws.cell(row=r, column=3, value=f"={nm}")
            c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
            c = ws.cell(row=r, column=4, value=qty)
            c.font = FONT_INPUT; c.fill = FILL_INPUT; c.alignment = CENTER
            c = ws.cell(row=r, column=5, value=f"=C{r}*D{r}")
            c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
            c = ws.cell(row=r, column=6, value=f"=E{r}*(1+Markup)")
            c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
            ws.cell(row=r, column=7, value="").font = FONT_NOTE
            for col in range(1, 8):
                ws.cell(row=r, column=col).border = BORDER
            mat_rows.append(r)
            r += 1

    # Material subtotal
    r += 1
    ws.cell(row=r, column=2, value="MATERIALS SUBTOTAL (with markup)").font = FONT_LABEL
    raw_sum_parts = "+".join(f"E{x}" for x in mat_rows)
    mark_sum_parts = "+".join(f"F{x}" for x in mat_rows)
    c = ws.cell(row=r, column=5, value=f"={raw_sum_parts}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
    c = ws.cell(row=r, column=6, value=f"={mark_sum_parts}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    c.fill = FILL_TOTAL
    mat_subtotal_cell = f"F{r}"

    # Labor section
    r += 2
    section_header(ws, "Labor", r, span=7, level=2)
    r += 1
    hdrs3 = ["", "Component", "Hours basis", "Hours", "Rate", "Cost", "Notes"]
    for i, h in enumerate(hdrs3, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = FONT_LABEL; c.fill = FILL_H2; c.alignment = CENTER; c.border = BORDER
    r += 1
    # per-1000 sqft labor
    ws.cell(row=r, column=2, value="Field labor (per 1,000 sq ft)").alignment = LEFT
    ws.cell(row=r, column=3, value=f"={phase_hours_name} hr/1,000 sf").alignment = CENTER
    c = ws.cell(row=r, column=4, value=f"={phase_hours_name}*(LawnSize/1000)")
    c.number_format = "0.00"; c.alignment = CENTER
    c = ws.cell(row=r, column=5, value="=LaborRate")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws.cell(row=r, column=6, value=f"=D{r}*E{r}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = BORDER
    field_lab_row = r
    r += 1
    # Mobilization labor
    ws.cell(row=r, column=2, value="Mobilization / setup").alignment = LEFT
    ws.cell(row=r, column=3, value="Fixed hrs/job").alignment = CENTER
    c = ws.cell(row=r, column=4, value=f"={phase_mob_name}")
    c.number_format = "0.00"; c.alignment = CENTER
    c = ws.cell(row=r, column=5, value="=LaborRate")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws.cell(row=r, column=6, value=f"=D{r}*E{r}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = BORDER
    mob_row = r

    r += 1
    ws.cell(row=r, column=2, value="LABOR SUBTOTAL").font = FONT_LABEL
    c = ws.cell(row=r, column=6, value=f"=F{field_lab_row}+F{mob_row}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    c.fill = FILL_TOTAL
    lab_subtotal_cell = f"F{r}"

    # Phase total
    r += 2
    ws.cell(row=r, column=2, value=f"{title.upper()} - PHASE TOTAL").font = Font(
        name="Calibri", size=14, bold=True, color="FFFFFF")
    ws.cell(row=r, column=2).fill = FILL_GRANDTOTAL
    ws.cell(row=r, column=2).alignment = LEFT
    for col in range(3, 6):
        ws.cell(row=r, column=col).fill = FILL_GRANDTOTAL
    c = ws.cell(row=r, column=6, value=f"={mat_subtotal_cell}+{lab_subtotal_cell}")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c.fill = FILL_GRANDTOTAL
    phase_total_cell_addr = f"'{name}'!$F${r}"

    # define per-phase named ranges for total / mat / labor
    short = name.split()[1]  # "1", "2", etc -> after "Phase"
    # actually phase names are "Phase 1 - Kill" etc, let's slug it
    slug = "P" + name.split()[1].rstrip("-")
    wb.defined_names[f"{slug}_Total"] = DefinedName(f"{slug}_Total", attr_text=phase_total_cell_addr)
    wb.defined_names[f"{slug}_Mat"]   = DefinedName(f"{slug}_Mat",   attr_text=f"'{name}'!${mat_subtotal_cell}")
    wb.defined_names[f"{slug}_Lab"]   = DefinedName(f"{slug}_Lab",   attr_text=f"'{name}'!${lab_subtotal_cell}")
    return phase_total_cell_addr


# =====================================================================
# Phase 1 -- Kill
# =====================================================================
build_phase(
    name="Phase 1 - Kill",
    title="Phase 1  |  Kill (Glyphosate Burn-Down)",
    materials=[
        ("Glyphosate", "Glyphosate (Roundup Super Concentrate 50%)", 2),
        ("Surfactant", "Non-Ionic Surfactant (tank mix)", 2),
    ],
    per_job_items=[
        ("Sprayer", "2-gallon pump sprayer", 1),
    ],
    phase_hours_name="P1Hrs",
    phase_mob_name="P1Mob",
    narrative=(
        "Two glyphosate applications spaced 10-14 days apart. App #1 burns down all existing turf "
        "and weeds. After kill, light watering is used to germinate dormant weed seeds; App #2 "
        "kills that flush so the seedbed is clean before Phase 2 prep."
    ),
)

# =====================================================================
# Phase 2 -- Prep
# =====================================================================
build_phase(
    name="Phase 2 - Prep",
    title="Phase 2  |  Prep (Scalp, Dethatch, Aerate)",
    materials=[],
    per_job_items=[
        ("Dethatcher", "Power dethatcher rental (4-hr)", 1),
        ("Aerator",    "Core aerator rental (4-hr)",    1),
    ],
    phase_hours_name="P2Hrs",
    phase_mob_name="P2Mob",
    narrative=(
        "Scalp the dead lawn to the lowest mower setting and bag clippings. Run the power "
        "dethatcher in two directions to lift dead thatch. Core-aerate to open the seedbed. "
        "Final cleanup leaves bare soil ready for seed-to-soil contact."
    ),
)

# =====================================================================
# Phase 3 -- Seed & Starter
# =====================================================================
build_phase(
    name="Phase 3 - Seed",
    title="Phase 3  |  Seed + Starter Fertilizer + Tenacity",
    materials=[
        ("Starter",    "Starter fertilizer (24-25-4)",       1),
        ("Seed",       "Valkyrie tall fescue seed",           1),
        ("Tenacity",   "Tenacity at-seeding pre-emergent",   1),
        ("Surfactant", "Non-Ionic Surfactant (Tenacity tank)",1),
        ("Peat",       "Peat moss top-dress",                 1),
    ],
    per_job_items=[],
    phase_hours_name="P3Hrs",
    phase_mob_name="P3Mob",
    narrative=(
        "Apply starter fertilizer, then broadcast Valkyrie seed at 10 lb / 1,000 sq ft in two "
        "directions. Tank-mix Tenacity with non-ionic surfactant and apply at 0.18 oz / 1,000 sq ft "
        "for 21 days of pre-emergent weed protection. Top-dress lightly with peat moss to retain "
        "moisture, then begin the watering schedule (light 2-3x daily until germination)."
    ),
)

# =====================================================================
# Phase 4 -- Establishment
# =====================================================================
build_phase(
    name="Phase 4 - Establish",
    title="Phase 4  |  Establishment (Follow-up Feed + 2nd Tenacity)",
    materials=[
        ("Milo",       "Milorganite 4-6 week follow-up feed", 1),
        ("Tenacity",   "Tenacity 2nd app (post 30 days)",      1),
        ("Surfactant", "Non-Ionic Surfactant (Tenacity tank)", 1),
    ],
    per_job_items=[],
    phase_hours_name="P4Hrs",
    phase_mob_name="P4Mob",
    narrative=(
        "Once new turf reaches the 4-week mark and has been mowed once, apply Milorganite for slow-"
        "release nitrogen. After 30 days from initial Tenacity, a second app extends pre-emergent "
        "protection while new grass continues filling in. Customer receives a written watering and "
        "mowing schedule for ongoing care."
    ),
)

# =====================================================================
# Tab: PROJECT TOTAL
# =====================================================================
ws_t = wb.create_sheet("Project Total")
title_bar(ws_t, "Project Total  |  All Phases Combined", span=5)
set_col_widths(ws_t, [4, 30, 18, 18, 18])

ws_t.cell(row=2, column=1, value="#").font = FONT_LABEL
ws_t.cell(row=2, column=2, value="Phase").font = FONT_LABEL
ws_t.cell(row=2, column=3, value="Materials (w/ markup)").font = FONT_LABEL
ws_t.cell(row=2, column=4, value="Labor").font = FONT_LABEL
ws_t.cell(row=2, column=5, value="Phase Total").font = FONT_LABEL
for col in range(1, 6):
    ws_t.cell(row=2, column=col).fill = FILL_H2
    ws_t.cell(row=2, column=col).alignment = CENTER
    ws_t.cell(row=2, column=col).border = BORDER

phase_meta = [
    ("1", "Phase 1 - Kill (glyphosate burn-down)",   "P1"),
    ("2", "Phase 2 - Prep (scalp, dethatch, aerate)", "P2"),
    ("3", "Phase 3 - Seed + Starter + Tenacity",       "P3"),
    ("4", "Phase 4 - Establishment + follow-up",       "P4"),
]

r = 3
phase_total_rows = []
for num, label, slug in phase_meta:
    ws_t.cell(row=r, column=1, value=num).alignment = CENTER
    ws_t.cell(row=r, column=2, value=label).alignment = LEFT
    c = ws_t.cell(row=r, column=3, value=f"={slug}_Mat*(1+Markup)")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_t.cell(row=r, column=4, value=f"={slug}_Lab")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_t.cell(row=r, column=5, value=f"={slug}_Total")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
    for col in range(1, 6):
        ws_t.cell(row=r, column=col).border = BORDER
    phase_total_rows.append(r)
    r += 1

# Project totals
r += 1
ws_t.cell(row=r, column=2, value="MATERIALS TOTAL (w/ markup)").font = FONT_LABEL
c = ws_t.cell(row=r, column=3,
              value=f"=SUM(C{phase_total_rows[0]}:C{phase_total_rows[-1]})")
c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
c.font = FONT_TOTAL; c.fill = FILL_TOTAL
r += 1
ws_t.cell(row=r, column=2, value="LABOR TOTAL").font = FONT_LABEL
c = ws_t.cell(row=r, column=4,
              value=f"=SUM(D{phase_total_rows[0]}:D{phase_total_rows[-1]})")
c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
c.font = FONT_TOTAL; c.fill = FILL_TOTAL
r += 1
ws_t.cell(row=r, column=2,
          value="PROJECT GRAND TOTAL").font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
ws_t.cell(row=r, column=2).fill = FILL_GRANDTOTAL
for col in [3, 4]:
    ws_t.cell(row=r, column=col).fill = FILL_GRANDTOTAL
c = ws_t.cell(row=r, column=5,
              value=f"=SUM(E{phase_total_rows[0]}:E{phase_total_rows[-1]})")
c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF"); c.fill = FILL_GRANDTOTAL
grand_total_cell = f"'Project Total'!$E${r}"
wb.defined_names["GrandTotal"] = DefinedName("GrandTotal", attr_text=grand_total_cell)

# Per-sqft tag
r += 2
ws_t.cell(row=r, column=2, value="Project cost per sq ft (reference)").font = FONT_NOTE
c = ws_t.cell(row=r, column=5, value="=GrandTotal/LawnSize")
c.number_format = '"$"#,##0.0000'; c.alignment = RIGHT; c.font = FONT_NOTE


# =====================================================================
# Tab 1 (filled in last): SUMMARY -- customer-facing
# =====================================================================
title_bar(ws_sum, "Clean and Green Lawn Care  |  Lawn Renovation Proposal", span=5)
set_col_widths(ws_sum, [4, 32, 20, 20, 20])

# Project header
ws_sum.cell(row=2, column=2, value="Project").font = FONT_LABEL
ws_sum.cell(row=2, column=3, value="Kill & Reseed Lawn Renovation").alignment = LEFT
ws_sum.cell(row=3, column=2, value="Method").font = FONT_LABEL
ws_sum.cell(row=3, column=3, value="The Lawn Care Nut (Allyn Hane) renovation protocol").alignment = LEFT
ws_sum.cell(row=4, column=2, value="Lawn Size").font = FONT_LABEL
c = ws_sum.cell(row=4, column=3, value="=LawnSize"); c.number_format = "#,##0\" sq ft\""
ws_sum.cell(row=5, column=2, value="Hero Seed").font = FONT_LABEL
ws_sum.cell(row=5, column=3,
            value="Valkyrie Tall Fescue (United Seeds) @ 10 lb / 1,000 sq ft").alignment = LEFT
ws_sum.cell(row=6, column=2, value="Date").font = FONT_LABEL
ws_sum.cell(row=6, column=3, value="").alignment = LEFT  # customer fills in

# Phase summary block
section_header(ws_sum, "Phase-by-Phase Investment", 8, span=5, level=1)
hdrs = ["#", "Phase", "Materials (w/ markup)", "Labor", "Phase Total"]
for i, h in enumerate(hdrs, 1):
    c = ws_sum.cell(row=9, column=i, value=h)
    c.font = FONT_LABEL; c.fill = FILL_H2; c.alignment = CENTER; c.border = BORDER

r = 10
for num, label, slug in phase_meta:
    ws_sum.cell(row=r, column=1, value=num).alignment = CENTER
    ws_sum.cell(row=r, column=2, value=label).alignment = LEFT
    c = ws_sum.cell(row=r, column=3, value=f"={slug}_Mat*(1+Markup)")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_sum.cell(row=r, column=4, value=f"={slug}_Lab")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
    c = ws_sum.cell(row=r, column=5, value=f"={slug}_Total")
    c.number_format = '"$"#,##0.00'; c.alignment = RIGHT; c.font = FONT_LABEL
    for col in range(1, 6):
        ws_sum.cell(row=r, column=col).border = BORDER
    r += 1

# Grand total on summary
r += 1
ws_sum.cell(row=r, column=2,
            value="TOTAL PROJECT INVESTMENT").font = Font(
    name="Calibri", size=14, bold=True, color="FFFFFF")
ws_sum.cell(row=r, column=2).fill = FILL_GRANDTOTAL
for col in [3, 4]:
    ws_sum.cell(row=r, column=col).fill = FILL_GRANDTOTAL
c = ws_sum.cell(row=r, column=5, value="=GrandTotal")
c.number_format = '"$"#,##0.00'; c.alignment = RIGHT
c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
c.fill = FILL_GRANDTOTAL
ws_sum.row_dimensions[r].height = 28

# What's included
r += 3
section_header(ws_sum, "What's Included", r, span=5, level=1)
r += 1
includes = [
    "Two glyphosate burn-down applications with pump-sprayer mix and labor",
    "Power dethatching and core aeration of the prepared seedbed",
    "Premium Valkyrie tall fescue seed, broadcast at 10 lb / 1,000 sq ft",
    "Scotts starter fertilizer (24-25-4) applied at seeding",
    "Tenacity (mesotrione) at-seeding application + 30-day follow-up",
    "Peat moss top-dressing for moisture retention",
    "Milorganite follow-up feeding at week 4-6",
    "Written watering and mowing schedule, plus 30-day callback guarantee",
]
for line in includes:
    ws_sum.cell(row=r, column=2, value="•  " + line).alignment = LEFT
    ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    r += 1

# Terms
r += 1
section_header(ws_sum, "Terms", r, span=5, level=1)
r += 1
terms = [
    "Quote valid for 30 days from date of issue.",
    "50% deposit due at signing; balance due upon completion of Phase 4.",
    "Customer is responsible for daily watering per the schedule provided.",
    "Material prices subject to change if vendor pricing moves more than 10% before scheduling.",
    "30-day callback: any bare patch larger than 4 sq ft will be re-seeded at no additional charge.",
]
for line in terms:
    ws_sum.cell(row=r, column=2, value="•  " + line).alignment = LEFT
    ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    r += 1

# Signature block
r += 2
section_header(ws_sum, "Acceptance & Signature", r, span=5, level=1)
r += 2
ws_sum.cell(row=r, column=2,
            value="By signing below, customer accepts the scope, timing, and total above.").font = FONT_NOTE
ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)

r += 2
sig_thick = Side(border_style="medium", color=CHARCOAL)
sig_border = Border(bottom=sig_thick)

# Customer
ws_sum.cell(row=r, column=2, value="Customer Signature").font = FONT_LABEL
ws_sum.cell(row=r, column=4, value="Date").font = FONT_LABEL
r += 1
for col in [2, 3]:
    ws_sum.cell(row=r, column=col).border = sig_border
ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
ws_sum.cell(row=r, column=4).border = sig_border
ws_sum.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
ws_sum.row_dimensions[r].height = 28

r += 1
ws_sum.cell(row=r, column=2, value="Printed Name").font = FONT_LABEL
r += 1
ws_sum.cell(row=r, column=2).border = sig_border
ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
ws_sum.row_dimensions[r].height = 24

# Company
r += 2
ws_sum.cell(row=r, column=2, value="Clean and Green Lawn Care - Authorized Representative").font = FONT_LABEL
ws_sum.cell(row=r, column=4, value="Date").font = FONT_LABEL
r += 1
for col in [2, 3]:
    ws_sum.cell(row=r, column=col).border = sig_border
ws_sum.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
ws_sum.cell(row=r, column=4).border = sig_border
ws_sum.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
ws_sum.row_dimensions[r].height = 28

# Reorder tabs: Summary, Assumptions, Materials Catalog, Phase 1-4, Project Total
order = ["Summary", "Assumptions", "Materials Catalog",
         "Phase 1 - Kill", "Phase 2 - Prep", "Phase 3 - Seed", "Phase 4 - Establish",
         "Project Total"]
wb._sheets = [wb[s] for s in order]

# Tab colors
wb["Summary"].sheet_properties.tabColor = GREEN_DARK
wb["Assumptions"].sheet_properties.tabColor = YELLOW
wb["Materials Catalog"].sheet_properties.tabColor = YELLOW
for s in ["Phase 1 - Kill", "Phase 2 - Prep", "Phase 3 - Seed", "Phase 4 - Establish"]:
    wb[s].sheet_properties.tabColor = GREEN_PRIMARY
wb["Project Total"].sheet_properties.tabColor = GREEN_DARK

out = "Clean_and_Green_KillAndReseed_Pricing.xlsx"
wb.save(out)
print(f"Wrote {out}")
