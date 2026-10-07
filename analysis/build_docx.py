"""Generate the Lot Due-Diligence report as a Word .docx (A4 portrait, page numbers)."""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FIG = "/projects/sandbox/work/fig"
OUTDIR = ("/projects/sandbox/https-www.zillow.com-homedetails-"
          "1147-1153-Waters-Edge-Cv-130-Tignall-GA-30668-465607553_z-")
OUT = os.path.join(OUTDIR, "Property LOT Due Diligence - Report.docx")

ACCENT = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x1F, 0x2B)
GREEN = RGBColor(0x1E, 0x6B, 0x3A)
GREY = RGBColor(0x55, 0x55, 0x55)

doc = Document()

# ---------------------------------------------------------------- page setup
sec = doc.sections[0]
sec.page_width = Cm(21.0)      # A4 portrait
sec.page_height = Cm(29.7)
for m in ("top_margin", "bottom_margin"):
    setattr(sec, m, Cm(1.8))
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(1.6)

st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(9.5)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.06

for nm, sz, col, before in (("Heading 1", 15, ACCENT, 14),
                            ("Heading 2", 12, ACCENT, 10),
                            ("Heading 3", 10.5, RGBColor(0x33, 0x33, 0x33), 8)):
    s = doc.styles[nm]
    s.font.name = "Calibri"
    s.font.size = Pt(sz)
    s.font.color.rgb = col
    s.font.bold = True
    s.paragraph_format.space_before = Pt(before)
    s.paragraph_format.space_after = Pt(3)
    s.paragraph_format.keep_with_next = True


# ---------------------------------------------------------------- helpers
def field(par, instr):
    r = par.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    return r


def add_footer():
    ft = sec.footer
    p = ft.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("1147 / 1153 Watersedge Cove, Tignall (Lincoln County), GA 30668  —  Page ")
    run.font.size = Pt(7.5); run.font.color.rgb = GREY
    field(p, "PAGE")
    r2 = p.add_run(" of "); r2.font.size = Pt(7.5); r2.font.color.rgb = GREY
    field(p, "NUMPAGES")
    for r in p.runs:
        r.font.size = Pt(7.5); r.font.color.rgb = GREY


def para(text, size=9.5, bold=False, italic=False, color=None, align=None,
         space_after=4, space_before=0, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    r.font.size = Pt(size); r.bold = bold; r.italic = italic
    if color:
        r.font.color.rgb = color
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    return p


def bullets(items, size=9.5, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.left_indent = Cm(0.65)
        if isinstance(it, tuple):
            r = p.add_run(it[0] + " "); r.bold = True; r.font.size = Pt(size)
            r2 = p.add_run(it[1]); r2.font.size = Pt(size)
        else:
            r = p.add_run(it); r.font.size = Pt(size)


TIER = {
    "D": ("DOCUMENTED", RGBColor(0x1E, 0x6B, 0x3A)),
    "M": ("MODEL-DERIVED ESTIMATE", RGBColor(0xB4, 0x6B, 0x00)),
    "V": ("REQUIRES PROFESSIONAL / OFFICIAL VERIFICATION", RGBColor(0xC0, 0x1F, 0x2B)),
}


def tier(code, note=""):
    lbl, col = TIER[code]
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("[" + lbl + "]")
    r.font.size = Pt(7.5); r.bold = True; r.font.color.rgb = col
    if note:
        r2 = p.add_run("  " + note)
        r2.font.size = Pt(7.5); r2.italic = True; r2.font.color.rgb = GREY
    return p


def shade(cell, hexcol):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hexcol)
    tcPr.append(sh)


def table(headers, rows, widths=None, fontsize=8.3, header_fill="1F4E79",
          zebra=True, align=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        r = p.add_run(str(h))
        r.bold = True; r.font.size = Pt(fontsize); r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(hdr[i], header_fill)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            txt = "" if v is None else str(v)
            bold = txt.startswith("**") and txt.endswith("**")
            if bold:
                txt = txt[2:-2]
            r = p.add_run(txt)
            r.font.size = Pt(fontsize); r.bold = bold
            if align and i in align:
                p.alignment = align[i]
            if zebra and ri % 2 == 1:
                shade(cells[i], "F2F5F9")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(fname, caption, width_cm=16.5):
    path = os.path.join(FIG, fname)
    if not os.path.exists(path):
        para("[figure missing: " + fname + "]", color=RED)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(caption)
    r.font.size = Pt(7.8); r.italic = True; r.font.color.rgb = GREY
    c.paragraph_format.space_after = Pt(8)


def callout(title, body, fill="FFF4E5", bordercol="E08C00"):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    shade(c, fill)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(1)
    r = p.add_run(title)
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x8A, 0x41, 0x00)
    p2 = c.add_paragraph()
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(body); r2.font.size = Pt(8.6)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)
    return t


def pagebreak():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


add_footer()
print("scaffold ready")

# =============================================================== COVER
para("COMPREHENSIVE LOT DUE-DILIGENCE REPORT", 20, bold=True, color=ACCENT,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_before=26, space_after=2)
para("1147 (Lot 130) and 1153 (Lot 131) Watersedge Cove", 14, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
para("Tignall, Lincoln County, Georgia 30668", 12,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
para("Two adjacent 1.07-acre homesites · 2.14 acres combined · Stillwater Coves",
     10, italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
para("Clarks Hill / J. Strom Thurmond Lake · Zillow ZPID 465607553",
     9.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

table(["Item", "Detail"],
      [["Report date", "7 October 2026"],
       ["Subject", "Pre-purchase land due diligence, 61-point checklist"],
       ["Reference point used", "33.94100 N, 82.56460 W (Watersedge Cove centreline)"],
       ["Jurisdiction", "Lincoln County, Georgia (FIPS 13181) — unincorporated"],
       ["Primary elevation source", "USGS 3DEP 1-metre lidar, GA_Statewide_2018, NAVD 88"],
       ["Primary flood source", "FEMA National Flood Hazard Layer, DFIRM 13181C"],
       ["Primary soil source", "USDA NRCS SSURGO via Soil Data Access"],
       ["Prepared by", "Automated desk study — not a survey, appraisal or legal opinion"]],
      widths=[4.6, 12.0], fontsize=9)

para("HOW TO READ THIS REPORT — EVIDENCE TIERS", 11, bold=True, color=ACCENT, space_before=14)
para("Every finding carries one of three labels. This distinction matters: a modelled contour "
     "is not a survey, and a soil-survey polygon is not a percolation test.", 9)
table(["Label", "Meaning", "Can you rely on it?"],
      [["**DOCUMENTED**", "Taken from an official or authoritative published source "
        "(federal/state agency dataset, statute, published rule, government website).",
        "Yes, subject to the source's own currency."],
       ["**MODEL-DERIVED ESTIMATE**", "Computed here from authoritative raw data "
        "(lidar elevation, soil polygons, reanalysis climate) using stated assumptions.",
        "Directionally reliable; NOT a substitute for survey/engineering."],
       ["**REQUIRES VERIFICATION**", "Not publicly obtainable, contradictory across sources, "
        "or legally requires a licensed professional or county sign-off.",
        "No. Treat as an open question."]],
      widths=[3.4, 8.2, 5.0], fontsize=8.2)

callout("CRITICAL LIMITATION — NO RECORDED PLAT WAS OBTAINED",
        "Lincoln County's parcel GIS services are access-restricted (HTTP 499 \"Token Required\"), and "
        "no public parcel polygon for Lot 130 or Lot 131 could be retrieved. Every statement in this "
        "report about lot boundaries, dimensions, frontage, depth, setbacks and buildable area rests on "
        "a MODELLED rectangle of 150 ft x 310 ft (1.07 ac). The recorded plat must be obtained from the "
        "Lincoln County Clerk of Superior Court before any of those figures are relied upon. "
        "Terrain, flood and soil findings are independent of that assumption and remain valid for the area.",
        fill="FDE7E9", bordercol="C01F2B")

pagebreak()

# =============================================================== EXEC SUMMARY
doc.add_heading("Executive Summary", level=1)
para("This report answers a 61-point due-diligence checklist for two adjacent 1.07-acre lots marketed "
     "as dockable waterfront on Clarks Hill (J. Strom Thurmond) Lake. The desk study used 1-metre "
     "federal lidar, the FEMA flood layer, the USDA soil survey, USACE shoreline policy and the "
     "road network to test the marketing claims against measurable data. Six findings materially "
     "affect value and buildability.", 9.5)

doc.add_heading("The six findings that matter", level=2)
table(["#", "Finding", "Why it matters", "Tier"],
      [["1", "**The lots are not at the water.** The road sits 385–399 ft; full pool is 330 ft. "
             "The 330 ft shoreline is 467–688 ft horizontally from the road and 53–65 ft below it. "
             "A 310 ft-deep lot ends at roughly 348–372 ft elevation.",
        "\"WATERFRONT\" and \"direct waterfront access\" cannot mean fee ownership to the water. "
        "There is almost certainly USACE-owned land between the rear lot line and the lake.", "M"],
       ["2", "**Dock rights are a federal permit, not a property right.** The Thurmond Shoreline "
             "Management Plan (1 May 2018) requires a 20 ft minimum shared boundary with federal land, "
             "and for subdivisions platted after 1 May 2018 authorises COMMUNITY DOCKS ONLY.",
        "\"DOCKABLE\" is an assertion until the Thurmond Project Office confirms the shoreline "
        "allocation and eligibility for this specific parcel. Permits are 5-year, non-transferable, "
        "and void on sale.", "D"],
       ["3", "**Part of the area is a Special Flood Hazard Area with no published BFE.** "
             "FEMA DFIRM 13181C shows Zone A along the lake margin and Zone X over the road corridor. "
             "Zone A carries no Base Flood Elevation.",
        "Without a BFE, the minimum finished-floor elevation must be set by the county floodplain "
        "administrator, and lenders/insurers may demand an Elevation Certificate.", "D"],
       ["4", "**The dominant soils are rated \"Very limited\" for septic and for dwellings.** "
             "Pacolet sandy clay loam 10–25% severely eroded, and Zion silt loam 10–25% "
             "(lithic bedrock at 77 cm), both rate Very limited; Cecil sandy loam 2–6% rates only "
             "Somewhat limited.",
        "There is no public sewer. A 4-bedroom conventional septic system may not be permittable "
        "across much of the slope. Georgia denies permits where bedrock is under 2 ft below the "
        "field bottom.", "D"],
       ["5", "**The asking price is unresolved.** Cached listings show $55,000, $169,900 and "
             "$144,000 (Lot 131 alone, off-market) under three different brokerages.",
        "A 3x spread means the lots have been listed, delisted and relisted. Do not negotiate from "
        "a cached number.", "V"],
       ["6", "**It is remote.** 17 mi to the county seat, 20 mi to a 25-bed critical-access "
             "hospital, 44 mi to a Walmart, 55 mi to full-service hospitals, 63 mi to the nearest "
             "commercial airport.",
        "Every daily necessity requires a car. Emergency medical response times and contractor "
        "availability should both be factored in.", "D"]],
      widths=[0.8, 6.0, 6.2, 0.9], fontsize=7.9)

doc.add_heading("Headline measurements", level=2)
table(["Parameter", "Value", "Source"],
      [["Combined area (as marketed)", "2.14 ac (2 x 1.07 ac)", "Listing"],
       ["Road (Watersedge Cove) elevation", "385.0 – 399.1 ft NAVD 88", "1 m lidar"],
       ["Lake full pool", "330 ft-msl", "USACE"],
       ["Vertical relief, road to full pool", "52.6 – 65.0 ft", "1 m lidar"],
       ["Horizontal distance, road to full pool", "467 – 688 ft (142 – 210 m)", "1 m lidar"],
       ["Average slope, road to lake", "9.0 – 11.7 % (5.1 – 6.7 deg)", "1 m lidar"],
       ["Steepest slope over any 50 ft", "18.6 – 24.5 % (10.5 – 13.7 deg)", "1 m lidar"],
       ["Study-area high / low point", "446.8 ft / 326.9 ft", "1 m lidar"],
       ["FEMA zones present", "A (SFHA, no BFE) and X (minimal)", "FEMA NFHL"],
       ["Dominant soils", "Pacolet, Cecil, Madison, Zion", "NRCS SSURGO"],
       ["Public sewer", "None — septic required", "Listings / county"],
       ["Annual mean temperature", "18.2 deg C", "ERA5 1995–2024"],
       ["Frost-free season", "~283 days", "ERA5 1995–2024"]],
      widths=[6.2, 5.6, 4.6], fontsize=8.4)

doc.add_heading("Verdict in one paragraph", level=2)
para("The land is physically buildable: average slopes of 6–13 % across the likely building envelope are "
     "ordinary for the Georgia Piedmont, the road corridor is outside the flood hazard area, public water "
     "is reported, and fibre broadband reaches parts of the ZIP code. What is NOT established is the "
     "premium the price implies. The two things a buyer is paying extra for — water frontage and a private "
     "dock — are the two things this desk study could not confirm, and federal policy since May 2018 has "
     "moved decisively against new private docks in newly platted subdivisions. Treat the purchase as a "
     "wooded upland homesite with lake proximity unless and until the Thurmond Project Office confirms "
     "dock eligibility in writing and the recorded plat confirms the rear boundary. Resolve the four "
     "gating items in Section M.58 before removing any due-diligence contingency.", 9.5)

pagebreak()

# =============================================================== SECTION A
doc.add_heading("A. Property Identification & Mapping", level=1)

doc.add_heading("1. Property identification", level=2)
table(["Attribute", "Finding", "Status"],
      [["Marketed address", "1147 (Lot 130) and 1153 (Lot 131) Watersedge Cove / "
        "Waters Edge Cv, Tignall, GA 30668", "Documented"],
       ["Lot numbers", "Lot 130 (addressed 1147) and Lot 131 (addressed 1153)", "Documented"],
       ["Subdivision / community", "Stillwater Coves", "Documented"],
       ["County", "**Lincoln County, Georgia** (FIPS 13181). Tignall is the postal address only — "
        "Tignall town lies in Wilkes County. The HOA's own website and county tax records place "
        "Stillwater Coves in Lincoln County.", "Documented"],
       ["City limits", "Unincorporated Lincoln County (not City of Lincolnton)", "Documented"],
       ["State", "Georgia", "Documented"],
       ["Parcel / APN number", "**NOT OBTAINED.** Lincoln County parcel GIS is token-restricted.",
        "Verify"],
       ["Lot acreage", "1.07 ac each; 2.14 ac combined (per listing; 1.07 ac = 46,609 sq ft)",
        "Listing — verify on plat"],
       ["Lot dimensions", "**NOT PUBLISHED.** Modelled herein as 150 ft x 310 ft.", "Verify"],
       ["Legal description", "**NOT OBTAINED** — in the recorded plat and deed.", "Verify"],
       ["Zillow ZPID", "465607553", "Documented"]],
      widths=[3.6, 10.0, 2.8], fontsize=8.3)
tier("V", "Parcel ID, dimensions and legal description must come from the recorded plat and deed.")

callout("COUNTY-OF-JURISDICTION WARNING",
        "The postal address says Tignall, which most people associate with Wilkes County. The parcel is in "
        "LINCOLN County. This determines which planning office, building department, tax assessor, "
        "environmental health office and school district apply. Confirm on the deed before filing any "
        "application — applying to the wrong county wastes weeks.")

doc.add_heading("2. Zillow / listing information", level=2)
para("Zillow actively blocks automated retrieval. Three attempts on the supplied URL returned "
     "HTTP 403 Forbidden, as did Redfin; realtor.com returned HTTP 429. No listing photographs, "
     "hero image or price history could be captured, and none are reproduced here "
     "(they are also third-party copyright material).", 9)
tier("V", "Price, status, photos and listing history must be read directly on Zillow or from the broker.")

para("Pricing signals recovered from search-engine caches are mutually inconsistent:", 9, bold=True)
table(["Source context", "Price shown", "Brokerage named", "Area"],
      [["realtor.com, Tignall homes under $100K", "$55,000", "Meybohm Real Estate – Evans", "2.14 ac"],
       ["realtor.com, Tignall land search", "$169,900", "—", "2.14 ac"],
       ["realtor.com, Tignall homes under $200K", "listed", "J Brand Realty, LLC", "2.14 ac"],
       ["realtor.com, 1153 Waters Edge Cv Lot 131 alone", "$144,000 (Off Market)", "—", "1.07 ac"],
       ["Agent profile page, 1147 Watersedge Cv", "$49,900", "—", "—"],
       ["Comparable: 1135 Watersedge Cv (Lot 128)", "~$145,250", "—", "1.07 ac"]],
      widths=[6.4, 3.4, 4.2, 2.4], fontsize=8.3)
para("Three brokerages and a roughly 3x price spread across cached pages indicate repeated "
     "list / delist / relist cycles rather than a single current price. The $55,000 and $169,900 "
     "figures both carry a 2.14-acre area, so both plausibly refer to this combined listing at "
     "different times.", 9)

para("Listing claims to be tested (verbatim claims paraphrased):", 9, bold=True)
bullets([("Two adjacent 1.07-acre lots, both dockable, direct waterfront access —",
          "tested in Sections B, F.24 and F.26. Not supported by terrain data as fee-owned frontage."),
         ("Gently sloping, mix of open space and mature hardwoods —",
          "broadly supported: 6–13 % average across the near-road third (Section B.5–B.7)."),
         ("Relatively flat with gradual slope to the lake once you get to the back corps line —",
          "this phrasing in the marketing material itself concedes a Corps boundary at the rear, "
          "consistent with finding #1."),
         ("Ideal for boating, fishing, swimming —",
          "lake recreation is genuinely excellent (Section F.28–F.29); private access is the open question.")])

doc.add_heading("3. Official plat & boundary", level=2)
para("No recorded plat was obtained. The following must be pulled and compared against each other:", 9)
bullets([("Recorded subdivision plat for Stillwater Coves —", "Lincoln County Clerk of Superior Court. "
          "Must show lot boundaries, adjacent lots, road right-of-way, the USACE boundary, easements "
          "and any platted building setback lines."),
         ("Plat recording date —", "DECISIVE for dock rights. Subdivisions platted after 1 May 2018 are "
          "limited to community docks under the Thurmond Shoreline Management Plan. Establish this date "
          "before anything else."),
         ("Deed and title commitment —", "legal description, easements, reservations, mineral rights, "
          "and any dock or shoreline covenants."),
         ("USACE boundary survey —", "the federal fee line adjacent to the rear boundary; obtain from "
          "the Thurmond Project Office.")])
tier("V", "Plat, boundary and easement verification requires county records and a licensed surveyor.")

doc.add_heading("4. GIS / Esri map", level=2)
para("Lincoln County publishes three ArcGIS feature services (parcels with WinGAP attributes, site "
     "address points, and road centrelines) through a CivicPlus-hosted organisation. All three returned "
     "error code 499 \"Token Required\" and are therefore not public. No county parcel polygon, parcel "
     "ID or assessed value could be retrieved programmatically.", 9)
para("Geographic reference actually used in this report, derived from OpenStreetMap road centreline "
     "data for Watersedge Cove (OSM way 409863360) and confirmed against the lidar surface:", 9)
table(["Feature", "Latitude (N)", "Longitude (W)", "Lidar elevation (ft NAVD 88)"],
      [["West cul-de-sac, head", "33.940373", "-82.566237", "392.7"],
       ["West cul-de-sac, north side", "33.940713", "-82.566049", "385.8"],
       ["Road station A", "33.940784", "-82.565716", "387.1"],
       ["Road station B", "33.940922", "-82.565062", "389.7"],
       ["Road station C (interpolated)", "33.941200", "-82.563800", "390.5"],
       ["Road station D", "33.941483", "-82.562494", "393.1"],
       ["Road east end", "33.941303", "-82.561594", "388.1"],
       ["Study-area high point", "~33.937", "~-82.564", "446.8"],
       ["Study-area low point (lake bed)", "~33.943", "~-82.566", "326.9"]],
      widths=[5.6, 3.3, 3.3, 4.3], fontsize=8.3)
para("Discrepancies identified between sources:", 9, bold=True)
bullets([("Street name —", "rendered variously as \"Watersedge Cv\", \"Waters Edge Cv\" and "
          "\"Watersedge Cove\". OSM and the county use Watersedge Cove."),
         ("Postal town vs. county —", "Tignall (Wilkes County post office) vs. Lincoln County parcel."),
         ("Address-to-lot mapping —", "1147 is given as Lot 130 and 1153 as Lot 131, but some cached "
          "pages label 1153 as \"Lot 131\" and others omit lot numbers. Confirm on the plat."),
         ("ZIP code —", "OSM returns 29840 (a South Carolina ZIP) for the road segment; listings use "
          "30668. The 29840 value is an OSM data error.")])
tier("V", "County GIS is not publicly accessible; all parcel-level attributes remain unverified.")

pagebreak()

# =============================================================== SECTION B
doc.add_heading("B. Topography & Terrain", level=1)
para("This is the strongest part of the dataset. The site is covered by USGS 3DEP 1-metre lidar "
     "(GA_Statewide_2018_B18_DRRA, vertical datum NAVD 88). A 2,665-point grid at roughly 15-metre "
     "spacing was sampled across a 1,070 m x 670 m study area and used for every figure and slope "
     "figure below.", 9)

doc.add_heading("5. True-scale topographic map", level=2)
figure("fig1_topo.png",
       "Fig. 1 — True-scale topographic map. Contours at 5 ft; the heavy blue line is the 330 ft "
       "full-pool shoreline; arrows show down-gradient drainage. The red rectangles are MODELLED "
       "lot outlines (150 x 310 ft), not surveyed boundaries.")
table(["Map element", "Value"],
      [["Elevation range in study area", "326.9 ft to 446.8 ft NAVD 88 (relief 119.9 ft)"],
       ["Highest point", "446.8 ft, on the ridge south of Watersedge Cove"],
       ["Lowest point", "326.9 ft, in the Fishing Creek arm north of the road (below full pool)"],
       ["Contour interval", "5 ft (labelled every 10 ft)"],
       ["Full-pool shoreline", "330 ft-msl (heavy blue)"],
       ["Flood-control pool", "335 ft-msl, top of spillway gates (dotted blue)"],
       ["Drainage direction", "Predominantly north / north-west, from the ridge toward the "
        "Fishing Creek arm of Thurmond Lake"],
       ["Grid SW corner", "33.93720 N, 82.56950 W"],
       ["Grid NE corner", "33.94320 N, 82.55990 W"]],
      widths=[5.4, 11.2], fontsize=8.4)
tier("M", "Contours are lidar-derived at roughly 15 m sample spacing. Not a boundary or topographic survey.")

doc.add_heading("6. Longitudinal terrain cross-section (road to lake)", level=2)
para("The checklist requires that any vertically exaggerated section be shown alongside a true 1:1 "
     "section, so that an exaggerated profile is not mistaken for real steepness. Both are given below.", 9)
figure("fig2_section.png",
       "Fig. 2 — Road-to-lake longitudinal section at road station A. Panel (a) is vertically "
       "exaggerated about 5:1, the convention used on most maps. Panel (b) is the same ground at "
       "true 1:1 scale. The visual difference is the point: this is a gentle to moderate slope, "
       "not a cliff.", width_cm=16.0)

para("Measured section statistics at six stations:", 9, bold=True)
table(["Station", "Road elev (ft)", "Shore elev (ft)", "Run (ft)", "Fall (ft)",
       "Avg slope", "Avg deg", "Max 50 ft", "Max deg"],
      [["W cul-de-sac head", "392.7", "332.4", "595", "60.3", "10.1 %", "5.8", "24.5 %", "13.7"],
       ["W cul-de-sac N side", "385.0", "332.4", "471", "52.6", "11.2 %", "6.4", "19.4 %", "11.0"],
       ["Road station A", "387.3", "332.4", "467", "54.9", "11.7 %", "6.7", "18.6 %", "10.5"],
       ["Road station B", "389.7", "332.4", "523", "57.3", "11.0 %", "6.3", "19.1 %", "10.8"],
       ["Road station C", "390.5", "328.7", "688", "61.8", "9.0 %", "5.1", "22.4 %", "12.7"],
       ["Road station D", "393.2", "328.7", "627", "65.0", "10.3 %", "5.9", "21.8 %", "12.3"]],
      widths=[3.5, 1.9, 1.9, 1.5, 1.5, 1.7, 1.3, 1.7, 1.3], fontsize=7.8)

para("Within the first 310 ft back from the road — the depth a 1.07-acre lot can actually occupy:", 9, bold=True)
table(["Position along frontage", "Elev at road (ft)", "Elev at 310 ft back (ft)", "Drop (ft)",
       "Avg slope", "Steepest 50 ft"],
      [["West", "387.3", "347.5", "39.7", "12.8 % (7.3 deg)", "15.2 % (8.7 deg)"],
       ["Middle", "389.7", "357.0", "32.7", "10.5 % (6.0 deg)", "18.0 % (10.2 deg)"],
       ["East", "390.5", "371.9", "18.6", "6.0 % (3.4 deg)", "8.5 % (4.8 deg)"]],
      widths=[3.6, 2.8, 3.3, 1.9, 2.6, 2.4], fontsize=8.2)

para("Depth back from the road at which the ground first reaches a given elevation — the single most "
     "important table in this report:", 9, bold=True)
table(["Target elevation", "West station", "Middle station", "East station"],
      [["380 ft", "60 ft", "135 ft", "145 ft"],
       ["370 ft", "140 ft", "195 ft", "345 ft"],
       ["360 ft", "210 ft", "260 ft", "515 ft"],
       ["350 ft", "290 ft", "800 ft", "beyond 900 ft"],
       ["340 ft", "490 ft", "beyond 900 ft", "beyond 900 ft"],
       ["335 ft (top of gates)", "595 ft", "beyond 900 ft", "beyond 900 ft"],
       ["**330 ft (FULL POOL)**", "**720 ft**", "**beyond 900 ft**", "**beyond 900 ft**"]],
      widths=[4.4, 4.0, 4.0, 4.2], fontsize=8.4)
para("Read the last row carefully. A lot 310 ft deep cannot reach water that is 720 ft or more away. "
     "Even a narrow, deep lot shape would have to be about 65 ft wide to reach full pool at 1.07 acres. "
     "The inescapable conclusion is that the rear boundary stops on the USACE line, well above and well "
     "back from the water.", 9, bold=True)
tier("M", "Derived from 1 m lidar along due-north transects. Actual lot orientation may differ; "
          "the order of magnitude will not.")

doc.add_heading("7. Terrain classification", level=2)
figure("fig4_slope.png",
       "Fig. 4 — Slope classification. Green is under 8 % (buildable with minimal grading); "
       "orange 10–20 %; red above 20 % (needs engineered solutions).")
table(["Classification band", "Applies to", "Assessment"],
      [["Flat (0–2 %)", "Very little of the site", "Not characteristic"],
       ["Nearly flat (2–5 %)", "Narrow benches near the road crown and the east frontage",
        "Present but limited"],
       ["**Gentle slope (5–10 %)**", "Much of the first 150–200 ft back from the road, "
        "and the east third of the frontage", "**Dominant in the building envelope**"],
       ["**Moderate slope (10–15 %)**", "West and middle frontage over the full 310 ft depth; "
        "general road-to-lake average", "**Dominant overall**"],
       ["Steep slope (15–25 %)", "Localised pitches; steepest measured 50 ft window is 24.5 %",
        "Present in bands, mostly below the likely rear lot line"],
       ["Very steep (over 25 %)", "Not measured anywhere in the study grid", "Absent"]],
      widths=[4.2, 6.4, 6.0], fontsize=8.3)
para("Construction implications:", 9, bold=True)
bullets([("Minimal grading —", "achievable only if the house is pulled close to the road on the east "
          "or middle frontage, where the first 100 ft averages under 8 %."),
         ("Significant grading —", "likely for any layout pushing more than about 150 ft back on the "
          "west frontage, where average slope is 12.8 %."),
         ("Cut-and-fill —", "expect a balanced cut-and-fill pad or a stepped foundation. A single-level "
          "slab across a 56 ft-deep footprint on 12 % ground implies roughly 6–7 ft of grade change "
          "across the pad."),
         ("Retaining walls —", "probable on the downhill side of a driveway or pad on the west frontage; "
          "possibly avoidable on the east frontage."),
         ("Extensive site engineering —", "not indicated by slope alone, but the soil ratings in "
          "Section D.14 push the overall requirement up.")])
tier("M", "Slope classification is lidar-derived. A licensed engineer must set actual grading and "
          "retaining requirements.")

doc.add_heading("8. 3D terrain model", level=2)
figure("fig3_3d.png",
       "Fig. 3 — 3-D terrain surface with the 330 ft full-pool plane in blue and Watersedge Cove in "
       "dark red. Vertical scale is exaggerated to make the cove geometry legible; refer to Fig. 2(b) "
       "for true steepness.")
para("The model shows the subdivision road running along the upper flank of a broad ridge, with the "
     "Fishing Creek arm of the lake occupying the low ground to the north and north-west. The lots sit "
     "on the north-facing slope between the two. A north aspect means less direct winter sun on the "
     "rear elevation and a cooler, damper micro-climate on the lower slope, which is relevant to both "
     "gardening (Section D.18) and to drying of the septic field.", 9)
tier("M", "Surface is interpolated from the sampled lidar grid.")

pagebreak()
print("sections A-B written")

# =============================================================== SECTION C
doc.add_heading("C. Flood, Elevation & Drainage", level=1)

doc.add_heading("9. FEMA flood-zone determination", level=2)
figure("fig5_fema.png",
       "Fig. 5 — FEMA National Flood Hazard Layer, DFIRM 13181C (Lincoln County, GA). Blue is "
       "Zone A (Special Flood Hazard Area, no BFE published); pale green is Zone X.")
table(["Attribute", "Value returned by the FEMA NFHL"],
      [["DFIRM identifier", "13181C (Lincoln County, Georgia)"],
       ["Zones intersecting the study area", "**Zone A** and **Zone X**"],
       ["Zone A — subtype", "None (no subtype assigned)"],
       ["Zone A — SFHA flag", "T (true — it IS a Special Flood Hazard Area)"],
       ["Zone A — STATIC_BFE", "**-9999, i.e. NO Base Flood Elevation published**"],
       ["Zone A — DEPTH", "-9999 (none published)"],
       ["Zone X — subtype", "AREA OF MINIMAL FLOOD HAZARD"],
       ["Zone X — SFHA flag", "F (false — outside the SFHA)"],
       ["Where Zone A sits", "The lake margin and the Fishing Creek arm, north/north-west of the road"],
       ["Where Zone X sits", "The road corridor and the upland, including the likely building envelope"]],
      widths=[5.4, 11.2], fontsize=8.4)
para("Zones NOT present: AE, AO, AH, VE. Zone VE does not occur inland. The absence of AE is "
     "significant — AE zones carry a published BFE, Zone A does not.", 9)
callout("WHAT \"ZONE A WITH NO BFE\" ACTUALLY MEANS FOR YOU",
        "Zone A is an approximate-study floodplain: FEMA has mapped the 1 %-annual-chance flood "
        "boundary but has not computed its elevation. Consequences: (1) the Lincoln County floodplain "
        "administrator, not FEMA, sets the required finished-floor elevation; (2) lenders and insurers "
        "will usually demand an Elevation Certificate prepared by a licensed surveyor; (3) flood "
        "insurance rating in an unnumbered A zone can be materially worse than in a numbered zone; "
        "(4) if any part of the lot or the proposed structure falls in Zone A, a site-specific "
        "engineering analysis may be required. Confirm the exact zone line against the recorded plat.")
tier("D", "FEMA National Flood Hazard Layer, queried directly 7 October 2026. "
          "Panel number and effective date must still be read off the printed FIRM panel.")
para("Still to obtain: the specific FIRM panel number and its effective date, plus the Flood Insurance "
     "Study report for Lincoln County (FEMA document 13181CV000A), and confirmation of whether the "
     "county participates in the NFIP in good standing.", 9, italic=True)

doc.add_heading("10. First Street / flood-risk information", level=2)
para("First Street Foundation data is licensed and was not retrievable without a subscription; no "
     "First Street flood factor, projected depth or 30-year risk figure is reported here. "
     "What can be said from the physical data is that the building envelope sits 48–65 ft above full "
     "pool and 18–42 ft above the top of the flood-control pool (335 ft), which is a very large "
     "freeboard by any standard.", 9)
tier("V", "First Street flood factor, flood depth and future projections not obtained.")

doc.add_heading("11. Ground elevations", level=2)
table(["Measure", "Value (ft NAVD 88)", "Relative to 330 ft full pool"],
      [["Study-area highest natural ground", "446.8", "+116.8 ft"],
       ["Road crown range along the frontage", "385.0 – 399.1", "+55.0 to +69.1 ft"],
       ["Rear of a 310 ft-deep lot, west frontage", "347.5", "+17.5 ft"],
       ["Rear of a 310 ft-deep lot, middle frontage", "357.0", "+27.0 ft"],
       ["Rear of a 310 ft-deep lot, east frontage", "371.9", "+41.9 ft"],
       ["Approximate mean of the building envelope", "~375 – 385", "+45 to +55 ft"],
       ["Study-area lowest ground (lake bed)", "326.9", "-3.1 ft (submerged at full pool)"]],
      widths=[6.4, 4.2, 6.0], fontsize=8.4)
tier("M", "1 m lidar, NAVD 88. An Elevation Certificate requires a licensed surveyor.")

doc.add_heading("12. BFE — Base Flood Elevation", level=2)
table(["Question", "Answer"],
      [["Does an official BFE exist for this location?",
        "**No.** FEMA returns STATIC_BFE = -9999 for the Zone A polygon."],
       ["Applicable FEMA flood elevation",
        "None published. The regulatory flood elevation must be determined by the county."],
       ["Useful federal reference elevations instead",
        "Full pool 330 ft; top of spillway gates (flood-control pool) 335 ft; "
        "conservation pool 312–330 ft; normal drawdown limit 305 ft; spillway crest 300 ft."],
       ["Relationship of the lot to those elevations",
        "The entire modelled building envelope lies above 345 ft, i.e. at least 10 ft above the "
        "top of the flood-control pool and at least 15 ft above full pool."],
       ["Difference between lowest likely ground on the lot and 335 ft",
        "Approximately +12.5 ft at the west rear corner, the least favourable modelled position."]],
      widths=[6.0, 10.6], fontsize=8.4)
para("Note the distinction: the reservoir's flood-control pool (335 ft) is an operational USACE "
     "elevation, not a FEMA regulatory BFE. They are different instruments and should not be "
     "substituted for one another in a permit application.", 9, italic=True)
tier("D", "USACE pool elevations documented. BFE genuinely does not exist for Zone A here.")

doc.add_heading("13. Finished-floor elevation", level=2)
para("Because no BFE is published, the minimum finished-floor elevation cannot be computed from "
     "FEMA data. The governing formula is:", 9)
para("        regulatory flood elevation  +  county freeboard  =  minimum finished-floor elevation",
     9.5, bold=True, color=ACCENT)
para("Georgia's model floodplain ordinance commonly applies 1–3 ft of freeboard, but Lincoln County's "
     "adopted figure must be read from its own ordinance. Given that the building envelope sits "
     "45–55 ft above full pool and comfortably outside Zone A, the practical answer is:", 9)
table(["Determination", "Finding"],
      [["Is existing terrain adequate?",
        "**Almost certainly yes** for a house sited in the Zone X portion near the road. "
        "No flood-driven elevation of the structure is indicated."],
       ["Is fill necessary for flood reasons?",
        "**No.** Fill may still be wanted for grading/pad reasons (Section K.44), which is a "
        "different question."],
       ["Must the house be elevated?",
        "Not for flood purposes, on current evidence."],
       ["Crawlspace / pier / slab restrictions?",
        "No flood-driven restriction expected outside the SFHA. Slope and soil drive this choice "
        "instead — see Section D.15."],
       ["If any part of the structure falls in Zone A",
        "Then the full floodplain-development permit path applies: county-determined flood elevation, "
        "freeboard, flood vents, and an Elevation Certificate."]],
      widths=[5.6, 11.0], fontsize=8.4)
callout("DO NOT ASSUME FILL IS PERMITTED",
        "Raising a pad above a flood elevation does not by itself make a project compliant. Placing fill "
        "in or near a floodplain, or within the Corps' flowage easement if one exists, can require "
        "separate authorisation and can be refused outright. Lincoln County floodplain administration "
        "and, where federal land or easement is involved, the USACE Thurmond Project Office must both "
        "be consulted before any fill is designed.")
tier("V", "Freeboard value and regulatory flood elevation must come from the Lincoln County "
          "floodplain administrator.")

doc.add_heading("Drainage note (expanded in Section K.45)", level=2)
para("The down-gradient arrows in Fig. 1 show a consistent north to north-west drainage pattern from "
     "the ridge, across the road corridor, down the lot slope and into the Fishing Creek arm. Two "
     "shallow swales are visible in the lidar surface on either side of the modelled lot pair. Any "
     "building pad placed across a swale will intercept upslope runoff and will need positive "
     "diversion. Because the lots sit below a ridge, they receive the runoff of the land above them; "
     "the volume is modest given the short contributing slope length, but it is not zero.", 9)
tier("M", "Flow directions computed from the lidar gradient field.")

pagebreak()

# =============================================================== SECTION D
doc.add_heading("D. Soil, Foundation, Septic & Garden", level=1)
para("Soil data below is USDA NRCS SSURGO, queried through Soil Data Access for the exact study-area "
     "polygon. Map units were then sampled at individual points along the road and slope to establish "
     "which soil sits where.", 9)

doc.add_heading("14. Soil characteristics", level=2)
para("Map units present in the study area, and where each one falls:", 9, bold=True)
table(["Map unit (mukey)", "Soil name", "Slope range", "Drainage", "Where sampled"],
      [["127112", "**Pacolet sandy clay loam, severely eroded**", "10–25 % (rep. 18 %)",
        "Well drained", "West cul-de-sac; mid-slope toward the lake"],
       ["127068", "**Cecil sandy loam**", "2–6 % (rep. 4 %)", "Well drained",
        "Middle and east road frontage"],
       ["127093", "Madison sandy loam", "2–6 % (rep. 4 %)", "Well drained", "Within study area"],
       ["127095", "Madison sandy loam", "10–25 % (rep. 18 %)", "Well drained", "Within study area"],
       ["127135", "**Zion silt loam**", "10–25 % (rep. 18 %)", "Well drained",
        "Lower slope near the lake arm"],
       ["127133", "Water", "—", "—", "The lake itself"]],
      widths=[2.4, 5.0, 2.8, 2.4, 4.0], fontsize=8.1)

para("Horizon-level properties of the major components:", 9, bold=True)
table(["Soil", "Depth (cm)", "Sand %", "Silt %", "Clay %", "OM %", "Ksat (um/s)", "pH (1:1)", "AWC"],
      [["Cecil", "0–15", "69.9", "20.2", "9.9", "0.50", "28", "5.3", "0.13"],
       ["Cecil", "15–107", "29.6", "15.6", "**54.8**", "0.10", "9", "5.0", "0.10"],
       ["Cecil", "107–203", "45.0", "20.1", "34.9", "0.10", "9", "5.0", "0.15"],
       ["Pacolet", "0–8", "58.0", "16.0", "26.0", "1.00", "9", "5.5", "0.11"],
       ["Pacolet", "8–48", "48.0", "12.0", "**40.0**", "0.25", "9", "5.0", "0.13"],
       ["Pacolet", "48–203", "60–65", "15–20", "15–25", "0.25", "9", "5.0", "0.12"],
       ["Madison", "0–13", "67.9", "19.6", "12.5", "1.25", "28", "**4.8**", "0.13"],
       ["Madison", "13–64", "51.8", "8.2", "40.0", "0.25", "9", "**4.8**", "0.16"],
       ["Zion", "0–25", "29.1", "**53.4**", "17.5", "1.25", "9", "5.3", "0.17"],
       ["Zion", "25–56", "23.3", "29.2", "**47.5**", "0.25", "**2**", "5.9", "0.15"],
       ["Zion", "56–89", "29.5", "28.0", "42.5", "0.25", "8", "6.2", "0.11"]],
      widths=[1.8, 2.0, 1.5, 1.5, 1.7, 1.4, 2.0, 1.7, 1.2], fontsize=7.8)

table(["Property", "Finding"],
      [["Soil series present", "Cecil, Pacolet, Madison (all Typic Kanhapludults, fine/kaolinitic, "
        "thermic) and Zion (Ultic Hapludalfs, fine, mixed, active)"],
       ["Texture", "Sandy loam to sandy clay loam surfaces over clay-rich subsoils "
        "(40–55 % clay). Zion has a silt loam surface."],
       ["Drainage class", "All major components **well drained**. None hydric."],
       ["Permeability (Ksat)", "Surface 28 um/s in Cecil/Madison, dropping to 9 um/s in subsoils. "
        "**Zion's 25–56 cm horizon is 2 um/s — very slow.**"],
       ["Bearing characteristics", "Kaolinitic clays give moderate, generally adequate bearing for "
        "residential loads; Zion's mixed/active mineralogy is less favourable."],
       ["Shrink/swell potential", "**Low** in Cecil, Pacolet and Madison (kaolinitic). "
        "**Higher in Zion** (mixed, active clay mineralogy)."],
       ["Depth to restrictive layer", "None recorded for Cecil, Pacolet or Madison. "
        "**Zion: lithic (indurated) bedrock at 77 cm — about 2.5 ft.**"],
       ["Depth to groundwater", "No seasonal high water table recorded for the major components "
        "(consistent with well-drained upland Piedmont soils)."],
       ["Erosion potential", "**High.** The Pacolet unit is explicitly mapped as "
        "\"severely eroded\" on 10–25 % slopes, and surface organic matter is low (0.25–1.25 %)."],
       ["Suitability for residential construction", "Mixed — see the ratings table immediately below."]],
      widths=[4.6, 12.0], fontsize=8.3)

para("NRCS engineering interpretations (the most decision-relevant table in this section):", 9, bold=True)
table(["Soil / map unit", "Dwellings w/o basement", "Dwellings w/ basement",
       "Septic absorption fields", "Shallow excavations", "Local roads & streets"],
      [["**Cecil, 2–6 %**", "Not limited", "Not limited", "Somewhat limited",
        "Somewhat limited", "Somewhat limited (GA)"],
       ["**Madison, 2–6 %**", "Not limited", "Not limited", "Somewhat limited",
        "Somewhat limited", "Somewhat limited (GA)"],
       ["**Madison, 10–25 %**", "**Very limited**", "**Very limited**", "**Very limited**",
        "**Very limited**", "**Very limited**"],
       ["**Pacolet, 10–25 %, severely eroded**", "**Very limited**", "**Very limited**",
        "**Very limited**", "**Very limited**", "**Very limited**"],
       ["**Zion, 10–25 %**", "**Very limited**", "**Very limited**", "**Very limited**",
        "**Very limited**", "**Very limited**"]],
      widths=[3.6, 2.6, 2.6, 2.6, 2.6, 2.6], fontsize=7.8)
para("Translation: the gentler Cecil and Madison units near the road are good building ground. "
     "Everything on the 10–25 % slope — which is where the lots extend toward the lake — carries the "
     "worst NRCS rating for houses, septic fields, excavations and roads simultaneously. The practical "
     "effect is to compress the usable part of each lot toward the road.", 9, bold=True)
tier("D", "USDA NRCS SSURGO via Soil Data Access, queried 7 October 2026. "
          "NRCS itself states survey-scale maps are insufficient for septic approval.")

doc.add_heading("15. Foundation suitability", level=2)
table(["Foundation type", "Assessment on this site", "Comment"],
      [["Slab-on-grade", "Feasible only on the flatter near-road bench",
        "On 10–13 % ground a slab needs a cut-and-fill pad; expect 6–7 ft of grade change across a "
        "56 ft-deep footprint."],
       ["Crawlspace", "**Likely the most economical fit**",
        "Absorbs moderate slope cheaply, keeps floor framing above grade, and suits the "
        "well-drained soils. Requires vented, properly graded perimeter."],
       ["Pier / beam", "Feasible", "Good on the steeper rear portion if the house is pushed back; "
        "minimises cut and preserves trees."],
       ["Basement (walk-out)", "Feasible and attractive on 10–13 % north-facing slope",
        "**But** NRCS rates the slope units \"Very limited\" for basements, and Cecil/Pacolet "
        "subsoils are 40–55 % clay — drainage and waterproofing must be engineered."],
       ["Stepped foundation", "Likely required if the footprint runs across the contour",
        "Orienting the long axis parallel to the contour reduces this substantially."],
       ["Retaining walls", "Probable on the west frontage, possibly avoidable on the east",
        "Walls over 4 ft generally require engineered design and a permit."]],
      widths=[3.0, 5.2, 8.4], fontsize=8.2)
para("A professional geotechnical investigation is advisable, and arguably necessary, for three "
     "specific reasons: (a) the Pacolet unit is severely eroded, so the residual horizon sequence on "
     "site may differ from the survey description; (b) Zion's lithic bedrock at about 2.5 ft would "
     "change both excavation cost and septic feasibility if it extends into the lot; and (c) the "
     "40–55 % clay subsoils control both bearing and drainage design.", 9)
tier("V", "Foundation design requires a licensed geotechnical engineer and a site-specific soil report.")

doc.add_heading("16. Septic suitability", level=2)
table(["Question", "Finding"],
      [["Is public sewer available?", "**No.** No sewer is reported for Stillwater Coves; "
        "listings describe public water only."],
       ["Is septic required?", "**Yes** — an on-site sewage management system is the only option."],
       ["Governing rules", "Georgia Rules for On-Site Sewage Management Systems, Chapter 511-3-1, "
        "administered by the county environmental health office under the Georgia DPH manual."],
       ["Soil suitability", "**Poor to marginal over most of the slope.** Pacolet, Zion and "
        "Madison 10–25 % all rate \"Very limited\" for absorption fields. Cecil 2–6 % and "
        "Madison 2–6 % rate \"Somewhat limited\"."],
       ["Perc / soil evaluation requirement", "A site-specific soil evaluation is mandatory. "
        "Georgia practice requires an original, stamped Level 3 or Level 4 soil report; NRCS survey "
        "maps are explicitly stated to be insufficient for approval."],
       ["Statutory disqualifiers", "Permits for conventional or chamber systems are **denied** where "
        "bedrock or impervious strata lie less than 2 ft below the absorption-field bottom, or where "
        "seasonal high groundwater is less than 2 ft below it (1 ft if Class I aerobic pretreatment "
        "is used)."],
       ["Specific risk here", "**Zion's indurated bedrock at 77 cm (2.5 ft)** would leave essentially "
        "no compliant depth beneath a field. If Zion extends into the rear of the lots, that ground "
        "is unusable for a conventional field."],
       ["Slow-permeability risk", "Zion's 25–56 cm horizon at Ksat 2 um/s is very slow; "
        "clay subsoils elsewhere sit at 9 um/s."],
       ["Does slope create problems?", "**Yes.** Absorption fields on 10–25 % slopes risk lateral "
        "breakout downslope, and the \"Very limited\" rating reflects exactly this."],
       ["Setback from the lake / water", "A substantial setback from the shoreline and from the "
        "USACE boundary will apply, in addition to well and property-line setbacks. "
        "The specific distances must be confirmed with county environmental health."]],
      widths=[4.6, 12.0], fontsize=8.3)
callout("BEDROOM COUNT IS A SOIL QUESTION, NOT AN ARCHITECTURAL ONE",
        "Georgia sizes absorption fields by bedroom count and soil group. A common county standard is "
        "roughly 3,000 sq ft of usable soil area per bedroom for the primary field plus the same again "
        "for a reserve field — about 24,000 sq ft of usable soil for a 4-bedroom house. A 1.07-acre lot "
        "contains 46,609 sq ft gross, so a 4-bedroom conventional system is arithmetically possible on a "
        "single lot but leaves little margin once the house, driveway, setbacks and the \"Very limited\" "
        "slope areas are excluded. THIS IS THE STRONGEST ARGUMENT FOR BUYING BOTH LOTS TOGETHER. "
        "A Level 3/4 soil report will give the real answer, and it may come back at 2 or 3 bedrooms "
        "for a conventional system, with 4 bedrooms achievable only via an engineered or "
        "advanced-treatment system at materially higher cost.")
tier("V", "Permittable bedroom count requires a stamped Level 3/4 soil report and county "
          "environmental health approval. Any number quoted before that is speculation.")

doc.add_heading("17. Gardening suitability", level=2)
table(["Property", "Measured value", "Gardening implication"],
      [["Surface pH", "4.8 (Madison) to 5.5 (Pacolet); 5.3 (Cecil, Zion)",
        "**Strongly to moderately acidic.** Most vegetables want 6.0–6.8. Liming is essential."],
       ["Surface organic matter", "0.50 % (Cecil), 1.00 % (Pacolet), 1.25 % (Madison, Zion)",
        "**Very low.** Target for productive beds is 3–5 %. Heavy compost addition needed."],
       ["Surface texture", "Sandy loam (Cecil, Madison, Pacolet); silt loam (Zion)",
        "Workable surface, easy to cultivate."],
       ["Subsoil clay", "40–55 % at 15–64 cm",
        "Root restriction and perched moisture below the surface horizon."],
       ["Available water capacity", "0.10–0.17 (low to moderate)",
        "Sandy surfaces drain fast; irrigation will be needed in summer."],
       ["Drainage", "Well drained throughout", "Good — no waterlogging risk on the upland."],
       ["Erosion", "Severely eroded Pacolet on 10–25 % slopes",
        "Bare cultivated soil on slope will erode. Terrace, mulch or use beds."]],
      widths=[3.2, 4.8, 8.6], fontsize=8.2)
para("Raised beds are advisable, for three converging reasons: the native surface organic matter is "
     "under 1.3 %, the subsoil is clay-rich and root-restricting, and the ground slopes enough that "
     "level beds also serve as erosion control. Imported topsoil and compost should be budgeted — "
     "do not plan on amending the native profile alone to vegetable-garden standard in one season.", 9)
tier("D", "Soil chemistry from SSURGO. A laboratory soil test through UGA Extension is the "
          "necessary next step before any amendment is purchased.")

doc.add_heading("18. Gardening and farming — what will actually grow", level=2)
para("Site climate, computed from ERA5 reanalysis at the exact coordinates for 1995–2024: "
     "annual mean 18.2 deg C, mean daily maximum 22.9 deg C, mean daily minimum 13.4 deg C, "
     "about 283 frost-free days, mean last spring frost 26 February and mean first autumn frost "
     "5 December, annual rainfall 1,070 mm. This is a long, hot, humid growing season with a mild "
     "winter — consistent with USDA hardiness zone 8a/8b.", 9)
figure("fig8_climate.png",
       "Fig. 8 — Monthly temperature and precipitation at the site, ERA5 reanalysis 1995–2024.",
       width_cm=15.0)
table(["Category", "Suitable plants", "Planting season", "Soil requirements"],
      [["**Vegetables (cool season)**",
        "Collards, kale, cabbage, broccoli, turnip, mustard, spinach, lettuce, carrot, beet, "
        "English pea, onion, garlic",
        "Late Aug – Oct for autumn/winter; Feb – Mar for spring",
        "Lime to pH 6.0–6.5; 5 cm compost; raised beds on slope"],
       ["**Vegetables (warm season)**",
        "Tomato, pepper, aubergine, okra, sweet potato, field pea, pole/bush bean, squash, "
        "cucumber, melon, sweetcorn",
        "After mid-Apr (soil above 15 deg C) through Jul",
        "pH 6.0–6.8; heavy compost; mulch and irrigation essential in Jun–Aug"],
       ["**Fruits**",
        "Rabbiteye blueberry (thrives at native pH 4.8–5.3), muscadine grape, fig, blackberry, "
        "persimmon, pomegranate; peach and plum with spray programme",
        "Dormant planting Dec – Feb",
        "**Blueberry is the standout — it wants the acid soil you already have.** "
        "Others need liming."],
       ["**Herbs**",
        "Rosemary, thyme, oregano, sage, basil (summer), parsley, chives, mint, lemongrass",
        "Spring after frost; perennials in autumn",
        "Sharp drainage; the sandy loam surface suits Mediterranean herbs well"],
       ["**Trees (shade / ornamental / nut)**",
        "Native oak, hickory, tulip poplar (already present), dogwood, redbud, crape myrtle, "
        "pecan, Chinese chestnut",
        "Dormant season Nov – Mar",
        "Tolerant of native acidity; preserve existing hardwoods for shade and erosion control"],
       ["**Not recommended**",
        "Apple and sweet cherry (insufficient winter chill, high disease pressure), "
        "most cool-climate brassica cultivars in summer",
        "—", "—"]],
      widths=[2.8, 5.4, 3.4, 5.0], fontsize=7.9)
para("Likely amendments, in priority order, based on the measured soil data:", 9, bold=True)
table(["Amendment", "Why this site needs it", "Priority"],
      [["**Agricultural lime**", "Surface pH 4.8–5.5 against a 6.0–6.8 target for most vegetables. "
        "Dolomitic lime also supplies magnesium.", "**Highest** — but rate MUST come from a soil test"],
       ["**Compost / organic matter**", "Surface OM is 0.50–1.25 % against a 3–5 % target.",
        "**Highest**"],
       ["Imported topsoil", "For raised beds, given the thin, eroded, low-OM native surface", "High"],
       ["Mulch", "Erosion control on 10–25 % slopes, plus summer moisture retention "
        "(52 days/yr above 32 deg C)", "High"],
       ["Balanced fertiliser", "Low native OM means low nutrient reserve and low CEC", "Moderate"],
       ["Sulphur", "**Only** for blueberry beds, to hold pH down near 4.5–5.5. "
        "Do not apply site-wide — the soil is already too acidic.", "Selective / low"]],
      widths=[3.2, 9.4, 4.0], fontsize=8.1)
para("Do not buy amendments on the strength of this table alone. The lime requirement in particular "
     "depends on buffer pH, which only a laboratory test can supply. UGA Extension soil testing is "
     "inexpensive and is the correct first expenditure.", 9, italic=True)
tier("M", "Plant lists are inferred from measured site climate and soil chemistry, cross-checked "
          "against standard Georgia Piedmont practice. Confirm with UGA Extension for Lincoln County.")

pagebreak()
print("sections C-D written")

# =============================================================== SECTION E
doc.add_heading("E. Utilities & Infrastructure", level=1)
callout("UTILITY SECTION HEALTH WARNING",
        "Every utility statement below rests on marketing copy for OTHER lots in Stillwater Coves, not on "
        "a utility-company confirmation for Lot 130/131. Note also that several listing descriptions in "
        "this ZIP code describe POINTE SHORES, a different community with different infrastructure "
        "(\"underground fiber, power, county garbage and water\"). Do not transfer Pointe Shores amenities "
        "to Stillwater Coves. Written availability letters from each provider, addressed to the specific "
        "parcel, are the only reliable evidence.")

doc.add_heading("19. Water", level=2)
table(["Question", "Finding", "Tier"],
      [["Municipal water", "Not applicable — outside Lincolnton city limits", "D"],
       ["Community / public water", "**Reported available.** Stillwater Coves marketing for other "
        "lots describes the community as having \"paved roads, underground utilities, public water\".", "V"],
       ["Private well", "Would be the fallback if public water does not actually reach the lot", "V"],
       ["Connection at the road", "Unconfirmed for this parcel", "V"],
       ["Connection requirements / cost", "Not published — tap fee, meter fee and service-line cost "
        "must be quoted by the water provider", "V"]],
      widths=[4.0, 11.6, 1.0], fontsize=8.3)
para("If a well were required instead, the measured geology gives a reasonable expectation: this is "
     "Georgia Piedmont terrain with residual kaolinitic soils over saprolite and crystalline bedrock "
     "(the Zion unit records indurated lithic bedrock at 77 cm, confirming shallow rock in places). "
     "Piedmont domestic wells typically draw from fractured-rock aquifers, and yield depends on "
     "intersecting fracture zones rather than on a continuous water table. Shallow saprolite water is "
     "more vulnerable to seasonal decline and to contamination from a nearby septic field; deeper "
     "fractured-bedrock completion generally gives better water quality and more reliable yield. "
     "A local licensed well driller should set the target depth — this is not determinable from "
     "desk data.", 9)
tier("V", "Well depth and yield require a licensed driller and, ideally, neighbouring well logs from "
          "the Georgia EPD well-log database.")

doc.add_heading("20. Sewer", level=2)
table(["Question", "Finding"],
      [["Municipal sewer availability", "**None.** No public sewer serves this area."],
       ["Sewer connection", "Not possible"],
       ["Septic requirement", "**Mandatory** — see Section D.16"],
       ["Distance to nearest sewer", "Lincolnton, approximately 17 miles by road. "
        "Extension is not a realistic option."]],
      widths=[5.0, 11.6], fontsize=8.4)
tier("D", "Confirmed by the absence of any sewer provider and by listing descriptions.")

doc.add_heading("21. Electricity", level=2)
table(["Question", "Finding"],
      [["Electric provider", "**Two serve Lincoln County: Rayle EMC and Georgia Power.** "
        "The county's economic-development page states that which provider applies depends on "
        "where the property sits. **The provider for this parcel is not established.**"],
       ["Market structure", "Georgia is a regulated market — the provider is fixed by address "
        "under the 1973 Territorial Service Act. There is no supplier choice."],
       ["Existing service to the lot", "Unconfirmed"],
       ["Power lines nearby", "The subdivision is described as having underground utilities, "
        "implying distribution is already in place along Watersedge Cove"],
       ["Underground vs. overhead", "Underground reported for the community; service drop "
        "arrangement for an individual lot must be confirmed"],
       ["Connection availability and cost", "Must be quoted by whichever of Rayle EMC or "
        "Georgia Power serves the parcel"]],
      widths=[4.4, 12.2], fontsize=8.4)
tier("V", "Identify the serving utility from the deed address and obtain a written service "
          "availability letter.")

doc.add_heading("22. Natural gas / propane", level=2)
table(["Question", "Finding"],
      [["Natural gas availability", "**Very unlikely.** No natural-gas distribution is indicated for "
        "this rural part of Lincoln County. Note that while Georgia's natural-gas market is "
        "deregulated at the supply level, that is irrelevant without physical distribution mains."],
       ["Propane availability", "**Expected to be the practical option** — standard for rural "
        "north-east Georgia. Served by regional propane dealers with owned or leased tanks."],
       ["Existing gas infrastructure", "None indicated on the lot"],
       ["Design implication", "Budget for a propane tank (above or below ground, subject to HOA "
        "architectural approval) if gas cooking, gas heating or a generator is wanted. "
        "Alternatively an all-electric design with a heat pump suits this mild climate well."]],
      widths=[4.4, 12.2], fontsize=8.4)
tier("V", "Confirm with local propane dealers and check HOA rules on tank placement and screening.")

doc.add_heading("23. Internet", level=2)
table(["Technology", "Availability in Tignall / 30668", "Notes"],
      [["**Fiber**", "**Yes, in parts** — Relyant Communications is reported as the fiber provider, "
        "with speeds up to about 1,000 Mbps", "**Parcel-level availability must be confirmed.** "
        "Fiber coverage in rural ZIP codes is patchy by road."],
       ["Cable", "Xfinity is reported as serving the Tignall area", "Coverage at the parcel unconfirmed"],
       ["DSL", "Available in only about 1.8 % of Tignall; AT&T is the main DSL provider",
        "Effectively not an option"],
       ["Fixed wireless", "T-Mobile Home Internet via 5G is listed as available",
        "Speeds vary with cellular conditions and tower distance"],
       ["Satellite", "Available (nationwide coverage)", "Usable fallback; higher latency"],
       ["Provider count", "Around 8–9 residential providers reported for the ZIP code, "
        "with roughly 98–99 % of the town covered by at least one", "Town-level, not parcel-level"],
       ["Reported maximum speed", "Up to about 963–1,000 Mbps in covered areas", "Advertised, not guaranteed"]],
      widths=[2.6, 7.4, 6.6], fontsize=8.2)
para("This matters more than usual here. The site is 44 miles from a Walmart and 55 miles from a "
     "full-service hospital; remote work, telehealth and online ordering substitute for physical "
     "proximity. Fiber at the parcel would meaningfully change the liveability calculus, and its "
     "absence would equally meaningfully reduce it. Verify by entering the exact address on each "
     "provider's availability checker and on the FCC National Broadband Map before committing.", 9)
tier("V", "All broadband figures are ZIP-code aggregates from commercial comparison sites. "
          "Parcel-level confirmation required.")

pagebreak()

# =============================================================== SECTION F
doc.add_heading("F. Lake, Dock & Waterfront", level=1)
para("This section contains the most consequential findings in the report, because the waterfront and "
     "dock claims are what distinguish this listing from ordinary rural acreage.", 9, bold=True)

doc.add_heading("The lake itself — reference data", level=2)
table(["Attribute", "Value", "Source"],
      [["Official name", "J. Strom Thurmond Lake (locally Clarks Hill Lake)", "USACE"],
       ["Operator", "US Army Corps of Engineers, Savannah District", "USACE"],
       ["Surface area", "approx. 71,100 acres", "Published"],
       ["Shoreline length", "approx. 1,200 miles", "USACE / HOA"],
       ["Standing", "Largest USACE lake east of the Mississippi River", "Published"],
       ["**Full pool**", "**330 ft-msl**", "**USACE Savannah District**"],
       ["Conservation storage", "Elevation 312 to 330 ft (1,045,000 acre-feet)", "USACE"],
       ["Flood-control storage", "Elevation 330 to 335 ft (390,000 acre-feet)", "USACE"],
       ["Top of spillway gates", "335 ft", "USACE"],
       ["Normal limit of drawdown", "305 ft", "USGS"],
       ["Spillway crest", "300 ft", "USGS"],
       ["Seasonal drawdown", "The project has a seasonal drawdown of the flood-control pool", "USACE"],
       ["Average depth", "approx. 37 ft", "Published"],
       ["Maximum depth", "approx. 180 ft (some sources cite 110 ft)", "Published — varies"]],
      widths=[4.6, 7.0, 5.0], fontsize=8.3)

doc.add_heading("24. Waterfront access", level=2)
table(["Access type", "Assessment"],
      [["Direct lake frontage in fee", "**Not supported by the terrain data.** "
        "Full pool (330 ft) lies 467–688 ft horizontally from the road and 53–65 ft below it; "
        "a 1.07-acre lot of ordinary proportions cannot span that distance (Section B.6)."],
       ["Frontage on the USACE boundary", "**This is the most probable configuration.** "
        "The listing's own wording — \"gradual slope to the lake once you get to the back corps "
        "line\" — concedes a Corps boundary at the rear of the lot."],
       ["Deeded lake access", "Unknown; would appear in the plat or CC&Rs"],
       ["Shared / community access", "Unknown for Stillwater Coves. Note that community docks are "
        "the only new-dock mechanism available to post-May-2018 subdivisions, so a community "
        "access point may be the actual amenity."],
       ["Easement access", "Unknown; check the title commitment"],
       ["Public access", "**Definitely available.** Hesters Ferry Campground is a county-managed "
        "former Corps park on Fishing Creek where it joins Thurmond Lake, close to the subdivision. "
        "Elijah Clark State Park is 21 miles by road."],
       ["Public right to use the shoreline", "**Important:** USACE states the Thurmond Lake shoreline "
        "is open to use by the general public. Even an adjacent owner with a dock permit gets no "
        "exclusive-use rights over the federal shoreline."]],
      widths=[4.4, 12.2], fontsize=8.3)
tier("M", "Conclusion derived from 1 m lidar geometry. The recorded plat and the USACE boundary "
          "survey will settle it definitively.")

doc.add_heading("25. Lakeside terrain — road to house to lake", level=2)
table(["Segment", "Distance", "Grade", "Character"],
      [["Road to house pad (first ~100 ft)", "approx. 100 ft", "6–13 %",
        "Gentle; driveway feasible without switchbacks"],
       ["House pad to rear lot line (to ~310 ft)", "approx. 210 ft", "6–13 %",
        "Gentle to moderate; walkable garden slope"],
       ["Rear lot line to full pool", "approx. 160–380 ft", "9–12 % average, "
        "up to 24.5 % in places", "**On USACE land.** Moderate, with steeper pitches"],
       ["**Total, road to water**", "**467–688 ft**", "**9.0–11.7 % average**",
        "**53–65 ft of vertical descent**"]],
      widths=[4.8, 3.2, 3.6, 5.0], fontsize=8.3)
para("Classification of lake access:", 9, bold=True)
bullets([("Gentle —", "yes, in gradient terms: 9–12 % average is a comfortable walking slope."),
         ("Moderate —", "yes, in the localised 50 ft windows reaching 18.6–24.5 %."),
         ("Steep —", "no. Nothing above 25 % was measured."),
         ("Requires stairs —", "**not for gradient, but plausibly for length.** The issue is not "
          "steepness, it is that the walk to the water is 470–690 ft each way with a 53–65 ft climb "
          "on return — roughly equivalent to five storeys. Carrying coolers, fishing gear or small "
          "children up that is a daily-life consideration, not a trivial one."),
         ("Requires retaining wall —", "possibly for the driveway and pad; not for lake access."),
         ("Requires significant grading —", "for the building pad on the west frontage, yes; "
          "for lake access, no — a path would suffice."),
         ("Golf-cart or ATV access —", "would materially improve usability, but any improvement on "
          "federal land requires a USACE permit.")])
tier("M", "Distances and grades from 1 m lidar.")

doc.add_heading("26. Dock feasibility — THE CRITICAL ITEM", level=2)
callout("THE 2018 RULE CHANGE THAT GOVERNS THIS PURCHASE",
        "The J. Strom Thurmond Lake Shoreline Management Plan was revised effective 1 May 2018. Two "
        "substantive changes bear directly on these lots: (1) a 20-FOOT MINIMUM SHARED BOUNDARY with "
        "federal land is now required to be eligible for new dock placement; and (2) FOR SUBDIVISIONS "
        "PLATTED AFTER 1 MAY 2018, ONLY COMMUNITY DOCKS WILL BE AUTHORISED, via agreements with the "
        "developer or the homeowners association. If Stillwater Coves — or the specific phase containing "
        "Lots 130 and 131 — was platted after that date, A PRIVATE INDIVIDUAL DOCK IS NOT AVAILABLE AT "
        "ANY PRICE. Establishing the plat recording date is therefore the single highest-value piece of "
        "due diligence on this property.",
        fill="FDE7E9", bordercol="C01F2B")
table(["Requirement / restriction", "What the record shows"],
      [["**Who authorises a dock**", "The US Army Corps of Engineers, Savannah District, through a "
        "Shoreline Use Permit. **Not** the seller, the HOA, the county or the listing agent."],
       ["**Shoreline allocation**", "Permits are possible only where the adjacent shoreline is "
        "allocated as a Limited Development Area. Approximately 18 % of the shoreline "
        "(about 211 miles) carries that allocation. **Permits are prohibited in some areas.** "
        "Allocation for this specific frontage is **unverified**."],
       ["**How to verify allocation**", "Contact the Thurmond Project Office, toll free "
        "800-533-3478. Rangers will determine zoning and permit possibility for a specific site."],
       ["**Access requirement**", "20 ft minimum shared boundary with federal land for new docks "
        "(2018 SMP)."],
       ["**Post-2018 plats**", "Community docks only."],
       ["**Permit term**", "Maximum five years."],
       ["**Transferability**", "**Non-transferable.** A permit becomes null and void at the moment "
        "the adjacent private property is sold or transferred. The new owner must contact the "
        "Thurmond Project Office to obtain a fresh five-year permit."],
       ["**Property rights conveyed**", "**None.** Permits grant no real estate rights and convey no "
        "private exclusive-use privileges on government property."],
       ["**Public access over the shoreline**", "Retained — the shoreline remains open to the "
        "general public."],
       ["**Other applicable law**", "Ownership, construction, operation, use and maintenance of a "
        "permitted facility remain subject to all applicable federal, state and local laws. "
        "Failure to comply may cause revocation."],
       ["**Dock size limitations**", "Governed by the SMP's dock facilities and requirements "
        "provisions. **Specific dimensions not retrieved** — the USACE PDFs returned HTTP 403."],
       ["**HOA restrictions**", "Stillwater Coves has an architectural approval process. "
        "Any dock-related covenant is **unverified** — obtain the CC&Rs."],
       ["**County / state regulations**", "Additional state (Georgia DNR) and county requirements may "
        "apply to shoreline work. **Unverified for this site.**"],
       ["**Riparian rights**", "Effectively displaced by federal ownership of the shoreline. "
        "This is a permit regime, not a riparian-rights regime."]],
      widths=[4.8, 11.8], fontsize=8.2)
callout("\"DOCKABLE\" AND \"VERBAL APPROVAL\" ARE NOT PERMITS",
        "The marketing for these lots asserts both are \"DOCKABLE\". A neighbouring Stillwater Coves "
        "listing goes further and cites \"Corps of Engineers verbal approval\". Treat both as unverified "
        "sales claims. A verbal indication from a ranger is not a Shoreline Use Permit, is not binding, "
        "and would in any event expire and need re-issuing on transfer of the property. Require written "
        "confirmation of (a) the shoreline allocation class for this frontage, (b) eligibility under the "
        "20 ft shared-boundary rule, and (c) whether the plat date limits you to a community dock. "
        "Make the contingency specific and in writing.")
tier("D", "USACE Savannah District published policy and the 2018 SMP news release. "
          "Site-specific eligibility is UNVERIFIED and must come from the Thurmond Project Office.")

doc.add_heading("27. Lake depth", level=2)
table(["Measure", "Value", "Confidence"],
      [["Average depth, whole lake", "approx. 37 ft", "Published"],
       ["Maximum depth, whole lake", "approx. 180 ft (lower figures of 110 ft also cited)",
        "Published — sources disagree"],
       ["Depth adjacent to the property", "**NOT DETERMINED.** No bathymetric survey of the "
        "Fishing Creek arm at this location was obtained.", "Not available"],
       ["Seasonal water-level variation", "Conservation pool spans 312–330 ft, i.e. **up to 18 ft** "
        "of operating range, and the project runs a seasonal drawdown of the flood-control pool. "
        "Normal drawdown limit is 305 ft.", "Documented (USACE)"]],
      widths=[4.2, 9.4, 3.0], fontsize=8.3)
callout("THE DRAWDOWN QUESTION — A COVE-SPECIFIC RISK",
        "An 18 ft conservation-pool range is a lot of vertical movement in a shallow cove at the head of "
        "a creek arm. In drought years, coves of this type at Thurmond routinely go dry or become mudflat, "
        "stranding docks and eliminating boat access for months. Because this lot sits on a creek-arm cove "
        "rather than the main lake body, this risk is above average. Obtain bathymetry near the frontage "
        "and ask the Thurmond Project Office and existing neighbours what the water does at elevation 320, "
        "315 and 312 ft. USACE publishes boat-ramp bottom elevations, which is a useful proxy for when "
        "access fails.")
tier("V", "Site-specific depth and drawdown behaviour require bathymetric data and local knowledge.")

doc.add_heading("28. Boating", level=2)
table(["Question", "Finding"],
      [["Are boats permitted?", "**Yes** — Thurmond is a major recreational boating lake"],
       ["Motorised boating", "Permitted"],
       ["Boat types", "Full range: powerboats, pontoons, sailing craft, personal watercraft, "
        "kayaks and canoes"],
       ["Wake boats", "No lake-wide prohibition identified; local wake, no-wake and idle zones apply "
        "near ramps, marinas, docks and swim areas — **verify current rules**"],
       ["Speed restrictions", "Zone-specific rather than lake-wide; no-wake zones apply in coves, "
        "around structures and in designated areas"],
       ["Boat ramps", "Numerous USACE and county ramps. Hesters Ferry (county-managed, close by) and "
        "Elijah Clark State Park (21 mi) are the nearest. **USACE publishes boat-ramp bottom "
        "elevations — essential reading for drought years.**"],
       ["Marina facilities", "Present on the lake; none identified immediately adjacent to "
        "Stillwater Coves"],
       ["Registration", "Vessels must be registered; Georgia or South Carolina registration applies "
        "depending on principal use — this is a boundary-water lake"],
       ["Practical constraint", "With no confirmed private dock and a 470–690 ft walk to the water, "
        "realistic boating access is by trailer to a public ramp, or via a community dock if "
        "Stillwater Coves has one"]],
      widths=[4.4, 12.2], fontsize=8.3)
tier("D", "Lake and ramp facts documented. Current speed/wake rules require confirmation with "
          "USACE and the relevant state DNR.")

doc.add_heading("29. Fishing", level=2)
table(["Question", "Finding"],
      [["Fish species present", "Largemouth bass, hybrid bass and striped bass, black crappie, "
        "white perch and yellow perch, chain pickerel, bluegill, shellcracker (redear sunfish), "
        "redbreast sunfish, white bass, and channel, blue and flathead catfish"],
       ["Invasive species", "**Spotted bass** are now present; Georgia DNR actively encourages "
        "harvesting them"],
       ["Are fish stocked?", "Striped and hybrid bass fisheries in this reservoir system are "
        "maintained by stocking — **confirm the current programme with Georgia DNR**"],
       ["Local reputation", "The Fishing Creek arm, where Hesters Ferry sits, is described as "
        "holding some of the best fishing on the lake — this is the arm adjoining the subdivision"],
       ["Seasonal pattern", "Crappie is the most frequently caught species in spring; largemouth "
        "also peak in spring; striped bass have overtaken largemouth in popularity"],
       ["Licence requirement", "**Yes.** A fishing licence is required. Because Thurmond is a "
        "Georgia–South Carolina boundary water, check which state's licence covers the water you "
        "intend to fish, and whether reciprocity applies — **verify before fishing**"],
       ["Example creel limits", "South Carolina: on Hartwell and Thurmond, no more than 10 striped "
        "or hybrid bass (or a combination) per day, of which only 3 may exceed 26 inches. "
        "Georgia: possession of more than 50 total individuals of listed game fish is unlawful, "
        "excepting channel and flathead catfish."],
       ["Seasonal restrictions and methods", "Species-specific; consult the current Georgia and "
        "South Carolina regulations"],
       ["Fishing from the property's shoreline", "**The shoreline is federal land open to the public, "
        "so it is not \"your\" shoreline.** Bank fishing is available, but subject to USACE rules "
        "and without exclusivity."]],
      widths=[4.4, 12.2], fontsize=8.2)
tier("D", "Species and example limits from Georgia DNR and South Carolina DNR. "
          "Limits change annually — always read the current year's regulations.")

pagebreak()
print("sections E-F written")

# =============================================================== SECTION G
doc.add_heading("G. HOA, Covenants & Land Use", level=1)

doc.add_heading("30. HOA", level=2)
table(["Question", "Finding", "Tier"],
      [["Is the property subject to an HOA?", "**Yes.** Stillwater Coves operates a homeowners "
        "association with a Board and standing Committees, and maintains a website "
        "(stillwatercovesgeorgia.com, hosted on an HOA-sites platform).", "D"],
       ["HOA name", "Stillwater Coves (property owners association)", "D"],
       ["Annual dues", "**NOT PUBLISHED.** No dues figure appears on the public site or in any "
        "listing retrieved.", "V"],
       ["Monthly equivalent", "Cannot be computed without the dues figure", "V"],
       ["Transfer fees", "Not published", "V"],
       ["Special assessments", "Not published — **ask specifically** whether any are pending, "
        "and request the last two years of minutes and budgets", "V"],
       ["Community character", "Gated, paved roads, underground utilities, public water; "
        "approximately 1,000 acres of wooded ridges and hilltops; generously sized homesites", "V"],
       ["Clubhouse / pool / tennis / fitness", "**No evidence of any.** None mentioned in HOA "
        "material or listings.", "V"],
       ["Parks / trails", "Not identified", "V"],
       ["Lake facilities / community dock", "**Not confirmed either way — and this is important.** "
        "Because post-May-2018 plats are limited to community docks, the existence or absence of a "
        "Stillwater Coves community dock may be the determining factor in lake access.", "V"],
       ["Boat / RV storage", "Not identified for Stillwater Coves. (Gated boat/RV storage is an "
        "amenity of POINTE SHORES, a different community — do not conflate.)", "V"]],
      widths=[4.0, 11.6, 1.0], fontsize=8.2)
para("A note on the HOA website: it is served over HTTP and presented an invalid TLS certificate on "
     "HTTPS. Content was readable but should not be treated as an authoritative legal source. "
     "Request documents directly from the association in writing.", 9, italic=True)

doc.add_heading("31. HOA covenants", level=2)
para("One concrete piece of evidence was recovered: Stillwater Coves publishes a **Plan Approval "
     "Checklist** requiring submission of owner, builder and house-plan details, and stating that "
     "**no clearing, grading or construction should commence before written approval is obtained**. "
     "An architectural review regime therefore exists and has teeth.", 9)
para("Governing documents to obtain in full before closing:", 9, bold=True)
table(["Document / provision", "Why it matters here specifically"],
      [["Declaration and CC&Rs", "The master document; establishes all other restrictions"],
       ["Architectural guidelines", "Confirmed to exist. Will drive minimum square footage, "
        "exterior materials, roof pitch, colours and plan approval timeline — all of which affect "
        "the build cost estimated in Section K.43"],
       ["Rules and regulations, plus all amendments", "Restrictions are frequently added by amendment"],
       ["**Minimum house size**", "Very common in lake communities and a direct cost driver. "
        "Surrounding new builds are 1,600–2,441 sq ft."],
       ["**Waterfront and dock restrictions**", "May be more restrictive than USACE policy. "
        "Also check whether the HOA holds or could hold a community dock agreement."],
       ["**Tree removal / clearing limits**", "Highly likely in a community marketing \"mature "
        "hardwoods\" and 1,000 acres of woodland. Could constrain the building pad and septic field."],
       ["Fence restrictions", "Affects garden, pet and poultry plans"],
       ["Outbuilding restrictions", "Affects workshop, greenhouse and storage plans"],
       ["RV and boat parking/storage restrictions", "Affects trailer-based lake access — which may be "
        "your primary access route"],
       ["**Livestock and poultry restrictions**", "Directly determines Section H.34"],
       ["**Short-term-rental restrictions**", "Directly determines Section I.35. HOA bans on STR are "
        "common and are enforceable even where county rules permit"],
       ["Time-to-build requirement", "Some lake communities require construction to start within a "
        "set period after purchase — critical if buying to hold"],
       ["Propane tank / generator rules", "Relevant given the absence of natural gas (Section E.22)"]],
      widths=[5.0, 11.6], fontsize=8.2)
tier("V", "CC&Rs and architectural guidelines were not obtained. Request them from the association "
          "and have counsel review before removing contingencies.")

doc.add_heading("32. County zoning", level=2)
para("Lincoln County maintains a Planning and Zoning function (separate from, but adjacent to, "
     "the City of Lincolnton's). A Lincoln County / City of Lincolnton Joint Comprehensive Plan "
     "exists, prepared through the Central Savannah River Area Regional Commission. The county's "
     "zoning ordinance text itself was not publicly retrievable.", 9)
table(["Zoning question", "Status"],
      [["Zoning designation for this parcel", "**NOT DETERMINED** — county GIS is token-restricted "
        "and the ordinance text is not online"],
       ["Permitted residential use", "Single-family residential is evidently permitted — the "
        "subdivision is platted and homes are under construction nearby"],
       ["Minimum lot size", "**Not determined.** Note that Georgia practice commonly applies a "
        "1-acre minimum for septic-served lots and 0.5 acre where public water and sewer exist. "
        "At 1.07 acres each, these lots sit just above a typical 1-acre septic threshold — "
        "a thin margin worth confirming."],
       ["Maximum building coverage", "Not determined"],
       ["Setbacks (front / side / rear)", "**Not determined.** This report models 35 ft front, "
        "15 ft side and 35 ft rear purely for illustration in Fig. 6."],
       ["Height restrictions", "Not determined"],
       ["Accessory structures", "Not determined"],
       ["ADU / in-law suite", "Not determined — see Section G.33"],
       ["Agricultural / hobby-farming provisions", "Not determined — see Section H.34"],
       ["Building permits and codes", "Georgia state minimum codes apply; administered locally"]],
      widths=[4.6, 12.0], fontsize=8.3)
callout("ACTION: ONE PHONE CALL RESOLVES MOST OF SECTION G",
        "Lincoln County Planning and Zoning can confirm zoning district, minimum lot size, setbacks, "
        "height limits, accessory-structure and ADU rules, and whether short-term rentals are regulated — "
        "in a single conversation, given the parcel ID. Obtain the parcel ID from the deed or the tax "
        "assessor first. Ask for the answers in writing or by email.")
tier("V", "All zoning parameters require confirmation from Lincoln County Planning and Zoning.")

doc.add_heading("33. Second in-law suite / ADU", level=2)
table(["Determinant", "Assessment"],
      [["Zoning", "**Unknown.** Whether a second independent dwelling unit is permitted, and whether "
        "it must be attached, is the governing question."],
       ["Maximum dwelling units per lot", "Unknown — typically one principal dwelling plus a "
        "restricted accessory unit, if any"],
       ["HOA", "Likely to have a view, given an active architectural review committee. Unknown."],
       ["Building code", "A second kitchen, separate entrance, bathroom and egress all trigger code "
        "requirements but are routinely achievable"],
       ["**Septic capacity — the real constraint**", "**This is where an ADU most likely fails.** "
        "An independent suite adds bedrooms, and bedrooms drive absorption-field area. On soils rated "
        "\"Very limited\" across much of the slope, adding 1–2 bedrooms of design load may exceed "
        "what the usable soil area can carry. A Level 3/4 soil report must size for the TOTAL "
        "bedroom count including the suite."],
       ["Water", "Public water (if confirmed) handles additional demand readily"],
       ["Parking", "Ample on 2.14 acres"],
       ["**Strategic advantage of two lots**", "Owning both lots materially improves the odds: an "
        "ADU could potentially sit on the second lot with its own septic system, subject to zoning "
        "and HOA rules on one-dwelling-per-lot."]],
      widths=[4.6, 12.0], fontsize=8.3)
tier("V", "Requires zoning confirmation, HOA review and a septic design sized for total bedrooms.")

pagebreak()

# =============================================================== SECTION H
doc.add_heading("H. Hobby Farming", level=1)

doc.add_heading("34. Poultry and animals", level=2)
para("Two independent permission layers apply, and **either one can prohibit**. County zoning in rural "
     "Georgia is typically permissive about small livestock; private HOA covenants in amenity lake "
     "communities are typically restrictive. In a conflict, the stricter controls.", 9)
table(["Animal", "County outlook (rural Lincoln County)", "HOA outlook (amenity lake community)",
       "Net assessment"],
      [["Chickens (hens)", "Likely permitted", "Frequently restricted or banned; roosters "
        "almost always banned", "**Check CC&Rs first**"],
       ["Ducks / geese", "Likely permitted", "Frequently restricted; also attract predators "
        "and foul water features", "Check CC&Rs"],
       ["Rabbits", "Likely permitted", "Often treated as pets if enclosed and not commercial",
        "Most likely to be acceptable"],
       ["Goats", "Possibly permitted on acreage", "Commonly banned as livestock",
        "Unlikely without specific allowance"],
       ["Bees", "Likely permitted; Georgia is beekeeping-friendly", "Variable; sometimes restricted "
        "on small lots", "Check CC&Rs; good fit with the site's fruit plantings"],
       ["Other small livestock", "Depends on acreage and zoning district", "Commonly banned",
        "Unlikely"],
       ["Horses / cattle", "Would require far more than 2.14 acres", "Almost certainly banned",
        "Not practical"]],
      widths=[2.6, 4.6, 5.0, 4.4], fontsize=8.1)
para("Site-specific practical factors, from the measured data:", 9, bold=True)
bullets([("Predation —", "the area is wooded with abundant wildlife; secure housing is essential "
          "for poultry."),
         ("Slope and erosion —", "animal runs on 10–25 % severely eroded Pacolet will strip "
          "vegetation and erode quickly. Site pens on the flatter near-road ground."),
         ("Septic field conflict —", "animal enclosures must not sit over the absorption field or "
          "its reserve area, which consumes the flattest usable ground."),
         ("Water quality —", "runoff from animal areas toward the lake is a regulatory sensitivity "
          "given federal shoreline ownership downslope."),
         ("Forage —", "native soils are acidic (pH 4.8–5.5) with low organic matter; pasture would "
          "need liming and seeding to be productive.")])
tier("V", "Requires both the Lincoln County zoning ordinance and the Stillwater Coves CC&Rs.")

pagebreak()

# =============================================================== SECTION I
doc.add_heading("I. Short-Term Rental", level=1)

doc.add_heading("35. STR feasibility", level=2)
table(["Determinant", "Finding", "Tier"],
      [["Georgia state law", "**There is no statewide short-term-rental law.** Whether a permit is "
        "needed, and how strict the rules are, is set by the city or county.", "D"],
       ["State taxes — unavoidable", "**4 % state sales tax** plus a **$5 per night hotel-motel fee** "
        "apply to every Georgia short-term rental, usually collected automatically by Airbnb or Vrbo. "
        "Stays exceeding 30 nights are generally excepted.", "D"],
       ["Lincoln County, Georgia rules", "**NOT DETERMINED.** No Lincoln County (Georgia) STR "
        "ordinance was located. Caution: searches surface STR rules for Lincoln County in other "
        "states — those are irrelevant here.", "V"],
       ["City rules", "Not applicable — the parcel is unincorporated", "D"],
       ["Zoning", "Unknown whether the applicable district permits transient lodging", "V"],
       ["**HOA restrictions**", "**The most likely binding constraint.** Amenity lake communities "
        "commonly prohibit or heavily restrict short-term letting, and such covenants are enforceable "
        "independently of county permissiveness. Stillwater Coves CC&Rs not obtained.", "V"],
       ["Minimum rental period", "Unknown; often imposed via HOA covenant (e.g. 30-day minimum)", "V"],
       ["STR permit / licence", "Unknown at county level", "V"],
       ["Occupancy limits", "Unknown. **Note:** occupancy would in any case be capped by the "
        "permitted septic bedroom count (Section D.16) — a 2-bedroom septic approval caps the "
        "rental proposition regardless of the house size.", "M"],
       ["Parking", "Not a constraint on 2.14 acres", "M"],
       ["Noise regulations", "County noise ordinance status unknown", "V"],
       ["Licensing / income tax", "Federal and Georgia income tax on rental income; "
        "local business licence may apply", "D"]],
      widths=[4.0, 11.6, 1.0], fontsize=8.2)
para("Commercial assessment, if an STR is part of the investment case: the location cuts both ways. "
     "Positives are genuine lake recreation, excellent fishing in the Fishing Creek arm, a quiet "
     "gated setting, proximity to Hesters Ferry and Elijah Clark State Park, and reported fibre "
     "broadband. Negatives are severe: 44 miles to a Walmart, 55 miles to a full-service hospital, "
     "63 miles to the nearest commercial airport, and — decisively — **no confirmed private dock and "
     "a 470–690 ft walk down a 53–65 ft drop to reach the water.** Lake-rental demand is driven "
     "overwhelmingly by dock and water access. Without a dock, this would compete as a rural cabin, "
     "not as a lake house, and should be underwritten accordingly.", 9)
callout("ORDER OF OPERATIONS FOR AN STR PLAN",
        "Resolve in this sequence, and stop at the first failure: (1) HOA CC&Rs — is short-term letting "
        "permitted at all? (2) Lincoln County — is a permit required and is the use allowed in the zoning "
        "district? (3) Septic — how many bedrooms will actually be approved, since that caps occupancy and "
        "therefore revenue? (4) Dock — is any form of water access available, since that sets the nightly "
        "rate? Only then model returns.")
tier("V", "STR feasibility is genuinely unresolved and depends primarily on documents not obtained.")

pagebreak()
print("sections G-I written")

# =============================================================== SECTION J
doc.add_heading("J. Environmental & Geological Due Diligence", level=1)
para("This section is deliberately short on findings and long on requests, because almost every item "
     "here is a document that either exists in someone's file or does not exist at all. Ask the seller, "
     "the county and the HOA directly for each one.", 9)

table(["#", "Item", "Status", "Where to obtain / what it would tell you"],
      [["36", "**Elevation Certificate**", "**Not obtained; may not exist**",
        "A licensed surveyor prepares it. Compare against: FEMA BFE (none published here), lowest "
        "adjacent grade, proposed finished floor, and flood-zone requirements. Likely required by "
        "lender/insurer if any part of the structure is in Zone A."],
       ["37", "**Boundary survey**", "**Not obtained**",
        "Essential. Establishes the true lot lines, acreage, frontage and — critically — the "
        "location of the USACE boundary at the rear."],
       ["37", "ALTA survey", "Not obtained",
        "Usually only for commercial transactions, but gives the most complete easement picture."],
       ["37", "Topographic survey", "Not obtained",
        "This report substitutes 1 m lidar, which is good for planning but is not a survey. "
        "A topo survey is needed for grading design and permits."],
       ["38", "**Geotechnical / soil report**", "**Not obtained; likely does not exist**",
        "Advisable given severely eroded Pacolet, clay subsoils at 40–55 %, and possible shallow "
        "bedrock. Needed for foundation and retaining-wall design."],
       ["39", "**Percolation / Level 3-4 soil report**", "**Not obtained — GATING ITEM**",
        "Determines whether a septic system is permittable and for how many bedrooms. "
        "Georgia requires an original stamped Level 3 or 4 report. Without it, the house size "
        "is unknown."],
       ["40", "Ground Penetrating Radar / subsurface investigation", "Not obtained; unlikely to exist",
        "Would map shallow bedrock and buried features. Normally unnecessary unless the "
        "geotechnical report flags rock."],
       ["41", "**Radon measurement**", "**Not determined**",
        "Lincoln County's EPA radon zone could not be reliably extracted (the EPA state map is "
        "image-based). Georgia contains only a handful of Zone 1 counties, so Zone 2 or 3 is more "
        "likely here — but Piedmont granitic/gneissic terrain is a recognised radon source. "
        "EPA's position is that elevated radon occurs in ALL three zones and every home should be "
        "tested. Budget for a test and for passive radon-resistant construction, which is cheap at "
        "build stage and expensive to retrofit."],
       ["42", "**Sinkholes / karst**", "**Risk assessed as very low**",
        "The soil taxonomy is decisive: Typic Kanhapludults (Cecil, Pacolet, Madison) and Ultic "
        "Hapludalfs (Zion) with indurated lithic bedrock are residual soils over crystalline "
        "Piedmont rock — granite, gneiss and schist. **This is not karst terrain.** Georgia's "
        "sinkhole problems occur in the limestone Coastal Plain, not the Piedmont. No known "
        "sinkholes, subsidence or geological hazards were identified for this area."]],
      widths=[0.8, 3.4, 2.8, 9.6], fontsize=7.9)
tier("V", "Items 36–41 all require licensed professionals or records not publicly available. "
          "Item 42 is a documented geological inference.")

callout("SEISMIC AND MINING NOTE",
        "Two further items worth a question. (1) Lincoln County is home to Graves Mountain, noted for "
        "rutile and other mineral crystals — ask whether mineral rights have been severed from the "
        "surface estate, which the title commitment will show. (2) Georgia's Piedmont has low but "
        "non-zero seismic hazard; standard Georgia code seismic provisions apply and no special design "
        "is expected.")

pagebreak()

# =============================================================== SECTION K
doc.add_heading("K. Residential Construction Feasibility", level=1)

doc.add_heading("43. House construction and building envelope", level=2)
figure("fig6_envelope.png",
       "Fig. 6 — Illustrative site layout across both lots, showing a 4-bedroom footprint, driveway, "
       "and primary plus reserve septic fields. Lot shape and setbacks are ASSUMED; confirm against "
       "the recorded plat, HOA guidelines and Lincoln County zoning.")
table(["Element", "Assessment on 2.14 combined acres"],
      [["Building envelope", "Adequate. The flattest and best-drained ground (Cecil sandy loam, "
        "2–6 %, rated \"Not limited\" for dwellings) lies in the first 100–150 ft back from the road, "
        "especially on the middle and east frontage."],
       ["Setbacks", "**Unknown** — modelled at 35 ft front, 15 ft side, 35 ft rear for illustration only"],
       ["House footprint", "A 2,800 sq ft 4-bedroom footprint (modelled 56 x 50 ft) fits comfortably. "
        "Orient the long axis PARALLEL to the contour to minimise cut, fill and stepped footings."],
       ["Garage", "Easily accommodated; a side-entry or lower-level garage suits the 6–13 % grade well"],
       ["Driveway", "Short and straightforward from Watersedge Cove at 6–13 %. No switchback needed. "
        "Note the paved community road is already in place."],
       ["Pool", "Physically possible on the near-road bench. Expect additional cost for retaining "
        "on sloping ground, and check HOA architectural rules."],
       ["Patio / outdoor living", "Well suited; a north-facing rear elevation gives shaded summer "
        "outdoor space, valuable given 52 days a year above 32 deg C"],
       ["**Septic field + reserve**", "**The binding constraint.** Roughly 12,000 sq ft primary plus "
        "12,000 sq ft reserve for a 4-bedroom design, and it must sit on usable soil — which excludes "
        "the \"Very limited\" 10–25 % slope areas. **This is the single strongest reason to buy both "
        "lots rather than one.**"],
       ["Well", "Not needed if public water is confirmed. If needed, must maintain statutory "
        "separation from the septic field, further consuming usable area."],
       ["Retaining walls", "Probable on the west frontage; likely avoidable on the east"],
       ["Lake setback", "Unknown; expect a setback from the USACE boundary and from the shoreline. "
        "Confirm with the county and USACE."],
       ["Drainage", "Manageable — see 45 below. Positive diversion of upslope runoff is needed."]],
      widths=[3.6, 13.0], fontsize=8.3)

doc.add_heading("Cost estimate — 4-bedroom home comparable to surrounding construction", level=2)
para("The checklist asks for an estimated cost to build a 4-bedroom home similar to the surrounding "
     "homes and complying with HOA rules. Start from the observed market, then add this site's "
     "specific burdens.", 9)
para("Observed comparable new construction in and near the community:", 9, bold=True)
table(["Property", "Size", "Beds / baths", "Asking price"],
      [["1165 Watersedge Cv (same street)", "1,600 sq ft", "4 bd / 3.5 ba", "$535,000"],
       ["1192 Sunset Cv", "1,600 sq ft", "4 bd / 3.5 ba", "$535,000"],
       ["1198 Sunset Cv", "2,441 sq ft", "—", "$585,000"],
       ["1021 Pointe S", "—", "4 bd / 3 ba", "$685,000"],
       ["Stillwater Coves (aggregated listing data)", "—", "—", "from approx. $528,000"]],
      widths=[6.0, 3.0, 3.6, 4.0], fontsize=8.3)
para("Published Georgia build-cost ranges for 2026: general new construction $150–$300 per sq ft; "
     "custom homes $185–$450+ per sq ft; custom with upgraded finishes $270–$350 per sq ft; "
     "Augusta-area contractor ranges quoted as low as $120–$220 per sq ft. The spread is wide because "
     "it depends entirely on finish level and site difficulty.", 9)
para("Line-item estimate for a 2,800 sq ft, 4-bedroom home on THIS site:", 9, bold=True)
table(["Cost element", "Low", "High", "Basis and site-specific driver"],
      [["Site work, clearing, grading, erosion control", "$25,000", "$70,000",
        "6–13 % slope; severely eroded Pacolet; HOA tree-protection likely; cut-and-fill pad"],
       ["Retaining walls (if required)", "$0", "$45,000",
        "Probable on west frontage; engineered design needed above 4 ft"],
       ["Driveway and culvert", "$8,000", "$22,000", "Short run from a paved community road"],
       ["**Septic system**", "**$15,000**", "**$55,000**",
        "**Conventional 4-bedroom at the low end; engineered or advanced-treatment system at the "
        "high end if the Level 3/4 report comes back marginal on \"Very limited\" soils**"],
       ["Water connection (public)", "$2,000", "$9,000", "Tap, meter and service line, if public "
        "water is confirmed"],
       ["Well (only if public water unavailable)", "$0", "$25,000", "Piedmont fractured-rock well"],
       ["Electrical service connection", "$2,000", "$12,000", "Rayle EMC or Georgia Power; "
        "underground service drop"],
       ["Propane tank and lines", "$0", "$6,000", "No natural gas available"],
       ["Foundation (crawlspace or stepped)", "$30,000", "$70,000",
        "Clay subsoils 40–55 %; slope-driven stepping; drainage detailing"],
       ["Shell, envelope, roof, windows", "$220,000", "$330,000", "2,800 sq ft"],
       ["MEP (mechanical, electrical, plumbing)", "$70,000", "$120,000",
        "Heat pump suits the climate; 52 days/yr above 32 deg C drives cooling load"],
       ["Interior finishes", "$120,000", "$280,000", "The main swing factor; HOA guidelines may "
        "mandate exterior materials"],
       ["Radon-resistant passive system", "$1,500", "$4,500", "Cheap at build; costly to retrofit"],
       ["Soft costs — design, engineering, permits, HOA review", "$25,000", "$60,000",
        "Includes geotechnical report, Level 3/4 soil report, topographic survey, "
        "boundary survey, HOA plan approval"],
       ["Contingency at 10–15 %", "$55,000", "$160,000", "Higher than usual: septic outcome and "
        "retaining requirements are both unresolved"],
       ["**ESTIMATED TOTAL (excluding land)**", "**approx. $575,000**", "**approx. $1,270,000**",
        "**Most likely central case: $700,000 – $850,000, i.e. roughly $250–$305 per sq ft**"]],
      widths=[5.0, 2.2, 2.2, 7.2], fontsize=7.8)
callout("READ THE COST ESTIMATE AS A RANGE, NOT A NUMBER",
        "The low end assumes a conventional septic approval, no retaining walls, confirmed public water "
        "and modest finishes. The high end assumes an engineered septic system, retaining walls, a well, "
        "and upgraded finishes. The two unresolved items that move the number most are THE SEPTIC "
        "DESIGN and THE RETAINING REQUIREMENT, and both are answered by the same two reports — a "
        "Level 3/4 soil evaluation and a geotechnical investigation. Commission both before signing a "
        "build contract. Note also that the comparable finished homes at $535,000–$685,000 broadly "
        "bracket the central case, which is a sanity check that the estimate is realistic.")
tier("M", "Estimate built from published 2026 Georgia cost ranges, observed local comparables, and "
          "this site's measured slope and soil constraints. Not a contractor quotation.")

doc.add_heading("44. Grading and fill", level=2)
table(["Question", "Finding"],
      [["Existing grades", "Road crown 385–399 ft; 6.0–12.8 % average across the first 310 ft; "
        "localised 50 ft windows to 18.0 %"],
       ["Required grading", "Moderate. A balanced cut-and-fill pad on the near-road bench is the "
        "efficient approach. Expect 6–7 ft of grade change across a 56 ft-deep footprint on "
        "12 % ground."],
       ["Estimated cut / fill", "Site-specific design required. Orienting the house parallel to the "
        "contour can roughly halve the earthwork relative to a perpendicular orientation."],
       ["Retaining-wall requirements", "Probable on the west frontage; walls above 4 ft typically "
        "need engineered design and a permit"],
       ["Imported fill likely?", "Possibly, for pad construction. **Not** for flood reasons — "
        "the envelope already sits 45–55 ft above full pool."],
       ["County approval", "A land-disturbance permit and erosion-and-sedimentation control plan "
        "should be expected. Georgia's E&S rules are strict near state waters, and the lake is "
        "downslope."],
       ["**Erosion control is a priority, not a formality**", "The Pacolet unit is mapped as "
        "SEVERELY ERODED on 10–25 % slopes with surface organic matter of only 0.25–1.00 %. "
        "Disturbed soil here will move. Silt fence, temporary seeding, matting and prompt permanent "
        "stabilisation are essential — and runoff drains toward federal land and the lake."]],
      widths=[4.4, 12.2], fontsize=8.3)
tier("M", "Grades measured from lidar. Earthwork quantities require a topographic survey and "
          "engineered grading plan.")

doc.add_heading("45. Drainage", level=2)
table(["Element", "Finding"],
      [["Natural drainage", "**Well drained** soils throughout; consistent north to north-west "
        "flow from the ridge toward the Fishing Creek arm"],
       ["Swales", "Two shallow swales visible in the lidar surface flanking the modelled lot pair. "
        "Avoid placing a pad across a swale, or provide positive diversion."],
       ["Culverts", "Existing driveway culvert requirements set by county road standards"],
       ["Stormwater flow", "Short contributing slope lengths mean modest volumes, but the lots sit "
        "BELOW a ridge and therefore receive upslope runoff"],
       ["Neighbouring runoff", "Upslope lots and the road corridor both shed toward these lots. "
        "Perimeter interception upslope of the house is advisable."],
       ["Lake drainage / outfall", "All site runoff ultimately reaches Thurmond Lake across "
        "federal land — a regulatory sensitivity for both construction-phase sediment and "
        "post-construction quality"],
       ["Erosion risk", "**High** on disturbed slope (see 44). Low once permanently stabilised, "
        "given well-drained soils."],
       ["Subsurface drainage", "Clay subsoils at 40–55 % will perch water above them during wet "
        "periods. Foundation drains and a properly detailed crawlspace or basement waterproofing "
        "system are important."]],
      widths=[4.0, 12.6], fontsize=8.3)
tier("M", "Flow paths from the lidar gradient field; soil drainage classes from SSURGO.")

pagebreak()

# =============================================================== SECTION L
doc.add_heading("L. Living & Location Analysis", level=1)

doc.add_heading("46. Climate (metric)", level=2)
para("Computed from ERA5 reanalysis at 33.9410 N, 82.5646 W for the 30-year period 1995–2024.", 9)
table(["Month", "Mean daily max (C)", "Mean daily min (C)", "Mean (C)", "Precipitation (mm)"],
      [["January", "12.3", "3.5", "7.9", "98.5"],
       ["February", "14.5", "4.9", "9.7", "100.5"],
       ["March", "18.5", "8.2", "13.4", "107.0"],
       ["April", "23.2", "12.2", "17.7", "85.7"],
       ["May", "27.6", "17.1", "22.3", "70.9"],
       ["June", "30.9", "21.2", "26.1", "88.8"],
       ["July", "**32.3**", "23.1", "**27.7**", "95.5"],
       ["August", "31.7", "22.7", "27.2", "84.3"],
       ["September", "28.6", "19.7", "24.1", "87.7"],
       ["October", "23.6", "14.2", "18.9", "64.7"],
       ["November", "17.7", "8.6", "13.2", "77.0"],
       ["December", "13.5", "5.1", "9.3", "109.0"],
       ["**ANNUAL**", "**22.9**", "**13.4**", "**18.2**", "**1,070 mm (42.1 in)**"]],
      widths=[3.0, 3.6, 3.6, 2.6, 3.8], fontsize=8.3)
table(["Climate measure", "Value"],
      [["Average annual temperature", "18.2 deg C"],
       ["Summer highs (July mean max)", "32.3 deg C"],
       ["Winter lows (January mean min)", "3.5 deg C"],
       ["Annual rainfall", "1,070 mm / 42.1 in, distributed year-round with a late-spring dip"],
       ["Humidity", "High, typical of the humid subtropical south-east "
        "(Koppen Cfa). Not separately quantified here."],
       ["Growing season", "approx. 283 frost-free days; mean last spring frost 26 February, "
        "mean first autumn frost 5 December"],
       ["Latest recorded spring frost", "27 March"],
       ["Earliest recorded autumn frost", "13 November"],
       ["Days per year at or above 32 deg C", "approx. 52"],
       ["Days per year at or above 35 deg C", "approx. 13"],
       ["Extreme maximum (1995–2024)", "41.9 deg C"],
       ["Extreme minimum (1995–2024)", "-12.3 deg C"],
       ["Mean annual extreme minimum", "-5.3 deg C (22.4 deg F) — consistent with USDA "
        "hardiness zone 8a/8b"]],
      widths=[6.4, 10.2], fontsize=8.3)
tier("M", "ERA5 reanalysis is a gridded model product, excellent for averages but smoothed at the "
          "extremes. True local extremes are likely slightly more severe than shown.")

doc.add_heading("47. Natural hazards", level=2)
table(["Hazard", "Assessment", "Basis"],
      [["**Flood**", "**Zone A (SFHA, no BFE) at the lake margin; Zone X over the road corridor and "
        "building envelope.** The envelope sits 45–55 ft above full pool.", "FEMA NFHL — documented"],
       ["Hurricane / wind", "Moderate. Roughly 150 miles inland; remnants of Atlantic systems bring "
        "wind and heavy rain rather than storm surge. Falling trees are the main risk on a wooded lot.",
        "Regional pattern"],
       ["Tornado", "Low to moderate. Georgia experiences tornadoes, most often in spring; "
        "the Piedmont is not a peak-frequency corridor.", "Regional pattern"],
       ["**Wildfire**", "**Moderate and worth taking seriously.** A heavily wooded lot with mature "
        "hardwoods and pine in the vicinity, in a region with dry autumn spells. Defensible-space "
        "design and ember-resistant detailing are prudent.", "Site character"],
       ["Lightning", "Moderate to high. The south-east has among the highest US lightning "
        "flash densities; frequent summer thunderstorms. Surge protection advisable.", "Regional"],
       ["Hail", "Low to moderate; associated with spring severe storms", "Regional"],
       ["Severe thunderstorms", "Frequent in spring and summer; 1,070 mm annual rainfall is "
        "delivered substantially by convective storms", "Measured climate"],
       ["Drought", "**Relevant, with a lake-specific twist.** South-east Georgia droughts recur, "
        "and the real consequence here is reservoir drawdown — up to 18 ft of conservation-pool "
        "range — which can strand a cove.", "USACE operating data"],
       ["Extreme heat", "52 days/yr at or above 32 deg C, 13 days/yr at or above 35 deg C, "
        "record 41.9 deg C. Cooling capacity and shade matter.", "Measured climate"],
       ["**Sinkhole / karst**", "**Negligible.** Residual soils over crystalline Piedmont bedrock; "
        "Georgia's karst is in the Coastal Plain.", "Soil taxonomy — documented"],
       ["Seismic", "Low. Standard Georgia code provisions apply.", "Regional"],
       ["**Erosion**", "**High on disturbed slope** — severely eroded Pacolet on 10–25 % slopes with "
        "0.25–1.00 % surface organic matter", "SSURGO — documented"]],
      widths=[2.8, 9.6, 4.2], fontsize=8.0)

doc.add_heading("48. Community and demographics", level=2)
table(["Measure", "Lincoln County, Georgia"],
      [["Population (July 2025 estimate)", "8,112"],
       ["Population (2020 Census)", "7,690"],
       ["Population (2010 Census)", "7,996"],
       ["Change 2020 to 2025", "**+5.5 %** — modest growth, reversing the 2010–2020 decline"],
       ["Persons under 5", "5.1 %"],
       ["Persons under 18", "18.9 %"],
       ["**Persons 65 and over**", "**28.1 %** — markedly older than the US average; "
        "consistent with a retirement and lake-recreation in-migration pattern"],
       ["Female", "51.5 %"],
       ["White alone", "71.2 % (69.6 % white alone, not Hispanic or Latino)"],
       ["Black alone", "24.7 %"],
       ["Asian alone", "**0.4 %**"],
       ["American Indian / Alaska Native", "0.4 %"],
       ["Native Hawaiian / Pacific Islander", "0.2 %"],
       ["Two or more races", "3.2 %"],
       ["Hispanic or Latino", "2.5 %"],
       ["Per capita income (2020–2024, 2024 dollars)", "$29,291"],
       ["Median household income, Lincolnton", "$43,935"],
       ["Poverty rate, Lincoln County", "approx. 17 % — about 25 % above the Georgia rate"],
       ["Median age, Lincolnton", "48.1 years"],
       ["Household composition", "Predominantly family households; detailed breakdown not retrieved"],
       ["Indian / South Asian population", "**Very small.** Asian alone is 0.4 % of roughly 8,100 "
        "people — on the order of 30 residents countywide. See Section L.52."]],
      widths=[6.4, 10.2], fontsize=8.3)
tier("D", "US Census Bureau QuickFacts for Lincoln County, Georgia, and city-level data for "
          "Lincolnton. The Census API itself was unreachable from this environment; "
          "figures are from published QuickFacts tables.")

doc.add_heading("49. Schools", level=2)
table(["Item", "Finding"],
      [["School district", "**Lincoln County School District**, based in Lincolnton"],
       ["Total enrolment", "approx. 1,261 students, PK and K–12"],
       ["Student-teacher ratio", "approx. 18 : 1"],
       ["Structure", "A single small district — effectively one elementary, one middle and one "
        "high school serve the whole county"],
       ["Distance", "**approx. 17 miles / 37 minutes** by road to Lincolnton. Expect a long "
        "school-bus ride."],
       ["State test proficiency", "approx. 55 % proficient in mathematics, approx. 42 % in reading"],
       ["Published ratings", "One review service rates county schools at an average 8/10, placing "
        "them in the top 30 % of Georgia public schools; another rates Lincolnton schools 7/10. "
        "Ratings services disagree — treat as indicative only."],
       ["Private options", "None identified within a reasonable radius"],
       ["Practical note", "A small district means small class sizes and a strongly community "
        "school culture, but also a narrow range of advanced courses, languages and "
        "extracurricular options compared with a metro district"]],
      widths=[4.4, 12.2], fontsize=8.3)
tier("D", "Enrolment and ratio from district data; ratings from commercial review services, "
          "which are not authoritative. Verify attendance zones with the district directly.")

doc.add_heading("50. Healthcare", level=2)
table(["Facility", "Type", "Distance", "Drive time"],
      [["**Wills Memorial Hospital**, Washington, GA", "25-bed critical-access hospital with "
        "24-hour emergency department, outpatient therapy and digital imaging",
        "**19.9 mi**", "**43 min**"],
       ["Local physicians, Lincolnton", "Primary and family medicine", "approx. 17 mi", "approx. 37 min"],
       ["**Piedmont Augusta Hospital**", "812-bed acute-care hospital; multi-campus system with "
        "heart and vascular centre; regional referral centre", "**54.5 mi**", "**1 h 38 min**"],
       ["Athens, GA medical facilities", "Regional hospitals", "61.4 mi", "1 h 33 min"],
       ["Emergency services", "Lincolnton offers 24-hour 911 emergency service", "—", "—"]],
      widths=[5.6, 6.0, 2.4, 2.6], fontsize=8.3)
callout("HEALTHCARE IS THE MOST SIGNIFICANT LIVEABILITY TRADE-OFF",
        "The nearest emergency department is 43 minutes away by car and is a 25-bed critical-access "
        "facility — appropriate for stabilisation, not for complex care. The nearest full-service "
        "hospital is 1 hour 38 minutes away. For a household with elderly members, chronic conditions, "
        "or anyone requiring time-critical intervention such as stroke or cardiac care, this is a "
        "material consideration. It is also consistent with the county's 28.1 % over-65 population: "
        "many residents accept the trade, but it should be an informed choice. Check air-ambulance "
        "coverage and response times.")
tier("D", "Facility types documented; distances are OSRM road-network calculations.")

doc.add_heading("51. Shopping and services", level=2)
table(["Destination", "Distance", "Drive time", "What is there"],
      [["Lincolnton (county seat)", "17.0 mi", "37 min", "Groceries, pharmacy, bank, fuel, "
        "county offices, local restaurants, schools"],
       ["Washington, GA", "19.9 mi", "43 min", "Hospital, additional retail and services"],
       ["**Thomson, GA — Walmart Supercenter**", "**43.5 mi**", "**1 h 13 min**",
        "Nearest major big-box retailer; full grocery and general merchandise"],
       ["Athens, GA", "61.4 mi", "1 h 33 min", "Full metropolitan retail, university city, "
        "restaurants, speciality shopping"],
       ["Augusta, GA", "54.5 mi", "1 h 38 min", "Full metropolitan retail, hospitals, "
        "**Indian groceries and restaurants**, hardware and home-improvement chains"],
       ["Elijah Clark State Park", "21.2 mi", "43 min", "Lake recreation, boat ramp, museum, camping"],
       ["Hesters Ferry Campground", "very close (adjacent area)", "a few minutes",
        "County-managed lake access, boat ramp, waterfront campsites "
        "(16 with water and electric, plus a primitive loop)"]],
      widths=[4.6, 2.4, 2.4, 7.2], fontsize=8.2)
para("Everyday groceries, banking, fuel and pharmacy require a 37-minute round trip of 34 miles. "
     "Hardware and home-improvement chains, and any significant retail choice, mean a 1 h 13 min to "
     "1 h 38 min drive. Plan on consolidated weekly or fortnightly shopping trips and heavy reliance "
     "on delivery — which makes the broadband question in Section E.23 more important than it would "
     "normally be.", 9)
tier("D", "All distances and times are OSRM road-network calculations from the site coordinates.")

doc.add_heading("52. Indian / South Asian community", level=2)
table(["Item", "Finding"],
      [["Nearest Indian groceries", "**Augusta, GA — approx. 54.5 mi / 1 h 38 min.** "
        "Indian and Asian grocers operate there (for example Radha Indian Asian Grocers)."],
       ["Nearest Indian restaurants", "Augusta — approx. 54.5 mi / 1 h 38 min"],
       ["Nearest Hindu temple", "**Hindu Temple Society of Augusta** — approx. 54.5 mi / 1 h 38 min"],
       ["Mosques", "Augusta metro; none identified closer"],
       ["South Asian community organisations", "Augusta metro area"],
       ["Local South Asian population", "**Effectively negligible.** Lincoln County is 0.4 % "
        "Asian alone — roughly 30 people countywide across all Asian origins."],
       ["Nearest significant Indian population centre", "**Metropolitan Atlanta — approx. 130 mi / "
        "3 h 9 min.** Georgia's Indian population of roughly 183,000 is concentrated in Atlanta's "
        "northern suburbs: Johns Creek (approx. 11,700), Atlanta (approx. 11,200) and "
        "Alpharetta (approx. 9,200)."],
       ["Augusta metro context", "Augusta is Georgia's second-largest metro at roughly 611,000 "
        "people, but its Asian population is a small fraction of the Atlanta total. "
        "Expect a modest but real community with temple and grocery infrastructure."]],
      widths=[4.6, 12.0], fontsize=8.3)
para("Practical reading: cultural, religious and grocery infrastructure exists, but every visit is a "
     "3-hour round trip to Augusta. For a weekend or retirement property this may be entirely "
     "acceptable; for a primary residence with school-age children for whom community and cultural "
     "continuity matter, it is a significant and recurring cost in time.", 9)
tier("D", "Census race data documented; facility locations documented; distances OSRM-computed.")

doc.add_heading("53. Entertainment and nightlife", level=2)
table(["Category", "Nearest meaningful option", "Distance"],
      [["Local restaurants", "Lincolnton", "17.0 mi / 37 min"],
       ["Restaurant choice and variety", "Augusta or Athens", "54.5 mi / 61.4 mi"],
       ["Bars", "Lincolnton (limited); Augusta or Athens for choice", "17 mi / 55–61 mi"],
       ["Cafes", "Lincolnton (limited)", "17.0 mi"],
       ["Cinemas", "**Thomson, Augusta or Athens — none locally**", "43.5–61.4 mi"],
       ["Shopping and entertainment districts", "Augusta; Athens (university city, strongest "
        "nightlife and music scene in range)", "54.5 mi / 61.4 mi"],
       ["Major nightlife", "Athens is the regional standout; Atlanta for anything larger",
        "61.4 mi / 130.5 mi"],
       ["Local character", "Outdoor-recreation oriented: boating, fishing, camping, state parks, "
        "Graves Mountain, and local live-music acts based in the county", "on site to 21 mi"]],
      widths=[4.2, 8.4, 4.0], fontsize=8.3)
para("This is a quiet rural lake community, not a location with nightlife. The genuine recreational "
     "offer is outdoor: a 71,100-acre lake with 1,200 miles of shoreline, excellent fishing, two "
     "nearby parks and extensive woodland. Lincoln County also markets itself on more than 160 "
     "buildings listed on the National Register of Historic Places.", 9)
tier("D", "Distances OSRM-computed; amenity characterisation from county and regional sources.")

doc.add_heading("54. Walkability", level=2)
table(["Score", "Value", "Comment"],
      [["Walk Score", "**Not published** for this rural location — Walk Score does not "
        "meaningfully cover unincorporated rural addresses", "Expect a score in the 0–10 range "
        "(\"Car-Dependent / almost all errands require a car\")"],
       ["Bike Score", "Not published", "Paved community roads are pleasant for recreational cycling; "
        "the surrounding rural highways have no shoulders or bike infrastructure"],
       ["Transit Score", "**Not applicable** — no fixed-route public transit serves the area",
        "Rural transit, if any, would be demand-response only"]],
      widths=[3.0, 7.0, 6.6], fontsize=8.3)
para("Practical assessment: **a car is mandatory for every single daily necessity.** "
     "There is no shop, school, medical facility, pharmacy or employment within walking distance. "
     "The nearest groceries are 17 miles away. A two-car household should be assumed, and the "
     "annual mileage and fuel cost of a 34-mile round trip for routine errands should be built "
     "into any budget comparison against a town location. Walking within the community itself, "
     "on paved low-traffic roads through woodland, is genuinely pleasant — but it is recreation, "
     "not transport.", 9)
tier("M", "Scores unavailable; assessment reasoned from measured distances.")

doc.add_heading("55. Airport", level=2)
table(["Airport", "Distance", "Drive time", "Role"],
      [["**Augusta Regional Airport (AGS)**", "**63.4 mi**", "**1 h 51 min**",
        "**Nearest commercial airport.** Regional service, primarily connecting through major "
        "hubs — limited non-stop destinations"],
       ["Columbia Metropolitan Airport (CAE), SC", "116.9 mi", "2 h 57 min",
        "Alternative regional airport"],
       ["**Hartsfield-Jackson Atlanta (ATL)**", "**130.5 mi**", "**3 h 9 min**",
        "**Nearest true international hub.** Extensive domestic and international non-stop "
        "network; the realistic choice for any long-haul or international travel"]],
      widths=[4.8, 2.4, 2.6, 6.8], fontsize=8.3)
para("For international travel — relevant given the cultural-proximity question in Section L.52 — "
     "the practical airport is Atlanta, a 3-hour drive each way. That makes a trip to India or "
     "Europe a 6-hour round-trip ground journey on top of the flight. Augusta Regional is "
     "convenient for domestic connecting travel but offers few non-stops.", 9)
tier("D", "Distances and drive times OSRM-computed; airport roles from published service patterns.")

figure("fig7_location.png",
       "Fig. 7 — Regional location and access. Mileages are road-network driving distances "
       "from the site.", width_cm=15.5)

pagebreak()
print("sections J-L written")

# =============================================================== SECTION M
doc.add_heading("M. Final Due-Diligence Summary", level=1)

doc.add_heading("56. Property strengths", level=2)
table(["Strength", "The evidence behind it"],
      [["**Flood risk at the building envelope is genuinely low**",
        "The road corridor and upland sit in FEMA Zone X (minimal hazard), 45–55 ft above the "
        "330 ft full pool and 10+ ft above the 335 ft flood-control pool. That is an unusually "
        "large freeboard for a lake property."],
       ["**The near-road building ground is good**",
        "Cecil sandy loam at 2–6 % slope is rated \"Not limited\" by NRCS for dwellings with and "
        "without basements — the best rating available."],
       ["**Terrain is workable, not difficult**",
        "6.0–12.8 % average across the likely envelope; nothing above 25 % measured anywhere. "
        "Ordinary Georgia Piedmont building terrain."],
       ["**Buying two lots together is a real advantage**",
        "2.14 acres gives genuine flexibility for septic primary plus reserve fields, a possible "
        "ADU, privacy, and a hedge against a restrictive soil report."],
       ["**Infrastructure already in place**",
        "Gated community with paved roads and underground utilities; public water reported. "
        "No road-building or long utility runs required."],
       ["**Broadband is credible**",
        "Fibre reported in parts of the ZIP code at up to approximately 1,000 Mbps, plus cable, "
        "5G fixed wireless and satellite. Around 8–9 providers serve the area."],
       ["**Outstanding lake recreation**",
        "71,100 acres and 1,200 miles of shoreline; the largest USACE lake east of the Mississippi. "
        "The adjoining Fishing Creek arm is locally reputed as some of the best fishing on the lake."],
       ["**Nearby public lake access regardless of dock outcome**",
        "Hesters Ferry Campground is immediately nearby with a boat ramp and waterfront sites; "
        "Elijah Clark State Park is 21 miles."],
       ["**Long growing season and blueberry-friendly soil**",
        "283 frost-free days; native pH 4.8–5.3 is ideal for rabbiteye blueberry and muscadine "
        "without amendment."],
       ["**No karst or sinkhole exposure**",
        "Residual soils over crystalline Piedmont bedrock. Georgia's sinkhole problems are "
        "confined to the limestone Coastal Plain."],
       ["**Schools are small and reasonably rated**",
        "18:1 ratio; one review service places the county in the top 30 % of Georgia public schools."],
       ["**Quiet, private, wooded setting**",
        "Approximately 1,000 acres of wooded ridges; mature hardwoods; low-traffic gated roads."]],
      widths=[5.0, 11.6], fontsize=8.2)

doc.add_heading("57. Property limitations and risks", level=2)
table(["Risk", "Severity", "Detail"],
      [["**Water frontage is probably not what it appears**", "**HIGH**",
        "Full pool is 467–688 ft horizontally and 53–65 ft vertically from the road. A 1.07-acre lot "
        "cannot reach it. Expect USACE land between the rear boundary and the water, and no fee "
        "ownership of frontage."],
       ["**Private dock may be legally unavailable**", "**HIGH**",
        "The 2018 Thurmond SMP permits only COMMUNITY docks for subdivisions platted after "
        "1 May 2018, and requires a 20 ft shared federal boundary for any new dock. Shoreline "
        "allocation for this frontage is unverified; only about 18 % of the lake shoreline is "
        "Limited Development Area."],
       ["**Septic capacity may limit house size**", "**HIGH**",
        "Pacolet, Zion and Madison 10–25 % all rate \"Very limited\" for absorption fields. "
        "Zion has indurated bedrock at 77 cm. Georgia denies permits where bedrock is under 2 ft "
        "below the field. A 4-bedroom conventional approval is not assured."],
       ["**Asking price is unresolved**", "**HIGH**",
        "Cached listings show $55,000, $144,000 and $169,900 across three brokerages."],
       ["Zone A SFHA present on part of the area", "MEDIUM",
        "No BFE published, so the county sets the flood elevation and an Elevation Certificate "
        "will likely be required if any structure falls in Zone A."],
       ["Reservoir drawdown in a creek-arm cove", "MEDIUM",
        "Up to 18 ft of conservation-pool range (312–330 ft) plus seasonal flood-pool drawdown. "
        "Shallow cove locations can become unusable in drought."],
       ["Long, steep walk to the water", "MEDIUM",
        "470–690 ft each way with a 53–65 ft climb on return — roughly five storeys. "
        "A daily-use constraint even if a dock is permitted."],
       ["High erosion potential during construction", "MEDIUM",
        "Severely eroded Pacolet on 10–25 % slopes, 0.25–1.00 % surface organic matter, "
        "and runoff draining to federal land and the lake."],
       ["Construction cost uncertainty", "MEDIUM",
        "Estimated $575k–$1.27m for 2,800 sq ft; the septic and retaining outcomes drive the spread."],
       ["Clay subsoils at 40–55 %", "MEDIUM",
        "Perched water above the clay; foundation drainage and basement waterproofing must be "
        "engineered properly."],
       ["HOA restrictions unknown but certain to exist", "MEDIUM",
        "Confirmed architectural approval regime with a written pre-construction requirement. "
        "Dues, assessments, minimum house size, tree-clearing, livestock and STR rules all unobtained."],
       ["Zoning parameters entirely unknown", "MEDIUM",
        "Setbacks, minimum lot size, height, accessory structures and ADU rules not determined. "
        "At 1.07 acres each, the lots sit just above a typical 1-acre septic minimum."],
       ["**Remoteness from healthcare**", "**MEDIUM-HIGH**",
        "43 min to a 25-bed critical-access ER; 1 h 38 min to a full-service hospital."],
       ["Remoteness from retail and services", "MEDIUM",
        "17 mi to groceries, 43.5 mi to a Walmart, 63 mi to a commercial airport, "
        "130 mi to an international hub."],
       ["Car dependence is absolute", "MEDIUM",
        "No transit, no walkable amenities, 34-mile round trip for routine errands."],
       ["Very small South Asian community locally", "MEDIUM (household-specific)",
        "0.4 % Asian countywide; nearest temple and groceries 1 h 38 min away in Augusta, "
        "nearest major community 3 h away in Atlanta."],
       ["STR viability doubtful", "MEDIUM",
        "HOA rules unknown, county rules unknown, occupancy capped by septic bedroom count, "
        "and no confirmed dock — which is what drives lake-rental rates."],
       ["Radon zone undetermined", "LOW-MEDIUM",
        "Could not be extracted reliably. Piedmont crystalline terrain is a recognised radon source. "
        "Test, and install a passive system at build stage."],
       ["Wildfire exposure on a wooded lot", "LOW-MEDIUM",
        "Mature hardwood and pine cover; defensible space advisable."],
       ["Parcel identity not independently verified", "LOW (but fundamental)",
        "County GIS is token-restricted; no parcel ID, dimensions or legal description obtained."]],
      widths=[4.6, 2.0, 10.0], fontsize=7.9)

doc.add_heading("58. Missing documentation — what to obtain before purchase", level=2)
para("These are grouped by who holds them. The four marked GATING should be resolved before any "
     "due-diligence contingency is released.", 9)
table(["#", "Document", "Hold by", "Priority"],
      [["1", "**Recorded subdivision plat, INCLUDING THE RECORDING DATE**",
        "Lincoln County Clerk of Superior Court", "**GATING**"],
       ["2", "**Written confirmation of shoreline allocation and dock eligibility for this parcel**",
        "USACE Thurmond Project Office, 800-533-3478", "**GATING**"],
       ["3", "**Level 3 or Level 4 soil report (stamped) and septic feasibility / bedroom count**",
        "Licensed soil classifier + Lincoln County environmental health", "**GATING**"],
       ["4", "**Current written asking price, status and listing history**", "Listing broker",
        "**GATING**"],
       ["5", "Boundary survey showing the USACE line at the rear", "Licensed Georgia surveyor", "High"],
       ["6", "HOA Declaration, CC&Rs, architectural guidelines and all amendments",
        "Stillwater Coves POA", "High"],
       ["7", "HOA budget, dues schedule, transfer fees, pending assessments, last 2 years of minutes",
        "Stillwater Coves POA", "High"],
       ["8", "Zoning confirmation: district, setbacks, min lot size, height, ADU, livestock, STR",
        "Lincoln County Planning & Zoning", "High"],
       ["9", "Title commitment — easements, reservations, mineral rights, dock covenants",
        "Title company / closing attorney", "High"],
       ["10", "Utility availability letters: water, electricity (Rayle EMC or Georgia Power), broadband",
        "Each utility, addressed to the parcel", "High"],
       ["11", "Floodplain determination, required flood elevation and freeboard",
        "Lincoln County floodplain administrator", "High"],
       ["12", "FIRM panel number and effective date; Flood Insurance Study 13181CV000A",
        "FEMA Map Service Center", "Medium"],
       ["13", "Geotechnical / soil engineering report", "Licensed geotechnical engineer", "Medium"],
       ["14", "Topographic survey for grading design", "Licensed surveyor", "Medium"],
       ["15", "Elevation Certificate (if any structure falls in Zone A)", "Licensed surveyor", "Medium"],
       ["16", "Existing dock permit status for the subdivision / any community dock agreement",
        "USACE + HOA", "Medium"],
       ["17", "Bathymetric data for the adjacent cove; boat-ramp bottom elevations",
        "USACE Savannah District", "Medium"],
       ["18", "Laboratory soil test for pH, buffer pH and nutrients", "UGA Extension, Lincoln County",
        "Low (post-purchase)"],
       ["19", "Radon test", "Certified radon tester", "Low (post-purchase)"],
       ["20", "Property tax history and current assessment", "Lincoln County Tax Assessor", "Medium"],
       ["21", "Building permit requirements and fee schedule", "Lincoln County building department",
        "Medium"],
       ["22", "Well log data for neighbouring properties (if a well may be needed)",
        "Georgia EPD well-log database", "Low"]],
      widths=[0.8, 6.4, 5.2, 2.2], fontsize=7.9)

doc.add_heading("59. Final construction feasibility", level=2)
table(["Can the lot reasonably accommodate...", "Verdict", "Reasoning"],
      [["**Main house**", "**YES**",
        "Cecil sandy loam at 2–6 % near the road is rated \"Not limited\" for dwellings. "
        "A 2,800 sq ft 4-bedroom footprint fits with room for setbacks."],
       ["**Second in-law suite / ADU**", "**CONDITIONAL**",
        "Physically yes on 2.14 acres. Constrained by (a) unknown zoning rules on dwelling units, "
        "(b) HOA architectural review, and (c) **septic capacity for the total bedroom count** — "
        "the real constraint."],
       ["**Garage**", "**YES**", "Easily; a side-entry or lower-level garage suits the 6–13 % grade."],
       ["**Septic system**", "**CONDITIONAL — the central open question**",
        "A system is certainly achievable somewhere on 2.14 acres. Whether a CONVENTIONAL "
        "4-BEDROOM system is permittable is unresolved: much of the slope rates \"Very limited\", "
        "and Zion's bedrock at 77 cm would disqualify a conventional field. May require an "
        "engineered or advanced-treatment system."],
       ["**Garden**", "**YES**",
        "283 frost-free days. Requires liming (native pH 4.8–5.5), heavy compost (native OM under "
        "1.3 %) and raised beds on slope. Blueberry and muscadine thrive in the native soil."],
       ["**Hobby farming**", "**CONDITIONAL**",
        "County zoning is likely permissive; **HOA covenants are the probable obstacle**. "
        "Rabbits and bees are most likely acceptable; goats and poultry commonly restricted in "
        "amenity lake communities."],
       ["**Dock**", "**DOUBTFUL — UNRESOLVED**",
        "Requires a USACE Shoreline Use Permit. Blocked outright if the plat postdates 1 May 2018 "
        "(community docks only), and requires a 20 ft shared federal boundary. Allocation for this "
        "frontage unverified. **Do not pay a dock premium without written confirmation.**"],
       ["**Boat access**", "**YES, but by trailer**",
        "Excellent public access at Hesters Ferry (adjacent) and Elijah Clark State Park (21 mi). "
        "Private water access from the lot is unconfirmed and involves a 470–690 ft walk."],
       ["**Short-term rental**", "**DOUBTFUL — UNRESOLVED**",
        "HOA rules unknown and commonly prohibitive; county rules unknown; occupancy capped by "
        "septic bedrooms; and without a dock the property competes as a rural cabin rather than "
        "a lake house."],
       ["**Pool / outdoor recreation**", "**YES**",
        "Feasible on the near-road bench with retaining allowance. Subject to HOA architectural "
        "approval. Outdoor living suits the climate, with a north-facing rear elevation giving "
        "summer shade."]],
      widths=[4.2, 3.0, 9.4], fontsize=8.0)

doc.add_heading("60. Key measurements summary", level=2)
table(["Parameter", "Finding", "Tier"],
      [["Lot size", "1.07 ac each; 2.14 ac combined (46,609 sq ft each) — per listing", "Listing"],
       ["Lot dimensions", "**Not published.** Modelled as 150 ft x 310 ft", "Verify"],
       ["Parcel ID / APN", "**Not obtained** — county GIS token-restricted", "Verify"],
       ["Lowest elevation (study area)", "326.9 ft NAVD 88 (lake bed, below full pool)", "Lidar"],
       ["Highest elevation (study area)", "446.8 ft NAVD 88", "Lidar"],
       ["Road frontage elevation", "385.0 – 399.1 ft NAVD 88", "Lidar"],
       ["Elevation at rear of a 310 ft lot", "347.5 ft (west) to 371.9 ft (east)", "Lidar"],
       ["Elevation difference, road to full pool", "52.6 – 65.0 ft", "Lidar"],
       ["Horizontal distance, road to full pool", "467 – 688 ft (142 – 210 m)", "Lidar"],
       ["Average slope, road to lake", "9.0 – 11.7 % (5.1 – 6.7 deg)", "Lidar"],
       ["Average slope within first 310 ft", "6.0 – 12.8 % (3.4 – 7.3 deg)", "Lidar"],
       ["**Maximum slope over any 50 ft**", "**18.6 – 24.5 % (10.5 – 13.7 deg)**", "Lidar"],
       ["Terrain classification", "Gentle to moderate slope; localised moderate-steep bands", "Lidar"],
       ["**FEMA zone**", "**Zone A (SFHA, no BFE) at lake margin; Zone X over building envelope.** "
        "DFIRM 13181C", "FEMA"],
       ["**BFE**", "**None published** (STATIC_BFE = -9999)", "FEMA"],
       ["Minimum finished-floor elevation", "**County-determined** — no BFE exists to compute from", "Verify"],
       ["Fill required?", "**Not for flood purposes.** Possibly for pad grading", "Model"],
       ["Sewer", "**None.** Septic mandatory", "Documented"],
       ["Water", "Public/community water **reported**; parcel-level unconfirmed", "Verify"],
       ["Electricity", "**Rayle EMC or Georgia Power** — which one depends on location", "Verify"],
       ["Gas", "**No natural gas.** Propane expected", "Verify"],
       ["Internet", "Fibre reported (Relyant, up to ~1,000 Mbps) plus cable, 5G, satellite; "
        "approx. 8–9 providers in ZIP", "Verify at parcel"],
       ["Septic required?", "**Yes.** Bedroom count unresolved — soils rate \"Very limited\" on slope", "Verify"],
       ["Soil", "Cecil sandy loam 2–6 % (good); Pacolet 10–25 % severely eroded, Zion silt loam "
        "10–25 % (bedrock at 77 cm), Madison — all \"Very limited\"", "NRCS"],
       ["Soil pH / organic matter", "pH 4.8 – 5.5; OM 0.25 – 1.25 % (both low)", "NRCS"],
       ["HOA", "**Yes** — Stillwater Coves POA, with architectural approval required before "
        "any clearing or grading", "Documented"],
       ["HOA fee", "**Not published**", "Verify"],
       ["Zoning", "**Not determined**", "Verify"],
       ["**Dock possible?**", "**UNRESOLVED.** USACE permit required; community-dock-only if plat "
        "postdates 1 May 2018; 20 ft federal boundary needed; allocation unverified", "Verify"],
       ["Lake full pool / operating range", "330 ft full pool; 312–330 ft conservation "
        "(18 ft range); 335 ft top of gates", "USACE"],
       ["Lake depth", "Lake average approx. 37 ft; maximum approx. 180 ft. "
        "**Depth at the property not determined**", "Published / Verify"],
       ["Boating", "Permitted; full range of craft; public ramps nearby including Hesters Ferry", "Documented"],
       ["Fishing", "Largemouth, striped and hybrid bass, crappie, perch, bluegill, catfish; "
        "licence required; boundary-water rules apply", "Documented"],
       ["Hobby farming", "County likely permissive; **HOA likely the obstacle**", "Verify"],
       ["Second in-law suite", "Conditional on zoning, HOA and septic capacity", "Verify"],
       ["STR", "Unresolved; no Georgia state STR law; 4 % sales tax + $5/night fee apply; "
        "HOA likely binding", "Verify"],
       ["Nearest hospital", "**Wills Memorial, Washington GA — 19.9 mi / 43 min** "
        "(25-bed critical access). Piedmont Augusta 54.5 mi / 1 h 38 min", "OSRM"],
       ["Nearest international airport", "**Hartsfield-Jackson Atlanta (ATL) — 130.5 mi / 3 h 9 min.** "
        "Nearest commercial: Augusta Regional (AGS) 63.4 mi / 1 h 51 min", "OSRM"],
       ["Schools", "Lincoln County School District, approx. 1,261 students, 18:1, 17 mi / 37 min", "Documented"],
       ["Annual mean temperature", "18.2 deg C (mean max 22.9, mean min 13.4)", "ERA5"],
       ["Annual precipitation", "1,070 mm (42.1 in)", "ERA5"],
       ["Growing season", "approx. 283 frost-free days (26 Feb – 5 Dec mean)", "ERA5"],
       ["Radon", "**Zone not determined.** Test regardless", "Verify"],
       ["Sinkholes / karst", "**Negligible risk** — crystalline Piedmont, not karst", "Documented"]],
      widths=[4.6, 10.4, 1.6], fontsize=7.9)

doc.add_heading("61. Visual deliverables — what is included and what is not", level=2)
table(["Required visual", "Status in this report"],
      [["Zillow title / hero image", "**NOT INCLUDED.** Zillow returned HTTP 403 Forbidden on every "
        "attempt; Redfin 403 and realtor.com 429. The images are also third-party copyright material "
        "and are not reproduced."],
       ["Zillow surrounding-area image", "**NOT INCLUDED** — same reason"],
       ["Official plat map", "**NOT INCLUDED** — not available online; obtain from the "
        "Lincoln County Clerk of Superior Court"],
       ["Esri / GIS parcel map", "**NOT INCLUDED** — Lincoln County ArcGIS services returned "
        "HTTP 499 \"Token Required\""],
       ["True-scale boundary and contour map", "**INCLUDED — Fig. 1**"],
       ["Road-to-lake elevation cross-section", "**INCLUDED — Fig. 2**, shown both vertically "
        "exaggerated and at true 1:1 scale as required"],
       ["3-D terrain map", "**INCLUDED — Fig. 3**"],
       ["Slope classification map", "**INCLUDED — Fig. 4** (additional)"],
       ["FEMA flood map", "**INCLUDED — Fig. 5**"],
       ["Lake / bathymetric map", "**NOT INCLUDED** — no bathymetric data obtained for this cove"],
       ["Location / access map", "**INCLUDED — Fig. 7**"],
       ["Building-envelope diagram", "**INCLUDED — Fig. 6**"],
       ["Climate chart", "**INCLUDED — Fig. 8** (additional)"]],
      widths=[5.6, 11.0], fontsize=8.2)

pagebreak()

# =============================================================== CLOSING
doc.add_heading("Recommended Next Actions", level=1)
para("In order. Stop and reassess if any of the first four produces an unfavourable answer.", 9)
table(["Step", "Action", "Who", "Why it is in this position"],
      [["1", "Call the USACE Thurmond Project Office on 800-533-3478. Ask for the shoreline "
             "allocation class adjacent to Lots 130 and 131, and whether a private dock is "
             "permittable. Get it in writing.", "You", "Free, fast, and it can kill or confirm the "
             "entire investment thesis in one call."],
       ["2", "Pull the recorded Stillwater Coves plat and note the RECORDING DATE.",
        "Clerk of Superior Court", "If it postdates 1 May 2018, private docks are off the table "
        "permanently."],
       ["3", "Get the current written asking price, status and listing history.", "Listing broker",
        "You cannot evaluate value against a 3x price spread from cached pages."],
       ["4", "Commission a Level 3 or 4 soil report and ask environmental health for the "
             "permittable bedroom count.", "Soil classifier + county",
        "Determines the house size, and therefore the whole budget."],
       ["5", "Request the full HOA document set: CC&Rs, architectural guidelines, budget, dues, "
             "assessments, minutes.", "Stillwater Coves POA",
        "Resolves STR, livestock, tree clearing, minimum house size and dock covenants at once."],
       ["6", "Confirm zoning district, setbacks, minimum lot size, ADU and STR rules.",
        "Lincoln County Planning & Zoning", "One conversation resolves most of Section G."],
       ["7", "Obtain utility availability letters for water, electricity and broadband, "
             "addressed to the parcel.", "Each utility",
        "Converts marketing claims into commitments."],
       ["8", "Order a boundary survey locating the USACE line at the rear.", "Licensed surveyor",
        "Establishes what you are actually buying."],
       ["9", "Obtain the county floodplain determination and required flood elevation.",
        "Floodplain administrator", "No BFE exists, so only the county can tell you."],
       ["10", "Commission a geotechnical report before finalising foundation design.",
        "Geotechnical engineer", "Clay subsoils, severely eroded surface, possible shallow rock."]],
      widths=[0.9, 6.6, 3.6, 5.5], fontsize=8.0)

doc.add_heading("Sources", level=1)
para("Primary datasets queried directly on 7 October 2026:", 9, bold=True)
bullets([
    ("USGS 3DEP elevation —", "elevation.nationalmap.gov and epqs.nationalmap.gov. "
     "1-metre lidar, GA_Statewide_2018_B18_DRRA, NAVD 88. 2,665 grid samples."),
    ("FEMA National Flood Hazard Layer —", "hazards.fema.gov, layer 28 (S_FLD_HAZ_AR). DFIRM 13181C."),
    ("USDA NRCS SSURGO —", "sdmdataaccess.sc.egov.usda.gov Soil Data Access (tabular and spatial)."),
    ("OpenStreetMap —", "Nominatim and the OSM API (way 409863360, Watersedge Cove). "
     "Data copyright OpenStreetMap contributors, ODbL 1.0."),
    ("OSRM —", "router.project-osrm.org for road-network distances and drive times."),
    ("ERA5 reanalysis —", "archive-api.open-meteo.com, daily 1995–2024 at the site coordinates."),
], size=8.6)
para("Published and agency sources consulted:", 9, bold=True)
bullets([
    ("USACE Savannah District —", "J. Strom Thurmond Dam and Lake: Shoreline Management "
     "(allocation maps, permit conditions, FAQ, how to apply), Boat Ramp Bottom Elevations, "
     "and the 2018 Shoreline Management Plan news release. sas.usace.army.mil"),
    ("USACE water management —", "water.sas.usace.army.mil (conservation and flood-control storage)."),
    ("USGS —", "pubs.usgs.gov water-data report (spillway and drawdown elevations)."),
    ("Georgia Department of Public Health —", "On-Site Sewage Management Systems Manual and "
     "Rules Chapter 511-3-1. dph.georgia.gov and rules.sos.georgia.gov"),
    ("US Census Bureau —", "QuickFacts, Lincoln County, Georgia. census.gov"),
    ("Georgia DNR / coastalgadnr.org and georgiawildlife.com —", "Clarks Hill fish species."),
    ("South Carolina DNR —", "dnr.sc.gov Thurmond Lake fishing regulations."),
    ("EPA —", "Map of Radon Zones and radon action levels. epa.gov"),
    ("Lincoln County, Georgia —", "lincolncountyga.org (healthcare, infrastructure, "
     "electric providers)."),
    ("Stillwater Coves POA —", "stillwatercovesgeorgia.com (community description, "
     "Plan Approval Checklist)."),
    ("CSRA Regional Commission / Georgia DCA —", "Lincoln County and City of Lincolnton "
     "Joint Comprehensive Plan."),
    ("Listing aggregators —", "realtor.com, redfin.com, trulia.com, georgiamls.com, homes.com "
     "(prices, lot descriptions, comparables). Accessed via search-engine cache; "
     "direct fetches were blocked."),
    ("Build-cost references —", "homeguide.com, ibuyer.com, realpha.com, jnjcustomhomes.com, "
     "homeblue.com, homebuilderdigest.com (2026 Georgia and Augusta-area ranges)."),
    ("Broadband references —", "broadbandnow.com, broadbandmap.com, bestneighborhood.org "
     "(Tignall ZIP-level availability)."),
    ("School references —", "niche.com, publicschoolreview.com, greatschools.org."),
], size=8.6)
para("Content from third-party commercial sources has been paraphrased and summarised rather than "
     "quoted, and was rephrased for compliance with licensing restrictions. Agency and statutory "
     "material is public domain. No listing photographs are reproduced.", 8.5, italic=True, color=GREY)

doc.add_heading("Limitations and Disclaimer", level=1)
para("This is an automated desk study compiled from public data on 7 October 2026. It is not a survey, "
     "an appraisal, a geotechnical or environmental assessment, a legal opinion, a flood determination "
     "or investment advice, and it must not be used in place of any of them.", 9, bold=True)
bullets([
    "No site visit was made. No part of the property was physically inspected.",
    "No recorded plat, deed, title commitment, survey or HOA document was obtained. "
    "All statements about lot boundaries, dimensions, frontage, setbacks and buildable area rest on "
    "a modelled 150 ft x 310 ft rectangle and may be materially wrong.",
    "Zillow, Redfin and realtor.com all blocked automated access. Price, status and listing history "
    "are unverified, and the figures found conflict with one another.",
    "Lincoln County's parcel GIS is access-restricted, so no parcel ID, assessed value or official "
    "boundary was retrieved.",
    "Elevations are lidar-derived to NAVD 88 and are suitable for planning only. Flood and permit "
    "decisions require a surveyed Elevation Certificate.",
    "Soil interpretations come from survey-scale SSURGO mapping. NRCS states explicitly that "
    "survey-scale maps are insufficient to approve property for on-site sewage systems; "
    "an on-site evaluation is required.",
    "Climate figures are from ERA5 reanalysis, a gridded model product that smooths extremes.",
    "Dock feasibility is the single largest unresolved question and depends on USACE determinations "
    "that only the Thurmond Project Office can make.",
    "Regulations, prices, agency policies and permit rules change. Verify currency of every item "
    "before relying on it.",
], size=8.8)
para("Engage a licensed Georgia surveyor, geotechnical engineer, soil classifier and real-estate "
     "attorney, and obtain written determinations from the USACE Thurmond Project Office, "
     "Lincoln County Planning & Zoning, Lincoln County environmental health and the county "
     "floodplain administrator, before committing funds.", 9, bold=True)

doc.save(OUT)
print("SAVED ->", OUT)
print("size:", round(os.path.getsize(OUT) / 1024, 1), "KB")
