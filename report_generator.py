"""
Generates WC Premium Cost Breakout and Broadmoor Comp Reimbursement PDFs.
"""

import io
import re as _re
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)

# Net rates per $100 of compensable payroll — from policy declarations
NET_RATES = {
    "CO8810": 0.07,
    "CO8869": 0.69,
    "CO9180": 4.00,
    "FL8742": 0.23,
    "FL9180": 2.57,
    "GA9180": 2.81,
    "VA8742": 0.10,
    "VA9180": 1.99,
    "WV8742": 0.08,
    "WV8810": 0.05,
    "WV9180": 1.36,
}

# State groupings in display order
STATES = [
    {
        "name": "Colorado",
        "codes": [
            {"code": "CO8810", "description": "CO8810  Clerical"},
            {"code": "CO8869", "description": "CO8869  Daycare: Prof/Clerical"},
            {"code": "CO9180", "description": "CO9180  Amusement Operation-Not Travel", "use_subgroups": True},
        ],
    },
    {
        "name": "Florida",
        "codes": [
            {"code": "FL8742", "description": "FL8742  Outside Salesperson"},
            {"code": "FL9180", "description": "FL9180  Amusement Operation-Not Travel"},
        ],
    },
    {
        "name": "Georgia",
        "codes": [
            {"code": "GA9180", "description": "GA9180  Amusement Operation-Not Travel"},
        ],
    },
    {
        "name": "Virginia",
        "codes": [
            {"code": "VA8742", "description": "VA8742  Outside Salesperson"},
            {"code": "VA9180", "description": "VA9180  Amusement Operation-Not Travel"},
        ],
    },
    {
        "name": "West Virginia",
        "codes": [
            {"code": "WV8742", "description": "WV8742  Outside Salesperson"},
            {"code": "WV8810", "description": "WV8810  Clerical"},
            {"code": "WV9180", "description": "WV9180  Amusement Operation-Not Travel"},
        ],
    },
]

# CO9180 sub-groups in display order
CO9180_SUBGROUPS_ORDER = [
    "Soaring Guides — Depts 15 & 20",
    "Outfitter Guides — Depts 9 & 21",
    "Falconry - Broadmoor — Dept 18",
]

SUBGROUPS_DEF = [
    {
        "label": "Soaring Guides (incl. Reservations & Mgrs)",
        "co9180_depts": {15, 20},
        "co8810_depts": {16, 17},
        "blended": True,   # no single rate — depts carry different rates
    },
    {
        "label": "Outfitter Guides",
        "co9180_depts": {9, 21},
        "co8810_depts": set(),
        "blended": False,
    },
    {
        "label": "Falconry - Broadmoor",
        "co9180_depts": {18},
        "co8810_depts": set(),
        "blended": False,
    },
]


# ── Shared helpers ────────────────────────────────────────────────────────────

def _fmt(v: float) -> str:
    return f"${v:,.2f}"


def _est_premium(compensable: float, rate: float) -> float:
    return round((compensable / 100.0) * rate, 2)


def _dept_num_local(dept: str):
    m = _re.match(r'^\s*(\d+)', str(dept))
    try:
        return int(m.group(1)) if m else None
    except (ValueError, AttributeError):
        return None


def _dept_totals_from_detail(dept_str: str, subgroup_detail: dict) -> tuple[float, float]:
    """Return (gross, excluded) for a department from subgroup_detail."""
    employees = subgroup_detail.get(dept_str, {})
    gross = sum(e["gross"] for e in employees.values())
    excl = sum(e["excluded"] for e in employees.values())
    return round(gross, 2), round(excl, 2)


def _meta_table(named_insured: str, policy_number: str, producer: str,
                policy_period: str, period_date: str) -> Table:
    meta_data = [
        ["Named Insured:", named_insured, "Producer:", producer],
        ["Policy Number:", policy_number, "Reporting Period End:", period_date],
        ["Policy Period:", policy_period, "", ""],
    ]
    t = Table(meta_data, colWidths=[1.2*inch, 2.3*inch, 1.3*inch, 2.0*inch])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _shared_colors():
    return {
        "HEADER_BG":        colors.HexColor("#1a3a5c"),
        "TOTAL_BG":         colors.HexColor("#dce6f0"),
        "GRAND_BG":         colors.HexColor("#b8cfe0"),
        "ALT_ROW":          colors.HexColor("#f4f7fb"),
        "SOARING_CO8810_BG": colors.HexColor("#fffacd"),
        "SUBGROUP_HDR_BG":  colors.HexColor("#2d5986"),
        "DEPT_HDR_BG":      colors.HexColor("#e4edf7"),
        "NOTE_BG":          colors.HexColor("#fffbe6"),
    }


# ── WC PREMIUM COST BREAKOUT ──────────────────────────────────────────────────

def generate_pdf(
    period_date: str,
    class_codes: list[dict],
    co9180_subgroups: list[dict],
    soaring_co8810: dict | None = None,
    subgroup_detail: dict | None = None,
) -> bytes:
    """
    Generate the WC Premium Cost Breakout PDF.

    CO8810 is displayed as the non-Soaring portion only (depts 16 & 17 are shown
    inside the CO9180 Soaring Guides subgroup for cost-allocation visibility and
    are not double-counted in state or grand totals).
    """
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        month_label = dt.strftime("%B %Y")
    except Exception:
        month_label = period_date

    code_map = {item["code"]: item for item in class_codes}
    sc = soaring_co8810 or {"gross": 0.0, "excluded": 0.0, "compensable": 0.0}
    detail = subgroup_detail or {}
    C = _shared_colors()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=0.6*inch, rightMargin=0.6*inch,
        topMargin=0.6*inch, bottomMargin=0.6*inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("T", parent=styles["Normal"],
                                 fontSize=14, fontName="Helvetica-Bold", spaceAfter=2)
    subtitle_style = ParagraphStyle("ST", parent=styles["Normal"],
                                    fontSize=10, spaceAfter=4)
    meta_label_style = ParagraphStyle("ML", parent=styles["Normal"],
                                      fontSize=9, spaceAfter=4)
    note_style = ParagraphStyle("N", parent=styles["Normal"],
                                fontSize=7.5, spaceAfter=6,
                                textColor=colors.HexColor("#444444"))
    state_style = ParagraphStyle("SH", parent=styles["Normal"],
                                 fontSize=10, fontName="Helvetica-Bold",
                                 spaceBefore=10, spaceAfter=4)

    col_headers = ["Class Code / Employee", "Rate\n/$100",
                   "Gross Payroll", "Excl. Payroll", "Compensable", "Est. Premium"]
    col_widths = [2.5*inch, 0.65*inch, 1.05*inch, 1.05*inch, 1.05*inch, 0.9*inch]

    story = []

    # ── Title block ───────────────────────────────────────────────────────────
    story.append(Paragraph("WC PREMIUM COST BREAKOUT", title_style))
    story.append(Paragraph("Workers Compensation — Estimated Premium by State and Class Code",
                            subtitle_style))
    story.append(_meta_table(
        "Greenbrier Outfitters", "37-664725-01-01; -01-02",
        "Ashley Cazire", "04/15/26 – 04/15/27", period_date,
    ))
    story.append(Spacer(1, 6))

    grand_gross = grand_excl = grand_comp = grand_prem = 0.0

    for state in STATES:
        story.append(Paragraph(state["name"], state_style))

        table_data = [col_headers]
        row_styles = []
        state_gross = state_excl = state_comp = state_prem = 0.0

        for entry in state["codes"]:
            code = entry["code"]
            cd = code_map.get(code, {"gross": 0.0, "excluded": 0.0, "compensable": 0.0})
            g   = float(cd.get("gross", 0))
            ex  = float(cd.get("excluded", 0))
            comp = float(cd.get("compensable", 0))
            rate = NET_RATES.get(code, 0.0)

            if entry.get("use_subgroups"):
                # ── CO9180 main row (full CO9180 amount) ─────────────────────
                prem = _est_premium(comp, rate)
                table_data.append([
                    entry["description"], f"{rate:.2f}",
                    _fmt(g), _fmt(ex), _fmt(comp), _fmt(prem),
                ])
                r = len(table_data) - 1
                row_styles.extend([
                    ("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"),
                ])

                # State totals: CO9180 full amount
                state_gross += g
                state_excl  += ex
                state_comp  += comp
                state_prem  += prem

                # ── CO9180 subgroups ──────────────────────────────────────────
                for sg_def in SUBGROUPS_DEF:
                    # Collect dept strings matching this subgroup
                    co9180_dept_strs = sorted(
                        d for d in detail if _dept_num_local(d) in sg_def["co9180_depts"]
                    )
                    co8810_dept_strs = sorted(
                        d for d in detail if _dept_num_local(d) in sg_def["co8810_depts"]
                    )
                    all_dept_strs = co9180_dept_strs + co8810_dept_strs

                    # Compute subgroup totals
                    sg_co9180_gross = sg_co9180_excl = sg_co9180_comp = 0.0
                    sg_co8810_gross = sg_co8810_excl = sg_co8810_comp = 0.0

                    if all_dept_strs:
                        for d in co9180_dept_strs:
                            dg, de = _dept_totals_from_detail(d, detail)
                            sg_co9180_gross += dg
                            sg_co9180_excl  += de
                            sg_co9180_comp  += round(dg - de, 2)
                        for d in co8810_dept_strs:
                            dg, de = _dept_totals_from_detail(d, detail)
                            sg_co8810_gross += dg
                            sg_co8810_excl  += de
                            sg_co8810_comp  += round(dg - de, 2)
                    else:
                        # Fall back to co9180_subgroups summary data
                        sg_key = next(
                            (k for k in CO9180_SUBGROUPS_ORDER if sg_def["label"].split(" (")[0] in k),
                            sg_def["label"]
                        )
                        sg_data = next(
                            (s for s in co9180_subgroups if sg_key in s.get("description", "")),
                            {"gross": 0.0, "excluded": 0.0, "compensable": 0.0},
                        )
                        sg_co9180_gross = float(sg_data.get("gross", 0))
                        sg_co9180_excl  = float(sg_data.get("excluded", 0))
                        sg_co9180_comp  = float(sg_data.get("compensable", 0))
                        if sg_def["blended"]:
                            sg_co8810_gross = float(sc.get("gross", 0))
                            sg_co8810_excl  = float(sc.get("excluded", 0))
                            sg_co8810_comp  = float(sc.get("compensable", 0))

                    sg_tot_gross = round(sg_co9180_gross + sg_co8810_gross, 2)
                    sg_tot_excl  = round(sg_co9180_excl  + sg_co8810_excl,  2)
                    sg_tot_comp  = round(sg_co9180_comp  + sg_co8810_comp,  2)
                    sg_tot_prem  = round(
                        _est_premium(sg_co9180_comp, NET_RATES["CO9180"]) +
                        _est_premium(sg_co8810_comp, NET_RATES["CO8810"]), 2
                    )

                    sg_rate_str = "" if sg_def["blended"] else f"{NET_RATES['CO9180']:.2f}"

                    # Subgroup header row (colored, with numbers)
                    sg_hdr_idx = len(table_data)
                    table_data.append([
                        f"  {sg_def['label']}", sg_rate_str,
                        _fmt(sg_tot_gross), _fmt(sg_tot_excl),
                        _fmt(sg_tot_comp), _fmt(sg_tot_prem),
                    ])
                    row_styles.extend([
                        ("BACKGROUND", (0, sg_hdr_idx), (-1, sg_hdr_idx), C["SUBGROUP_HDR_BG"]),
                        ("TEXTCOLOR",  (0, sg_hdr_idx), (-1, sg_hdr_idx), colors.white),
                        ("FONTNAME",   (0, sg_hdr_idx), (-1, sg_hdr_idx), "Helvetica-Bold"),
                    ])

                    if all_dept_strs:
                        # Dept headers + employee rows
                        for dept_str in all_dept_strs:
                            is_co8810 = _dept_num_local(dept_str) in sg_def["co8810_depts"]
                            dept_rate = NET_RATES["CO8810"] if is_co8810 else NET_RATES["CO9180"]
                            dept_label = f"    {dept_str}" + (" (CO8810)" if is_co8810 else "")
                            dept_hdr_idx = len(table_data)
                            table_data.append([dept_label, "", "", "", "", ""])
                            dept_bg = C["SOARING_CO8810_BG"] if is_co8810 else C["DEPT_HDR_BG"]
                            row_styles.extend([
                                ("BACKGROUND", (0, dept_hdr_idx), (-1, dept_hdr_idx), dept_bg),
                                ("FONTNAME",   (0, dept_hdr_idx), (-1, dept_hdr_idx), "Helvetica-Bold"),
                                ("FONTSIZE",   (0, dept_hdr_idx), (-1, dept_hdr_idx), 7.5),
                                ("ITALIC",     (0, dept_hdr_idx), (0, dept_hdr_idx), 1),
                            ])

                            employees = detail.get(dept_str, {})
                            for emp_name in sorted(employees.keys()):
                                emp = employees[emp_name]
                                eg    = emp["gross"]
                                eex   = emp["excluded"]
                                ecomp = round(eg - eex, 2)
                                eprem = _est_premium(ecomp, dept_rate)
                                emp_idx = len(table_data)
                                table_data.append([
                                    f"      {emp_name}", "",
                                    _fmt(eg), _fmt(eex), _fmt(ecomp), _fmt(eprem),
                                ])
                                if is_co8810:
                                    row_styles.append((
                                        "BACKGROUND", (0, emp_idx), (-1, emp_idx),
                                        C["SOARING_CO8810_BG"]
                                    ))

                    # Subgroup total row
                    sg_tot_idx = len(table_data)
                    tot_label = f"  {sg_def['label'].split(' (')[0]} Total"
                    if sg_def["blended"]:
                        tot_label += " *"
                    table_data.append([
                        tot_label, sg_rate_str,
                        _fmt(sg_tot_gross), _fmt(sg_tot_excl),
                        _fmt(sg_tot_comp), _fmt(sg_tot_prem),
                    ])
                    row_styles.extend([
                        ("BACKGROUND", (0, sg_tot_idx), (-1, sg_tot_idx), C["TOTAL_BG"]),
                        ("FONTNAME",   (0, sg_tot_idx), (-1, sg_tot_idx), "Helvetica-Bold"),
                    ])

            else:
                # ── Standard class code row ───────────────────────────────────
                # CO8810: display non-Soaring portion only (soaring shown in CO9180 section)
                if code == "CO8810" and soaring_co8810:
                    g    = round(g    - float(sc.get("gross", 0)), 2)
                    ex   = round(ex   - float(sc.get("excluded", 0)), 2)
                    comp = round(comp - float(sc.get("compensable", 0)), 2)

                prem = _est_premium(comp, rate)
                table_data.append([
                    entry["description"], f"{rate:.2f}",
                    _fmt(g), _fmt(ex), _fmt(comp), _fmt(prem),
                ])
                state_gross += g
                state_excl  += ex
                state_comp  += comp
                state_prem  += prem

        # State subtotal row
        sub_idx = len(table_data)
        table_data.append([
            f"{state['name']} Subtotal", "",
            _fmt(round(state_gross, 2)), _fmt(round(state_excl, 2)),
            _fmt(round(state_comp,  2)), _fmt(round(state_prem, 2)),
        ])
        row_styles.extend([
            ("BACKGROUND", (0, sub_idx), (-1, sub_idx), C["TOTAL_BG"]),
            ("FONTNAME",   (0, sub_idx), (-1, sub_idx), "Helvetica-Bold"),
        ])

        grand_gross += state_gross
        grand_excl  += state_excl
        grand_comp  += state_comp
        grand_prem  += state_prem

        base_style = [
            ("BACKGROUND", (0, 0), (-1, 0), C["HEADER_BG"]),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("ALIGN",      (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN",      (0, 0), (0, -1),  "LEFT"),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("GRID",       (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
        ] + row_styles

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(base_style))
        story.append(t)
        story.append(Spacer(1, 6))

    # Grand total
    grand_data = [
        ["", "Gross Payroll", "Excl. Payroll", "Compensable", "Est. Premium"],
        ["GRAND TOTAL",
         _fmt(round(grand_gross, 2)), _fmt(round(grand_excl, 2)),
         _fmt(round(grand_comp,  2)), _fmt(round(grand_prem, 2))],
    ]
    grand_widths = [2.5*inch, 1.15*inch, 1.15*inch, 1.15*inch, 1.0*inch]
    gt = Table(grand_data, colWidths=grand_widths)
    gt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C["HEADER_BG"]),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, 1), C["GRAND_BG"]),
        ("FONTNAME",   (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("ALIGN",      (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN",      (0, 0), (0, -1),  "LEFT"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("GRID",       (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
    ]))
    story.append(gt)
    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "Note: Estimated premiums use net rates from policy declarations and are for internal "
        "reference only; final amounts subject to audit. The Soaring Guides sub-group includes "
        "Depts 16 (Reservations) and 17 (Managers), which are reported under CO8810 on the "
        "carrier form — included here for display only, not double-counted. "
        "* Soaring Guides sub-group premium is blended: Depts 15 & 20 rated at CO9180 "
        "($4.00/$100); Depts 16 & 17 rated at CO8810 ($0.07/$100).",
        ParagraphStyle("NF", parent=styles["Normal"], fontSize=7.5,
                       textColor=colors.HexColor("#555555")),
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ── BROADMOOR COMP REIMBURSEMENT ──────────────────────────────────────────────

# Soaring depts for the Broadmoor report
BROADMOOR_CO8810_DEPTS = {16, 17}
BROADMOOR_CO9180_DEPTS = {15, 20}


def generate_broadmoor_pdf(
    period_date: str,
    subgroup_detail: dict | None = None,
    soaring_co8810: dict | None = None,
    co9180_subgroups: list | None = None,
) -> bytes:
    """
    Generate the Broadmoor Comp Reimbursement PDF.
    Covers Soaring depts only: CO8810 (16, 17) and CO9180 (15, 20).
    Shows dept-level aggregates; individual employee rows are not included.
    """
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
    except Exception:
        dt = None

    detail = subgroup_detail or {}
    C = _shared_colors()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=0.6*inch, rightMargin=0.6*inch,
        topMargin=0.6*inch, bottomMargin=0.6*inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("T", parent=styles["Normal"],
                                 fontSize=14, fontName="Helvetica-Bold", spaceAfter=2)
    subtitle_style = ParagraphStyle("ST", parent=styles["Normal"],
                                    fontSize=10, spaceAfter=4)
    note_style = ParagraphStyle("N", parent=styles["Normal"],
                                fontSize=7.5, textColor=colors.HexColor("#555555"),
                                backColor=C["NOTE_BG"], borderPadding=4)

    col_headers = ["Department / Class Code", "Rate\n/$100",
                   "Gross Payroll", "Excl. Payroll", "Compensable", "Est. Premium"]
    col_widths = [2.5*inch, 0.65*inch, 1.05*inch, 1.05*inch, 1.05*inch, 0.9*inch]

    story = []

    # ── Title block ───────────────────────────────────────────────────────────
    story.append(Paragraph("BROADMOOR COMP REIMBURSEMENT", title_style))
    story.append(Paragraph("Soaring Adventure — Internal Workers Compensation Cost Allocation",
                            subtitle_style))
    story.append(_meta_table(
        "Greenbrier Outfitters", "37-664725-01-01; -01-02",
        "Ashley Cazire", "04/15/26 – 04/15/27", period_date,
    ))
    story.append(Spacer(1, 8))

    # ── Helper: dept totals ───────────────────────────────────────────────────
    def dept_rows_for(dept_nums: set, rate: float):
        """Return list of (dept_str, gross, excl, comp, prem) for matching depts."""
        rows = []
        for d in sorted(detail.keys(), key=lambda x: _dept_num_local(x) or 999):
            dn = _dept_num_local(d)
            if dn in dept_nums:
                dg, de = _dept_totals_from_detail(d, detail)
                dc = round(dg - de, 2)
                dp = _est_premium(dc, rate)
                rows.append((d, dg, de, dc, dp))
        return rows

    # ── CO8810 section ────────────────────────────────────────────────────────
    co8810_rate = NET_RATES["CO8810"]
    co8810_dept_rows = dept_rows_for(BROADMOOR_CO8810_DEPTS, co8810_rate)

    # CO8810 total from soaring_co8810 (pre-computed) or sum of dept rows
    if soaring_co8810 and (soaring_co8810.get("gross", 0) or co8810_dept_rows):
        if soaring_co8810.get("gross", 0):
            co8810_tot_g    = float(soaring_co8810["gross"])
            co8810_tot_ex   = float(soaring_co8810["excluded"])
            co8810_tot_comp = float(soaring_co8810["compensable"])
        else:
            co8810_tot_g    = sum(r[1] for r in co8810_dept_rows)
            co8810_tot_ex   = sum(r[2] for r in co8810_dept_rows)
            co8810_tot_comp = sum(r[3] for r in co8810_dept_rows)
    else:
        co8810_tot_g    = sum(r[1] for r in co8810_dept_rows)
        co8810_tot_ex   = sum(r[2] for r in co8810_dept_rows)
        co8810_tot_comp = sum(r[3] for r in co8810_dept_rows)

    co8810_tot_prem = _est_premium(round(co8810_tot_comp, 2), co8810_rate)

    # CO9180 section
    co9180_rate = NET_RATES["CO9180"]
    co9180_dept_rows = dept_rows_for(BROADMOOR_CO9180_DEPTS, co9180_rate)

    # CO9180 Soaring total from co9180_subgroups or sum of dept rows
    sg_soaring = next(
        (s for s in (co9180_subgroups or [])
         if "Soaring" in s.get("description", "")),
        None,
    )
    if sg_soaring and sg_soaring.get("gross", 0) and not co9180_dept_rows:
        co9180_tot_g    = float(sg_soaring["gross"])
        co9180_tot_ex   = float(sg_soaring["excluded"])
        co9180_tot_comp = float(sg_soaring["compensable"])
    else:
        co9180_tot_g    = sum(r[1] for r in co9180_dept_rows)
        co9180_tot_ex   = sum(r[2] for r in co9180_dept_rows)
        co9180_tot_comp = sum(r[3] for r in co9180_dept_rows)

    co9180_tot_prem = _est_premium(round(co9180_tot_comp, 2), co9180_rate)

    # Grand totals
    grand_g    = round(co8810_tot_g    + co9180_tot_g,    2)
    grand_ex   = round(co8810_tot_ex   + co9180_tot_ex,   2)
    grand_comp = round(co8810_tot_comp + co9180_tot_comp, 2)
    grand_prem = round(co8810_tot_prem + co9180_tot_prem, 2)

    # ── Build table ───────────────────────────────────────────────────────────
    table_data = [col_headers]
    row_styles = []

    def _section_header(label: str):
        """Bold left-spanning section label row (e.g. 'CO8810 — Clerical')."""
        idx = len(table_data)
        table_data.append([label, "", "", "", "", ""])
        row_styles.extend([
            ("FONTNAME",  (0, idx), (-1, idx), "Helvetica-Bold"),
            ("FONTSIZE",  (0, idx), (-1, idx), 9),
            ("SPAN",      (0, idx), (-1, idx)),
            ("TOPPADDING",    (0, idx), (-1, idx), 5),
            ("BOTTOMPADDING", (0, idx), (-1, idx), 5),
        ])

    def _class_row(label: str, rate: float, g: float, ex: float, comp: float, bold=False):
        idx = len(table_data)
        table_data.append([
            label, f"{rate:.2f}",
            _fmt(round(g, 2)), _fmt(round(ex, 2)),
            _fmt(round(comp, 2)), _fmt(_est_premium(round(comp, 2), rate)),
        ])
        if bold:
            row_styles.extend([
                ("FONTNAME", (0, idx), (-1, idx), "Helvetica-Bold"),
            ])

    def _dept_row(label: str, g: float, ex: float, comp: float, prem: float, yellow=False):
        idx = len(table_data)
        table_data.append([
            f"  {label}", "",
            _fmt(round(g, 2)), _fmt(round(ex, 2)),
            _fmt(round(comp, 2)), _fmt(round(prem, 2)),
        ])
        if yellow:
            row_styles.append(("BACKGROUND", (0, idx), (-1, idx), C["SOARING_CO8810_BG"]))

    def _subtotal_row(label: str, rate: float, g: float, ex: float, comp: float):
        idx = len(table_data)
        table_data.append([
            label, f"{rate:.2f}",
            _fmt(round(g, 2)), _fmt(round(ex, 2)),
            _fmt(round(comp, 2)), _fmt(_est_premium(round(comp, 2), rate)),
        ])
        row_styles.extend([
            ("BACKGROUND", (0, idx), (-1, idx), C["TOTAL_BG"]),
            ("FONTNAME",   (0, idx), (-1, idx), "Helvetica-Bold"),
        ])

    # CO8810 section
    _section_header("CO8810 — Clerical")
    _class_row("CO8810  Clerical", co8810_rate, co8810_tot_g, co8810_tot_ex, co8810_tot_comp)
    for dept_str, dg, de, dc, dp in co8810_dept_rows:
        _dept_row(dept_str, dg, de, dc, dp, yellow=True)
    _subtotal_row("CO8810 Subtotal", co8810_rate, co8810_tot_g, co8810_tot_ex, co8810_tot_comp)

    # Spacer row
    table_data.append(["", "", "", "", "", ""])
    row_styles.append(("TOPPADDING", (0, len(table_data)-1), (-1, len(table_data)-1), 3))
    row_styles.append(("BOTTOMPADDING", (0, len(table_data)-1), (-1, len(table_data)-1), 3))

    # CO9180 section
    _section_header("CO9180 — Amusement Operation-Not Travel")
    _class_row("CO9180  Amusement Operation-Not Travel", co9180_rate,
               co9180_tot_g, co9180_tot_ex, co9180_tot_comp)
    for dept_str, dg, de, dc, dp in co9180_dept_rows:
        _dept_row(dept_str, dg, de, dc, dp, yellow=False)
    _subtotal_row("CO9180 Subtotal", co9180_rate, co9180_tot_g, co9180_tot_ex, co9180_tot_comp)

    # Grand total
    grand_idx = len(table_data)
    table_data.append([
        "GRAND TOTAL", "",
        _fmt(grand_g), _fmt(grand_ex),
        _fmt(grand_comp), _fmt(grand_prem),
    ])
    row_styles.extend([
        ("BACKGROUND", (0, grand_idx), (-1, grand_idx), C["GRAND_BG"]),
        ("FONTNAME",   (0, grand_idx), (-1, grand_idx), "Helvetica-Bold"),
    ])

    base_style = [
        ("BACKGROUND", (0, 0), (-1, 0), C["HEADER_BG"]),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 8),
        ("ALIGN",      (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN",      (0, 0), (0, -1),  "LEFT"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("GRID",       (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
    ] + row_styles

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(base_style))
    story.append(t)
    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "NOTE: Estimated premiums use net rates from policy declarations and are for "
        "internal reference only; final amounts subject to audit.",
        note_style,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()
