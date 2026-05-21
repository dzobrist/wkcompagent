"""
Generates the Workers' Compensation Premium Breakout PDF report.
Matches the format of WC_Premium_Breakout_[Month]_[Year].pdf.
"""

import io
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
        "name": "COLORADO",
        "codes": [
            {"code": "CO8810", "description": "Clerical (CO8810)"},
            {"code": "CO8869", "description": "Daycare: Prof/Clerical (CO8869)"},
            {"code": "CO9180", "description": None, "use_subgroups": True},
        ],
    },
    {
        "name": "FLORIDA",
        "codes": [
            {"code": "FL8742", "description": "Outside Salesperson (FL8742)"},
            {"code": "FL9180", "description": "Amusement Operation-Not Travel (FL9180)"},
        ],
    },
    {
        "name": "GEORGIA",
        "codes": [
            {"code": "GA9180", "description": "Amusement Operation-Not Travel (GA9180)"},
        ],
    },
    {
        "name": "VIRGINIA",
        "codes": [
            {"code": "VA8742", "description": "Outside Salesperson (VA8742)"},
            {"code": "VA9180", "description": "Amusement Operation-Not Travel (VA9180)"},
        ],
    },
    {
        "name": "WEST VIRGINIA",
        "codes": [
            {"code": "WV8742", "description": "Outside Salesperson (WV8742)"},
            {"code": "WV8810", "description": "Clerical (WV8810)"},
            {"code": "WV9180", "description": "Amusement Operation-Not Travel (WV9180)"},
        ],
    },
]

# CO9180 sub-groups for internal detail
CO9180_SUBGROUPS_ORDER = [
    "Soaring Guides — Depts 15 & 20",
    "Outfitter Guides — Depts 9 & 21",
    "Falconry - Broadmoor — Dept 18",
]


def _fmt_dollar(v: float) -> str:
    return f"${v:,.2f}"


def _est_premium(compensable: float, rate: float) -> float:
    return round((compensable / 100) * rate, 2)


def generate_pdf(
    period_date: str,
    class_codes: list[dict],
    co9180_subgroups: list[dict],
    soaring_co8810: dict | None = None,
) -> bytes:
    """
    Generate the WC Premium Breakout PDF.

    Args:
        period_date: e.g. "4/30/2026"
        class_codes: list of {code, gross, excluded, compensable}
        co9180_subgroups: list of {description, gross, excluded, compensable}
        soaring_co8810: {gross, excluded, compensable} for depts 16 & 17 (internal visibility)
    Returns:
        PDF as bytes
    """
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        month_label = dt.strftime("%B %Y")
    except Exception:
        month_label = period_date

    # Index class codes by code
    code_map = {item["code"]: item for item in class_codes}

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=0.6 * inch,
        rightMargin=0.6 * inch,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle", parent=styles["Normal"],
        fontSize=14, fontName="Helvetica-Bold", spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        "Subtitle", parent=styles["Normal"],
        fontSize=11, fontName="Helvetica-Bold", spaceAfter=6
    )
    meta_style = ParagraphStyle(
        "Meta", parent=styles["Normal"],
        fontSize=9, spaceAfter=4
    )
    note_style = ParagraphStyle(
        "Note", parent=styles["Normal"],
        fontSize=7.5, spaceAfter=8, textColor=colors.HexColor("#444444")
    )
    state_style = ParagraphStyle(
        "StateHeader", parent=styles["Normal"],
        fontSize=10, fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4
    )

    HEADER_BG = colors.HexColor("#1a3a5c")
    TOTAL_BG = colors.HexColor("#dce6f0")
    GRAND_BG = colors.HexColor("#b8cfe0")
    ALT_ROW = colors.HexColor("#f4f7fb")
    SOARING_CO8810_BG = colors.HexColor("#fffacd")  # yellow highlight — internal visibility row

    col_headers = ["Description", "Net Rate\n(per $100)", "Gross Payroll",
                   "Excluded Payroll", "Compensable\nPayroll", "Est. Premium"]
    col_widths = [2.4 * inch, 0.75 * inch, 1.1 * inch, 1.1 * inch, 1.1 * inch, 0.9 * inch]

    def header_row():
        return col_headers

    def data_row(description, code, gross, excluded, compensable, rate=None):
        if rate is None:
            rate = NET_RATES.get(code, 0.0)
        premium = _est_premium(compensable, rate)
        return [description, f"{rate:.2f}", _fmt_dollar(gross),
                _fmt_dollar(excluded), _fmt_dollar(compensable), _fmt_dollar(premium)]

    def total_row(label, gross, excluded, compensable, bold=False):
        premium = sum(
            _est_premium(c.get("compensable", 0), NET_RATES.get(c["code"], 0))
            for c in class_codes
            if label.upper() in c["code"][:2].upper() or label == "ALL"
        )
        return [label, "", _fmt_dollar(gross), _fmt_dollar(excluded),
                _fmt_dollar(compensable), _fmt_dollar(premium)]

    story = []

    # Title block
    story.append(Paragraph("Workers' Compensation Premium Breakout", title_style))
    story.append(Paragraph(f"Greenbrier Outfitters — {month_label}", subtitle_style))

    meta_data = [
        ["Named Insured:", "Greenbrier Outfitters", "Producer:", "Ashley Cazire"],
        ["Policy Number:", "37-664725-01-01; -01-02", "Reporting Period:", period_date],
        ["Policy Period:", "04/15/26 – 04/15/27", "", ""],
    ]
    meta_table = Table(meta_data, colWidths=[1.2*inch, 2.3*inch, 1.1*inch, 2.0*inch])
    meta_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 4))

    note = (
        "Estimated premium = (Compensable Payroll ÷ 100) × Net Rate. "
        "Compensable Payroll = Gross Payroll minus Excluded Payroll per WC reporting rules. "
        "Colorado CO9180 is reported as a single class code on the carrier form; "
        "sub-group detail is for internal reference only. "
        "* Yellow row (Soaring Mgmt & Reservations) also appears on the CO8810 carrier line — "
        "shown here for cost-allocation visibility only; not double-counted in the grand total."
    )
    story.append(Paragraph(note, note_style))

    grand_gross = grand_excluded = grand_compensable = grand_premium = 0.0

    for state in STATES:
        story.append(Paragraph(state["name"], state_style))

        table_data = [header_row()]
        state_gross = state_excluded = state_compensable = state_premium = 0.0

        row_styles = []
        data_row_count = 0

        for entry in state["codes"]:
            code = entry["code"]
            cd = code_map.get(code, {"gross": 0.0, "excluded": 0.0, "compensable": 0.0})

            if entry.get("use_subgroups") and co9180_subgroups:
                # ── Soaring section: CO9180 guides + CO8810 mgmt/reservations ──
                soaring_9180_data = next(
                    (s for s in co9180_subgroups if s["description"] == "Soaring Guides — Depts 15 & 20"),
                    {"gross": 0.0, "excluded": 0.0, "compensable": 0.0}
                )
                sg_9180_g    = float(soaring_9180_data.get("gross", 0))
                sg_9180_ex   = float(soaring_9180_data.get("excluded", 0))
                sg_9180_comp = float(soaring_9180_data.get("compensable", 0))

                sc = soaring_co8810 or {"gross": 0.0, "excluded": 0.0, "compensable": 0.0}
                sg_8810_g    = float(sc.get("gross", 0))
                sg_8810_ex   = float(sc.get("excluded", 0))
                sg_8810_comp = float(sc.get("compensable", 0))

                # Row 1 — Soaring Guides CO9180 (depts 15 & 20)
                table_data.append(data_row("Soaring Guides — Depts 15 & 20 (CO9180)", "CO9180",
                                           sg_9180_g, sg_9180_ex, sg_9180_comp))
                data_row_count += 1

                # Row 2 — Soaring Mgmt & Reservations CO8810 (depts 16 & 17) — yellow
                soaring_8810_row_idx = len(table_data)
                table_data.append(data_row(
                    "Soaring Mgmt & Reservations — Depts 16 & 17 (CO8810)*",
                    "CO8810", sg_8810_g, sg_8810_ex, sg_8810_comp
                ))
                row_styles.append(("BACKGROUND", (0, soaring_8810_row_idx),
                                   (-1, soaring_8810_row_idx), SOARING_CO8810_BG))
                data_row_count += 1

                # Row 3 — Soaring Total rollup
                soaring_total_row_idx = len(table_data)
                soaring_tot_g    = round(sg_9180_g + sg_8810_g, 2)
                soaring_tot_ex   = round(sg_9180_ex + sg_8810_ex, 2)
                soaring_tot_comp = round(sg_9180_comp + sg_8810_comp, 2)
                soaring_tot_prem = round(
                    _est_premium(sg_9180_comp, NET_RATES["CO9180"]) +
                    _est_premium(sg_8810_comp, NET_RATES["CO8810"]), 2
                )
                table_data.append([
                    "  Soaring Total",
                    "", _fmt_dollar(soaring_tot_g), _fmt_dollar(soaring_tot_ex),
                    _fmt_dollar(soaring_tot_comp), _fmt_dollar(soaring_tot_prem),
                ])
                row_styles.append(("FONTNAME", (0, soaring_total_row_idx),
                                   (-1, soaring_total_row_idx), "Helvetica-Bold"))
                row_styles.append(("TEXTCOLOR", (0, soaring_total_row_idx),
                                   (-1, soaring_total_row_idx), colors.HexColor("#1a3a5c")))
                data_row_count += 1

                # Remaining CO9180 sub-groups (Outfitter Guides, Falconry)
                for sg in CO9180_SUBGROUPS_ORDER:
                    if sg == "Soaring Guides — Depts 15 & 20":
                        continue  # already rendered above
                    sg_data = next((s for s in co9180_subgroups if s["description"] == sg),
                                   {"gross": 0.0, "excluded": 0.0, "compensable": 0.0})
                    g = float(sg_data.get("gross", 0))
                    ex = float(sg_data.get("excluded", 0))
                    comp = float(sg_data.get("compensable", 0))
                    table_data.append(data_row(sg, code, g, ex, comp))
                    data_row_count += 1
                    if data_row_count % 2 == 0:
                        row_styles.append(("BACKGROUND", (0, data_row_count),
                                           (-1, data_row_count), ALT_ROW))
            else:
                g = float(cd.get("gross", 0))
                ex = float(cd.get("excluded", 0))
                comp = float(cd.get("compensable", 0))
                table_data.append(data_row(entry["description"], code, g, ex, comp))
                data_row_count += 1
                if data_row_count % 2 == 0:
                    row_styles.append(("BACKGROUND", (0, data_row_count), (-1, data_row_count), ALT_ROW))

            rate = NET_RATES.get(code, 0.0)
            state_gross += float(cd.get("gross", 0))
            state_excluded += float(cd.get("excluded", 0))
            state_compensable += float(cd.get("compensable", 0))
            state_premium += _est_premium(float(cd.get("compensable", 0)), rate)

        # State total row
        state_total_row_idx = len(table_data)
        table_data.append([
            f"TOTAL {state['name']}",
            "",
            _fmt_dollar(state_gross),
            _fmt_dollar(state_excluded),
            _fmt_dollar(state_compensable),
            _fmt_dollar(state_premium),
        ])

        grand_gross += state_gross
        grand_excluded += state_excluded
        grand_compensable += state_compensable
        grand_premium += state_premium

        base_style = [
            # Header
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
            # Total row
            ("BACKGROUND", (0, state_total_row_idx), (-1, state_total_row_idx), TOTAL_BG),
            ("FONTNAME", (0, state_total_row_idx), (-1, state_total_row_idx), "Helvetica-Bold"),
        ] + row_styles

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(base_style))
        story.append(t)
        story.append(Spacer(1, 6))

    # Grand total table
    story.append(Spacer(1, 4))
    grand_data = [
        ["", "Gross Payroll", "Excluded Payroll", "Compensable Payroll", "Est. Total Premium"],
        [
            "ALL STATES COMBINED",
            _fmt_dollar(grand_gross),
            _fmt_dollar(grand_excluded),
            _fmt_dollar(grand_compensable),
            _fmt_dollar(grand_premium),
        ],
    ]
    grand_widths = [2.4*inch, 1.3*inch, 1.3*inch, 1.3*inch, 1.1*inch]
    grand_table = Table(grand_data, colWidths=grand_widths)
    grand_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, 1), GRAND_BG),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
    ]))
    story.append(grand_table)

    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "Estimated premiums use net rates from the current policy declarations and are subject to "
        "final audit adjustment. Contact Ashley Cazire with questions.",
        note_style
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()
