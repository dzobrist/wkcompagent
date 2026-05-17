import io
import os
import re
from datetime import date, timedelta

import openpyxl

from classifier import classify_payroll, load_mapping_rules
from email_service import send_error_email, send_result_email
from report_generator import generate_pdf

KNOWLEDGE_DIR = os.path.join(os.path.dirname(__file__), "knowledge")
MAPPING_RULES_PATH = os.path.join(KNOWLEDGE_DIR, "Mapping_Rules.xlsx")
FORM_TEMPLATE_PATH = os.path.join(KNOWLEDGE_DIR, "Form_Template.xlsx")

# Class codes in the exact order they appear on the form (rows 17–27)
CLASS_CODE_ORDER = [
    "CO8810", "CO8869", "CO9180",
    "FL8742", "FL9180",
    "GA9180",
    "VA8742", "VA9180",
    "WV8742", "WV8810", "WV9180",
]
CODE_TO_ROW = {code: 17 + i for i, code in enumerate(CLASS_CODE_ORDER)}

def _build_summary(result: dict, period_date: str) -> str:
    """Build the plain-text summary from pre-calculated classification results."""
    from datetime import datetime
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        month_label = dt.strftime("%B %Y")
    except Exception:
        month_label = period_date

    overrides = result.get("overrides_applied", [])
    exceptions = result.get("exceptions", [])
    dq_notes = result.get("dq_notes", [])
    totals = result["totals"]

    reassign_count = sum(1 for o in overrides if "Reassign" in o)
    exclude_count = sum(1 for o in overrides if "Exclude" in o)

    lines = [
        f"PAYROLL REPORT SUMMARY — {month_label}",
        f"Reporting period end: {period_date}",
        "",
        "Source file: payroll export",
        f"Detail rows processed: {result['rows_processed']}",
        f"Earning entries processed: {result['earning_entries']}",
        f"Overrides applied: {len(overrides)} ({reassign_count} Reassign, {exclude_count} Exclude)",
        f"Exceptions: {len(exceptions)}",
        "",
        "TOTALS:",
        f"  Gross Payroll       ${totals['gross']:>12,.2f}",
        f"  Excluded Payroll    ${totals['excluded']:>12,.2f}",
        f"  Compensable Payroll ${totals['compensable']:>12,.2f}",
        "",
        "BY CLASS CODE:",
    ]
    for cc in result["class_codes"]:
        lines.append(
            f"  {cc['code']}  Gross ${cc['gross']:,.2f}  "
            f"Excluded ${cc['excluded']:,.2f}  Compensable ${cc['compensable']:,.2f}"
        )

    lines += ["", "OVERRIDES APPLIED:"]
    if overrides:
        for o in overrides:
            lines.append(f"  {o}")
    else:
        lines.append("  (none)")

    lines += ["", "EXCEPTIONS (not included in totals):"]
    if exceptions:
        for e in exceptions:
            lines.append(f"  {e}")
    else:
        lines.append("  (none)")

    if dq_notes:
        lines += ["", "DATA QUALITY NOTES:"]
        for n in dq_notes:
            lines.append(f"  {n}")

    lines += ["", "VARIANCE CHECK:", "  No prior month context available."]
    return "\n".join(lines)


def _fill_form_template(class_codes: list[dict], period_date: str) -> bytes:
    wb = openpyxl.load_workbook(FORM_TEMPLATE_PATH)
    ws = wb.active

    ws["B12"] = period_date

    totals = {"gross": 0.0, "excluded": 0.0, "compensable": 0.0}
    submitted_codes = {item["code"]: item for item in class_codes}

    for code in CLASS_CODE_ORDER:
        row = CODE_TO_ROW[code]
        data = submitted_codes.get(code, {"gross": 0.0, "excluded": 0.0, "compensable": 0.0})
        gross = round(float(data.get("gross", 0)), 2)
        excluded = round(float(data.get("excluded", 0)), 2)
        compensable = round(float(data.get("compensable", 0)), 2)
        ws.cell(row=row, column=6).value = gross
        ws.cell(row=row, column=7).value = excluded
        ws.cell(row=row, column=8).value = compensable
        totals["gross"] += gross
        totals["excluded"] += excluded
        totals["compensable"] += compensable

    # Write numeric totals to row 59 (overwrite any formulas)
    ws["F59"] = round(totals["gross"], 2)
    ws["G59"] = round(totals["excluded"], 2)
    ws["H59"] = round(totals["compensable"], 2)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _extract_period_date(text: str) -> str | None:
    match = re.search(r"\b(\d{1,2}/\d{1,2}/\d{4})\b", text)
    return match.group(1) if match else None


async def process_payroll(xlsx_bytes: bytes, period_date: str) -> tuple[bytes, bytes, str]:
    """
    Classify payroll deterministically in Python and return
    (filled_form_bytes, pdf_bytes, summary_text).
    No LLM is used for arithmetic — all calculations are done in classifier.py.
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    if "Earnings by Department" not in wb.sheetnames:
        raise ValueError("Sheet 'Earnings by Department' not found in the uploaded file.")

    ws = wb["Earnings by Department"]
    mapping_rules = load_mapping_rules(MAPPING_RULES_PATH)
    result = classify_payroll(ws, mapping_rules)

    # Log any exceptions / overrides for visibility in Railway logs
    for exc in result.get("exceptions", []):
        print(f"  EXCEPTION: {exc}")
    for ov in result.get("overrides_applied", []):
        print(f"  OVERRIDE:  {ov}")

    form_bytes = _fill_form_template(result["class_codes"], period_date)
    pdf_bytes = generate_pdf(
        period_date=period_date,
        class_codes=result["class_codes"],
        co9180_subgroups=result.get("co9180_subgroups", []),
    )
    summary = _build_summary(result, period_date)
    return form_bytes, pdf_bytes, summary


def _coordinator_emails() -> list[str]:
    """Return list of coordinator emails from comma-separated COORDINATOR_EMAIL env var."""
    raw = os.getenv("COORDINATOR_EMAIL", "")
    return [e.strip() for e in raw.split(",") if e.strip()]


async def handle_inbound(payload: dict):
    """Orchestrate the full pipeline for an inbound coordinator email."""
    coordinators = _coordinator_emails()
    from_email = payload.get("from_email", "").lower()

    if not any(c.lower() in from_email for c in coordinators):
        print(f"Ignoring inbound email from unexpected sender: {from_email}")
        return

    # Find xlsx attachment
    xlsx_bytes = None
    xlsx_filename = None
    for att in payload.get("attachment_files", []):
        if att["filename"].lower().endswith(".xlsx"):
            xlsx_bytes = att["content"]
            xlsx_filename = att["filename"]
            break

    if not xlsx_bytes:
        print("Inbound email from coordinator had no .xlsx attachment — ignoring")
        return

    # Determine period date from email text, or default to last month end
    combined_text = payload.get("subject", "") + " " + payload.get("body", "")
    period_date = _extract_period_date(combined_text)
    if not period_date:
        today = date.today()
        last_day = today.replace(day=1) - timedelta(days=1)
        period_date = f"{last_day.month}/{last_day.day}/{last_day.year}"
        print(f"No period date found in email — defaulting to {period_date}")

    print(f"Processing payroll for period ending {period_date}, file: {xlsx_filename}")

    try:
        form_bytes, pdf_bytes, summary = await process_payroll(xlsx_bytes, period_date)
        for recipient in coordinators:
            await send_result_email(recipient, period_date, form_bytes, pdf_bytes, summary)
        print(f"Completed and emailed report for period {period_date} to {coordinators}")
    except Exception as exc:
        error_msg = f"Error processing payroll for period {period_date}:\n\n{exc}"
        print(error_msg)
        for recipient in coordinators:
            await send_error_email(recipient, period_date, str(exc))
