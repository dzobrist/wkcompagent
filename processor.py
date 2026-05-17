import io
import os
import re
from datetime import date, timedelta

import anthropic
import openpyxl

from email_service import send_error_email, send_result_email

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

SYSTEM_PROMPT = """You are a payroll reporting assistant for Greenbrier Outfitters and affiliated properties. Your only job is to take a monthly payroll export and produce a completed Applied Underwriters Workers' Compensation Monthly Payroll Report.

## Context you should always assume

- The user is a non-technical employee running this monthly process. Keep your language plain and direct.
- The reporting carrier is Applied Underwriters. The named insured is Greenbrier Outfitters. The producer is Ashley Cazire. The policy number is 37-664725-01-01; -01-02. The policy period is 04/15/26–04/15/27.
- The classification rules are provided in the MAPPING RULES section of this message.

## Source file format (Earnings by Department)

The payroll export has a specific, slightly awkward structure you must handle correctly:

- Rows 1–6: report header metadata (company, date range). Skip.
- Row 7: column headers
- Rows 8 through approximately 85: **employee detail rows — process these**
- Rows ~86 onward: "Company Totals" and "Department Totals" summary sections — **IGNORE these**.

**Stop reading detail rows when** you hit two consecutive blank rows in column A, OR when column A contains "Company Totals" / "Department Totals", whichever comes first.

**Each detail row has up to 8 earning entries** spread across the columns:
- Column A: Employee Name
- Column B: Department
- Columns C–F: Earning 1 (Type, Hours, Rate, Amount)
- Columns G–J: Earning 2 (Type, Hours, Rate, Amount)
- Columns K–N: Earning 3
- Columns O–R: Earning 4
- Columns S–V: Earning 5
- Columns W–Z: Earning 6
- Columns AA–AD: Earning 7
- Columns AE–AH: Earning 8

You **must iterate all 8 earning slots** for every employee row.

## Mapping_Rules structure

The mapping rules are provided below in four sections. Use them exactly as described.

- **Tab 3 — Department to Class Code:** Department name → Class Code lookup. Default for all employees.
- **Tab 4 — Earning Type Rules:** Earning type → gross treatment + excluded treatment. Defines compensable vs. excluded.
- **Tab 5 — Class Codes Reference:** The 11 class codes that appear on the form, in form order.
- **Tab 6 — Employee Overrides:** Per-employee rules that take precedence over Tab 3.

## Your process — every time

**Step 1.** Read all mapping rules provided.

**Step 2.** Read the payroll data provided. Verify the column layout matches expected structure. If anything has changed, stop and include the issue in your exceptions.

**Step 3.** For each detail row, determine its disposition:
1. **Check Employee Overrides (Tab 6) first.** Match on exact Employee Name (case, punctuation, spacing).
   - Among name matches, prefer a row where Department also matches exactly. If found, that row wins.
   - If no department-specific row matches, use the first row with a blank Department field.
   - If Action = Exclude → skip the entire source row.
   - If Action = Reassign → use the Class Code from the override row.
2. **Otherwise, look up the Department in Tab 3** to get the Class Code.
3. If neither matches, the row is an exception.

**Step 4.** Apply earning rules for each non-empty earning slot in non-excluded rows:
- **Regular, Paid Time Off, Paid Sick Leave (HFWA):** Gross += Amount. Excluded += 0.
- **Service Chgs:** Gross += Amount. Excluded += Amount.
- **Misc pay:** Gross += Amount. Excluded += Amount.
- **Nonqualified Overtime:** Gross += Amount. Excluded += (Amount × 1/3).
- Always cross-check against Tab 4 — it overrides these defaults if different.

**Step 5.** Aggregate by Class Code. For each of the 11 class codes:
- Gross Payroll = sum of all gross contributions mapped to that code
- Excluded Payroll = sum of all excluded contributions mapped to that code
- Compensable Payroll = Gross − Excluded

**Step 6.** Call the submit_payroll_report tool with your results. Include all 11 class codes in the class_codes array (use 0.0 for codes with no payroll). Include a full summary string matching the format below:

```
PAYROLL REPORT SUMMARY — [Month Year]
Reporting period end: [date]

Source file: payroll export
Detail rows processed: [count]
Earning entries processed: [count]
Overrides applied: [count] ([N] Reassign, [N] Exclude)
Exceptions: [count]

TOTALS:
Gross Payroll       $[amount]
Excluded Payroll    $[amount]
Compensable Payroll $[amount]

BY CLASS CODE:
[code]  Gross $[amt]  Excluded $[amt]  Compensable $[amt]
... (all 11)

OVERRIDES APPLIED:
[For each: employee name, department, action, class code if Reassign, reason]

EXCEPTIONS (not included in totals):
[For each: row number, employee name, department, earning type, amount, reason]

DATA QUALITY NOTES:
[Anything unusual. Omit section if none.]

VARIANCE CHECK:
No prior month context available.
```

## Hard rules — never break these

1. Never invent payroll data. If a field is missing or unclear, report it as an exception.
2. Never silently drop a row or earning entry.
3. Never include summary rows ("Company Totals", "Department Totals") in calculations.
4. Always iterate all 8 earning slots per employee row.
5. Override precedence is fixed. Employee Overrides always beat the department lookup.
6. Exclude overrides drop the whole source row, not just some earnings.
7. Name and Department matching is exact.
8. Row 59 totals must equal the per-class-code sums — verify before calling the tool.
9. Never apply a partial override.
"""


def _extract_payroll_rows(ws) -> list[dict]:
    rows = []
    consecutive_blanks = 0

    for row_idx in range(8, ws.max_row + 1):
        employee_name = ws.cell(row=row_idx, column=1).value

        if employee_name is None:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                break
            continue

        employee_name = str(employee_name).strip()
        if employee_name in ("Company Totals", "Department Totals"):
            break

        consecutive_blanks = 0
        department = str(ws.cell(row=row_idx, column=2).value or "").strip()

        earnings = []
        for slot in range(8):
            col = 3 + (slot * 4)  # C=3, G=7, K=11, O=15, S=19, W=23, AA=27, AE=31
            earn_type = ws.cell(row=row_idx, column=col).value
            if earn_type:
                hours = ws.cell(row=row_idx, column=col + 1).value
                rate = ws.cell(row=row_idx, column=col + 2).value
                amount = ws.cell(row=row_idx, column=col + 3).value
                if amount is not None:
                    earnings.append(
                        {"type": str(earn_type).strip(), "hours": hours, "rate": rate, "amount": float(amount)}
                    )

        rows.append({"row": row_idx, "employee": employee_name, "department": department, "earnings": earnings})

    return rows


def _payroll_rows_to_text(rows: list[dict]) -> str:
    lines = [f"PAYROLL DATA — {len(rows)} employee detail rows\n"]
    for r in rows:
        parts = []
        for e in r["earnings"]:
            parts.append(f"{e['type']}: {e['hours']}h @ ${e['rate']} = ${e['amount']:.2f}")
        earnings_str = " | ".join(parts) if parts else "(no earnings)"
        lines.append(f"Row {r['row']}: {r['employee']} | Dept: {r['department']} | {earnings_str}")
    return "\n".join(lines)


def _read_mapping_rules() -> str:
    wb = openpyxl.load_workbook(MAPPING_RULES_PATH)
    sections = []

    tab_labels = {
        2: "TAB 3 — Department to Class Code",
        3: "TAB 4 — Earning Type Rules",
        4: "TAB 5 — Class Codes Reference",
        5: "TAB 6 — Employee Overrides",
    }

    for tab_idx, label in tab_labels.items():
        if tab_idx >= len(wb.worksheets):
            continue
        ws = wb.worksheets[tab_idx]
        rows_text = []
        for row in ws.iter_rows(min_row=1, values_only=True):
            if any(cell is not None for cell in row):
                rows_text.append(" | ".join(str(c) if c is not None else "" for c in row))
        sections.append(f"{label}:\n" + "\n".join(rows_text))

    return "\n\n".join(sections)


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


async def process_payroll(xlsx_bytes: bytes, period_date: str) -> tuple[bytes, str]:
    """Run the payroll xlsx through Claude and return (filled_form_bytes, summary_text)."""
    client = anthropic.AsyncAnthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    if "Earnings by Department" not in wb.sheetnames:
        raise ValueError("Sheet 'Earnings by Department' not found in the uploaded file.")

    ws = wb["Earnings by Department"]
    rows = _extract_payroll_rows(ws)
    payroll_text = _payroll_rows_to_text(rows)
    mapping_text = _read_mapping_rules()

    tools = [
        {
            "name": "submit_payroll_report",
            "description": "Submit the final calculated payroll report with class code breakdowns.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "period_end": {"type": "string", "description": "Reporting period end date"},
                    "class_codes": {
                        "type": "array",
                        "description": "All 11 class codes with their calculated payroll figures",
                        "items": {
                            "type": "object",
                            "properties": {
                                "code": {"type": "string"},
                                "gross": {"type": "number"},
                                "excluded": {"type": "number"},
                                "compensable": {"type": "number"},
                            },
                            "required": ["code", "gross", "excluded", "compensable"],
                        },
                    },
                    "overrides_applied": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Human-readable list of each override that fired",
                    },
                    "exceptions": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Rows/earning entries not included in totals, with reasons",
                    },
                    "data_quality_notes": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "summary": {
                        "type": "string",
                        "description": "Full summary text as specified in the system prompt",
                    },
                },
                "required": ["period_end", "class_codes", "summary"],
            },
        }
    ]

    response = await client.messages.create(
        model=os.getenv("CLAUDE_MODEL", "claude-opus-4-7"),
        max_tokens=8096,
        system=[
            {
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        tools=tools,
        tool_choice={"type": "any"},
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": f"MAPPING RULES\n\n{mapping_text}",
                        "cache_control": {"type": "ephemeral"},
                    },
                    {
                        "type": "text",
                        "text": f"Reporting period end: {period_date}\n\n{payroll_text}",
                    },
                ],
            }
        ],
    )

    tool_result = None
    for block in response.content:
        if block.type == "tool_use" and block.name == "submit_payroll_report":
            tool_result = block.input
            break

    if not tool_result:
        raise ValueError("Claude did not return a payroll report. Check API logs.")

    form_bytes = _fill_form_template(tool_result["class_codes"], period_date)
    return form_bytes, tool_result["summary"]


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
        form_bytes, summary = await process_payroll(xlsx_bytes, period_date)
        for recipient in coordinators:
            await send_result_email(recipient, period_date, form_bytes, summary)
        print(f"Completed and emailed report for period {period_date} to {coordinators}")
    except Exception as exc:
        error_msg = f"Error processing payroll for period {period_date}:\n\n{exc}"
        print(error_msg)
        for recipient in coordinators:
            await send_error_email(recipient, period_date, str(exc))
