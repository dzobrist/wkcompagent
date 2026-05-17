"""
Deterministic payroll classification engine.
All arithmetic happens here in Python — Claude is not used for calculations.
"""

import re
from typing import Optional
import openpyxl

CLASS_CODE_ORDER = [
    "CO8810", "CO8869", "CO9180",
    "FL8742", "FL9180",
    "GA9180",
    "VA8742", "VA9180",
    "WV8742", "WV8810", "WV9180",
]

CO9180_SUBGROUP_DEPTS = {
    "Soaring Guides — Depts 15 & 20": {15, 20},
    "Outfitter Guides — Depts 9 & 21": {9, 21},
    "Falconry - Broadmoor — Dept 18": {18},
}


def _dept_num(dept: str) -> Optional[int]:
    m = re.match(r'^\s*(\d+)', dept)
    try:
        return int(m.group(1)) if m else None
    except (ValueError, AttributeError):
        return None


def _co9180_subgroup(dept: str) -> Optional[str]:
    num = _dept_num(dept)
    if num is None:
        return None
    for name, nums in CO9180_SUBGROUP_DEPTS.items():
        if num in nums:
            return name
    return None


def load_mapping_rules(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)

    # Tab 3 (index 2): Department → Class Code
    dept_to_code: dict[str, str] = {}
    ws3 = wb.worksheets[2]
    for row in ws3.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            dept_to_code[str(row[0]).strip()] = str(row[1]).strip()

    # Tab 4 (index 3): Earning Type → excluded treatment
    # Expected columns: Earning Type | (Gross col) | Excluded Treatment
    # Gross is always full for included entries; we only need excluded logic.
    earning_rules: dict[str, str] = {}  # earn_type → excluded_mode
    ws4 = wb.worksheets[3]
    headers = [str(c).strip().lower() if c else "" for c in next(ws4.iter_rows(min_row=1, max_row=1, values_only=True))]
    try:
        excl_col = next(i for i, h in enumerate(headers) if "excl" in h or "excluded" in h)
    except StopIteration:
        excl_col = 2  # fallback: third column

    for row in ws4.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        earn_type = str(row[0]).strip()
        excl_val = str(row[excl_col]).strip().lower() if row[excl_col] is not None else "none"
        earning_rules[earn_type] = excl_val

    # Tab 6 (index 5): Employee Overrides
    overrides: list[dict] = []
    ws6 = wb.worksheets[5]
    for row in ws6.iter_rows(min_row=2, values_only=True):
        if not row[0] or not row[3]:
            continue
        dept_raw = str(row[1]).strip() if row[1] else None
        # Skip rows where department is a placeholder (brackets, "n/a", etc.)
        if dept_raw and re.match(r'^\[.*\]$|^n/?a$|^tbd$|^placeholder$', dept_raw, re.IGNORECASE):
            dept_raw = None
        overrides.append({
            "name": str(row[0]).strip(),
            "dept": dept_raw if dept_raw else None,
            "class_code": str(row[2]).strip() if row[2] else None,
            "action": str(row[3]).strip(),
            "reason": str(row[4]).strip() if row[4] else "",
        })

    return {
        "dept_to_code": dept_to_code,
        "earning_rules": earning_rules,
        "overrides": overrides,
    }


def _excluded_amount(amount: float, mode: str) -> float:
    mode = mode.lower().strip()
    if not mode or mode in ("none", "0", "zero", "$0", "no", "false"):
        return 0.0
    if mode in ("full", "all", "yes", "100%", "100", "true"):
        return amount
    if "third" in mode or "1/3" in mode or "33" in mode or "premium" in mode:
        return round(amount / 3, 2)
    try:
        pct = float(mode.strip("%")) / 100
        return round(amount * pct, 2)
    except ValueError:
        return 0.0


def _find_override(name: str, dept: str, overrides: list) -> Optional[dict]:
    name_matches = [o for o in overrides if o["name"] == name]
    if not name_matches:
        return None
    # Prefer exact dept match
    dept_match = next((o for o in name_matches if o["dept"] and o["dept"] == dept), None)
    if dept_match:
        return dept_match
    # Fall back to blank-dept (applies to all depts)
    return next((o for o in name_matches if not o["dept"]), None)


def extract_payroll_rows(ws) -> list[dict]:
    rows = []
    consecutive_blanks = 0
    for row_idx in range(8, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=1).value
        if val is None:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                break
            continue
        name = str(val).strip()
        if name in ("Company Totals", "Department Totals"):
            break
        consecutive_blanks = 0
        dept = str(ws.cell(row=row_idx, column=2).value or "").strip()
        earnings = []
        for slot in range(8):
            col = 3 + slot * 4
            earn_type = ws.cell(row=row_idx, column=col).value
            if earn_type:
                amount = ws.cell(row=row_idx, column=col + 3).value
                if amount is not None:
                    try:
                        earnings.append({"type": str(earn_type).strip(), "amount": float(amount)})
                    except (ValueError, TypeError):
                        pass
        rows.append({"row": row_idx, "employee": name, "department": dept, "earnings": earnings})
    return rows


def classify_payroll(payroll_ws, mapping_rules: dict) -> dict:
    dept_to_code = mapping_rules["dept_to_code"]
    earning_rules = mapping_rules["earning_rules"]
    overrides = mapping_rules["overrides"]

    rows = extract_payroll_rows(payroll_ws)

    totals = {code: {"gross": 0.0, "excluded": 0.0} for code in CLASS_CODE_ORDER}
    co9180_subs = {name: {"gross": 0.0, "excluded": 0.0} for name in CO9180_SUBGROUP_DEPTS}

    overrides_applied = []
    exceptions = []
    dq_notes = []
    earning_count = 0
    override_names_applied = set()

    for row in rows:
        name = row["employee"]
        dept = row["department"]

        override = _find_override(name, dept, overrides)

        if override:
            action = override["action"].strip().lower()
            if action == "exclude":
                overrides_applied.append(
                    f"{name} / {dept} / Exclude — {override['reason']}"
                )
                continue
            elif action == "reassign":
                class_code = override["class_code"]
                overrides_applied.append(
                    f"{name} / {dept} / Reassign → {class_code} — {override['reason']}"
                )
            else:
                class_code = dept_to_code.get(dept)
        else:
            class_code = dept_to_code.get(dept)

        if class_code is None:
            exceptions.append(
                f"Row {row['row']}: {name} / dept '{dept}' — no class code mapping; row excluded"
            )
            continue

        if class_code not in totals:
            exceptions.append(
                f"Row {row['row']}: {name} / unknown class code '{class_code}'"
            )
            continue

        for earning in row["earnings"]:
            earning_count += 1
            earn_type = earning["type"]
            amount = earning["amount"]

            if earn_type not in earning_rules:
                exceptions.append(
                    f"Row {row['row']}: {name} — unrecognized earning type '{earn_type}' "
                    f"(${amount:,.2f}); excluded from totals"
                )
                continue

            excl = _excluded_amount(amount, earning_rules[earn_type])
            totals[class_code]["gross"] = round(totals[class_code]["gross"] + amount, 2)
            totals[class_code]["excluded"] = round(totals[class_code]["excluded"] + excl, 2)

            if class_code == "CO9180":
                sg = _co9180_subgroup(dept)
                if sg and sg in co9180_subs:
                    co9180_subs[sg]["gross"] = round(co9180_subs[sg]["gross"] + amount, 2)
                    co9180_subs[sg]["excluded"] = round(co9180_subs[sg]["excluded"] + excl, 2)

    # Build output
    class_codes = []
    for code in CLASS_CODE_ORDER:
        g = round(totals[code]["gross"], 2)
        ex = round(totals[code]["excluded"], 2)
        class_codes.append({"code": code, "gross": g, "excluded": ex, "compensable": round(g - ex, 2)})

    co9180_subgroups = []
    for sg_name in CO9180_SUBGROUP_DEPTS:
        g = round(co9180_subs[sg_name]["gross"], 2)
        ex = round(co9180_subs[sg_name]["excluded"], 2)
        co9180_subgroups.append({"description": sg_name, "gross": g, "excluded": ex, "compensable": round(g - ex, 2)})

    total_gross = round(sum(c["gross"] for c in class_codes), 2)
    total_excl = round(sum(c["excluded"] for c in class_codes), 2)

    return {
        "rows_processed": len(rows),
        "earning_entries": earning_count,
        "class_codes": class_codes,
        "co9180_subgroups": co9180_subgroups,
        "totals": {"gross": total_gross, "excluded": total_excl, "compensable": round(total_gross - total_excl, 2)},
        "overrides_applied": overrides_applied,
        "exceptions": exceptions,
        "dq_notes": dq_notes,
    }
