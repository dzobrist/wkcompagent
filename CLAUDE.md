# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working in this repository.

## What This Agent Does

Automated Workers' Compensation payroll reporting agent:
1. On the 20th of each month at 8:00 AM Mountain Time, emails the payroll coordinator requesting that month's payroll file
2. Receives the coordinator's reply (xlsx attachment) via SendGrid Inbound Parse webhook
3. Classifies payroll rows using Claude (claude-opus-4-5), aggregates dollar amounts in Python
4. Produces three reports: a WC Premium Cost Breakout PDF, a Broadmoor Comp Reimbursement PDF, and an Applied Underwriters carrier form (xlsx)
5. Emails both reports back to the coordinator

## Manual Testing

Use the `/trigger` slash command to fire the monthly request email without waiting for the scheduler:

```bash
curl -s -X POST https://wkcompagent-production.up.railway.app/trigger | python3 -m json.tool
```

Check server health and environment variable status:
```bash
curl https://wkcompagent-production.up.railway.app/status
```

Local development:
```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

## Architecture

### Email Flow
- **Outbound** (request + results): SendGrid API, FROM `wkcomp@resortoutfitters.com` (hardcoded — this domain is SendGrid-authenticated)
- **Inbound** (coordinator reply): SendGrid Inbound Parse → POST `/inbound` webhook
- **Reply-To**: `wkcomp@wkcomp.resortoutfitters.com` (MX record → SendGrid)
- **Period date** embedded in subject as `[period:M/D/YYYY]` tag — travels with coordinator's reply so the inbound handler knows which period to report on

### Processing Pipeline (`processor.py`)
1. `extract_payroll_rows()` in `classifier.py` reads the xlsx (9 earning slots per row, cols C–AL)
2. `load_mapping_rules()` in `classifier.py` reads `knowledge/Mapping_Rules.xlsx`
3. `_classify_with_claude()` — Claude classifies each row (class code, overrides, exceptions) but does NOT compute dollars
4. `_aggregate()` — Python applies earning rules and sums all dollar amounts deterministically
5. `_fill_form_template()` — writes to `knowledge/Form_Template.xlsx` (Applied Underwriters form)
6. `generate_pdf()` in `report_generator.py` — produces the WC Premium Cost Breakout PDF
7. `generate_broadmoor_pdf()` in `report_generator.py` — produces the Broadmoor Comp Reimbursement PDF (Soaring depts 15, 16, 17, 20 only)

**Critical design principle**: Claude classifies, Python does all arithmetic. This ensures output matches the Claude Project and is deterministic.

### Classification (`classifier.py`)
- `load_mapping_rules(path)` loads three sheets from `knowledge/Mapping_Rules.xlsx`:
  - **"Department to Class Code"**: col0=dept, col1=state, col2=class_code
  - **"Earning Type Rules"**: col0=type, col2=excluded_treatment (none/full/1/3/percentage)
  - **"Employee Overrides"**: col0=name, col1=dept, col2=class_code, col3=action (Reassign/Exclude), col4=reason
- Header rows are detected by content (not fixed row numbers) — robust against future spreadsheet changes
- Rows >120 chars in the key column are skipped (trailing notes/instructions in the sheet)

### CO9180 Soaring Subgroups (Internal PDF Only)
- CO9180 Soaring: depts 15 & 20
- CO8810 Soaring (yellow highlight): depts 16 & 17 — these wages appear on the CO8810 carrier form line but are shown separately in the PDF for cost-allocation visibility
- CO9180 Outfitter Guides: depts 9 & 21
- CO9180 Falconry: dept 18
- The grand total is calculated from the 11 carrier form lines only — the Soaring CO8810 sub-row is informational and not double-counted

### Report Generation (`report_generator.py`)
- `NET_RATES` dict — per-$100 rates from policy declarations; **update here when rates change**
- `generate_pdf(period_date, class_codes, co9180_subgroups, soaring_co8810, subgroup_detail)` — full employee-level detail if available, falls back to subgroup totals
- PDF colors: header=`#1a3a5c`, soaring CO8810 highlight=`#fffacd` (yellow)

### Email Service (`email_service.py`)
- Uses Claude Haiku (`claude-haiku-4-5`) for jokes and "today in history" facts (cheap/fast)
- Request email includes ASCII "Hello" banner + on-this-day history fact + HR/payroll joke
- Result and error emails include a joke at the top

### Scheduler (`main.py`)
- Cron: day 20, hour 8, `America/Denver` timezone
- Override via env vars: `TRIGGER_DAY`, `TRIGGER_HOUR`

## Environment Variables

Required:
- `ANTHROPIC_API_KEY`
- `SENDGRID_API_KEY`
- `FROM_EMAIL` (must match SendGrid-authenticated domain)
- `COORDINATOR_EMAIL` (comma-separated for multiple recipients)

Optional:
- `ADMIN_EMAIL` — receives error notifications
- `TRIGGER_DAY` (default: 20)
- `TRIGGER_HOUR` (default: 8)
- `CLAUDE_MODEL` (default: claude-opus-4-5)

## Knowledge Files

All in `knowledge/`:
- `Mapping_Rules.xlsx` — **primary config**: dept→class code, earning rules, employee overrides. Update this file to change classification behavior; no code changes needed.
- `Form_Template.xlsx` — Applied Underwriters carrier form template (rows 17–27 = class codes, row 59 = totals)
- `SOP_Monthly_Payroll_Reporting.md` — SOP provided as context to Claude during classification
- `WC_Premium_Breakout_Sample.pdf` — Sample report for reference

## Deployment

Hosted on Railway. Push to main branch triggers redeploy. After pushing changes, wait ~2 minutes for Railway to redeploy before re-triggering.

Production URL: `https://wkcompagent-production.up.railway.app`
