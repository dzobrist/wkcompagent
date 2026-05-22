# Standard Operating Procedure
## Monthly Payroll Reporting — Applied Underwriters Workers' Compensation

**Owner:** [Insert role/team name]
**Frequency:** Monthly, within 7 days of period end
**Estimated time per run:** 5 minutes
**Last updated:** May 2026

---

## Purpose

This SOP describes how to generate the monthly payroll report required by Applied Underwriters for Greenbrier Outfitters' workers' compensation coverage. The process uses a dedicated Claude Project to read the payroll system's "Earnings by Department" report, apply exclusion rules, aggregate wages by state and class code, populate the carrier's reporting form, and produce a premium cost breakout by state and class code.

## Who can run this

Any employee with access to:
- The payroll system (to export the "Earnings by Department" report)
- The Claude Project named **"WK Comp Monthly Payroll Reporting"**
- The shared folder where completed forms are filed

No spreadsheet, accounting, or technical expertise required.

---

## The monthly procedure

### Step 1 — Export payroll data

Log into the payroll system and run the **"Earnings by Department"** report for the prior calendar month.

- Set the check date range to cover the full reporting month (e.g., 4/1/2026 to 4/30/2026)
- Export as Excel (.xlsx)
- Save the file locally with a clear name: `Earning_by_Dept_April_2026.xlsx`

### Step 2 — Open the Claude Project

1. Go to claude.ai
2. Open the Project titled **"WK Comp Monthly Payroll Reporting"**
3. Click **New chat**

### Step 3 — Submit the request

Type a message like this and attach the payroll export from Step 1:

> Running April 2026 payroll. Reporting period end date: 4/30/2026.
> Payroll export attached.

### Step 4 — Review Claude's output

Claude will return **two downloadable files** plus a summary:

1. **The filled Applied Underwriters reporting form** — `Payroll_Report_[Month]_[Year].xlsx`
2. **The WC premium cost breakout** — `WC_Premium_Breakout_[Month]_[Year].pdf`
   - Breaks out compensable payroll and estimated premium by state and class code
   - Colorado CO9180 is further split into: Soaring Guides (Depts 15 & 20), Outfitter Guides (Depts 9 & 21), and Falconry - Broadmoor (Dept 18)
   - Estimated premiums use the net rates from the policy declarations page and are for internal reference; final amounts are subject to audit
3. **A text summary** showing total Gross, Excluded, and Compensable payroll plus a per-class-code breakdown
4. **An overrides applied list** — any employees whose wages were reclassified or excluded per the Employee Overrides tab
5. **An exceptions list** — any rows that did not match a mapping rule and were not included

**Always review the exceptions and overrides list before submitting.** Common legitimate exceptions: a new department added since last month, a new earning type used for the first time. Common problems: an entire department got dropped, totals are dramatically different from prior month, an override was applied to the wrong person because of a name typo, or a department-specific override didn't fire because the department text in the rules tab doesn't exactly match the source. If anything looks wrong, escalate to [Insert escalation contact] before submitting.

### Step 5 — File and submit

1. Download both files
2. Save the carrier form to: `[Insert shared drive path]\Payroll Reports\[YEAR]\[MONTH]_Payroll_Report.xlsx`
3. Save the premium breakout to: `[Insert shared drive path]\Payroll Reports\[YEAR]\[MONTH]_WC_Premium_Breakout.pdf`
4. Submit the carrier form to Applied Underwriters per current instructions: [Insert submission method]
5. Confirm the submission in [Insert tracking system or log]

---

## How exceptions, exclusions, and overrides work

The reported "Compensable Payroll" is **Gross Payroll minus Excluded Payroll**. Per current rules in the Mapping Rules file:

- **Compensable (counted in full):** Regular wages, Paid Time Off, Paid Sick Leave (HFWA)
- **Excluded in full:** Service Charges (tips), Misc pay
- **Partially excluded:** Nonqualified Overtime — only the premium portion (1/3 of the amount, representing the "half" above straight time in time-and-a-half) is excluded

Any earning type Claude doesn't recognize gets reported as an exception rather than guessed at. Same for any department that isn't in the Department → Class Code lookup.

### Employee-level overrides

The **Employee Overrides** tab in `Mapping_Rules.xlsx` supports two actions that take precedence over the department lookup:

- **Reassign** — use a different class code than the employee's department default. Used when an employee's role doesn't match their payroll department (e.g., an outside salesperson posted to a guide department).
- **Exclude** — drop all of the employee's wages from the report. Used for officers and owners with valid workers' comp exclusion elections on file with the carrier.

Each row has an **optional Department column**. This determines how broadly the rule applies:

- **Department blank** — the rule applies to *every* department this employee appears in. Use this when an employee should always be classified the same way regardless of how payroll posted them.
- **Department specified** — the rule applies *only* to earnings posted under that specific department. Use this when an employee splits time between departments and only some of their earnings should be reclassified or excluded.

Lookup precedence for each detail row:

1. Employee Overrides — match on exact Employee Name. If multiple override rows match the same name:
   - A row with a matching Department wins over a row with blank Department.
   - If two rows tie, the first one in the tab wins.
2. If override matches → use that rule (Reassign uses the override's Class Code; Exclude drops the row).
3. Otherwise → Department → Class Code lookup.
4. If neither matches → exception.

**Exclude entries require a documented carrier exclusion election.** Excluding wages that should be reported creates premium audit exposure. Confirm with Ashley Cazire (producer) before adding an Exclude row.

**Department-specific overrides require exact text match.** "1 - OFFICE" does not match "1 - Office" or "OFFICE" or "1 -OFFICE". Copy the department text directly from the payroll export to be safe.

---

## Troubleshooting

**Claude says it can't find a column it expected.** The payroll report layout may have changed. Compare the new export's row 7 headers to the layout described in the Mapping Rules → "Source File Structure" tab. Notify [Insert escalation contact] so the Project can be updated.

**A new department appears in this month's report.** Add a row to the Mapping Rules workbook → "Department to Class Code" tab with the correct state and class code, save, and re-upload to Project knowledge. Then re-run.

**A new earning type appears.** Same process — add a row to "Earning Type Rules" tab.

**An employee should be classified differently than their department default — across all departments.** Add a row to "Employee Overrides" with Action=Reassign, the correct class code, and Department left blank.

**An employee should be classified differently — but only for one of their departments.** Add a row to "Employee Overrides" with Action=Reassign, the target class code, and the exact source department text in the Department column. Their other departments will continue to follow the default class code mapping.

**An officer or owner needs to be excluded from the report.** Confirm with Ashley Cazire that an exclusion election is on file with Applied Underwriters. Then add a row to "Employee Overrides" with Action=Exclude, Class Code blank, and Department blank (excludes all their earnings). Document the election reference in the Reason column.

**An employee changes name in payroll (marriage, correction, etc.).** Update the matching row(s) in "Employee Overrides" to the new exact name. Name match is case- and punctuation-sensitive; a stale entry will silently stop applying.

**An override doesn't seem to be applying.** Most common cause: name or department mismatch. Open the source file and check the exact text in column A (name) and column B (department) for the relevant row, then copy/paste into the override tab.

**Totals look dramatically different from last month.** Compare against the prior month's filed report. Large swings can be legitimate (seasonal hiring) or signal a problem (duplicate file, wrong period, missed exclusion). Verify before submitting.

**The premium breakout rates look wrong.** The PDF uses net rates hardcoded from the policy declarations page. If rates change mid-policy (endorsement, audit adjustment), the Project system prompt must be updated with the new rates. Notify [Insert escalation contact].

**Claude returns an error or the output looks malformed.** Start a new chat in the same Project and try again. If the issue persists, escalate.

---

## Maintenance — for the SOP owner

This section is for the person responsible for keeping the Project current.

**When to update Mapping Rules:**
- New department appears in payroll → Tab 3 (Department to Class Code)
- New earning type appears in payroll → Tab 4 (Earning Type Rules)
- Carrier changes definition of compensable vs. excluded payroll → Tab 4
- Policy adds or removes a class code → Tab 5 (Class Codes Reference) AND form template
- Employee role changes such that their class code should differ from their department's default → Tab 6 (Employee Overrides) with Action=Reassign
- Employee splits time across departments and only some earnings should be reclassified → Tab 6 with Action=Reassign and a specific Department value
- Officer or owner files a workers' comp exclusion election with the carrier → Tab 6 with Action=Exclude
- An override no longer applies (employee left, election rescinded, role changed back, department changed in payroll) → delete or edit the row in Tab 6
- A best-guess mapping turns out to be wrong (correct it as soon as you discover the error)

**When to update the Project system prompt:**
- The reporting form template changes
- The summary output format requirements change
- The escalation contact changes
- The override action types or precedence rules change
- The Employee Overrides schema changes (e.g., adding new columns)
- **Policy rates change** (endorsement, renewal, audit adjustment) → update the rates used in the premium breakout PDF

**Annual review:** Verify policy period dates, producer name, named insured, the full set of class codes on the form against your current policy declarations, **the net rates used in the premium breakout PDF**, and every Employee Overrides row (especially Exclude rows — confirm each election is still active with the carrier; and department-specific Reassign rows — confirm the source department text still matches what payroll exports).

---

## Reference files (in the Claude Project)

| File | Purpose |
|------|---------|
| `Applied_Underwriters_Form_Template.xls` | Blank reporting form Claude fills each month |
| `Mapping_Rules.xlsx` | Department → class code, earning type rules, and employee overrides |
| `Sample_Completed_Form.xlsx` | Worked example (created after first successful run) |
| `SOP_Monthly_Payroll_Reporting.md` | This document |

---

## Contacts

| Role | Name | Email |
|------|------|-------|
| Process owner | [Insert] | [Insert] |
| Escalation / questions | [Insert] | [Insert] |
| Applied Underwriters producer | Ashley Cazire | [Insert] |
| Backup runner | [Insert] | [Insert] |
