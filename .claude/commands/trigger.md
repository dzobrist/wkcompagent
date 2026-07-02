Trigger the WK Comp monthly payroll request email.

Current month (auto-detected):
```bash
curl -s -X POST https://wkcompagent-production.up.railway.app/trigger | python3 -m json.tool
```

Specific period (pass the month-end date as a query param):
```bash
curl -s -X POST "https://wkcompagent-production.up.railway.app/trigger?period=6/30/2026" | python3 -m json.tool
```

Report back the status and period returned.
