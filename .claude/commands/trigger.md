Trigger the WK Comp monthly payroll request email by running:

```bash
curl -s -X POST https://wkcompagent-production.up.railway.app/trigger | python3 -m json.tool
```

Report back the status and period returned.
