import calendar
import os
from contextlib import asynccontextmanager
from datetime import date

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI, Request
from fastapi.responses import JSONResponse

load_dotenv()

from email_service import send_request_email
from processor import handle_inbound

scheduler = AsyncIOScheduler()


def _current_month_end_str() -> str:
    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]
    return f"{today.month}/{last_day}/{today.year}"


async def trigger_monthly():
    period = _current_month_end_str()
    raw = os.getenv("COORDINATOR_EMAIL", "")
    coordinators = [e.strip() for e in raw.split(",") if e.strip()]
    if not coordinators:
        print("ERROR: COORDINATOR_EMAIL not set — cannot send request email")
        return
    print(f"Sending payroll request email for period ending {period} to {coordinators}")
    for coordinator in coordinators:
        await send_request_email(coordinator, period)


@asynccontextmanager
async def lifespan(app: FastAPI):
    trigger_day = int(os.getenv("TRIGGER_DAY", "20"))
    trigger_hour = int(os.getenv("TRIGGER_HOUR", "8"))
    scheduler.add_job(
        trigger_monthly,
        "cron",
        day=trigger_day,
        hour=trigger_hour,
        minute=0,
        timezone="America/Denver",
    )
    scheduler.start()
    print(f"Scheduler started: fires on day {trigger_day} of each month at {trigger_hour}:00 ET")
    yield
    scheduler.shutdown()


app = FastAPI(title="WK Comp Agent", lifespan=lifespan)


@app.get("/")
def health():
    return {"status": "ok", "service": "WK Comp Agent"}


@app.get("/status")
def status():
    """Report which environment variables are configured (values redacted)."""
    required = ["ANTHROPIC_API_KEY", "SENDGRID_API_KEY", "FROM_EMAIL", "COORDINATOR_EMAIL"]
    optional = ["ADMIN_EMAIL", "TRIGGER_DAY", "TRIGGER_HOUR", "CLAUDE_MODEL"]
    return {
        "required": {k: "set" if os.getenv(k) else "MISSING" for k in required},
        "optional": {k: os.getenv(k, "(not set)") if k not in ("ANTHROPIC_API_KEY", "SENDGRID_API_KEY") else ("set" if os.getenv(k) else "MISSING") for k in optional},
        "from_email_actual": os.getenv("FROM_EMAIL", "wkcomp@wkcomp.resortoutfitters.com"),
    }


@app.post("/trigger")
async def trigger():
    """Manually kick off the monthly payroll request email."""
    import traceback
    try:
        await trigger_monthly()
        return {"status": "triggered", "period": _current_month_end_str()}
    except Exception as exc:
        return JSONResponse(status_code=500, content={"error": str(exc), "detail": traceback.format_exc()})


@app.post("/inbound")
async def inbound(request: Request, background_tasks: BackgroundTasks):
    """
    SendGrid Inbound Parse webhook.
    Configure this URL in SendGrid: https://<your-railway-app>.up.railway.app/inbound
    """
    form_data = await request.form()

    # Extract fields while form is in scope
    payload = {
        "from_email": form_data.get("from", ""),
        "subject": form_data.get("subject", ""),
        "body": form_data.get("text", "") or form_data.get("html", ""),
        "num_attachments": int(form_data.get("attachments", 0)),
        "attachment_files": [],
    }

    for i in range(1, payload["num_attachments"] + 1):
        attachment = form_data.get(f"attachment{i}")
        if attachment and hasattr(attachment, "read"):
            content = await attachment.read()
            payload["attachment_files"].append(
                {"filename": attachment.filename or f"attachment{i}", "content": content}
            )

    background_tasks.add_task(handle_inbound, payload)
    return JSONResponse({"status": "received"})
