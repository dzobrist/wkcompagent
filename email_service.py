import base64
import os
from datetime import datetime

import anthropic
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    Email,
    FileContent,
    FileName,
    FileType,
    Mail,
)

# Must be the SendGrid-authenticated domain
AUTHENTICATED_FROM = "wkcomp@resortoutfitters.com"
# Replies route here via the MX record → SendGrid inbound parse
INBOUND_EMAIL = "wkcomp@wkcomp.resortoutfitters.com"


def _sg_client() -> SendGridAPIClient:
    return SendGridAPIClient(os.getenv("SENDGRID_API_KEY"))


def _claude(prompt: str, max_tokens: int = 150) -> str:
    """Call Claude Haiku with a prompt and return the text response."""
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    response = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def _get_joke() -> str:
    """Ask Claude for a fresh payroll or HR joke."""
    try:
        return _claude(
            "Give me one short, funny payroll or HR joke — a single punchline, "
            "no setup explanation, no commentary, just the joke itself. "
            "Keep it clean and workplace-appropriate."
        )
    except Exception:
        return "Why did the payroll clerk quit? They just couldn't make ends meet."


def _get_today_in_history() -> str:
    """Ask Claude for a funny or surprising thing that happened on today's date in history."""
    try:
        today = datetime.now().strftime("%B %d")
        return _claude(
            f"What is one genuinely funny, strange, or surprising thing that happened on {today} "
            f"in history? Give me just the fact itself — one or two sentences, no intro like "
            f"'On this day' or 'Here is a fact'. Make it entertaining. Keep it workplace-appropriate.",
            max_tokens=120,
        )
    except Exception:
        return "History took the day off — probably a payroll error."


def _month_label(period_date: str) -> str:
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        return dt.strftime("%B %Y")
    except Exception:
        return period_date


async def send_request_email(to_email: str, period_date: str):
    """Send the monthly payroll file request to the coordinator."""
    month_label = _month_label(period_date)
    joke = _get_joke()
    history = _get_today_in_history()
    today_label = datetime.now().strftime("%B %d")

    message = Mail(
        from_email=AUTHENTICATED_FROM,
        to_emails=to_email,
        subject=f"Workers Comp Payroll Report Request — {month_label} [period:{period_date}]",
        plain_text_content=(
            f"******************************************\n"
            f"*                                        *\n"
            f"*         H E L L O  !  👋               *\n"
            f"*                                        *\n"
            f"******************************************\n\n"
            f"😄 Joke of the day:\n{joke}\n\n"
            f"📅 On this day ({today_label}) in history:\n{history}\n\n"
            f"---\n\n"
            f"It's time to run the Workers Comp payroll report for {month_label}.\n\n"
            f"  1. Log into the payroll system\n"
            f"  2. Run the Earnings by Department report\n"
            f"  3. Set check dates: full month of {month_label}\n"
            f"  4. Export as Excel (.xlsx)\n"
            f"  5. Reply to this email with the file attached\n\n"
            f"Thank you"
        ),
    )
    # Set Reply-To directly in the request body (SDK property causes 400)
    sg = _sg_client()
    body = message.get()
    body["reply_to"] = {"email": INBOUND_EMAIL}
    sg.client.mail.send.post(request_body=body)
    print(f"Request email sent to {to_email} for period {period_date}")


async def send_result_email(
    to_email: str, period_date: str,
    form_bytes: bytes, pdf_bytes: bytes, broadmoor_bytes: bytes, summary: str
):
    """Send all three WC reports to the coordinator."""
    month_label = _month_label(period_date)
    joke = _get_joke()

    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        month_str = dt.strftime("%B_%Y")
    except Exception:
        month_str = "Report"

    body = (
        f"😄 {joke}\n\n"
        f"---\n\n"
        f"Workers Comp reports for {month_label} are attached.\n\n"
        f"  1. WC_Premium_Breakout_{month_str}.pdf — state-by-state breakout with estimated premiums\n"
        f"  2. Broadmoor_Comp_Reimbursement_{month_str}.pdf — Soaring Adventure internal cost allocation\n"
        f"  3. Payroll_Report_{month_str}.xlsx — Applied Underwriters carrier submission form\n\n"
        f"Review the summary below before submitting the carrier form to Applied Underwriters. "
        f"After the reports are verified, please submit to ecprdata@auw.com\n\n"
        f"{'=' * 60}\n\n"
        f"{summary}"
    )

    message = Mail(
        from_email=AUTHENTICATED_FROM,
        to_emails=to_email,
        subject=f"Workers Comp Reports Ready — {month_label}",
        plain_text_content=body,
    )

    # Attachment 1: WC Premium Cost Breakout PDF
    message.attachment = Attachment(
        file_content=FileContent(base64.b64encode(pdf_bytes).decode()),
        file_name=FileName(f"WC_Premium_Breakout_{month_str}.pdf"),
        file_type=FileType("application/pdf"),
        disposition=Disposition("attachment"),
    )

    # Attachment 2: Broadmoor Comp Reimbursement PDF
    message.attachment = Attachment(
        file_content=FileContent(base64.b64encode(broadmoor_bytes).decode()),
        file_name=FileName(f"Broadmoor_Comp_Reimbursement_{month_str}.pdf"),
        file_type=FileType("application/pdf"),
        disposition=Disposition("attachment"),
    )

    # Attachment 3: Applied Underwriters xlsx form
    message.attachment = Attachment(
        file_content=FileContent(base64.b64encode(form_bytes).decode()),
        file_name=FileName(f"Payroll_Report_{month_str}.xlsx"),
        file_type=FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        disposition=Disposition("attachment"),
    )

    _sg_client().send(message)
    print(f"Result email with 3 attachments sent to {to_email} for period {period_date}")


async def send_error_email(to_email: str, period_date: str, error_detail: str):
    """Notify the coordinator that processing failed."""
    month_label = _month_label(period_date)
    admin_email = os.getenv("ADMIN_EMAIL", AUTHENTICATED_FROM)
    joke = _get_joke()

    body = (
        f"😄 {joke}\n\n"
        f"---\n\n"
        f"The Workers Comp agent could not process the payroll report for {month_label}.\n\n"
        f"Error details:\n{error_detail}\n\n"
        f"Please forward your payroll file to {admin_email} for manual processing."
    )

    for recipient in {to_email, admin_email}:
        message = Mail(
            from_email=AUTHENTICATED_FROM,
            to_emails=recipient,
            subject=f"WK Comp Agent Error — {month_label}",
            plain_text_content=body,
        )
        _sg_client().send(message)
