import base64
import os
from datetime import datetime

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    Email,
    FileContent,
    FileName,
    FileType,
    Mail,
    To,
)


def _sg_client() -> SendGridAPIClient:
    return SendGridAPIClient(os.getenv("SENDGRID_API_KEY"))


def _month_label(period_date: str) -> str:
    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        return dt.strftime("%B %Y")
    except Exception:
        return period_date


INBOUND_EMAIL = "wkcomp@wkcomp.resortoutfitters.com"
AUTHENTICATED_FROM = "wkcomp@resortoutfitters.com"  # SendGrid authenticated domain
FROM_EMAIL = os.getenv("FROM_EMAIL", AUTHENTICATED_FROM)


async def send_request_email(to_email: str, period_date: str):
    """Send the monthly payroll file request to the coordinator."""
    month_label = _month_label(period_date)
    from_email = FROM_EMAIL

    message = Mail()
    message.from_email = Email(from_email, "WK Comp Agent")
    message.reply_to = Email(INBOUND_EMAIL, "WK Comp Agent")
    message.to = To(to_email)
    message.subject = f"Workers Comp Payroll Report Request — {month_label}"
    message.plain_text_content = (
        f"Hi,\n\n"
        f"It's time to run the Workers Comp payroll report for {month_label}.\n\n"
        f"  1. Log into the payroll system\n"
        f"  2. Run the Earnings by Department report\n"
        f"  3. Set check dates: full month of {month_label}\n"
        f"  4. Export as Excel (.xlsx)\n"
        f"  5. Reply to this email with the file attached\n\n"
        f"Reporting period end date: {period_date}\n\n"
        f"Thank you"
    )

    _sg_client().send(message)
    print(f"Request email sent to {to_email} for period {period_date}")


async def send_result_email(
    to_email: str, period_date: str, form_bytes: bytes, pdf_bytes: bytes, summary: str
):
    """Send the WC Premium Breakout PDF and the Applied Underwriters form to the coordinator."""
    month_label = _month_label(period_date)
    from_email = FROM_EMAIL

    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        month_str = dt.strftime("%B_%Y")
    except Exception:
        month_str = "Report"

    body = (
        f"Workers Comp reports for {month_label} are attached.\n\n"
        f"  1. WC_Premium_Breakout_{month_str}.pdf — state-by-state breakout with estimated premiums\n"
        f"  2. Payroll_Report_{month_str}.xlsx — Applied Underwriters carrier submission form\n\n"
        f"Review the summary below before submitting the carrier form to Applied Underwriters.\n\n"
        f"{'=' * 60}\n\n"
        f"{summary}"
    )

    message = Mail(
        from_email=from_email,
        to_emails=to_email,
        subject=f"Workers Comp Reports Ready — {month_label}",
        plain_text_content=body,
    )

    # Attachment 1: PDF breakout report
    message.attachment = Attachment(
        file_content=FileContent(base64.b64encode(pdf_bytes).decode()),
        file_name=FileName(f"WC_Premium_Breakout_{month_str}.pdf"),
        file_type=FileType("application/pdf"),
        disposition=Disposition("attachment"),
    )

    # Attachment 2: Applied Underwriters xlsx form
    message.attachment = Attachment(
        file_content=FileContent(base64.b64encode(form_bytes).decode()),
        file_name=FileName(f"Payroll_Report_{month_str}.xlsx"),
        file_type=FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        disposition=Disposition("attachment"),
    )

    _sg_client().send(message)
    print(f"Result email with 2 attachments sent to {to_email} for period {period_date}")


async def send_error_email(to_email: str, period_date: str, error_detail: str):
    """Notify the coordinator that processing failed."""
    month_label = _month_label(period_date)
    from_email = FROM_EMAIL
    admin_email = os.getenv("ADMIN_EMAIL", from_email)

    body = (
        f"The Workers Comp agent could not process the payroll report for {month_label}.\n\n"
        f"Error details:\n{error_detail}\n\n"
        f"Please forward your payroll file to {admin_email} for manual processing."
    )

    for recipient in {to_email, admin_email}:
        message = Mail(
            from_email=from_email,
            to_emails=recipient,
            subject=f"WK Comp Agent Error — {month_label}",
            plain_text_content=body,
        )
        _sg_client().send(message)
