import base64
import os
from datetime import datetime

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    FileContent,
    FileName,
    FileType,
    Header,
    Mail,
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


async def send_request_email(to_email: str, period_date: str):
    """Send the monthly payroll file request to the coordinator."""
    month_label = _month_label(period_date)
    from_email = os.getenv("FROM_EMAIL", "wkcomp@resortoutfitters.com")

    message = Mail(
        from_email=from_email,
        to_emails=to_email,
        subject=f"Workers Comp Payroll Report Request — {month_label}",
        plain_text_content=(
            f"Hi,\n\n"
            f"Please reply to this email with the Earnings by Department payroll export for {month_label}.\n\n"
            f"  - Log into the payroll system\n"
            f"  - Run the Earnings by Department report\n"
            f"  - Set check dates: full month of {month_label}\n"
            f"  - Export as Excel (.xlsx)\n"
            f"  - Reply to this email with the file attached\n\n"
            f"Reporting period end date: {period_date}\n\n"
            f"Thank you"
        ),
    )
    # Reply-To routes coordinator replies to the SendGrid inbound parse address
    message.header = Header("Reply-To", INBOUND_EMAIL)

    _sg_client().send(message)
    print(f"Request email sent to {to_email} for period {period_date}")


async def send_result_email(to_email: str, period_date: str, form_bytes: bytes, summary: str):
    """Send the completed Applied Underwriters form back to the coordinator."""
    month_label = _month_label(period_date)
    from_email = os.getenv("FROM_EMAIL", "wkcomp@resortoutfitters.com")

    try:
        dt = datetime.strptime(period_date, "%m/%d/%Y")
        file_name = f"Payroll_Report_{dt.strftime('%B_%Y')}.xlsx"
    except Exception:
        file_name = "Payroll_Report.xlsx"

    encoded = base64.b64encode(form_bytes).decode()
    attachment = Attachment(
        file_content=FileContent(encoded),
        file_name=FileName(file_name),
        file_type=FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        disposition=Disposition("attachment"),
    )

    body = (
        f"Workers Comp Payroll Report for {month_label} is attached.\n\n"
        f"Review the exceptions and overrides list before submitting to Applied Underwriters.\n\n"
        f"{'=' * 60}\n\n"
        f"{summary}"
    )

    message = Mail(
        from_email=from_email,
        to_emails=to_email,
        subject=f"Workers Comp Report Ready — {month_label}",
        plain_text_content=body,
    )
    message.attachment = attachment

    _sg_client().send(message)
    print(f"Result email sent to {to_email} for period {period_date}")


async def send_error_email(to_email: str, period_date: str, error_detail: str):
    """Notify the coordinator that processing failed."""
    month_label = _month_label(period_date)
    from_email = os.getenv("FROM_EMAIL", "wkcomp@resortoutfitters.com")
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
