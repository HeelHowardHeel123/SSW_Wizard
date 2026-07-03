"""SendGrid email alerts for auto-generated parsers."""

import re
import os
import base64

SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY", "")
ALERT_EMAIL      = os.environ.get("ALERT_EMAIL", "sward@tpc.us")
FROM_EMAIL       = os.environ.get("FROM_EMAIL", ALERT_EMAIL)


def send_parser_alert(
    company_name: str,
    file_count: int,
    row_count: int,
    code: str,
    is_update: bool = False,
) -> bool:
    """Email the generated parser code to ALERT_EMAIL via SendGrid.

    is_update=True when a previous generated parser failed mid-batch and
    a new version was generated. The email subject reflects this.

    Returns True if sent successfully, False on any failure.
    """
    if not SENDGRID_API_KEY:
        return False

    try:
        from sendgrid import SendGridAPIClient
        from sendgrid.helpers.mail import (
            Attachment, Disposition, FileContent, FileName, FileType, Mail,
        )

        slug        = re.sub(r"[^a-z0-9]+", "_", company_name.lower()).strip("_")
        attach_name = f"{slug}_fringe.py"
        verb        = "Updated parser" if is_update else "New parser"

        if is_update:
            body = (
                f"UPDATED parser for: {company_name}\n"
                f"(The previous generated parser failed mid-batch — new version attached)\n\n"
                f"Batch summary:\n"
                f"  Files processed with new version: {file_count}\n"
                f"  Rows extracted:                   {row_count}\n\n"
            )
        else:
            body = (
                f"New payroll company detected: {company_name}\n\n"
                f"Batch summary:\n"
                f"  Files processed (AI): {file_count}\n"
                f"  Rows extracted:       {row_count}\n\n"
            )

        body += (
            f"The generated parser is attached as {attach_name}.\n\n"
            f"Next steps:\n"
            f"  1. Review the attached code\n"
            f"  2. Test it on a sample PDF from {company_name}\n"
            f"  3. Save to backend/parsers/{slug}/fringe.py\n"
            f"  4. Commit and push — Railway auto-deploys\n\n"
            f"Until then, {company_name} PDFs will use AI extraction each batch."
        )

        recipients = list(dict.fromkeys([ALERT_EMAIL, FROM_EMAIL]))  # dedup, preserve order
        msg = Mail(
            from_email=FROM_EMAIL,
            to_emails=recipients,
            subject=f"[TPC Wizard] {verb}: {company_name}",
            plain_text_content=body,
        )
        msg.attachment = Attachment(
            FileContent(base64.b64encode(code.encode()).decode()),
            FileName(attach_name),
            FileType("text/x-python"),
            Disposition("attachment"),
        )

        SendGridAPIClient(SENDGRID_API_KEY).send(msg)
        return True

    except Exception:
        return False
