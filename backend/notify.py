"""SendGrid email alerts for auto-generated parsers and run summaries."""

import re
import os
import base64
from datetime import datetime, timezone

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

        msg = Mail(
            from_email=FROM_EMAIL,
            to_emails=FROM_EMAIL,
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


_ENDPOINT_LABELS = {
    "fringe":        "Fringe / Payroll",
    "freelance":     "Freelance Invoices",
    "hours_letters": "Hours Letters",
}


def send_run_summary(
    endpoint: str,
    file_summaries: list[dict],
    issues: list[str],
    new_parsers: list[dict] | None = None,
) -> bool:
    """Email a run summary after any extraction endpoint completes.

    file_summaries: [{filename, company, rows, issues}] — same shape every endpoint returns.
    new_parsers:    alert_queue entries from fringe runs [{company_name, file_count, row_count, is_update}].
    """
    if not SENDGRID_API_KEY:
        return False

    try:
        from sendgrid import SendGridAPIClient
        from sendgrid.helpers.mail import Mail

        label       = _ENDPOINT_LABELS.get(endpoint, endpoint)
        total_files = len(file_summaries)
        total_rows  = sum(f.get("rows", 0) for f in file_summaries)
        issue_count = len(issues)
        timestamp   = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

        subject = f"[TPC Wizard] {label} — {timestamp} — {total_files} file(s), {total_rows} row(s)"
        if issue_count:
            subject += f" ({issue_count} issue(s))"

        lines = [
            f"Run summary: {label}",
            f"Timestamp:   {timestamp}",
            f"Files:       {total_files}",
            f"Rows:        {total_rows}",
        ]
        if issue_count:
            lines.append(f"Issues:      {issue_count}")
        lines.append("")

        if file_summaries:
            lines.append("FILE DETAILS")
            lines.append("-" * 60)
            for f in file_summaries:
                lines.append(f"  {f.get('filename', '?')}")
                lines.append(f"    Parser / source : {f.get('company', 'unknown')}")
                lines.append(f"    Rows extracted  : {f.get('rows', 0)}")
                for err in f.get("issues", []):
                    lines.append(f"    WARNING: {err}")
            lines.append("")

        if new_parsers:
            lines.append("NEW / UPDATED PARSERS")
            lines.append("-" * 60)
            for p in new_parsers:
                verb = "Updated" if p.get("is_update") else "New"
                lines.append(
                    f"  {verb}: {p['company_name']} — "
                    f"{p['file_count']} file(s), {p['row_count']} row(s)"
                )
            lines.append("  (Parser code emailed separately with attachment)")
            lines.append("")

        if issues:
            lines.append("ALL ISSUES")
            lines.append("-" * 60)
            for issue in issues:
                lines.append(f"  • {issue}")
        else:
            lines.append("No issues.")

        msg = Mail(
            from_email=FROM_EMAIL,
            to_emails=FROM_EMAIL,
            subject=subject,
            plain_text_content="\n".join(lines),
        )
        SendGridAPIClient(SENDGRID_API_KEY).send(msg)
        return True

    except Exception:
        return False
