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
            f"  3. Save as the next fringe_NNN.py in backend/parsers/{slug}/\n"
            f"     (e.g. fringe_001.py if the folder is empty, fringe_002.py alongside an existing fringe_001.py)\n"
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
    "fringe":              "Fringe / Payroll",
    "freelance":           "Freelance Invoices",
    "hours_letters":       "Hours Letters",
    "talent":              "Talent & Extras",
    "billings":            "Billings",
    "agency_subvendors":   "Agency Sub-Vendors",
    "agency_hours":        "Agency Hours",
    "retainer_billings":   "Retainer Billings",
}


def send_run_summary(
    project_title: str,
    workbook_type: str,
    runs: list[dict],
) -> bool:
    """Email a consolidated run summary after workbook assembly.

    runs: [{endpoint, files: [{filename, company, rows, issues}], issues}]
    """
    if not SENDGRID_API_KEY:
        return False

    try:
        from sendgrid import SendGridAPIClient
        from sendgrid.helpers.mail import Mail

        timestamp   = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
        total_rows  = sum(f.get("rows") or 0 for r in runs for f in r.get("files", []))
        all_issues  = [i for r in runs for i in r.get("issues", [])]
        issue_count = len(all_issues)

        title_part = project_title.strip() if project_title.strip() else "TPC Wizard Run"
        subject = f"[TPC Wizard] {title_part} — {timestamp} — {total_rows} row(s)"
        if issue_count:
            subject += f" ({issue_count} issue(s))"

        lines = [
            f"Project:   {title_part}" + (f" ({workbook_type})" if workbook_type else ""),
            f"Timestamp: {timestamp}",
            f"Total rows: {total_rows}",
        ]
        if issue_count:
            lines.append(f"Issues:    {issue_count}")
        lines.append("")

        for run in runs:
            endpoint = run.get("endpoint", "")
            label    = _ENDPOINT_LABELS.get(endpoint, endpoint)
            files    = run.get("files", [])
            run_rows = sum(f.get("rows") or 0 for f in files)

            lines.append(f"{label.upper()}  —  {len(files)} file(s), {run_rows} row(s)")
            lines.append("-" * 60)
            for f in files:
                lines.append(f"  {f.get('filename', '?')}")
                lines.append(f"    Source  : {f.get('company', 'unknown')}")
                row_display = f.get("rows")
                lines.append(f"    Rows    : {row_display if row_display is not None else '—'}")
                for err in f.get("issues", []):
                    lines.append(f"    WARNING : {err}")
            lines.append("")

        if all_issues:
            lines.append("ALL ISSUES")
            lines.append("-" * 60)
            for issue in all_issues:
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
