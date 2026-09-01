"""SendGrid email alerts for run summaries."""

import os
from datetime import datetime, timezone

SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY", "")
ALERT_EMAIL      = os.environ.get("ALERT_EMAIL", "sward@tpc.us")
FROM_EMAIL       = os.environ.get("FROM_EMAIL", ALERT_EMAIL)


_ENDPOINT_LABELS = {
    "fringe":              "Fringe / Payroll",
    "freelance":           "Freelance Invoices",
    "hours_letters":       "Hours Letters",
    "talent":              "Talent & Extras",
    "billings":            "Billings",
    "agency_subvendors":   "Agency Sub-Vendors",
    "agency_hours":        "Agency Hours",
    "retainer_billings":   "Retainer Billings",
    "residency_docs":      "Residency Documents",
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
        print("[send_run_summary] SENDGRID_API_KEY is not set -- skipping send", flush=True)
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

    except Exception as e:
        body = getattr(e, "body", None)
        if isinstance(body, bytes):
            body = body.decode(errors="replace")
        print(f"[send_run_summary] send failed: {e!r}" + (f" -- body: {body}" if body else ""), flush=True)
        return False
