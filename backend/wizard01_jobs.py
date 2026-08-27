"""
Wizard 01 -- job storage, background processing, and cleanup.

Fully self-contained (own OpenAI/Anthropic client construction, own
APP_SHARED_SECRET read) rather than importing from main.py, so that main.py
can import from this module without any circular dependency -- exactly how
it already imports talent_extractor.py/payroll_reconciler.py one-directionally.

Job layout on the Railway Volume (WIZARD01_STORAGE_PATH, default
/data/wizard01_jobs):

  <job_id>/
    uploads/       -- raw uploaded files, deleted immediately once processing
                       completes (never held longer than needed)
    output/        -- per-type folders + Not Readable/, deleted immediately
                       once zipped
    result.zip     -- persists until expires_at
    status.json    -- persists until expires_at (same clock as the zip --
                       one expiry, not two)

No background-job or scheduled-cleanup infrastructure exists elsewhere in
this backend (confirmed before writing this), so both are new here rather
than extensions of something pre-existing. Job execution runs in-process via
asyncio.create_task -- if the service restarts mid-job, that job is lost
(no auto-resume). Acceptable for an internal tool at this scale; status is
read from the Volume-backed JSON sidecar rather than an in-memory dict
specifically so this failure mode stays visible/consistent rather than
silently wrong, and so polling works correctly regardless of worker count.
"""

import os
import re
import json
import uuid
import shutil
import asyncio
import zipfile
import csv
import io
from datetime import datetime, timedelta, timezone

import openai
import anthropic

import pdf_namer
from pdf_namer import BatchContext, DocResult

WIZARD01_STORAGE_PATH = os.environ.get("WIZARD01_STORAGE_PATH", "/data/wizard01_jobs")
APP_SHARED_SECRET = os.environ.get("APP_SHARED_SECRET", "")
JOB_RETENTION = timedelta(days=2)
MAX_CONCURRENT_FILES = 5

_JOB_ID_RE = re.compile(r"^[A-Za-z0-9_-]+$")


def _client():
    key = os.environ.get("OPENAI_API_KEY", "")
    if not key:
        raise ValueError("OPENAI_API_KEY is not set.")
    return openai.OpenAI(api_key=key)


def _anthropic_client():
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    return anthropic.Anthropic(api_key=key) if key else None


def _job_dir(job_id: str) -> str:
    if not _JOB_ID_RE.match(job_id):
        raise ValueError("invalid job id")
    return os.path.join(WIZARD01_STORAGE_PATH, job_id)


def _status_path(job_id: str) -> str:
    return os.path.join(_job_dir(job_id), "status.json")


def _zip_path(job_id: str) -> str:
    return os.path.join(_job_dir(job_id), "result.zip")


def _uploads_dir(job_id: str) -> str:
    return os.path.join(_job_dir(job_id), "uploads")


def _output_dir(job_id: str) -> str:
    return os.path.join(_job_dir(job_id), "output")


# ── Status read/write ─────────────────────────────────────────────────────────

def read_status(job_id: str) -> dict | None:
    path = _status_path(job_id)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _write_status(job_id: str, status: dict) -> None:
    path = _status_path(job_id)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(status, f)
    os.replace(tmp, path)


def is_expired(status: dict) -> bool:
    try:
        expires_at = datetime.fromisoformat(status["expires_at"])
    except Exception:
        return False
    return datetime.now(timezone.utc) >= expires_at


# ── Job creation ──────────────────────────────────────────────────────────────

def create_job(batch: BatchContext) -> str:
    """Creates the job directory + initial status.json and returns the new
    job id (random UUID4 -- deliberately unguessable/non-sequential, unlike
    the human-chosen ids used for /templates, since jobs are many and
    ephemeral rather than few and admin-managed). Caller still needs to
    stream the uploaded files into uploads_dir(job_id) before starting
    processing."""
    job_id = uuid.uuid4().hex
    os.makedirs(_uploads_dir(job_id), exist_ok=True)
    os.makedirs(_output_dir(job_id), exist_ok=True)
    now = datetime.now(timezone.utc)
    status = {
        "job_id": job_id,
        "created_at": now.isoformat(),
        "expires_at": (now + JOB_RETENTION).isoformat(),
        "state": "running",
        "total": 0,
        "processed": 0,
        "renamed": 0,
        "not_readable": 0,
        "needs_review": 0,
        "log": [],
        "batch_context": {
            "client_name": batch.client_name, "agency_name": batch.agency_name,
            "prodco_name": batch.prodco_name, "received_from": batch.received_from,
            "vendor_naming": batch.vendor_naming,
        },
    }
    _write_status(job_id, status)
    return job_id


def uploads_dir(job_id: str) -> str:
    """Public accessor for main.py's upload-streaming endpoint."""
    return _uploads_dir(job_id)


# ── Background processing ─────────────────────────────────────────────────────

async def run_job(job_id: str, uploaded: list[tuple[str, str]]) -> None:
    """uploaded: list of (file_id, original_filename) already written to
    uploads_dir(job_id)/<file_id>__<original_filename> by the caller. Runs
    entirely in the background -- the request that triggered this has
    already returned by the time most of this executes."""
    status = read_status(job_id)
    if status is None:
        return
    status["total"] = len(uploaded)
    _write_status(job_id, status)

    lock = asyncio.Lock()
    sem = asyncio.Semaphore(MAX_CONCURRENT_FILES)
    loop = asyncio.get_running_loop()
    openai_client = _client()
    anthropic_client = _anthropic_client()
    batch = BatchContext(**status["batch_context"])
    unable_counters: dict[str, int] = {}

    async def _update(entries: list[dict]):
        """One call per UPLOADED file, regardless of how many output
        documents it produced -- processed increments once (it tracks
        upload progress), while renamed/not_readable/needs_review and the
        log each get one entry per OUTPUT document, since Residency can
        split one upload into several (a batch-scanned stack of IDs)."""
        async with lock:
            s = read_status(job_id)
            s["processed"] += 1
            for entry in entries:
                bucket = entry["bucket"]
                if bucket == pdf_namer.BUCKET_RENAMED:
                    s["renamed"] += 1
                elif bucket == pdf_namer.BUCKET_NOT_READABLE:
                    s["not_readable"] += 1
                else:
                    s["needs_review"] += 1
                s["log"].append(entry)
            _write_status(job_id, s)

    async def _process_one(file_id: str, original_filename: str):
        async with sem:
            src_path = os.path.join(_uploads_dir(job_id), f"{file_id}__{original_filename}")
            # No extension whitelist here -- pdf_namer.load_source_as_pdf_bytes
            # accepts a PDF or effectively any image format Pillow can open
            # (PNG/JPG/BMP/GIF/TIFF/WEBP/HEIC-HEIF and more), regardless of
            # the file's own extension. Anything genuinely unreadable raises
            # there and lands in Not Readable via process_one_document's own
            # exception handling below -- no separate check needed here.
            try:
                with open(src_path, "rb") as f:
                    raw_bytes = f.read()
            except Exception as e:
                entry = {
                    "file_id": uuid.uuid4().hex, "filename": original_filename,
                    "bucket": pdf_namer.BUCKET_NOT_READABLE, "doc_type": "unknown",
                    "new_name": None, "reason_code": "unreadable",
                    "reason_detail": f"Could not read uploaded file: {e}",
                }
                await _update([entry])
                return

            results = await pdf_namer.process_one_document(
                raw_bytes, original_filename, batch, openai_client, anthropic_client, loop,
            )

            entries = []
            for result in results:
                final_name = None
                try:
                    if result.bucket == pdf_namer.BUCKET_RENAMED:
                        # Proof of Payments uses "_002"/"_003" collision
                        # suffixes instead of the " (2)"/" (3)" style used
                        # everywhere else, per explicit instruction.
                        suffix_style = "underscore" if result.doc_type == "proof_of_payment" else "paren"
                        final_name = _write_output(
                            job_id, result.subfolder, result.new_filename, result.output_pdf_bytes,
                            suffix_style=suffix_style,
                        )
                    elif result.bucket == pdf_namer.BUCKET_UNABLE_TO_RENAME:
                        async with lock:
                            n = unable_counters.get(result.subfolder, 0) + 1
                            unable_counters[result.subfolder] = n
                        # "aaa" prefix floats these to the top of the folder
                        # alphabetically, ahead of every real renamed file.
                        final_name = _write_output(job_id, result.subfolder, f"aaaUnable_To_Rename_{n:03d}.pdf", result.output_pdf_bytes)
                    else:
                        final_name = _write_not_readable(job_id, None, original_filename, data=result.output_pdf_bytes or raw_bytes)
                except Exception as e:
                    result.reason_detail = f"{result.reason_detail} (also failed to write output: {e})".strip()

                entries.append({
                    "file_id": uuid.uuid4().hex, "filename": original_filename,
                    "bucket": result.bucket, "doc_type": result.doc_type,
                    "new_name": final_name, "reason_code": result.reason_code,
                    "reason_detail": result.reason_detail,
                })

            await _update(entries)

            try:
                os.remove(src_path)
            except Exception:
                pass

    await asyncio.gather(*[_process_one(fid, fn) for fid, fn in uploaded])

    try:
        _finalize_job(job_id)
    except Exception as e:
        s = read_status(job_id)
        if s is not None:
            s["state"] = "error"
            s["log"].append({
                "file_id": "", "filename": "", "bucket": "error", "doc_type": "",
                "new_name": None, "reason_code": "provider_error",
                "reason_detail": f"Failed to finalize job: {e}",
            })
            _write_status(job_id, s)


def _resolve_collision_free_path(dest_dir: str, dest_filename: str, suffix_style: str = "paren") -> str:
    """suffix_style "paren" (default, used everywhere else in this tool)
    produces "Name (2).pdf", "Name (3).pdf", ...; "underscore" (Proof of
    Payments only, per explicit instruction) produces "Name_002.pdf",
    "Name_003.pdf", ... instead."""
    os.makedirs(dest_dir, exist_ok=True)
    base, ext = os.path.splitext(dest_filename)
    candidate = dest_filename
    n = 2
    while os.path.exists(os.path.join(dest_dir, candidate)):
        if suffix_style == "underscore":
            candidate = f"{base}_{n:03d}{ext}"
        else:
            candidate = f"{base} ({n}){ext}"
        n += 1
    return os.path.join(dest_dir, candidate)


def _write_output(job_id: str, subfolder: str, filename: str, data: bytes, suffix_style: str = "paren") -> str:
    dest_dir = os.path.join(_output_dir(job_id), subfolder)
    final_path = _resolve_collision_free_path(dest_dir, filename, suffix_style=suffix_style)
    with open(final_path, "wb") as f:
        f.write(data)
    return os.path.basename(final_path)


def _write_not_readable(job_id: str, src_path: str | None, original_filename: str, data: bytes = None) -> str:
    dest_dir = os.path.join(_output_dir(job_id), pdf_namer.FOLDER_NOT_READABLE)
    final_path = _resolve_collision_free_path(dest_dir, original_filename)
    if data is not None:
        with open(final_path, "wb") as f:
            f.write(data)
    elif src_path and os.path.exists(src_path):
        shutil.copyfile(src_path, final_path)
    return os.path.basename(final_path)


def _finalize_job(job_id: str) -> None:
    """Zips output/ (with a _manifest.csv from the log), writes result.zip,
    deletes uploads/ and output/ (only the zip + status persist from here),
    and flips state to done."""
    status = read_status(job_id)
    output_dir = _output_dir(job_id)

    manifest_buf = io.StringIO()
    writer = csv.writer(manifest_buf)
    writer.writerow(["original_name", "new_name", "bucket", "reason_code", "reason_detail"])
    for entry in status["log"]:
        writer.writerow([
            entry.get("filename", ""), entry.get("new_name") or "", entry.get("bucket", ""),
            entry.get("reason_code", ""), entry.get("reason_detail", ""),
        ])
    with open(os.path.join(output_dir, "_manifest.csv"), "w", encoding="utf-8", newline="") as f:
        f.write(manifest_buf.getvalue())

    zip_path = _zip_path(job_id)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(output_dir):
            for fn in files:
                full = os.path.join(root, fn)
                arcname = os.path.relpath(full, output_dir)
                zf.write(full, arcname)

    shutil.rmtree(_uploads_dir(job_id), ignore_errors=True)
    shutil.rmtree(output_dir, ignore_errors=True)

    status["state"] = "done"
    _write_status(job_id, status)


# ── Cleanup ────────────────────────────────────────────────────────────────────

def sweep_expired_jobs() -> None:
    """No scheduler infra exists in this backend, so this runs opportunistically
    at the top of the job-creation endpoint (the natural highest-traffic entry
    point) rather than on a cron. Deletes any job whose expires_at has passed,
    whatever state it's in."""
    if not os.path.isdir(WIZARD01_STORAGE_PATH):
        return
    for name in os.listdir(WIZARD01_STORAGE_PATH):
        job_dir = os.path.join(WIZARD01_STORAGE_PATH, name)
        if not os.path.isdir(job_dir):
            continue
        status = read_status(name)
        if status is None or is_expired(status):
            shutil.rmtree(job_dir, ignore_errors=True)


def delete_job(job_id: str) -> bool:
    job_dir = _job_dir(job_id)
    if not os.path.isdir(job_dir):
        return False
    shutil.rmtree(job_dir, ignore_errors=True)
    return True
