#!/usr/bin/env python3
"""Archive LinkedIn recruiting emails from Gmail Inbox to local storage."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from pathlib import Path
from typing import Any, Iterable, Sequence

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
ARCHIVE_ROOT = Path("RecruitingInboxArchive")
DB_PATH = ARCHIVE_ROOT / "state.db"
EXCEL_PATH = ARCHIVE_ROOT / "candidates.xlsx"
ROLE_MAP_SHEET = "_role_map"
ALL_CANDIDATES_SHEET = "_all_candidates"
QUERY_BASE = (
    'in:inbox from:jobs-listings@linkedin.com '
    '(subject:"New application:" OR subject:"Your job is posted:")'
)


@dataclass
class SubjectInfo:
    """Represents parsed LinkedIn recruiting subject metadata."""

    kind: str
    role: str
    candidate: str | None


@dataclass
class AttachmentItem:
    """Represents an extracted attachment from a Gmail message."""

    filename: str
    mime_type: str
    data: bytes


@dataclass
class MessageContext:
    """Canonical parsed context for one message."""

    gmail_id: str
    subject: str
    received_dt: datetime
    role: str
    candidate: str | None
    short_hash: str
    subject_kind: str
    body_text: str
    raw_bytes: bytes
    attachments: list[AttachmentItem]


@dataclass
class RuntimeStats:
    """Processing counters shown in final summary."""

    matched: int = 0
    skipped: int = 0
    processed: int = 0
    failed: int = 0


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(description="LinkedIn-only Gmail Inbox archiver")
    parser.add_argument("--start-date", help="Start date in YYYY-MM-DD")
    parser.add_argument("--end-date", help="End date in YYYY-MM-DD")
    parser.add_argument("--credentials", default="credentials.json", help="OAuth client secrets file")
    parser.add_argument("--token", default="token.json", help="OAuth token cache path")
    parser.add_argument("--dry-run", action="store_true", help="Fetch and simulate writes without persisting")
    parser.add_argument("--per-role-sheets", action="store_true", help="Also write per-role sheets")
    parser.add_argument("--log-level", default="INFO", help="Logging level")
    return parser.parse_args()


def parse_iso_date(value: str) -> date:
    """Parse YYYY-MM-DD into date."""
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid date '{value}', expected YYYY-MM-DD") from exc


def compute_date_range(start_arg: str | None, end_arg: str | None) -> tuple[date, date]:
    """Compute inclusive date range according to spec."""
    today = datetime.now(timezone.utc).date()

    if start_arg and end_arg:
        start_date = parse_iso_date(start_arg)
        end_date = parse_iso_date(end_arg)
    elif start_arg and not end_arg:
        start_date = parse_iso_date(start_arg)
        end_date = today
    elif not start_arg and end_arg:
        start_date = today - timedelta(days=30)
        end_date = parse_iso_date(end_arg)
    else:
        end_date = today
        start_date = today - timedelta(days=30)

    if start_date > end_date:
        raise ValueError("start-date must be <= end-date")

    return start_date, end_date


def build_gmail_query(start_date: date, end_date: date) -> str:
    """Build Gmail query with inclusive range by shifting `before` by +1 day."""
    before_exclusive = end_date + timedelta(days=1)
    after_s = start_date.strftime("%Y/%m/%d")
    before_s = before_exclusive.strftime("%Y/%m/%d")
    return f"{QUERY_BASE} after:{after_s} before:{before_s}"


def sanitize_fs_name(value: str, max_len: int = 100) -> str:
    """Sanitize filesystem name for cross-platform compatibility."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    cleaned = cleaned or "unnamed"
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(". ")
    return cleaned or "unnamed"


def parse_subject(subject: str) -> SubjectInfo | None:
    """Parse supported LinkedIn recruiting subjects."""
    app_match = re.fullmatch(r"New application:\s*(?P<role>.+?)\s+from\s+(?P<candidate>.+)", subject.strip())
    if app_match:
        return SubjectInfo(
            kind="application",
            role=app_match.group("role").strip(),
            candidate=app_match.group("candidate").strip(),
        )

    post_match = re.fullmatch(r"Your job is posted:\s*\[(?P<role>.+?)\]\s*", subject.strip())
    if post_match:
        return SubjectInfo(kind="job_posted", role=post_match.group("role").strip(), candidate=None)

    return None


def get_gmail_service(credentials_path: Path, token_path: Path, dry_run: bool):
    """Build authenticated Gmail API service using OAuth desktop flow."""
    creds: Credentials | None = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)
        if not dry_run:
            token_path.write_text(creds.to_json(), encoding="utf-8")
        else:
            logging.info("DRY RUN: skipped token persistence to %s", token_path)

    return build("gmail", "v1", credentials=creds)


def fetch_messages(service: Any, query: str) -> list[dict[str, Any]]:
    """Fetch all message IDs matching query."""
    logging.info("Fetching Gmail message list")
    messages: list[dict[str, Any]] = []
    page_token: str | None = None

    while True:
        response = (
            service.users()
            .messages()
            .list(userId="me", q=query, pageToken=page_token, maxResults=500)
            .execute()
        )
        messages.extend(response.get("messages", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break

    return messages


def extract_header(payload_headers: Sequence[dict[str, str]], name: str) -> str:
    """Extract header value case-insensitively from Gmail payload headers."""
    lower_name = name.lower()
    for h in payload_headers:
        if h.get("name", "").lower() == lower_name:
            return h.get("value", "")
    return ""


def decode_gmail_b64(data: str) -> bytes:
    """Decode Gmail URL-safe base64 data."""
    padding = '=' * (-len(data) % 4)
    return base64.urlsafe_b64decode(data + padding)


def collect_parts(payload: dict[str, Any]) -> Iterable[dict[str, Any]]:
    """Yield payload part tree depth-first."""
    yield payload
    for part in payload.get("parts", []) or []:
        yield from collect_parts(part)


def extract_body_text(payload: dict[str, Any]) -> str:
    """Extract plaintext body from Gmail payload tree."""
    for part in collect_parts(payload):
        if part.get("mimeType") == "text/plain":
            body_data = (part.get("body") or {}).get("data")
            if body_data:
                return decode_gmail_b64(body_data).decode("utf-8", errors="replace")
    for part in collect_parts(payload):
        if part.get("mimeType") == "text/html":
            body_data = (part.get("body") or {}).get("data")
            if body_data:
                html = decode_gmail_b64(body_data).decode("utf-8", errors="replace")
                return re.sub(r"<[^>]+>", "", html)
    return ""


def fetch_attachments(service: Any, payload: dict[str, Any], message_id: str) -> list[AttachmentItem]:
    """Fetch attachments referenced by payload parts."""
    attachments: list[AttachmentItem] = []
    for part in collect_parts(payload):
        filename = part.get("filename") or ""
        body = part.get("body") or {}
        att_id = body.get("attachmentId")
        mime = part.get("mimeType") or "application/octet-stream"
        if filename and att_id:
            att_resp = (
                service.users()
                .messages()
                .attachments()
                .get(userId="me", messageId=message_id, id=att_id)
                .execute()
            )
            data = decode_gmail_b64(att_resp["data"])
            attachments.append(AttachmentItem(filename=filename, mime_type=mime, data=data))
    return attachments


def to_email_message(raw_bytes: bytes) -> EmailMessage:
    """Parse raw RFC822 bytes into EmailMessage."""
    parsed = BytesParser(policy=policy.default).parsebytes(raw_bytes)
    return parsed


def build_context(service: Any, gmail_id: str) -> MessageContext | None:
    """Fetch and parse one Gmail message to processing context."""
    resp = (
        service.users()
        .messages()
        .get(userId="me", id=gmail_id, format="full")
        .execute()
    )
    payload = resp.get("payload", {})
    headers = payload.get("headers", [])

    subject = extract_header(headers, "Subject")
    parsed_subject = parse_subject(subject)
    if not parsed_subject:
        logging.info("Skipping message %s due to unsupported subject", gmail_id)
        return None

    internal_ts = int(resp.get("internalDate", "0")) / 1000
    received_dt = datetime.fromtimestamp(internal_ts, tz=timezone.utc)

    raw_resp = (
        service.users()
        .messages()
        .get(userId="me", id=gmail_id, format="raw")
        .execute()
    )
    raw_bytes = decode_gmail_b64(raw_resp["raw"])
    _ = to_email_message(raw_bytes)

    attachments = fetch_attachments(service, payload, gmail_id)
    short_hash = hashlib.sha1(gmail_id.encode("utf-8")).hexdigest()[:8]
    body_text = extract_body_text(payload)

    return MessageContext(
        gmail_id=gmail_id,
        subject=subject,
        received_dt=received_dt,
        role=parsed_subject.role,
        candidate=parsed_subject.candidate,
        short_hash=short_hash,
        subject_kind=parsed_subject.kind,
        body_text=body_text,
        raw_bytes=raw_bytes,
        attachments=attachments,
    )


def init_db(db_path: Path) -> None:
    """Create SQLite table if needed."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS processed_messages(
                gmail_id TEXT PRIMARY KEY,
                processed_at TEXT
            )
            """
        )
        conn.commit()


def already_processed(db_path: Path, gmail_id: str) -> bool:
    """Check if message ID already processed."""
    if not db_path.exists():
        return False
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT 1 FROM processed_messages WHERE gmail_id = ?", (gmail_id,)
        ).fetchone()
    return bool(row)


def update_db(db_path: Path, gmail_id: str) -> None:
    """Mark Gmail ID as processed in SQLite."""
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO processed_messages(gmail_id, processed_at) VALUES (?, ?)",
            (gmail_id, datetime.now(timezone.utc).isoformat()),
        )
        conn.commit()


def resolve_paths(ctx: MessageContext) -> dict[str, Path]:
    """Resolve canonical output paths for this message."""
    role_fs = sanitize_fs_name(ctx.role, max_len=60)
    month_dir = ctx.received_dt.strftime("%Y-%m")
    date_s = ctx.received_dt.strftime("%Y-%m-%d")
    base_dir = ARCHIVE_ROOT / role_fs / month_dir

    stem = f"{date_s}__msg-{ctx.short_hash}"
    return {
        "base_dir": base_dir,
        "raw": base_dir / "raw" / f"{stem}.eml",
        "body": base_dir / "bodies" / f"{stem}.txt",
        "meta": base_dir / "meta" / f"{stem}.json",
        "resumes_dir": base_dir / "resumes",
    }


def build_resume_filename(ctx: MessageContext, original_name: str) -> str:
    """Build normalized resume filename."""
    date_s = ctx.received_dt.strftime("%Y-%m-%d")
    candidate = sanitize_fs_name(ctx.candidate or "UnknownCandidate", max_len=60)
    original = sanitize_fs_name(original_name, max_len=120)
    return sanitize_fs_name(
        f"{candidate}__{date_s}__msg-{ctx.short_hash}__{original}", max_len=220
    )


def save_files(ctx: MessageContext, paths: dict[str, Path], dry_run: bool) -> list[str]:
    """Persist message files; returns list of saved resume filenames."""
    saved_resume_names: list[str] = []

    resume_names = [build_resume_filename(ctx, a.filename) for a in ctx.attachments]

    if dry_run:
        return resume_names

    for dirname in ("raw", "bodies", "meta", "resumes"):
        (paths["base_dir"] / dirname).mkdir(parents=True, exist_ok=True)

    paths["raw"].write_bytes(ctx.raw_bytes)
    paths["body"].write_text(ctx.body_text, encoding="utf-8")

    metadata = {
        "gmail_id": ctx.gmail_id,
        "subject": ctx.subject,
        "received_date_utc": ctx.received_dt.isoformat(),
        "role": ctx.role,
        "candidate": ctx.candidate,
        "kind": ctx.subject_kind,
        "attachments": [a.filename for a in ctx.attachments],
    }
    paths["meta"].write_text(json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8")

    for attachment, resume_name in zip(ctx.attachments, resume_names):
        target = paths["resumes_dir"] / resume_name
        target.write_bytes(attachment.data)
        saved_resume_names.append(resume_name)

    return saved_resume_names


def update_excel(rows: list[dict[str, str]], excel_path: Path, per_role_sheets: bool, dry_run: bool) -> None:
    """Upsert candidate rows into workbook sheets."""
    if not rows:
        return
    if dry_run:
        return

    excel_path.parent.mkdir(parents=True, exist_ok=True)

    all_cols = ["role", "candidate_name", "applied_date", "resume_filename", "gmail_id"]

    if excel_path.exists():
        existing_all = pd.read_excel(excel_path, sheet_name=ALL_CANDIDATES_SHEET)
        existing_map = pd.read_excel(excel_path, sheet_name=ROLE_MAP_SHEET)
    else:
        existing_all = pd.DataFrame(columns=all_cols)
        existing_map = pd.DataFrame(columns=["original_role", "sheet_name"])

    new_df = pd.DataFrame(rows, columns=all_cols)
    merged_all = pd.concat([existing_all, new_df], ignore_index=True)
    merged_all = merged_all.drop_duplicates(subset=["gmail_id", "resume_filename"], keep="last")

    role_map = existing_map.copy()
    role_sheet_records: dict[str, str] = {}

    def role_to_sheet_name(role: str) -> str:
        base = sanitize_fs_name(role, max_len=31)
        base = re.sub(r"[\[\]:*?/\\]", "_", base).strip() or "role"
        return base[:31]

    for role in merged_all["role"].dropna().unique():
        if not ((role_map["original_role"] == role).any()):
            sheet = role_to_sheet_name(str(role))
            existing_sheets = set(role_map["sheet_name"].tolist())
            candidate = sheet
            idx = 1
            while candidate in existing_sheets:
                suffix = f"_{idx}"
                candidate = (sheet[: 31 - len(suffix)] + suffix).strip()
                idx += 1
            role_map = pd.concat(
                [
                    role_map,
                    pd.DataFrame([{"original_role": role, "sheet_name": candidate}]),
                ],
                ignore_index=True,
            )

    role_map = role_map.drop_duplicates(subset=["original_role"], keep="last")

    for _, row in role_map.iterrows():
        role_sheet_records[str(row["original_role"])] = str(row["sheet_name"])

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        merged_all.to_excel(writer, sheet_name=ALL_CANDIDATES_SHEET, index=False)
        role_map.to_excel(writer, sheet_name=ROLE_MAP_SHEET, index=False)
        if per_role_sheets:
            for role, sheet in role_sheet_records.items():
                role_df = merged_all[merged_all["role"] == role]
                role_df.to_excel(writer, sheet_name=sheet, index=False)


def print_dry_run_message(ctx: MessageContext, paths: dict[str, Path], simulated_resumes: Sequence[str]) -> None:
    """Print required dry-run per-email simulation block."""
    print("----------------------------------------")
    print("Email:")
    print(f"  Gmail ID: {ctx.gmail_id}")
    print(f"  Subject: {ctx.subject}")
    print(f"  Role: {ctx.role}")
    print(f"  Candidate: {ctx.candidate or 'N/A'}")
    print(f"  Received Date: {ctx.received_dt.strftime('%Y-%m-%d')}")
    print("  Attachments:")
    if simulated_resumes:
        for name in simulated_resumes:
            print(f"    - {name}")
    else:
        print("    - (none)")
    target = f"RecruitingInboxArchive/{sanitize_fs_name(ctx.role, max_len=60)}/{ctx.received_dt.strftime('%Y-%m')}/"
    print(f"  Target Folder: {target}")
    print("----------------------------------------")


def handle_message(
    service: Any,
    message_id: str,
    db_path: Path,
    dry_run: bool,
) -> tuple[str, list[dict[str, str]]]:
    """Handle one message and return status plus excel candidate rows."""
    if already_processed(db_path, message_id):
        return "skipped", []

    ctx = build_context(service, message_id)
    if ctx is None:
        return "ignored", []

    paths = resolve_paths(ctx)
    resume_names = save_files(ctx, paths, dry_run=dry_run)

    if dry_run:
        print_dry_run_message(ctx, paths, resume_names)

    candidate_rows: list[dict[str, str]] = []
    if ctx.subject_kind == "application":
        applied_date = ctx.received_dt.strftime("%Y-%m-%d")
        if resume_names:
            for resume_name in resume_names:
                candidate_rows.append(
                    {
                        "role": ctx.role,
                        "candidate_name": ctx.candidate or "",
                        "applied_date": applied_date,
                        "resume_filename": resume_name,
                        "gmail_id": ctx.gmail_id,
                    }
                )
        else:
            candidate_rows.append(
                {
                    "role": ctx.role,
                    "candidate_name": ctx.candidate or "",
                    "applied_date": applied_date,
                    "resume_filename": "",
                    "gmail_id": ctx.gmail_id,
                }
            )

    if not dry_run:
        update_db(db_path, ctx.gmail_id)

    return "processed", candidate_rows


def configure_logging(level: str) -> None:
    """Configure root logging."""
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def main() -> int:
    """Program entrypoint."""
    args = parse_args()
    configure_logging(args.log_level)

    try:
        start_date, end_date = compute_date_range(args.start_date, args.end_date)
    except Exception as exc:
        logging.error("Invalid date range: %s", exc)
        return 2

    query = build_gmail_query(start_date, end_date)
    logging.info("Using Gmail query: %s", query)

    credentials_path = Path(args.credentials)
    token_path = Path(args.token)
    if not credentials_path.exists():
        logging.error("Credentials file not found: %s", credentials_path)
        return 2

    if not args.dry_run:
        init_db(DB_PATH)

    try:
        service = get_gmail_service(credentials_path, token_path, args.dry_run)
        messages = fetch_messages(service, query)
    except HttpError as exc:
        logging.error("Gmail API error: %s", exc)
        return 1

    stats = RuntimeStats(matched=len(messages))
    excel_rows: list[dict[str, str]] = []

    for message in messages:
        gmail_id = message.get("id")
        if not gmail_id:
            continue
        try:
            status, rows = handle_message(service, gmail_id, DB_PATH, args.dry_run)
            if status == "skipped":
                stats.skipped += 1
            elif status == "processed":
                stats.processed += 1
                excel_rows.extend(rows)
        except Exception as exc:  # per-message crash guard
            stats.failed += 1
            logging.exception("Failed to process message %s: %s", gmail_id, exc)

    try:
        update_excel(excel_rows, EXCEL_PATH, args.per_role_sheets, args.dry_run)
    except Exception as exc:
        logging.error("Failed to update workbook: %s", exc)
        return 1

    print("\n========== RUN SUMMARY ==========")
    if args.dry_run:
        print("*** DRY RUN MODE ACTIVE ***")
    print(f"Total matched: {stats.matched}")
    print(f"Total skipped (already processed): {stats.skipped}")
    print(f"Total processed: {stats.processed}")
    print(f"Total failed: {stats.failed}")
    print(f"Date range used: {start_date.isoformat()} to {end_date.isoformat()} (inclusive)")
    print(f"Gmail query used: {query}")
    print("=================================")

    return 0 if stats.failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
