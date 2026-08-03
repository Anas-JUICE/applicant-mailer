"""
emailer.py
----------
All the "brains" of the mail merge tool, kept separate from the Streamlit UI
(app.py) so it can be tested/imported independently.

Responsibilities:
  1. Reading an .xlsx workbook that may have multiple sheets/tabs, each with
     a DIFFERENT and CHANGING set of columns (because it's a Google Form
     response sheet and the form owner can add/remove/reorder questions).
  2. Auto-detecting which column is the email address and which is the
     applicant's name, on a per-sheet basis, without assuming fixed column
     positions or exact header text.
  3. Building one unified "recipients" table across all sheets.
  4. Tracking who has already been emailed ("sent log") so re-uploading an
     updated sheet only flags genuinely new people by default.
  5. Personalizing subject/body text with {name} / {first_name} / {email}.
  6. Actually sending the emails over SMTP (Gmail / Outlook / Yahoo / custom).
"""

from __future__ import annotations

import io
import json
import re
import smtplib
import ssl
import time
from dataclasses import dataclass, field
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

# --------------------------------------------------------------------------- #
# Column auto-detection
# --------------------------------------------------------------------------- #

# Checked in order, most specific first. The FIRST header that contains one
# of these substrings (after normalizing whitespace/case) wins.
EMAIL_HEADER_PRIORITY = [
    "e-mail address",
    "email address",
    "e-mail",
    "email",
]

# "Name" is trickier because forms often also have "Name of Institution",
# "Country/City name", etc. We first look for very specific, high-confidence
# phrases, then fall back to any column containing "name" that isn't
# obviously about something else.
NAME_HEADER_PRIORITY = [
    "full name of the advisor",
    "applicant full name",
    "applicant name",
    "full name",
]
NAME_HEADER_EXCLUDE = [
    "institution", "school", "university", "organi", "committee",
    "country", "conference", "event", "delegation size", "team",
]

TIMESTAMP_HEADER_PRIORITY = ["timestamp", "submitted", "date submitted"]


def _normalize(header) -> str:
    """Collapse whitespace/newlines and lowercase a header cell."""
    return re.sub(r"\s+", " ", str(header)).strip().lower()


def detect_email_column(headers: list) -> Optional[int]:
    """Return the index of the most likely email column, or None."""
    norm = [_normalize(h) for h in headers]
    for keyword in EMAIL_HEADER_PRIORITY:
        for i, h in enumerate(norm):
            if keyword in h:
                return i
    return None


def detect_name_column(headers: list) -> Optional[int]:
    """Return the index of the most likely full-name column, or None."""
    norm = [_normalize(h) for h in headers]
    for keyword in NAME_HEADER_PRIORITY:
        for i, h in enumerate(norm):
            if keyword in h:
                return i
    # Fallback: loosest match, but skip headers that are clearly about
    # something else (institution name, country name, etc.)
    for i, h in enumerate(norm):
        if "name" in h and not any(bad in h for bad in NAME_HEADER_EXCLUDE):
            return i
    return None


def detect_timestamp_column(headers: list) -> Optional[int]:
    norm = [_normalize(h) for h in headers]
    for keyword in TIMESTAMP_HEADER_PRIORITY:
        for i, h in enumerate(norm):
            if keyword in h:
                return i
    return None


def _find_header_row(raw: pd.DataFrame, max_scan_rows: int = 10) -> Optional[int]:
    """
    Scan the first few rows of a sheet (read with header=None) to find the
    row that actually contains a recognizable email-address header. Handles
    sheets where a title row or blank row sits above the real header.
    """
    scan_limit = min(max_scan_rows, len(raw))
    for row_idx in range(scan_limit):
        row_vals = raw.iloc[row_idx].tolist()
        if detect_email_column(row_vals) is not None:
            return row_idx
    return None


EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(value: str) -> bool:
    return bool(value) and bool(EMAIL_REGEX.match(str(value).strip()))


# --------------------------------------------------------------------------- #
# Loading the workbook -> one unified recipients DataFrame
# --------------------------------------------------------------------------- #

@dataclass
class SheetParseResult:
    sheet_name: str
    used: bool
    reason: str = ""
    email_header: str = ""
    name_header: str = ""
    row_count: int = 0


def load_recipients_from_workbook(file_bytes: bytes):
    """
    Parse every sheet in the workbook. For each sheet, auto-detect the
    email/name/timestamp columns and pull out clean recipient rows.

    Returns:
        recipients_df: pd.DataFrame with columns
            ['Category', 'Name', 'Email', 'SubmittedAt', 'Key']
        parse_report: list[SheetParseResult]  (for a "here's what I found /
            skipped and why" summary shown in the UI)
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    all_rows = []
    report: list[SheetParseResult] = []

    for sheet_name in xls.sheet_names:
        raw = xls.parse(sheet_name, header=None, dtype=str)

        if raw.empty:
            report.append(SheetParseResult(sheet_name, False, "Sheet is empty"))
            continue

        header_row = _find_header_row(raw)
        if header_row is None:
            report.append(
                SheetParseResult(sheet_name, False, "No email-address column found")
            )
            continue

        df = xls.parse(sheet_name, header=header_row)
        df = df.dropna(how="all")  # drop fully-blank rows
        headers = list(df.columns)

        email_idx = detect_email_column(headers)
        name_idx = detect_name_column(headers)
        ts_idx = detect_timestamp_column(headers)

        if email_idx is None or df.shape[0] == 0:
            report.append(
                SheetParseResult(sheet_name, False, "No email-address column found")
            )
            continue

        email_col = headers[email_idx]
        name_col = headers[name_idx] if name_idx is not None else None
        ts_col = headers[ts_idx] if ts_idx is not None else None

        kept = 0
        for _, row in df.iterrows():
            email_val = row.get(email_col)
            if pd.isna(email_val):
                continue
            email_val = str(email_val).strip()
            if not is_valid_email(email_val):
                continue

            name_val = ""
            if name_col is not None:
                nv = row.get(name_col)
                if not pd.isna(nv):
                    name_val = str(nv).strip()

            ts_val = ""
            if ts_col is not None:
                tv = row.get(ts_col)
                if not pd.isna(tv):
                    ts_val = str(tv)

            all_rows.append(
                {
                    "Category": sheet_name,
                    "Name": name_val,
                    "Email": email_val,
                    "SubmittedAt": ts_val,
                    "Key": f"{sheet_name}||{email_val.lower()}",
                }
            )
            kept += 1

        report.append(
            SheetParseResult(
                sheet_name,
                True,
                f"Found {kept} recipient(s)",
                email_header=email_col,
                name_header=name_col or "(not found — will use fallback greeting)",
                row_count=kept,
            )
        )

    if all_rows:
        recipients_df = pd.DataFrame(all_rows)
        # If the same person filled the same form's tab twice, keep only
        # their most recent submission (last row wins).
        recipients_df = recipients_df.drop_duplicates(subset="Key", keep="last")
        recipients_df = recipients_df.reset_index(drop=True)
    else:
        recipients_df = pd.DataFrame(
            columns=["Category", "Name", "Email", "SubmittedAt", "Key"]
        )

    return recipients_df, report


# --------------------------------------------------------------------------- #
# Sent-log persistence (so re-uploading the sheet only flags NEW people)
# --------------------------------------------------------------------------- #

def load_sent_log(path: str) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text())
    except (json.JSONDecodeError, OSError):
        return {}


def save_sent_log(path: str, log: dict) -> None:
    Path(path).write_text(json.dumps(log, indent=2, default=str))


def mark_sent(log: dict, key: str, subject: str) -> None:
    entry = log.get(key, {"times_sent": 0, "history": []})
    entry["times_sent"] = entry.get("times_sent", 0) + 1
    entry["last_sent"] = pd.Timestamp.now().isoformat(timespec="seconds")
    entry["last_subject"] = subject
    log[key] = entry


# --------------------------------------------------------------------------- #
# Personalization
# --------------------------------------------------------------------------- #

def first_name_of(full_name: str) -> str:
    full_name = (full_name or "").strip()
    return full_name.split()[0] if full_name else ""


def personalize(template: str, name: str, email: str, fallback: str = "there") -> str:
    display_name = name.strip() if name and name.strip() else fallback
    first = first_name_of(name) or fallback
    out = template
    out = out.replace("{name}", display_name)
    out = out.replace("{first_name}", first)
    out = out.replace("{email}", email)
    return out


# --------------------------------------------------------------------------- #
# SMTP sending
# --------------------------------------------------------------------------- #

SMTP_PRESETS = {
    "Gmail": {"host": "smtp.gmail.com", "port": 587, "security": "STARTTLS"},
    "Outlook / Office365": {"host": "smtp.office365.com", "port": 587, "security": "STARTTLS"},
    "Yahoo": {"host": "smtp.mail.yahoo.com", "port": 587, "security": "STARTTLS"},
    "Other": {"host": "", "port": 587, "security": "STARTTLS"},
}


@dataclass
class SendResult:
    email: str
    success: bool
    error: str = ""


class SMTPSession:
    """
    Thin wrapper that keeps one SMTP connection open for a whole batch
    (much faster than reconnecting per email) and transparently reconnects
    once if the server drops the connection mid-batch.
    """

    def __init__(self, host: str, port: int, security: str, username: str, password: str):
        self.host = host
        self.port = port
        self.security = security
        self.username = username
        self.password = password
        self.server = None

    def _connect(self):
        if self.security == "SSL":
            context = ssl.create_default_context()
            self.server = smtplib.SMTP_SSL(self.host, self.port, context=context, timeout=30)
        else:
            self.server = smtplib.SMTP(self.host, self.port, timeout=30)
            if self.security == "STARTTLS":
                self.server.starttls(context=ssl.create_default_context())
        self.server.login(self.username, self.password)

    def __enter__(self):
        self._connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.server is not None:
            try:
                self.server.quit()
            except Exception:
                pass

    def send(
        self,
        sender_email: str,
        sender_name: str,
        to_email: str,
        subject: str,
        body: str,
        is_html: bool = False,
        reply_to: str = "",
    ) -> SendResult:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = formataddr((sender_name, sender_email)) if sender_name else sender_email
        msg["To"] = to_email
        if reply_to:
            msg["Reply-To"] = reply_to

        if is_html:
            msg.set_content("This email requires an HTML-capable mail client.")
            msg.add_alternative(body, subtype="html")
        else:
            msg.set_content(body)

        try:
            self.server.send_message(msg)
            return SendResult(to_email, True)
        except smtplib.SMTPServerDisconnected:
            # try once to reconnect and resend
            try:
                self._connect()
                self.server.send_message(msg)
                return SendResult(to_email, True)
            except Exception as e:
                return SendResult(to_email, False, str(e))
        except Exception as e:
            return SendResult(to_email, False, str(e))


def send_batch(
    smtp_host: str,
    smtp_port: int,
    security: str,
    sender_email: str,
    sender_password: str,
    sender_name: str,
    reply_to: str,
    recipients: list,  # list of dicts: {key, name, email, subject_template, body_template}
    is_html: bool,
    delay_seconds: float,
    progress_callback: Optional[Callable[[int, int, SendResult], None]] = None,
) -> list:
    """Send a batch of personalized emails, returns list of SendResult."""
    results = []
    with SMTPSession(smtp_host, smtp_port, security, sender_email, sender_password) as session:
        total = len(recipients)
        for i, r in enumerate(recipients, start=1):
            subject = personalize(r["subject_template"], r["name"], r["email"])
            body = personalize(r["body_template"], r["name"], r["email"])
            result = session.send(
                sender_email, sender_name, r["email"], subject, body,
                is_html=is_html, reply_to=reply_to,
            )
            results.append(result)
            if progress_callback:
                progress_callback(i, total, result)
            if delay_seconds > 0 and i < total:
                time.sleep(delay_seconds)
    return results
