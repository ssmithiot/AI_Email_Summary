import json
import os
import sqlite3
import threading
import uuid
import webbrowser
from collections import defaultdict
from datetime import datetime, timedelta

import win32com.client
import win32gui
from dotenv import load_dotenv
from flask import Flask, Response, jsonify, render_template, request, stream_with_context
from openai import OpenAI

from rules_engine import CONDITION_LABELS, VALUE_CONDITIONS, apply_rules, load_rules, save_rules, strip_prefixes
from summarize_inbox import build_email_record, format_emails_for_claude, get_outlook_emails, normalize_message_id, normalize_subject

load_dotenv()
app = Flask(__name__)
WATCHING_DB = os.path.join(os.path.dirname(__file__), "watching.db")
GENERIC_SUBJECTS = {"", "(no subject)", "hi", "hello", "thanks", "thank you", "question", "follow up"}

PROMPT_TEMPLATE = """You are reviewing {count} emails from an Outlook inbox, already sorted by a user-defined priority scoring system (highest score = most important to the user).

Each email may carry tags: [URGENT] = Outlook High Importance; [UNREAD]; [HAS-ATTACHMENT]; [PRIORITY-SCORE:N].
The "Rules:" line lists which user rules matched that email - use this to understand why it scored highly.

Return all of these sections in this exact order, even if some are empty:
1. **Urgent / High Priority** - [URGENT] tagged emails first, then the next highest-scoring emails. Include sender and subject.
2. **Quick Stats** - total, unread count, urgent count, how many need action
3. **Action Items** - emails needing a reply or decision, highest priority first, with brief recommended next steps
4. **Key Topics** - group remaining emails by theme (meetings, reports, FYI, etc.)
5. **Can Ignore / Low Priority** - newsletters, notifications, automated mail

Format rules:
- Use markdown headings for each section.
- Use bullet points under every section.
- If a section has nothing meaningful, write a short bullet like `- None flagged`.
- In Action Items, say why the email matters and what the user should do next.
- Do not stop after the first section.

Whenever you mention a specific email, append a citation in the exact format `[Email N]` using that email's number from the inbox list.
Be concise. Use bullets. Respect the priority scoring - emails with higher scores matter more to this user.

--- INBOX EMAILS (sorted by priority score, highest first) ---
{emails}"""


def build_mode_config(args):
    mode = args.get("mode", "quantity")
    if mode == "unread":
        return {"mode": "unread", "label": "unread emails"}
    if mode == "date":
        now = datetime.now()
        range_type = args.get("range", "7days")
        if range_type == "today":
            since = now.replace(hour=0, minute=0, second=0, microsecond=0)
            label = "emails from today"
        elif range_type == "30days":
            since = now - timedelta(days=30)
            label = "emails from the last 30 days"
        elif range_type == "custom":
            date_str = args.get("custom_date", "")
            try:
                since = datetime.strptime(date_str, "%Y-%m-%d")
                label = f"emails since {date_str}"
            except ValueError:
                since = now - timedelta(days=7)
                label = "emails from the last 7 days"
        else:
            since = now - timedelta(days=7)
            label = "emails from the last 7 days"
        return {"mode": "date", "since": since, "label": label}

    try:
        count = max(1, int(args.get("count", 30)))
    except ValueError:
        count = 30
    return {"mode": "quantity", "count": count, "label": f"{count} most recent emails"}


def get_outlook_namespace():
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def get_db_connection():
    conn = sqlite3.connect(WATCHING_DB)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _normalise_address_list(raw_value):
    parts = str(raw_value or "").replace(",", ";").split(";")
    return {part.strip().lower() for part in parts if part.strip()}


def _normalise_identity_tokens(*values):
    tokens = set()
    for value in values:
        text = str(value or "").strip().lower()
        if not text:
            continue
        tokens.add(text)
        for part in text.replace(",", ";").split(";"):
            part = part.strip()
            if part:
                tokens.add(part)
        if "@" not in text:
            compact = " ".join(text.split())
            if compact:
                tokens.add(compact)
    return tokens


def _safe_datetime(value):
    return value if hasattr(value, "strftime") else None


def _bool_flag(value):
    return 1 if value else 0


def _json_dumps(values):
    return json.dumps(list(values or []))


def _json_loads(value):
    if not value:
        return []
    try:
        loaded = json.loads(value)
    except json.JSONDecodeError:
        return []
    return loaded if isinstance(loaded, list) else []


def _placeholder_email(entry_id, sender="Unknown", subject="(No Subject)", received=""):
    return {
        "entry_id": entry_id,
        "from": sender or "Unknown",
        "from_email": "",
        "subject": subject or "(No Subject)",
        "normalized_subject": normalize_subject(subject),
        "received": received or "",
        "received_sort": received or "",
        "unread": False,
        "urgent": False,
        "has_attachment": False,
        "conversation_id": "",
        "conversation_topic": "",
        "internet_message_id": "",
        "in_reply_to": "",
        "references": [],
    }


def get_email_by_entry_id(entry_id):
    ns = get_outlook_namespace()
    msg = ns.GetItemFromID(entry_id)
    return build_email_record(msg)


def _email_has_identifiers(email):
    return bool(email.get("conversation_id") or email.get("internet_message_id") or email.get("in_reply_to") or email.get("references"))


def _within_subject_fallback_window(email_received_sort, thread_received_sort, days=14):
    patterns = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M")

    def _parse(value):
        for pattern in patterns:
            try:
                return datetime.strptime(value, pattern)
            except (TypeError, ValueError):
                continue
        return None

    email_dt = _parse(email_received_sort)
    thread_dt = _parse(thread_received_sort)
    if not email_dt or not thread_dt:
        return False
    return abs((email_dt - thread_dt).days) <= days


def _build_sent_lookup(sent_items, limit=1500):
    sent_index = {}
    for index, msg in enumerate(sent_items):
        if index >= limit:
            break
        try:
            subject_key = strip_prefixes(msg.Subject or "")
            sent_on = _safe_datetime(getattr(msg, "SentOn", None))
            if not subject_key or not sent_on:
                continue
            recipient_tokens = _normalise_identity_tokens(msg.To)
            try:
                for recipient in msg.Recipients:
                    recipient_tokens.update(
                        _normalise_identity_tokens(
                            getattr(recipient, "Address", ""),
                            getattr(recipient, "Name", ""),
                        )
                    )
            except Exception:
                pass
            sent_index.setdefault(subject_key, []).append(
                {
                    "entry_id": getattr(msg, "EntryID", ""),
                    "subject": msg.Subject or "",
                    "sent_on_dt": sent_on,
                    "sent_on": sent_on.strftime("%Y-%m-%d %H:%M"),
                    "to": msg.To or "",
                    "to_set": _normalise_address_list(msg.To),
                    "recipient_tokens": recipient_tokens,
                }
            )
        except Exception:
            continue
    return sent_index


def _find_reply_match(ns, sent_index, entry_id, fallback_subject=""):
    original = ns.GetItemFromID(entry_id)
    subject_key = strip_prefixes(getattr(original, "Subject", "") or fallback_subject)
    received = _safe_datetime(getattr(original, "ReceivedTime", None))
    sender = (getattr(original, "SenderEmailAddress", "") or "").lower()
    sender_name = (getattr(original, "SenderName", "") or "").lower()
    sender_tokens = _normalise_identity_tokens(sender, sender_name)

    for candidate in sent_index.get(subject_key, []):
        if received and candidate["sent_on_dt"] < received:
            continue
        if sender and candidate["to_set"] and sender in candidate["to_set"]:
            return {"entry_id": candidate["entry_id"], "subject": candidate["subject"], "sent_on": candidate["sent_on"], "to": candidate["to"]}
        if sender_tokens and candidate.get("recipient_tokens") and sender_tokens.intersection(candidate["recipient_tokens"]):
            return {"entry_id": candidate["entry_id"], "subject": candidate["subject"], "sent_on": candidate["sent_on"], "to": candidate["to"]}
        if sender and candidate["to_set"]:
            continue
        return {"entry_id": candidate["entry_id"], "subject": candidate["subject"], "sent_on": candidate["sent_on"], "to": candidate["to"]}
    return None


def init_watching_db():
    conn = get_db_connection()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS watched_threads (
                thread_id INTEGER PRIMARY KEY AUTOINCREMENT,
                seed_entry_id TEXT NOT NULL UNIQUE,
                conversation_id TEXT,
                normalized_subject TEXT NOT NULL,
                subject TEXT NOT NULL,
                latest_entry_id TEXT NOT NULL,
                latest_subject TEXT NOT NULL,
                latest_from TEXT NOT NULL,
                latest_received TEXT NOT NULL,
                latest_received_sort TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS watched_messages (
                entry_id TEXT PRIMARY KEY,
                thread_id INTEGER NOT NULL,
                sender TEXT NOT NULL,
                sender_email TEXT,
                subject TEXT NOT NULL,
                normalized_subject TEXT NOT NULL,
                received TEXT NOT NULL,
                received_sort TEXT NOT NULL,
                conversation_id TEXT,
                internet_message_id TEXT,
                in_reply_to TEXT,
                references_json TEXT NOT NULL DEFAULT '[]',
                unread INTEGER NOT NULL DEFAULT 0,
                urgent INTEGER NOT NULL DEFAULT 0,
                has_attachment INTEGER NOT NULL DEFAULT 0,
                is_seed INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(thread_id) REFERENCES watched_threads(thread_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_watched_threads_latest ON watched_threads(latest_received_sort DESC, thread_id DESC);
            CREATE INDEX IF NOT EXISTS idx_watched_threads_subject ON watched_threads(normalized_subject);
            CREATE INDEX IF NOT EXISTS idx_watched_threads_conversation ON watched_threads(conversation_id);
            CREATE INDEX IF NOT EXISTS idx_watched_messages_thread ON watched_messages(thread_id, received_sort DESC);
            CREATE INDEX IF NOT EXISTS idx_watched_messages_message_id ON watched_messages(internet_message_id);
            CREATE INDEX IF NOT EXISTS idx_watched_messages_conversation ON watched_messages(conversation_id);
            CREATE INDEX IF NOT EXISTS idx_watched_messages_subject ON watched_messages(normalized_subject);
            """
        )
        _migrate_legacy_watching(conn)
        conn.commit()
    finally:
        conn.close()


def _migrate_legacy_watching(conn):
    table_exists = conn.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'watching_items'").fetchone()
    if not table_exists:
        return

    has_threads = conn.execute("SELECT 1 FROM watched_threads LIMIT 1").fetchone()
    if has_threads:
        return

    rows = conn.execute("SELECT entry_id, sender, subject, received FROM watching_items ORDER BY order_index ASC, rowid ASC").fetchall()
    if not rows:
        return

    state = _load_watching_state(conn)
    for row in rows:
        try:
            email = get_email_by_entry_id(row["entry_id"])
        except Exception:
            email = _placeholder_email(row["entry_id"], sender=row["sender"], subject=row["subject"], received=row["received"])
        _subscribe_email_to_watching(conn, state, email)

def _load_watching_state(conn):
    threads = {}
    conversation_map = defaultdict(set)
    message_id_map = defaultdict(set)
    subject_map = defaultdict(set)

    thread_rows = conn.execute(
        """
        SELECT thread_id, seed_entry_id, conversation_id, normalized_subject, subject,
               latest_entry_id, latest_subject, latest_from, latest_received,
               latest_received_sort
        FROM watched_threads
        """
    ).fetchall()
    for row in thread_rows:
        thread_id = row["thread_id"]
        thread = {
            "thread_id": thread_id,
            "seed_entry_id": row["seed_entry_id"],
            "conversation_id": row["conversation_id"] or "",
            "normalized_subject": row["normalized_subject"] or "",
            "subject": row["subject"] or "(No Subject)",
            "latest_entry_id": row["latest_entry_id"],
            "latest_subject": row["latest_subject"],
            "latest_from": row["latest_from"],
            "latest_received": row["latest_received"],
            "latest_received_sort": row["latest_received_sort"],
            "participant_emails": set(),
            "message_ids": set(),
            "has_identifier": bool(row["conversation_id"]),
            "message_count": 0,
        }
        threads[thread_id] = thread
        if thread["conversation_id"]:
            conversation_map[thread["conversation_id"]].add(thread_id)
        if thread["normalized_subject"]:
            subject_map[thread["normalized_subject"]].add(thread_id)

    message_rows = conn.execute(
        """
        SELECT thread_id, sender_email, normalized_subject, conversation_id,
               internet_message_id, in_reply_to, references_json
        FROM watched_messages
        """
    ).fetchall()
    for row in message_rows:
        thread = threads.get(row["thread_id"])
        if not thread:
            continue
        thread["message_count"] += 1
        if row["sender_email"]:
            thread["participant_emails"].add(row["sender_email"].lower())
        if row["conversation_id"]:
            conversation_map[row["conversation_id"]].add(row["thread_id"])
            thread["has_identifier"] = True
        if row["internet_message_id"]:
            thread["message_ids"].add(row["internet_message_id"])
            message_id_map[row["internet_message_id"]].add(row["thread_id"])
            thread["has_identifier"] = True
        if row["in_reply_to"] or _json_loads(row["references_json"]):
            thread["has_identifier"] = True
        if row["normalized_subject"]:
            subject_map[row["normalized_subject"]].add(row["thread_id"])

    return {
        "threads": threads,
        "conversation_map": conversation_map,
        "message_id_map": message_id_map,
        "subject_map": subject_map,
    }


def _register_email_in_state(state, thread_id, email, is_seed=False):
    thread = state["threads"][thread_id]
    conversation_id = email.get("conversation_id") or ""
    message_id = normalize_message_id(email.get("internet_message_id"))
    normalized = email.get("normalized_subject") or normalize_subject(email.get("subject"))
    sender_email = (email.get("from_email") or "").lower()

    if conversation_id:
        state["conversation_map"][conversation_id].add(thread_id)
        thread["conversation_id"] = thread["conversation_id"] or conversation_id
    if normalized:
        state["subject_map"][normalized].add(thread_id)
        thread["normalized_subject"] = normalized
    if sender_email:
        thread["participant_emails"].add(sender_email)
    if message_id:
        state["message_id_map"][message_id].add(thread_id)
        thread["message_ids"].add(message_id)

    thread["has_identifier"] = thread["has_identifier"] or _email_has_identifiers(email)

    received_sort = email.get("received_sort") or email.get("received") or ""
    if received_sort >= (thread.get("latest_received_sort") or ""):
        thread["latest_entry_id"] = email.get("entry_id") or thread["latest_entry_id"]
        thread["latest_subject"] = email.get("subject") or thread["latest_subject"]
        thread["latest_from"] = email.get("from") or thread["latest_from"]
        thread["latest_received"] = email.get("received") or thread["latest_received"]
        thread["latest_received_sort"] = received_sort
    if is_seed:
        thread["seed_entry_id"] = email.get("entry_id") or thread["seed_entry_id"]
        thread["subject"] = email.get("subject") or thread["subject"]


def _create_thread(conn, state, email):
    conn.execute(
        """
        INSERT INTO watched_threads (
            seed_entry_id, conversation_id, normalized_subject, subject,
            latest_entry_id, latest_subject, latest_from, latest_received,
            latest_received_sort
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            email.get("entry_id"),
            email.get("conversation_id") or None,
            email.get("normalized_subject") or normalize_subject(email.get("subject")),
            email.get("subject") or "(No Subject)",
            email.get("entry_id"),
            email.get("subject") or "(No Subject)",
            email.get("from") or "Unknown",
            email.get("received") or "",
            email.get("received_sort") or email.get("received") or "",
        ),
    )
    thread_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    state["threads"][thread_id] = {
        "thread_id": thread_id,
        "seed_entry_id": email.get("entry_id"),
        "conversation_id": "",
        "normalized_subject": email.get("normalized_subject") or normalize_subject(email.get("subject")),
        "subject": email.get("subject") or "(No Subject)",
        "latest_entry_id": email.get("entry_id"),
        "latest_subject": email.get("subject") or "(No Subject)",
        "latest_from": email.get("from") or "Unknown",
        "latest_received": email.get("received") or "",
        "latest_received_sort": email.get("received_sort") or email.get("received") or "",
        "participant_emails": set(),
        "message_ids": set(),
        "has_identifier": False,
        "message_count": 0,
    }
    _register_email_in_state(state, thread_id, email, is_seed=True)
    return thread_id


def _refresh_thread_row(conn, state, thread_id):
    thread = state["threads"][thread_id]
    latest_message = conn.execute(
        """
        SELECT entry_id, sender, subject, received, received_sort
        FROM watched_messages
        WHERE thread_id = ?
        ORDER BY received_sort DESC, rowid DESC
        LIMIT 1
        """,
        (thread_id,),
    ).fetchone()

    if latest_message:
        thread["latest_entry_id"] = latest_message["entry_id"]
        thread["latest_subject"] = latest_message["subject"]
        thread["latest_from"] = latest_message["sender"]
        thread["latest_received"] = latest_message["received"]
        thread["latest_received_sort"] = latest_message["received_sort"]

    conn.execute(
        """
        UPDATE watched_threads
        SET seed_entry_id = ?,
            conversation_id = ?,
            normalized_subject = ?,
            subject = ?,
            latest_entry_id = ?,
            latest_subject = ?,
            latest_from = ?,
            latest_received = ?,
            latest_received_sort = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE thread_id = ?
        """,
        (
            thread.get("seed_entry_id") or thread.get("latest_entry_id"),
            thread.get("conversation_id") or None,
            thread.get("normalized_subject") or "",
            thread.get("subject") or thread.get("latest_subject") or "(No Subject)",
            thread.get("latest_entry_id") or thread.get("seed_entry_id"),
            thread.get("latest_subject") or thread.get("subject") or "(No Subject)",
            thread.get("latest_from") or "Unknown",
            thread.get("latest_received") or "",
            thread.get("latest_received_sort") or "",
            thread_id,
        ),
    )


def _upsert_message(conn, state, thread_id, email, is_seed=False):
    entry_id = email.get("entry_id")
    if not entry_id:
        return False

    existing = conn.execute("SELECT thread_id FROM watched_messages WHERE entry_id = ?", (entry_id,)).fetchone()
    conn.execute(
        """
        INSERT INTO watched_messages (
            entry_id, thread_id, sender, sender_email, subject, normalized_subject,
            received, received_sort, conversation_id, internet_message_id,
            in_reply_to, references_json, unread, urgent, has_attachment, is_seed,
            updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(entry_id) DO UPDATE SET
            thread_id = excluded.thread_id,
            sender = excluded.sender,
            sender_email = excluded.sender_email,
            subject = excluded.subject,
            normalized_subject = excluded.normalized_subject,
            received = excluded.received,
            received_sort = excluded.received_sort,
            conversation_id = excluded.conversation_id,
            internet_message_id = excluded.internet_message_id,
            in_reply_to = excluded.in_reply_to,
            references_json = excluded.references_json,
            unread = excluded.unread,
            urgent = excluded.urgent,
            has_attachment = excluded.has_attachment,
            is_seed = CASE WHEN watched_messages.is_seed = 1 THEN 1 ELSE excluded.is_seed END,
            updated_at = CURRENT_TIMESTAMP
        """,
        (
            entry_id,
            thread_id,
            email.get("from") or "Unknown",
            (email.get("from_email") or "").lower(),
            email.get("subject") or "(No Subject)",
            email.get("normalized_subject") or normalize_subject(email.get("subject")),
            email.get("received") or "",
            email.get("received_sort") or email.get("received") or "",
            email.get("conversation_id") or None,
            normalize_message_id(email.get("internet_message_id")) or None,
            normalize_message_id(email.get("in_reply_to")) or None,
            _json_dumps(email.get("references") or []),
            _bool_flag(email.get("unread")),
            _bool_flag(email.get("urgent")),
            _bool_flag(email.get("has_attachment")),
            _bool_flag(is_seed),
        ),
    )
    if not existing:
        state["threads"][thread_id]["message_count"] += 1
    _register_email_in_state(state, thread_id, email, is_seed=is_seed)
    _refresh_thread_row(conn, state, thread_id)
    return existing is None or existing["thread_id"] != thread_id

def _match_thread_for_email(email, state):
    conversation_id = email.get("conversation_id") or ""
    if conversation_id:
        candidates = state["conversation_map"].get(conversation_id, set())
        if len(candidates) == 1:
            return next(iter(candidates))
        if len(candidates) > 1:
            return None

    in_reply_to = normalize_message_id(email.get("in_reply_to"))
    if in_reply_to:
        candidates = state["message_id_map"].get(in_reply_to, set())
        if len(candidates) == 1:
            return next(iter(candidates))
        if len(candidates) > 1:
            return None

    reference_candidates = set()
    for reference in email.get("references") or []:
        normalized_ref = normalize_message_id(reference)
        if normalized_ref:
            reference_candidates.update(state["message_id_map"].get(normalized_ref, set()))
    if len(reference_candidates) == 1:
        return next(iter(reference_candidates))
    if len(reference_candidates) > 1:
        return None

    normalized = email.get("normalized_subject") or normalize_subject(email.get("subject"))
    if not normalized or normalized in GENERIC_SUBJECTS or len(normalized) < 8:
        return None

    candidates = state["subject_map"].get(normalized, set())
    if len(candidates) != 1:
        return None

    thread_id = next(iter(candidates))
    thread = state["threads"].get(thread_id)
    if not thread:
        return None
    if not _within_subject_fallback_window(email.get("received_sort"), thread.get("latest_received_sort")):
        return None

    sender_email = (email.get("from_email") or "").lower()
    if sender_email and sender_email in thread["participant_emails"]:
        return thread_id
    if not _email_has_identifiers(email) and not thread["has_identifier"]:
        return thread_id
    return None


def _subscribe_email_to_watching(conn, state, email):
    entry_id = email.get("entry_id")
    if not entry_id:
        return None, False

    existing = conn.execute("SELECT thread_id FROM watched_messages WHERE entry_id = ?", (entry_id,)).fetchone()
    if existing:
        thread_id = existing["thread_id"]
        _upsert_message(conn, state, thread_id, email)
        return thread_id, False

    thread_id = _match_thread_for_email(email, state)
    created = False
    if thread_id is None:
        thread_id = _create_thread(conn, state, email)
        created = True

    _upsert_message(conn, state, thread_id, email, is_seed=created)
    return thread_id, created


def refresh_watched_threads(scanned_emails):
    conn = get_db_connection()
    try:
        state = _load_watching_state(conn)
        changed = False
        for email in scanned_emails:
            thread_id = _match_thread_for_email(email, state)
            if thread_id is None:
                continue
            changed = _upsert_message(conn, state, thread_id, email) or changed
        if changed:
            conn.commit()
        return changed
    finally:
        conn.close()


def list_watching_threads():
    conn = get_db_connection()
    try:
        thread_rows = conn.execute(
            """
            SELECT thread_id, seed_entry_id, conversation_id, normalized_subject, subject,
                   latest_entry_id, latest_subject, latest_from, latest_received,
                   latest_received_sort
            FROM watched_threads
            ORDER BY latest_received_sort DESC, thread_id DESC
            """
        ).fetchall()
        message_rows = conn.execute(
            """
            SELECT entry_id, thread_id, sender, subject, received, received_sort,
                   unread, urgent, has_attachment
            FROM watched_messages
            ORDER BY received_sort DESC, rowid DESC
            """
        ).fetchall()
    finally:
        conn.close()

    messages_by_thread = defaultdict(list)
    for row in message_rows:
        messages_by_thread[row["thread_id"]].append(
            {
                "entry_id": row["entry_id"],
                "from": row["sender"],
                "subject": row["subject"],
                "received": row["received"],
                "received_sort": row["received_sort"],
                "unread": bool(row["unread"]),
                "urgent": bool(row["urgent"]),
                "has_attachment": bool(row["has_attachment"]),
            }
        )

    threads = []
    for row in thread_rows:
        messages = messages_by_thread.get(row["thread_id"], [])
        threads.append(
            {
                "thread_id": row["thread_id"],
                "seed_entry_id": row["seed_entry_id"],
                "conversation_id": row["conversation_id"] or "",
                "normalized_subject": row["normalized_subject"] or "",
                "subject": row["subject"] or "(No Subject)",
                "latest_entry_id": row["latest_entry_id"],
                "latest_subject": row["latest_subject"] or row["subject"] or "(No Subject)",
                "latest_from": row["latest_from"] or "Unknown",
                "latest_received": row["latest_received"] or "",
                "latest_received_sort": row["latest_received_sort"] or "",
                "message_count": len(messages),
                "messages": messages,
            }
        )
    return threads


def remove_watching_thread(thread_id):
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM watched_threads WHERE thread_id = ?", (thread_id,))
        conn.commit()
    finally:
        conn.close()


def remove_watching_by_entry(entry_id):
    conn = get_db_connection()
    try:
        row = conn.execute("SELECT thread_id FROM watched_threads WHERE seed_entry_id = ?", (entry_id,)).fetchone()
        if row:
            conn.execute("DELETE FROM watched_threads WHERE thread_id = ?", (row["thread_id"],))
            conn.commit()
    finally:
        conn.close()


def clear_watching_items():
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM watched_threads")
        conn.commit()
    finally:
        conn.close()


init_watching_db()


def _extract_response_text(response):
    text = getattr(response, "output_text", "") or ""
    if text:
        return text

    parts = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            if getattr(content, "type", "") in {"output_text", "text"}:
                value = getattr(content, "text", "") or getattr(content, "value", "")
                if value:
                    parts.append(value)
    return "".join(parts)


def _build_summary_email_payload(emails, compact=False):
    prepared = []
    for email in emails:
        clone = dict(email)
        body = clone.get("body", "") or ""
        if compact:
            max_preview = 140
        elif len(emails) >= 30:
            max_preview = 220
        elif len(emails) >= 20:
            max_preview = 320
        else:
            max_preview = 450
        if len(body) > max_preview:
            clone["body"] = body[:max_preview].rstrip() + "..."
        prepared.append(clone)
    return prepared


def _request_summary_text(client, prompt):
    response = client.responses.create(model="gpt-5", input=prompt, max_output_tokens=3000)
    return _extract_response_text(response).strip()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/summarize")
def summarize():
    def generate():
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            yield f"data: {json.dumps({'type': 'error', 'text': 'OPENAI_API_KEY not set in .env file.'})}\n\n"
            return

        mode_config = build_mode_config(request.args)

        try:
            status = json.dumps({"type": "status", "text": f"Connecting to Outlook ({mode_config['label']})..."})
            yield f"data: {status}\n\n"
            emails = get_outlook_emails(mode_config)
        except RuntimeError as exc:
            yield f"data: {json.dumps({'type': 'error', 'text': str(exc)})}\n\n"
            return

        if not emails:
            yield f"data: {json.dumps({'type': 'error', 'text': 'No emails found for that selection.'})}\n\n"
            return

        emails = apply_rules(emails, load_rules())
        refresh_watched_threads(emails)

        found_msg = json.dumps({"type": "status", "text": f"Found {len(emails)} emails. Summarizing with OpenAI..."})
        yield f"data: {found_msg}\n\n"

        client = OpenAI(api_key=api_key)
        try:
            email_text = format_emails_for_claude(_build_summary_email_payload(emails, compact=False))
            prompt = PROMPT_TEMPLATE.format(count=len(emails), emails=email_text)
            summary_text = _request_summary_text(client, prompt)
            if not summary_text:
                compact_text = format_emails_for_claude(_build_summary_email_payload(emails, compact=True))
                compact_prompt = PROMPT_TEMPLATE.format(count=len(emails), emails=compact_text)
                summary_text = _request_summary_text(client, compact_prompt)
            if not summary_text:
                yield f"data: {json.dumps({'type': 'error', 'text': 'OpenAI returned no summary text.'})}\n\n"
                return
            yield f"data: {json.dumps({'type': 'text', 'text': summary_text})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'text': f'OpenAI API error: {exc}'})}\n\n"
            return

        card_data = [
            {
                "index": email["index"],
                "entry_id": email["entry_id"],
                "from": email["from"],
                "from_email": email.get("from_email", ""),
                "subject": email["subject"],
                "normalized_subject": email.get("normalized_subject", ""),
                "received": email["received"],
                "received_sort": email.get("received_sort", email["received"]),
                "unread": email["unread"],
                "urgent": email["urgent"],
                "has_attachment": email.get("has_attachment", False),
                "rule_score": email.get("rule_score", 0),
                "matched_rules": email.get("matched_rules", []),
                "conversation_id": email.get("conversation_id", ""),
                "internet_message_id": email.get("internet_message_id", ""),
                "in_reply_to": email.get("in_reply_to", ""),
                "references": email.get("references", []),
            }
            for email in emails
        ]
        yield f"data: {json.dumps({'type': 'emails', 'emails': card_data})}\n\n"
        yield f"data: {json.dumps({'type': 'done'})}\n\n"

    return Response(stream_with_context(generate()), mimetype="text/event-stream", headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/rules", methods=["GET"])
def get_rules():
    config = load_rules()
    config["_condition_labels"] = CONDITION_LABELS
    config["_value_conditions"] = list(VALUE_CONDITIONS)
    return jsonify(config)


@app.route("/rules", methods=["POST"])
def post_rules():
    data = request.json
    config = load_rules()
    if "my_email" in data:
        config["my_email"] = data["my_email"].strip().lower()
    if "my_name" in data:
        config["my_name"] = data["my_name"].strip()
    if "rules" in data:
        config["rules"] = data["rules"]
    if "custom_rules" in data:
        config["custom_rules"] = data["custom_rules"]
    save_rules(config)
    return jsonify({"ok": True})


@app.route("/rules/custom", methods=["POST"])
def add_custom_rule():
    data = request.json
    config = load_rules()
    new_rule = {
        "id": str(uuid.uuid4())[:8],
        "name": data.get("name", "Custom rule"),
        "description": data.get("description", ""),
        "enabled": True,
        "score": int(data.get("score", 50)),
        "builtin": False,
        "conditions": data.get("conditions", []),
    }
    config.setdefault("custom_rules", []).append(new_rule)
    save_rules(config)
    return jsonify({"ok": True, "rule": new_rule})


@app.route("/rules/custom/<rule_id>", methods=["DELETE"])
def delete_custom_rule(rule_id):
    config = load_rules()
    config["custom_rules"] = [rule for rule in config.get("custom_rules", []) if rule["id"] != rule_id]
    save_rules(config)
    return jsonify({"ok": True})


@app.route("/watching", methods=["GET"])
def get_watching():
    threads = list_watching_threads()
    return jsonify({"ok": True, "threads": threads, "items": threads})


@app.route("/watching", methods=["POST"])
def add_watching():
    item = request.json or {}
    entry_id = item.get("entry_id")
    if not entry_id:
        return jsonify({"ok": False, "error": "Missing entry_id"}), 400

    email = {
        "entry_id": entry_id,
        "from": item.get("from") or "Unknown",
        "from_email": item.get("from_email") or "",
        "subject": item.get("subject") or "(No Subject)",
        "normalized_subject": item.get("normalized_subject") or normalize_subject(item.get("subject")),
        "received": item.get("received") or "",
        "received_sort": item.get("received_sort") or item.get("received") or "",
        "unread": bool(item.get("unread", False)),
        "urgent": bool(item.get("urgent", False)),
        "has_attachment": bool(item.get("has_attachment", False)),
        "conversation_id": item.get("conversation_id") or "",
        "conversation_topic": item.get("conversation_topic") or "",
        "internet_message_id": normalize_message_id(item.get("internet_message_id")) or "",
        "in_reply_to": normalize_message_id(item.get("in_reply_to")) or "",
        "references": [normalize_message_id(ref) for ref in item.get("references") or [] if normalize_message_id(ref)],
    }
    if not _email_has_identifiers(email):
        try:
            email = get_email_by_entry_id(entry_id)
        except Exception:
            pass

    conn = get_db_connection()
    try:
        state = _load_watching_state(conn)
        thread_id, created = _subscribe_email_to_watching(conn, state, email)
        conn.commit()
    finally:
        conn.close()

    threads = list_watching_threads()
    message = "Subscribed to watched thread" if created else "Updated watched thread"
    return jsonify({"ok": True, "thread_id": thread_id, "created": created, "message": message, "threads": threads, "items": threads})


@app.route("/watching/thread/<int:thread_id>", methods=["DELETE"])
def delete_watching_thread(thread_id):
    remove_watching_thread(thread_id)
    threads = list_watching_threads()
    return jsonify({"ok": True, "threads": threads, "items": threads})


@app.route("/watching/<path:entry_id>", methods=["DELETE"])
def delete_watching_legacy(entry_id):
    remove_watching_by_entry(entry_id)
    threads = list_watching_threads()
    return jsonify({"ok": True, "threads": threads, "items": threads})


@app.route("/watching/clear", methods=["POST"])
def clear_watching():
    clear_watching_items()
    return jsonify({"ok": True, "threads": [], "items": []})


@app.route("/check-replied", methods=["POST"])
def check_replied():
    items = request.json.get("emails", [])
    try:
        ns = get_outlook_namespace()
        sent = ns.GetDefaultFolder(5)
        sent_items = sent.Items
        sent_items.Sort("[SentOn]", True)
        sent_index = _build_sent_lookup(sent_items)

        results = {}
        for item in items:
            entry_id = item.get("entry_id")
            if not entry_id:
                continue
            try:
                results[entry_id] = _find_reply_match(ns, sent_index, entry_id, item.get("subject", ""))
            except Exception:
                results[entry_id] = None

        return jsonify({"ok": True, "results": results})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/flag-email", methods=["POST"])
def flag_email():
    entry_id = request.json.get("entry_id")
    try:
        ns = get_outlook_namespace()
        msg = ns.GetItemFromID(entry_id)
        msg.FlagStatus = 1
        msg.Save()
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/mark-email-read", methods=["POST"])
def mark_email_read():
    entry_id = request.json.get("entry_id")
    try:
        ns = get_outlook_namespace()
        msg = ns.GetItemFromID(entry_id)
        msg.UnRead = False
        msg.Save()
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/open-email", methods=["POST"])
def open_email():
    entry_id = request.json.get("entry_id")
    if not entry_id:
        return jsonify({"ok": False, "error": "Missing entry_id"}), 400

    try:
        ns = get_outlook_namespace()
        msg = ns.GetItemFromID(entry_id)
        msg.Display()
        inspector = msg.GetInspector
        inspector.Activate()

        caption = getattr(inspector, "Caption", "") or getattr(msg, "Subject", "")
        if caption:
            def enum_handler(hwnd, matches):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = win32gui.GetWindowText(hwnd)
                if title == caption or caption in title:
                    matches.append(hwnd)

            matches = []
            win32gui.EnumWindows(enum_handler, matches)
            if matches:
                win32gui.SetForegroundWindow(matches[0])
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/mark-read", methods=["POST"])
def mark_all_read():
    try:
        ns = get_outlook_namespace()
        inbox = ns.GetDefaultFolder(6)
        count = 0
        for msg in inbox.Items:
            try:
                if msg.UnRead:
                    msg.UnRead = False
                    count += 1
            except Exception:
                continue
        return jsonify({"ok": True, "count": count})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/my-identity")
def my_identity():
    try:
        ns = get_outlook_namespace()
        user = ns.CurrentUser
        return jsonify({"ok": True, "email": (user.Address or "").lower(), "name": user.Name or ""})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


def open_browser():
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    threading.Timer(1.0, open_browser).start()
    app.run(debug=False, port=5000)
