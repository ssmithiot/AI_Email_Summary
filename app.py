import json
import os
import queue
import sqlite3
import threading
import uuid
import webbrowser
from collections import defaultdict
from datetime import datetime, timedelta

import pywintypes
import win32com.client
import win32gui
from dotenv import load_dotenv
from flask import Flask, Response, jsonify, render_template, request, stream_with_context
from openai import OpenAI

from rules_engine import CONDITION_LABELS, VALUE_CONDITIONS, apply_rules, load_rules, save_rules, strip_prefixes
from summarize_inbox import (
    build_email_record,
    format_emails_for_claude,
    get_outlook_emails,
    get_outlook_namespace,
    normalize_message_id,
    normalize_subject,
)

BASE_DIR = os.path.dirname(__file__)
load_dotenv(os.path.join(BASE_DIR, ".env"))
app = Flask(__name__)
WATCHING_DB = os.path.join(BASE_DIR, "watching.db")
GENERIC_SUBJECTS = {"", "(no subject)", "hi", "hello", "thanks", "thank you", "question", "follow up"}

PROMPT_TEMPLATE = """You are reviewing {count} emails from an Outlook inbox, already sorted by a user-defined priority scoring system (highest score = most important to the user).

Each email may carry tags: [URGENT] = Outlook High Importance; [UNREAD]; [HAS-ATTACHMENT]; [PRIORITY-SCORE:N].
The "Rules:" line lists which user rules matched that email - use this to understand why it scored highly.
Some emails belong to the same conversation thread. When that happens, treat the thread as one ongoing conversation instead of repeating the same point for each message.

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
- In Action Items, give plain-English advice about what the user should do next.
- If an unread email looks like a real human message, assume it deserves review unless the content clearly looks disposable.
- Avoid robotic wording like "None flagged" when there are unread human emails that probably need a glance.
- If several messages are clearly from the same thread, mention the thread once and note how many recent messages it contains.
- Do not stop after the first section.

Whenever you mention a specific email, append a citation in the exact format `[Email N]` using that email's number from the inbox list.
Be concise. Use bullets. Respect the priority scoring - emails with higher scores matter more to this user. Sound like a sharp executive assistant, not a log parser.

--- INBOX EMAILS (sorted by priority score, highest first) ---
{emails}"""


def _thread_key(email):
    conversation_id = (email.get("conversation_id") or "").strip()
    if conversation_id:
        return f"cid:{conversation_id}"

    normalized_subject = email.get("normalized_subject") or normalize_subject(email.get("subject"))
    sender = (email.get("from_email") or email.get("from") or "").strip().lower()
    if normalized_subject and sender:
        return f"subj:{normalized_subject}|sender:{sender}"
    if normalized_subject:
        return f"subj:{normalized_subject}"
    return f"entry:{email.get('entry_id')}"


def _annotate_threads(emails):
    grouped = defaultdict(list)
    for email in emails:
        grouped[_thread_key(email)].append(email)

    for items in grouped.values():
        sorted_items = sorted(
            items,
            key=lambda item: (item.get("received_sort") or "", item.get("index", 0)),
            reverse=True,
        )
        latest = sorted_items[0]
        count = len(items)
        for position, item in enumerate(sorted_items, start=1):
            item["thread_message_count"] = count
            item["thread_is_primary"] = item["entry_id"] == latest.get("entry_id")
            item["thread_latest_subject"] = latest.get("subject") or item.get("subject") or "(No Subject)"
            item["thread_latest_received"] = latest.get("received") or item.get("received") or ""
            item["thread_position"] = position
    return emails


def _display_emails(emails):
    primary_emails = [email for email in emails if email.get("thread_is_primary", True)]
    grouped = defaultdict(list)
    for email in primary_emails:
        normalized_subject = email.get("normalized_subject") or normalize_subject(email.get("subject"))
        sender = (email.get("from_email") or email.get("from") or "").strip().lower()
        if normalized_subject and normalized_subject not in GENERIC_SUBJECTS:
            grouped[f"{normalized_subject}|{sender}"].append(email)
        else:
            grouped[f"entry:{email.get('entry_id')}"].append(email)

    deduped = []
    for items in grouped.values():
        best = sorted(
            items,
            key=lambda item: (
                item.get("rule_score", 0),
                bool(item.get("unread")),
                item.get("received_sort") or "",
                item.get("index", 0),
            ),
            reverse=True,
        )[0]
        deduped.append(dict(best))

    deduped.sort(key=lambda item: item.get("index", 0))
    for new_index, email in enumerate(deduped, start=1):
        email["original_index"] = email.get("index", new_index)
        email["index"] = new_index
    return deduped


def _build_local_summary(emails):
    urgent = [email for email in emails if email.get("urgent")]
    high_priority = [email for email in emails if email.get("rule_score", 0) > 0][:5]
    unread_count = sum(1 for email in emails if email.get("unread"))
    action_items = [
        email for email in emails
        if email.get("unread") or email.get("rule_score", 0) >= 50 or email.get("urgent")
    ][:5]

    sender_counts = {}
    for email in emails:
        sender = email.get("from") or "Unknown"
        sender_counts[sender] = sender_counts.get(sender, 0) + 1
    top_senders = sorted(sender_counts.items(), key=lambda item: item[1], reverse=True)[:3]

    lines = ["## Urgent / High Priority"]
    if urgent:
        for email in urgent[:5]:
            thread_note = f" ({email['thread_message_count']} recent messages)" if email.get("thread_message_count", 1) > 1 else ""
            lines.append(f"- {email['from']} - {email['subject']}{thread_note} [Email {email['index']}]")
    elif high_priority:
        for email in high_priority:
            thread_note = f" ({email['thread_message_count']} recent messages)" if email.get("thread_message_count", 1) > 1 else ""
            lines.append(f"- {email['from']} - {email['subject']}{thread_note} [Email {email['index']}]")
    else:
        lines.append("- None flagged")

    lines.append("")
    lines.append("## Quick Stats")
    lines.append(f"- Total emails: {len(emails)}")
    lines.append(f"- Unread emails: {unread_count}")
    lines.append(f"- Urgent emails: {len(urgent)}")
    lines.append(f"- Needs attention: {len(action_items)}")

    lines.append("")
    lines.append("## Action Items")
    if action_items:
        for email in action_items:
            reason = "urgent" if email.get("urgent") else "unread" if email.get("unread") else "high priority"
            if email.get("thread_message_count", 1) > 1:
                lines.append(
                    f"- Review the {email['thread_message_count']}-message thread about \"{email['thread_latest_subject']}\" from {email['from']}. "
                    f"It looks {reason} and probably needs a quick human check. [Email {email['index']}]"
                )
            else:
                lines.append(
                    f"- Review \"{email['subject']}\" from {email['from']}. It looks {reason} and likely needs a quick human check. [Email {email['index']}]"
                )
    else:
        lines.append("- None flagged")

    lines.append("")
    lines.append("## Key Topics")
    if top_senders:
        for sender, count in top_senders:
            lines.append(f"- {sender}: {count} email{'s' if count != 1 else ''}")
    else:
        lines.append("- None flagged")

    lines.append("")
    lines.append("## Can Ignore / Low Priority")
    low_priority = [email for email in emails if not email.get("urgent") and email.get("rule_score", 0) == 0][:5]
    if low_priority:
        for email in low_priority:
            thread_note = f" ({email['thread_message_count']} in thread)" if email.get("thread_message_count", 1) > 1 else ""
            lines.append(f"- {email['subject']} from {email['from']}{thread_note} [Email {email['index']}]")
    else:
        lines.append("- None flagged")

    return "\n".join(lines)


def build_mode_config(args):
    mode = args.get("mode", "quantity")
    include_subfolders = str(args.get("include_subfolders", "")).lower() in {"1", "true", "yes", "on"}
    include_all_inboxes = str(args.get("include_all_inboxes", "")).lower() in {"1", "true", "yes", "on"}
    try:
        scan_cap = max(25, min(1000, int(args.get("scan_cap", 100))))
    except ValueError:
        scan_cap = 100
    filter_text = (args.get("filter_text", "") or "").strip()
    filter_from = (args.get("filter_from", "") or "").strip()
    filter_to = (args.get("filter_to", "") or "").strip()
    filter_subject = (args.get("filter_subject", "") or "").strip()
    unread_only = str(args.get("filter_unread", "")).lower() in {"1", "true", "yes", "on"}
    attachments_only = str(args.get("filter_attach", "")).lower() in {"1", "true", "yes", "on"}

    def apply_common_filters(label):
        if include_all_inboxes:
            label += " across all account inboxes"
        if filter_from:
            label += f" from '{filter_from}'"
        if filter_to:
            label += f" to '{filter_to}'"
        if filter_subject:
            label += f" subject '{filter_subject}'"
        if filter_text:
            label += f" keyword '{filter_text}'"
        if unread_only:
            label += " (unread only)"
        if attachments_only:
            label += " (attachments only)"
        return label

    def scanning_status(label):
        if mode == "date":
            return f"Connecting to Outlook ({label})..."
        return f"Connecting to Outlook ({label}). Scanning up to {scan_cap} recent email{'s' if scan_cap != 1 else ''} for matches..."

    if mode == "unread":
        label = "unread emails"
        if include_subfolders:
            label += " including Inbox subfolders"
        return {
            "mode": "unread",
            "label": apply_common_filters(label),
            "status_text": scanning_status(apply_common_filters(label)),
            "include_subfolders": include_subfolders,
            "include_all_inboxes": include_all_inboxes,
            "filter_text": filter_text,
            "filter_from": filter_from,
            "filter_to": filter_to,
            "filter_subject": filter_subject,
            "filter_unread": unread_only,
            "filter_attach": attachments_only,
            "scan_cap": scan_cap,
        }
    if mode == "date":
        now = datetime.now()
        range_type = args.get("range", "7days")
        if range_type == "today":
            since = now.replace(hour=0, minute=0, second=0, microsecond=0)
            until = None
            label = "emails from today"
        elif range_type == "yesterday":
            since = (now - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
            until = now.replace(hour=0, minute=0, second=0, microsecond=0)
            label = "emails from yesterday"
        elif range_type == "3days":
            since = now - timedelta(days=3)
            until = None
            label = "emails from the last 3 days"
        elif range_type == "custom":
            start_str = args.get("custom_start", "")
            end_str = args.get("custom_end", "")
            try:
                since = datetime.strptime(start_str, "%Y-%m-%d")
                until = datetime.strptime(end_str, "%Y-%m-%d") + timedelta(days=1)
                if until <= since:
                    raise ValueError
                label = f"emails from {start_str} to {end_str}"
            except ValueError:
                since = now - timedelta(days=7)
                until = None
                label = "emails from the last 7 days"
        else:
            since = now - timedelta(days=7)
            until = None
            label = "emails from the last 7 days"
        return {
            "mode": "date",
            "since": since,
            "until": until,
            "label": apply_common_filters(label),
            "status_text": scanning_status(apply_common_filters(label)),
            "include_subfolders": include_subfolders,
            "include_all_inboxes": include_all_inboxes,
            "filter_text": filter_text,
            "filter_from": filter_from,
            "filter_to": filter_to,
            "filter_subject": filter_subject,
            "filter_unread": unread_only,
            "filter_attach": attachments_only,
            "scan_cap": scan_cap,
        }

    try:
        count = max(1, int(args.get("count", 30)))
    except ValueError:
        count = 30
    label = f"{count} most recent emails"
    if include_subfolders:
        label += " including Inbox subfolders"
    return {
        "mode": "quantity",
        "count": count,
        "label": apply_common_filters(label),
        "status_text": scanning_status(apply_common_filters(label)),
        "include_subfolders": include_subfolders,
        "include_all_inboxes": include_all_inboxes,
        "filter_text": filter_text,
        "filter_from": filter_from,
        "filter_to": filter_to,
        "filter_subject": filter_subject,
        "filter_unread": unread_only,
        "filter_attach": attachments_only,
        "scan_cap": scan_cap,
    }


def _friendly_outlook_error(exc, action="talk to Outlook"):
    if isinstance(exc, pywintypes.com_error):
        return (
            f"Outlook was busy while trying to {action}. "
            "Please make sure Outlook is open, wait a few seconds, and try again."
        )
    return str(exc)


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
                order_index INTEGER NOT NULL DEFAULT 0,
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

            CREATE TABLE IF NOT EXISTS search_suggestions_cache (
                value TEXT PRIMARY KEY,
                source TEXT NOT NULL DEFAULT 'unknown',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        thread_columns = {row["name"] for row in conn.execute("PRAGMA table_info(watched_threads)").fetchall()}
        if "order_index" not in thread_columns:
            conn.execute("ALTER TABLE watched_threads ADD COLUMN order_index INTEGER NOT NULL DEFAULT 0")
        zero_order_rows = conn.execute(
            """
            SELECT thread_id
            FROM watched_threads
            WHERE COALESCE(order_index, 0) = 0
            ORDER BY created_at ASC, thread_id ASC
            """
        ).fetchall()
        if zero_order_rows:
            next_order = conn.execute("SELECT COALESCE(MAX(order_index), 0) FROM watched_threads").fetchone()[0]
            conn.executemany(
                "UPDATE watched_threads SET order_index = ? WHERE thread_id = ?",
                [(next_order + index, row["thread_id"]) for index, row in enumerate(zero_order_rows, start=1)],
            )
        _migrate_legacy_watching(conn)
        conn.commit()
    finally:
        conn.close()


def _cache_suggestion_values(values, source="outlook"):
    conn = get_db_connection()
    try:
        conn.executemany(
            """
            INSERT INTO search_suggestions_cache (value, source, updated_at)
            VALUES (?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(value) DO UPDATE SET
                source = excluded.source,
                updated_at = CURRENT_TIMESTAMP
            """,
            [(value, source) for value in values if str(value or "").strip()],
        )
        conn.commit()
    finally:
        conn.close()


def _load_cached_suggestions(limit=300, query=""):
    conn = get_db_connection()
    try:
        if query:
            needle = f"%{query.lower()}%"
            rows = conn.execute(
                """
                SELECT value
                FROM search_suggestions_cache
                WHERE lower(value) LIKE ?
                ORDER BY value COLLATE NOCASE
                LIMIT ?
                """,
                (needle, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT value
                FROM search_suggestions_cache
                ORDER BY value COLLATE NOCASE
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [row["value"] for row in rows]
    finally:
        conn.close()


def _build_search_suggestions_from_outlook(limit_contacts=300, limit_items=200):
    ns = get_outlook_namespace()
    suggestions = set()

    def add_value(value):
        text = str(value or "").strip()
        if text and len(text) >= 2:
            suggestions.add(text)

    try:
        contacts = ns.GetDefaultFolder(10).Items
        for index, item in enumerate(contacts):
            if index >= limit_contacts:
                break
            add_value(getattr(item, "FullName", ""))
            add_value(getattr(item, "CompanyName", ""))
            add_value(getattr(item, "Email1Address", ""))
            add_value(getattr(item, "Email2Address", ""))
            add_value(getattr(item, "Email3Address", ""))
        if suggestions:
            _cache_suggestion_values(suggestions, source="contacts")
    except Exception:
        pass

    for folder_id, field_names, sort_field in (
        (6, ("SenderName", "SenderEmailAddress", "To"), "[ReceivedTime]"),
        (5, ("To",), "[SentOn]"),
    ):
        local_values = set()
        try:
            items = ns.GetDefaultFolder(folder_id).Items
            items.Sort(sort_field, True)
            for index, item in enumerate(items):
                if index >= limit_items:
                    break
                for field_name in field_names:
                    add_value(getattr(item, field_name, ""))
                    local_values.add(str(getattr(item, field_name, "") or "").strip())
        except Exception:
            continue
        if local_values:
            _cache_suggestion_values(local_values, source="mail")

    return sorted(suggestions.union(_load_cached_suggestions(limit=limit_contacts + limit_items)), key=lambda value: value.lower())[:300]


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
               order_index, latest_entry_id, latest_subject, latest_from, latest_received,
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
            "order_index": row["order_index"],
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
    next_order = conn.execute("SELECT COALESCE(MAX(order_index), 0) + 1 FROM watched_threads").fetchone()[0]
    conn.execute(
        """
        INSERT INTO watched_threads (
            seed_entry_id, conversation_id, normalized_subject, subject, order_index,
            latest_entry_id, latest_subject, latest_from, latest_received,
            latest_received_sort
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            email.get("entry_id"),
            email.get("conversation_id") or None,
            email.get("normalized_subject") or normalize_subject(email.get("subject")),
            email.get("subject") or "(No Subject)",
            next_order,
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
        "order_index": next_order,
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
            order_index = ?,
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
            thread.get("order_index") or 0,
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
                   order_index,
                   latest_entry_id, latest_subject, latest_from, latest_received,
                   latest_received_sort
            FROM watched_threads
            ORDER BY order_index ASC, thread_id ASC
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
                "order_index": row["order_index"],
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
        _normalize_watching_order(conn)
        conn.commit()
    finally:
        conn.close()


def remove_watching_by_entry(entry_id):
    conn = get_db_connection()
    try:
        row = conn.execute("SELECT thread_id FROM watched_threads WHERE seed_entry_id = ?", (entry_id,)).fetchone()
        if row:
            conn.execute("DELETE FROM watched_threads WHERE thread_id = ?", (row["thread_id"],))
            _normalize_watching_order(conn)
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


def _normalize_watching_order(conn):
    rows = conn.execute(
        "SELECT thread_id FROM watched_threads ORDER BY order_index ASC, thread_id ASC"
    ).fetchall()
    conn.executemany(
        "UPDATE watched_threads SET order_index = ?, updated_at = CURRENT_TIMESTAMP WHERE thread_id = ?",
        [(index, row["thread_id"]) for index, row in enumerate(rows, start=1)],
    )


def reorder_watching_threads(thread_ids):
    cleaned_ids = []
    seen = set()
    for raw_id in thread_ids or []:
        try:
            thread_id = int(raw_id)
        except (TypeError, ValueError):
            continue
        if thread_id in seen:
            continue
        seen.add(thread_id)
        cleaned_ids.append(thread_id)

    conn = get_db_connection()
    try:
        existing_ids = [row["thread_id"] for row in conn.execute(
            "SELECT thread_id FROM watched_threads ORDER BY order_index ASC, thread_id ASC"
        ).fetchall()]
        if not existing_ids:
            return []

        remaining_ids = [thread_id for thread_id in existing_ids if thread_id not in seen]
        final_ids = cleaned_ids + remaining_ids
        conn.executemany(
            "UPDATE watched_threads SET order_index = ?, updated_at = CURRENT_TIMESTAMP WHERE thread_id = ?",
            [(index, thread_id) for index, thread_id in enumerate(final_ids, start=1)],
        )
        conn.commit()
    finally:
        conn.close()

    return list_watching_threads()


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
            status = json.dumps({"type": "status", "text": mode_config.get("status_text") or f"Connecting to Outlook ({mode_config['label']})..."})
            yield f"data: {status}\n\n"

            result_queue = queue.Queue()

            def read_outlook_worker():
                try:
                    result_queue.put(("ok", get_outlook_emails(mode_config)))
                except Exception as exc:
                    result_queue.put(("error", exc))

            worker = threading.Thread(target=read_outlook_worker, daemon=True)
            worker.start()

            wait_seconds = 0
            emails = None
            while True:
                try:
                    kind, payload = result_queue.get(timeout=2)
                    if kind == "error":
                        raise payload
                    emails = payload
                    break
                except queue.Empty:
                    wait_seconds += 2
                    if wait_seconds == 10:
                        yield f"data: {json.dumps({'type': 'status', 'text': 'Still waiting on Outlook. If Outlook has a popup or security prompt, bring it to the front.'})}\n\n"
                    elif wait_seconds >= 30:
                        raise RuntimeError(
                            "Outlook did not respond within 30 seconds. Bring Outlook to the front, close any popups, and try again."
                        )
        except RuntimeError as exc:
            yield f"data: {json.dumps({'type': 'error', 'text': str(exc)})}\n\n"
            return

        if not emails:
            yield f"data: {json.dumps({'type': 'error', 'text': 'No emails found for that selection.'})}\n\n"
            return

        emails = apply_rules(emails, load_rules())
        emails = _annotate_threads(emails)
        refresh_watched_threads(emails)
        display_emails = _display_emails(emails)

        found_msg = json.dumps({"type": "status", "text": f"Found {len(emails)} emails across {len(display_emails)} active threads. Summarizing with OpenAI..."})
        yield f"data: {found_msg}\n\n"

        client = OpenAI(api_key=api_key)
        try:
            email_text = format_emails_for_claude(_build_summary_email_payload(display_emails, compact=False))
            prompt = PROMPT_TEMPLATE.format(count=len(display_emails), emails=email_text)
            summary_text = _request_summary_text(client, prompt)
            if not summary_text:
                compact_text = format_emails_for_claude(_build_summary_email_payload(display_emails, compact=True))
                compact_prompt = PROMPT_TEMPLATE.format(count=len(display_emails), emails=compact_text)
                summary_text = _request_summary_text(client, compact_prompt)
            if not summary_text:
                summary_text = _build_local_summary(display_emails)
                yield f"data: {json.dumps({'type': 'status', 'text': 'OpenAI was quiet, so a local fallback summary was generated.'})}\n\n"
            yield f"data: {json.dumps({'type': 'text', 'text': summary_text})}\n\n"
        except Exception as exc:
            summary_text = _build_local_summary(display_emails)
            yield f"data: {json.dumps({'type': 'status', 'text': f'OpenAI had trouble ({exc}). Showing a local fallback summary instead.'})}\n\n"
            yield f"data: {json.dumps({'type': 'text', 'text': summary_text})}\n\n"

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
                "thread_message_count": email.get("thread_message_count", 1),
                "thread_latest_subject": email.get("thread_latest_subject", email["subject"]),
                "thread_latest_received": email.get("thread_latest_received", email["received"]),
            }
            for email in display_emails
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


@app.route("/watching/reorder", methods=["POST"])
def reorder_watching():
    payload = request.json or {}
    threads = reorder_watching_threads(payload.get("thread_ids") or [])
    return jsonify({"ok": True, "threads": threads, "items": threads})


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
        return jsonify({"ok": False, "error": _friendly_outlook_error(exc, "open that email")}), 500


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
        return jsonify({"ok": False, "error": _friendly_outlook_error(exc, "mark emails as read")}), 500


@app.route("/my-identity")
def my_identity():
    try:
        ns = get_outlook_namespace()
        user = ns.CurrentUser
        return jsonify({"ok": True, "email": (user.Address or "").lower(), "name": user.Name or ""})
    except Exception as exc:
        return jsonify({"ok": False, "error": _friendly_outlook_error(exc, "read your Outlook identity")}), 500


@app.route("/search-suggestions")
def search_suggestions():
    try:
        query = (request.args.get("q", "") or "").strip()
        refresh = str(request.args.get("refresh", "")).lower() in {"1", "true", "yes", "on"}
        cached = _load_cached_suggestions(limit=300, query=query)
        if refresh or not cached:
            _build_search_suggestions_from_outlook()
            cached = _load_cached_suggestions(limit=300, query=query)
        values = cached
        return jsonify({"ok": True, "values": values})
    except Exception as exc:
        return jsonify({"ok": False, "error": _friendly_outlook_error(exc, "read Outlook contacts")}), 500


@app.route("/favicon.ico")
def favicon():
    return ("", 204)


def open_browser():
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    threading.Timer(1.0, open_browser).start()
    app.run(debug=False, port=5000)
