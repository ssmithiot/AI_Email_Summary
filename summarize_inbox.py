import os
import re
import sys
import time
from datetime import datetime, timedelta

import pythoncom
import pywintypes
import win32com.client
from openai import OpenAI
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(__file__)
load_dotenv(os.path.join(BASE_DIR, ".env"))

BODY_PREVIEW_LEN = 600  # Characters of body to include per email
TRANSPORT_HEADERS_PROP = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
MESSAGE_ID_PATTERN = re.compile(r"<[^>]+>")
WHITESPACE_PATTERN = re.compile(r"\s+")
HEADER_PATTERN_TEMPLATE = r"(?im)^{}:\s*(.+?)(?:\r?\n[ \t].+?)*$"


def _split_addresses(value):
    return [part.strip().lower() for part in value.split(";") if part.strip()]


def normalize_subject(subject):
    base = re.sub(r"^(?:\s*(?:re|fw|fwd)\s*:\s*)+", "", str(subject or ""), flags=re.IGNORECASE)
    return WHITESPACE_PATTERN.sub(" ", base).strip().lower()


def normalize_message_id(value):
    if not value:
        return ""
    text = str(value).strip()
    match = MESSAGE_ID_PATTERN.search(text)
    if match:
        text = match.group(0)
    if not text:
        return ""
    if not text.startswith("<"):
        text = f"<{text}>"
    if not text.endswith(">"):
        text = f"{text}>"
    return text.lower()


def extract_message_ids(value):
    return [normalize_message_id(match) for match in MESSAGE_ID_PATTERN.findall(str(value or "")) if normalize_message_id(match)]


def _header_value(headers, name):
    if not headers:
        return ""
    pattern = re.compile(HEADER_PATTERN_TEMPLATE.format(re.escape(name)))
    match = pattern.search(headers)
    if not match:
        return ""
    return WHITESPACE_PATTERN.sub(" ", match.group(0).split(":", 1)[1]).strip()


def _get_transport_headers(msg):
    try:
        return msg.PropertyAccessor.GetProperty(TRANSPORT_HEADERS_PROP) or ""
    except Exception:
        return ""


def _coerce_datetime(value):
    if not hasattr(value, "year"):
        return None
    try:
        if hasattr(value, "tzinfo") and value.tzinfo is not None:
            value = value.replace(tzinfo=None)
        return datetime(
            value.year,
            value.month,
            value.day,
            getattr(value, "hour", 0),
            getattr(value, "minute", 0),
            getattr(value, "second", 0),
            getattr(value, "microsecond", 0),
        )
    except Exception:
        return None


def _matches_text_filter(record, filter_text):
    if not filter_text:
        return True
    needle = filter_text.lower()
    haystacks = [
        record.get("from", ""),
        record.get("from_email", ""),
        record.get("subject", ""),
        record.get("body", ""),
        record.get("full_body", ""),
    ]
    return any(needle in str(haystack or "").lower() for haystack in haystacks)


def _matches_structured_filters(record, filter_from="", filter_to="", filter_subject=""):
    from_value = str(record.get("from", "") or "").lower()
    from_email = str(record.get("from_email", "") or "").lower()
    to_value = str(record.get("to_recipients", "") or "").lower()
    cc_value = str(record.get("cc_recipients", "") or "").lower()
    subject_value = str(record.get("subject", "") or "").lower()

    if filter_from:
        needle = filter_from.lower()
        if needle not in from_value and needle not in from_email:
            return False

    if filter_to:
        needle = filter_to.lower()
        if needle not in to_value and needle not in cc_value:
            return False

    if filter_subject:
        needle = filter_subject.lower()
        if needle not in subject_value:
            return False

    return True


def _iter_mail_folders(folder):
    yield folder
    try:
        for child in folder.Folders:
            yield from _iter_mail_folders(child)
    except Exception:
        return


def _inbox_folders(namespace, include_all_inboxes=False):
    if not include_all_inboxes:
        return [namespace.GetDefaultFolder(6)]

    folders = []
    seen = set()
    for index in range(1, namespace.Folders.Count + 1):
        try:
            store_root = namespace.Folders.Item(index)
            inbox = store_root.Folders["Inbox"]
            key = getattr(inbox, "EntryID", None) or f"idx:{index}"
            if key in seen:
                continue
            seen.add(key)
            folders.append(inbox)
        except Exception:
            continue

    return folders or [namespace.GetDefaultFolder(6)]


def list_available_inboxes():
    namespace = get_outlook_namespace()
    mailboxes = []
    for index in range(1, namespace.Folders.Count + 1):
        try:
            store_root = namespace.Folders.Item(index)
            inbox = store_root.Folders["Inbox"]
            store_id = getattr(store_root, "EntryID", "") or ""
            mailboxes.append(
                {
                    "id": store_id,
                    "name": str(getattr(store_root, "Name", "") or f"Mailbox {index}"),
                    "inbox_name": str(getattr(inbox, "Name", "") or "Inbox"),
                    "item_count": int(getattr(inbox.Items, "Count", 0) or 0),
                }
            )
        except Exception:
            continue
    return mailboxes


def _recipient_strings(msg):
    to_str = ""
    cc_str = ""
    try:
        for recipient in msg.Recipients:
            addr = (recipient.Address or "").lower()
            if recipient.Type == 1:  # olTo
                to_str += addr + ";"
            elif recipient.Type == 2:  # olCC
                cc_str += addr + ";"
    except Exception:
        to_str = (msg.To or "").lower()
        cc_str = (msg.CC or "").lower()
    return to_str, cc_str


def build_email_record(msg, index=None):
    received = getattr(msg, "ReceivedTime", "")
    if hasattr(received, "strftime"):
        received_str = received.strftime("%Y-%m-%d %H:%M")
        received_sort = received.strftime("%Y-%m-%d %H:%M:%S")
    else:
        received_str = str(received)
        received_sort = received_str

    body = (getattr(msg, "Body", "") or "").strip()
    body_preview = body[:BODY_PREVIEW_LEN] + ("..." if len(body) > BODY_PREVIEW_LEN else "")
    importance = getattr(msg, "Importance", 1)
    to_str, cc_str = _recipient_strings(msg)
    headers = _get_transport_headers(msg)

    internet_message_id = normalize_message_id(
        _header_value(headers, "Message-ID")
        or _header_value(headers, "Message-Id")
        or getattr(msg, "InternetMessageID", "")
    )
    in_reply_to = normalize_message_id(_header_value(headers, "In-Reply-To"))
    references = extract_message_ids(_header_value(headers, "References"))
    subject = getattr(msg, "Subject", "") or "(No Subject)"

    record = {
        "entry_id": getattr(msg, "EntryID", ""),
        "from": getattr(msg, "SenderName", "") or getattr(msg, "SenderEmailAddress", "") or "Unknown",
        "from_email": (getattr(msg, "SenderEmailAddress", "") or "").lower(),
        "to_recipients": to_str,
        "to_recipient_list": _split_addresses(to_str),
        "cc_recipients": cc_str,
        "cc_recipient_list": _split_addresses(cc_str),
        "has_attachment": getattr(getattr(msg, "Attachments", None), "Count", 0) > 0,
        "subject": subject,
        "normalized_subject": normalize_subject(subject),
        "received": received_str,
        "received_sort": received_sort,
        "unread": bool(getattr(msg, "UnRead", False)),
        "urgent": importance == 2,
        "body": body_preview,
        "full_body": body,
        "conversation_id": str(getattr(msg, "ConversationID", "") or "").strip(),
        "conversation_topic": str(getattr(msg, "ConversationTopic", "") or "").strip(),
        "internet_message_id": internet_message_id,
        "in_reply_to": in_reply_to,
        "references": references,
    }
    if index is not None:
        record["index"] = index
    return record


def pick_mode():
    print()
    print("=" * 60)
    print("  OUTLOOK INBOX SUMMARIZER")
    print("=" * 60)
    print("  1. Unread emails only")
    print("  2. All emails by date range")
    print("  3. Most recent (default: 30)")
    print("=" * 60)
    choice = input("  Choose [1/2/3] or press Enter for default (3): ").strip()

    if choice == "1":
        return {"mode": "unread", "label": "unread emails"}

    if choice == "2":
        print()
        print("  Date range options:")
        print("  a. Today")
        print("  b. Yesterday")
        print("  c. Last 3 days")
        print("  d. Last 7 days")
        print("  e. Custom date range (YYYY-MM-DD)")
        sub = input("  Choose [a/b/c/d/e]: ").strip().lower()
        now = datetime.now()
        if sub == "a":
            since = now.replace(hour=0, minute=0, second=0, microsecond=0)
            until = None
            label = "emails from today"
        elif sub == "b":
            since = (now - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
            until = now.replace(hour=0, minute=0, second=0, microsecond=0)
            label = "emails from yesterday"
        elif sub == "c":
            since = now - timedelta(days=3)
            until = None
            label = "emails from the last 3 days"
        elif sub == "d":
            since = now - timedelta(days=7)
            until = None
            label = "emails from the last 7 days"
        elif sub == "e":
            raw_start = input("  Enter start date (YYYY-MM-DD): ").strip()
            raw_end = input("  Enter end date (YYYY-MM-DD): ").strip()
            try:
                since = datetime.strptime(raw_start, "%Y-%m-%d")
                until = datetime.strptime(raw_end, "%Y-%m-%d") + timedelta(days=1)
                if until <= since:
                    raise ValueError
                label = f"emails from {raw_start} to {raw_end}"
            except ValueError:
                print("  Invalid date range, defaulting to last 7 days.")
                since = now - timedelta(days=7)
                until = None
                label = "emails from the last 7 days"
        else:
            since = now - timedelta(days=7)
            until = None
            label = "emails from the last 7 days"
        return {"mode": "date", "since": since, "until": until, "label": label}

    raw = input("  How many recent emails? [press Enter for 30]: ").strip()
    count = int(raw) if raw.isdigit() else 30
    return {"mode": "quantity", "count": count, "label": f"{count} most recent emails"}


def get_outlook_namespace():
    # Prefer attaching to a running Outlook session, but fall back to a direct
    # COM activation when Outlook is open yet not exposed through GetActiveObject.
    pythoncom.CoInitialize()
    errors = []

    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        return outlook.GetNamespace("MAPI")
    except Exception as exc:
        errors.append(f"attach failed: {exc}")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook.GetNamespace("MAPI")
    except Exception as exc:
        errors.append(f"dispatch failed: {exc}")

    raise RuntimeError(
        "Outlook could not be reached through Windows automation. Outlook appears to be installed, but its COM/MAPI "
        "session is not available right now. Make sure classic Outlook is fully open to your inbox, no hidden prompts "
        "are waiting, and Outlook and this app are running at the same privilege level. "
        f"Details: {'; '.join(errors)}"
    )


def get_outlook_emails(mode_config):
    def _read_messages():
        try:
            namespace = get_outlook_namespace()
        except RuntimeError:
            raise
        selected_mailbox_ids = {value for value in mode_config.get("mailbox_ids", []) if value}
        include_all_inboxes = bool(mode_config.get("include_all_inboxes")) or bool(selected_mailbox_ids)
        inboxes = _inbox_folders(namespace, include_all_inboxes=include_all_inboxes)
        if selected_mailbox_ids:
            filtered_inboxes = []
            for inbox in inboxes:
                parent = getattr(inbox, "Parent", None)
                parent_id = getattr(parent, "EntryID", "") if parent is not None else ""
                if parent_id in selected_mailbox_ids:
                    filtered_inboxes.append(inbox)
            inboxes = filtered_inboxes
        mode = mode_config["mode"]
        max_count = mode_config.get("count", 500)
        since = mode_config.get("since")
        until = mode_config.get("until")
        filter_text = mode_config.get("filter_text", "")
        filter_from = mode_config.get("filter_from", "")
        filter_to = mode_config.get("filter_to", "")
        filter_subject = mode_config.get("filter_subject", "")
        unread_only = bool(mode_config.get("filter_unread"))
        attachments_only = bool(mode_config.get("filter_attach"))
        include_subfolders = bool(mode_config.get("include_subfolders"))
        scan_cap = max(25, int(mode_config.get("scan_cap", 100) or 100))
        folders = []
        for inbox in inboxes:
            if include_subfolders:
                folders.extend(list(_iter_mail_folders(inbox)))
            else:
                folders.append(inbox)

        emails = []
        stop_early = False
        inspected_count = 0
        for folder in folders:
            try:
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True)  # Newest first
            except Exception:
                continue

            for msg in messages:
                try:
                    if mode in {"quantity", "unread"}:
                        if inspected_count >= scan_cap:
                            stop_early = True
                            break
                        inspected_count += 1
                    if mode == "unread" and not msg.UnRead:
                        continue
                    received = getattr(msg, "ReceivedTime", None)
                    received_dt = _coerce_datetime(received)
                    if received_dt and mode == "date" and until and received_dt >= until:
                        continue
                    if received_dt and mode == "date" and since and received_dt < since:
                        continue
                    record = build_email_record(msg)
                    if unread_only and not record.get("unread"):
                        continue
                    if attachments_only and not record.get("has_attachment"):
                        continue
                    if not _matches_structured_filters(record, filter_from, filter_to, filter_subject):
                        continue
                    if not _matches_text_filter(record, filter_text):
                        continue
                    emails.append(record)
                    if mode == "quantity" and len(emails) >= max_count:
                        stop_early = True
                        break
                except Exception:
                    continue
            if stop_early:
                break

        emails.sort(key=lambda item: item.get("received_sort", ""), reverse=True)
        if mode == "quantity":
            emails = emails[:max_count]
        for index, email in enumerate(emails, start=1):
            email["index"] = index

        return emails

    last_exc = None
    for attempt in range(2):
        try:
            return _read_messages()
        except pywintypes.com_error as exc:
            last_exc = exc
            if attempt == 0:
                time.sleep(1.5)
                continue
            break

    if last_exc is not None:
        raise RuntimeError(
            "Outlook is busy or temporarily unavailable right now. Please make sure Outlook is open, wait a few seconds, and try again."
        )

    raise RuntimeError("Outlook could not be read right now. Please try again.")


def format_emails_for_claude(emails):
    lines = []
    for email in emails:
        tags = []
        if email.get("urgent"):
            tags.append("URGENT")
        if email.get("unread"):
            tags.append("UNREAD")
        if email.get("has_attachment"):
            tags.append("HAS-ATTACHMENT")
        if email.get("rule_score", 0) > 0:
            tags.append(f"PRIORITY-SCORE:{email['rule_score']}")
        tag_str = (" [" + " | ".join(tags) + "]") if tags else ""

        lines.append(f"--- Email {email['index']}{tag_str} ---")
        lines.append(f"From:     {email['from']}")
        lines.append(f"Subject:  {email['subject']}")
        lines.append(f"Received: {email['received']}")
        if email.get("matched_rules"):
            lines.append(f"Rules:    {', '.join(email['matched_rules'])}")
        lines.append(f"Preview:  {email['body']}")
        lines.append("")
    return "\n".join(lines)


def extract_response_text(response):
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


def build_summary_email_payload(emails, compact=False):
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


def request_summary_text(client, prompt):
    response = client.responses.create(
        model="gpt-5",
        input=prompt,
        max_output_tokens=3000,
    )
    return extract_response_text(response).strip()


def summarize_with_claude(email_text, total_count, emails=None):
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("ERROR: OPENAI_API_KEY not set.")
        print("Add it to the .env file in this folder:  OPENAI_API_KEY=sk-...")
        sys.exit(1)

    client = OpenAI(api_key=api_key)

    prompt = f"""You are reviewing {total_count} emails from an Outlook inbox. Emails tagged [URGENT] have been marked High Importance by the sender in Outlook.

Return all of these sections in this exact order, even if some are empty:
1. **Urgent / High Priority** - list every [URGENT] tagged email first, then the next highest-priority emails.
2. **Quick Stats** - total emails, how many are unread, how many are urgent, how many need action.
3. **Action Items** - emails that need a reply or action, listed by priority, with a brief recommended next step.
4. **Key Topics** - group the remaining emails by theme (e.g. meetings, reports, FYI).
5. **Can Ignore / Low Priority** - newsletters, notifications, or low-priority FYIs.

Formatting rules:
- Use markdown headings for each section.
- Use bullet points under every section.
- If a section has nothing meaningful, write a short bullet like `- None flagged`.
- In Action Items, say why the email matters and what the user should do next.
- Do not stop after the first section.

Whenever you mention a specific email, append a citation in the exact format [Email N].
Be concise. Use bullet points. Always lead with urgent items.

--- INBOX EMAILS ---
{email_text}"""

    print(f"\nSummarizing {total_count} emails with OpenAI...\n")
    print("=" * 60)
    print(f"  INBOX SUMMARY  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    summary_text = request_summary_text(client, prompt)
    if not summary_text and emails:
        compact_email_text = format_emails_for_claude(build_summary_email_payload(emails, compact=True))
        compact_prompt = prompt.replace(email_text, compact_email_text)
        summary_text = request_summary_text(client, compact_prompt)
    print(summary_text, end="", flush=True)

    print("\n" + "=" * 60)


def main():
    mode_config = pick_mode()
    print(f"\nConnecting to Outlook ({mode_config['label']})...")
    emails = get_outlook_emails(mode_config)

    if not emails:
        print("No emails found for that selection.")
        sys.exit(0)

    print(f"Found {len(emails)} emails.")
    email_text = format_emails_for_claude(emails)
    summarize_with_claude(email_text, len(emails), emails)


if __name__ == "__main__":
    main()
