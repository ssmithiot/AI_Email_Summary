import os
import re
import sys
import time
from datetime import datetime, timedelta

import pywintypes
import win32com.client
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

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
        print("  b. Last 7 days")
        print("  c. Last 30 days")
        print("  d. Custom date (YYYY-MM-DD)")
        sub = input("  Choose [a/b/c/d]: ").strip().lower()
        now = datetime.now()
        if sub == "a":
            since = now.replace(hour=0, minute=0, second=0, microsecond=0)
            label = "emails from today"
        elif sub == "b":
            since = now - timedelta(days=7)
            label = "emails from the last 7 days"
        elif sub == "c":
            since = now - timedelta(days=30)
            label = "emails from the last 30 days"
        elif sub == "d":
            raw = input("  Enter start date (YYYY-MM-DD): ").strip()
            try:
                since = datetime.strptime(raw, "%Y-%m-%d")
                label = f"emails since {raw}"
            except ValueError:
                print("  Invalid date format, defaulting to last 7 days.")
                since = now - timedelta(days=7)
                label = "emails from the last 7 days"
        else:
            since = now - timedelta(days=7)
            label = "emails from the last 7 days"
        return {"mode": "date", "since": since, "label": label}

    raw = input("  How many recent emails? [press Enter for 30]: ").strip()
    count = int(raw) if raw.isdigit() else 30
    return {"mode": "quantity", "count": count, "label": f"{count} most recent emails"}


def get_outlook_emails(mode_config):
    def _read_messages():
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            raise RuntimeError(f"Could not connect to Outlook. Is it open and logged in?\n{exc}")

        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Newest first

        mode = mode_config["mode"]
        max_count = mode_config.get("count", 500)
        since = mode_config.get("since")

        emails = []
        count = 0
        for msg in messages:
            if mode == "quantity" and count >= max_count:
                break
            try:
                if mode == "unread" and not msg.UnRead:
                    continue
                received = getattr(msg, "ReceivedTime", None)
                if hasattr(received, "strftime") and mode == "date" and since and received < since:
                    break
                record = build_email_record(msg, index=count + 1)
                emails.append(record)
                count += 1
            except Exception:
                continue

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
