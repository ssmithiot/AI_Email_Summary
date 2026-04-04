"""
rules_engine.py
---------------
Evaluates user-defined rules against emails and produces a priority score.
Each email gets a numeric score — higher = more important.

Score sources:
  - Outlook High Importance flag  → +200  (always floats to top)
  - Matched user rules            → +rule.score each
  - Unread                        → +5    (tiebreaker)

Rule condition types
--------------------
to_only_me          I am the sole To: recipient (no one else in To or CC)
to_me               My email appears in the To: field
cc_only_me          I appear in CC but NOT in To
at_mention          @MyName appears anywhere in the body
from_address        Sender email equals/contains a value
from_domain         Sender email domain matches a value
subject_contains    Subject contains a keyword (case-insensitive)
body_contains       Body contains a keyword (case-insensitive)
has_attachment      Email has one or more attachments

All conditions in a rule use AND logic (all must match).
Rules themselves are additive — a matching rule adds its score.
"""

import json
import os
import re

RULES_FILE = os.path.join(os.path.dirname(__file__), "rules.json")

DEFAULT_CONFIG = {
    "my_email": "",
    "my_name": "",
    "rules": [
        {
            "id": "to_only_me",
            "name": "Sent only to me (sole recipient)",
            "description": "I am the only person in To — no CC, no other recipients",
            "enabled": True,
            "score": 100,
            "builtin": True,
            "conditions": [{"type": "to_only_me"}]
        },
        {
            "id": "at_mention",
            "name": "@mention of my name in body",
            "description": "@MyName appears somewhere in the email body",
            "enabled": True,
            "score": 80,
            "builtin": True,
            "conditions": [{"type": "at_mention"}]
        },
        {
            "id": "to_me_direct",
            "name": "My email is in the To field",
            "description": "I am a direct To recipient (may have others)",
            "enabled": True,
            "score": 60,
            "builtin": True,
            "conditions": [{"type": "to_me"}]
        },
        {
            "id": "has_attachment",
            "name": "Has attachments",
            "description": "Email contains one or more attachments",
            "enabled": True,
            "score": 30,
            "builtin": True,
            "conditions": [{"type": "has_attachment"}]
        },
        {
            "id": "cc_only_me",
            "name": "I am CC'd only (not in To)",
            "description": "I'm in CC but not a primary To recipient",
            "enabled": True,
            "score": 15,
            "builtin": True,
            "conditions": [{"type": "cc_only_me"}]
        }
    ],
    "custom_rules": []
}

# Condition types that require a user-supplied value
VALUE_CONDITIONS = {"from_address", "from_domain", "subject_contains", "body_contains"}

# Human-readable labels for the UI
CONDITION_LABELS = {
    "to_only_me":       "Sent only to me (sole recipient)",
    "to_me":            "My email is in the To field",
    "cc_only_me":       "I am CC'd but not in To",
    "at_mention":       "Body contains @mention of my name",
    "from_address":     "From address contains",
    "from_domain":      "From domain equals",
    "subject_contains": "Subject contains",
    "body_contains":    "Body contains",
    "has_attachment":   "Has attachments",
}


# ---------------------------------------------------------------------------
# Load / save
# ---------------------------------------------------------------------------

def load_rules():
    if not os.path.exists(RULES_FILE):
        save_rules(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    with open(RULES_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_rules(config):
    with open(RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


# ---------------------------------------------------------------------------
# Condition evaluation
# ---------------------------------------------------------------------------

def _normalise(s):
    return (s or "").strip().lower()


def _normalise_list(values):
    return [_normalise(value) for value in (values or []) if _normalise(value)]


def evaluate_condition(condition, email, my_email, my_name):
    ctype = condition.get("type", "")
    value = _normalise(condition.get("value", ""))

    to_str      = _normalise(email.get("to_recipients", ""))
    cc_str      = _normalise(email.get("cc_recipients", ""))
    to_list     = _normalise_list(email.get("to_recipient_list"))
    cc_list     = _normalise_list(email.get("cc_recipient_list"))
    subject     = _normalise(email.get("subject", ""))
    body        = _normalise(email.get("full_body") or email.get("body", ""))
    from_addr   = _normalise(email.get("from_email", ""))
    has_attach  = bool(email.get("has_attachment", False))

    if ctype == "to_only_me":
        return bool(my_email and len(to_list) == 1 and to_list[0] == my_email and not cc_list)

    elif ctype == "to_me":
        return bool(my_email and my_email in to_list)

    elif ctype == "cc_only_me":
        return bool(my_email and my_email in cc_list and my_email not in to_list)

    elif ctype == "at_mention":
        if not my_name:
            return False
        pattern = r"@" + re.escape(my_name)
        return bool(re.search(pattern, body, re.IGNORECASE))

    elif ctype == "from_address":
        return bool(value and value in from_addr)

    elif ctype == "from_domain":
        if not value or "@" not in from_addr:
            return False
        sender_domain = from_addr.split("@")[-1]
        return sender_domain == value

    elif ctype == "subject_contains":
        return bool(value and value in subject)

    elif ctype == "body_contains":
        return bool(value and value in body)

    elif ctype == "has_attachment":
        return has_attach

    return False


def evaluate_rule(rule, email, my_email, my_name):
    """All conditions must match (AND logic)."""
    conditions = rule.get("conditions", [])
    if not conditions:
        return False
    return all(evaluate_condition(c, email, my_email, my_name) for c in conditions)


# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------

def score_email(email, config):
    """Return (total_score, list_of_matched_rule_names)."""
    my_email = _normalise(config.get("my_email", ""))
    my_name  = _normalise(config.get("my_name", ""))

    total   = 0
    matched = []

    # Outlook High Importance always wins
    if email.get("urgent"):
        total += 200
        matched.append("Marked High Importance in Outlook")

    all_rules = list(config.get("rules", [])) + list(config.get("custom_rules", []))
    for rule in all_rules:
        if not rule.get("enabled"):
            continue
        if evaluate_rule(rule, email, my_email, my_name):
            total += rule.get("score", 0)
            matched.append(rule["name"])

    # Tiny tiebreaker for unread
    if email.get("unread"):
        total += 5

    return total, matched


def apply_rules(emails, config):
    """
    Score every email, sort descending by score, re-index,
    and attach rule_score / matched_rules to each dict.
    """
    scored = []
    for email in emails:
        e = dict(email)
        e["rule_score"], e["matched_rules"] = score_email(e, config)
        scored.append(e)

    scored.sort(key=lambda e: e["rule_score"], reverse=True)

    for i, e in enumerate(scored):
        e["index"] = i + 1

    return scored


# ---------------------------------------------------------------------------
# Sent-items reply detection
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r"^(re|fw|fwd)\s*:\s*", re.IGNORECASE)

def strip_prefixes(subject):
    s = (subject or "").strip()
    while True:
        new = _PREFIX_RE.sub("", s).strip()
        if new == s:
            break
        s = new
    return s.lower()
