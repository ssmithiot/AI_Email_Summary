import unittest

from rules_engine import evaluate_condition


class RulesEngineTests(unittest.TestCase):
    def test_body_contains_checks_full_body_not_preview_only(self):
        email = {
            "body": "short preview",
            "full_body": "short preview ... final approval keyword lives here",
        }

        result = evaluate_condition(
            {"type": "body_contains", "value": "approval keyword"},
            email,
            my_email="",
            my_name="",
        )

        self.assertTrue(result)

    def test_to_me_uses_exact_recipient_matches(self):
        email = {
            "to_recipients": "joann@example.com;",
            "to_recipient_list": ["joann@example.com"],
            "cc_recipients": "",
            "cc_recipient_list": [],
        }

        result = evaluate_condition(
            {"type": "to_me"},
            email,
            my_email="ann@example.com",
            my_name="",
        )

        self.assertFalse(result)

    def test_to_only_me_requires_single_exact_to_recipient(self):
        email = {
            "to_recipients": "ann@example.com;other@example.com;",
            "to_recipient_list": ["ann@example.com", "other@example.com"],
            "cc_recipients": "",
            "cc_recipient_list": [],
        }

        result = evaluate_condition(
            {"type": "to_only_me"},
            email,
            my_email="ann@example.com",
            my_name="",
        )

        self.assertFalse(result)

    def test_at_mention_uses_full_body(self):
        email = {
            "body": "preview text",
            "full_body": "FYI only. @Steph please review the attached plan.",
        }

        result = evaluate_condition(
            {"type": "at_mention"},
            email,
            my_email="",
            my_name="Steph",
        )

        self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()
