from __future__ import annotations

import unittest

from export_jira_work_items import _stable_resolved_since


class ExportJiraWorkItemsTests(unittest.TestCase):
    def test_stable_resolved_since_when_final_status_is_resolved(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "In Progress", "toString": "Resolved!"}],
                    }
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "2026-02-01")

    def test_stable_resolved_since_resets_when_reopened(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "In Progress", "toString": "Resolved!"}],
                    },
                    {
                        "created": "2026-02-02T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "Resolved!", "toString": "Reopened"}],
                    },
                    {
                        "created": "2026-02-03T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "Reopened", "toString": "Resolved!"}],
                    },
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "2026-02-03")

    def test_stable_resolved_since_empty_when_never_resolved(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "To Do", "toString": "In Progress"}],
                    }
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "")


if __name__ == "__main__":
    unittest.main()
