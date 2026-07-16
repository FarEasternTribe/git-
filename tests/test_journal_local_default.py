from __future__ import annotations

import sys
import unittest
from unittest.mock import patch

import orchestrator_agent
import summarize_note5


class JournalLocalDefaultTests(unittest.TestCase):
    def test_vague_request_requires_clarification(self) -> None:
        route = orchestrator_agent.route_request("ようやくお願い", execute=True)

        self.assertEqual("確認Agent", route.agent)
        self.assertEqual([], route.command)
        self.assertIsNotNone(route.clarification_question)
        self.assertIn("何をどこまで", route.clarification_question or "")

    def test_clear_journal_request_does_not_require_clarification(self) -> None:
        route = orchestrator_agent.route_request("日誌を更新して", execute=True)

        self.assertEqual("日誌Agent", route.agent)
        self.assertIsNone(route.clarification_question)

    def test_pdf_summary_requires_per_paper_summary_validation(self) -> None:
        route = orchestrator_agent.route_request(r"C:\papers\sample.pdf を要約して", execute=True)

        self.assertEqual("PaperIndexAgent", route.agent)
        self.assertIn("--require-summary", route.command)

    def test_natural_language_journal_route_is_local(self) -> None:
        route = orchestrator_agent.journal_route("日誌を更新して", execute=True)

        self.assertFalse(route.requires_external_send)
        self.assertIn("--local-summary", route.command)
        self.assertIn("--skip-ask", route.command)
        self.assertNotIn("--api-summary", route.command)

    def test_api_summary_requires_explicit_request(self) -> None:
        route = orchestrator_agent.journal_route(
            "OpenAI API使用 api-summary 日誌を更新して",
            execute=True,
        )

        self.assertTrue(route.requires_external_send)
        self.assertIn("--api-summary", route.command)

    def test_cli_defaults_to_local_summary(self) -> None:
        with patch.object(sys, "argv", ["summarize_note5.py"]):
            args = summarize_note5.parse_args()

        self.assertTrue(args.local_summary)

    def test_cli_api_summary_is_opt_in(self) -> None:
        with patch.object(sys, "argv", ["summarize_note5.py", "--api-summary"]):
            args = summarize_note5.parse_args()

        self.assertFalse(args.local_summary)


if __name__ == "__main__":
    unittest.main()
