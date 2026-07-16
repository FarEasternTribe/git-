from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from paper_index_agent import validate_linked_summary


class PaperWorkflowCompletionTests(unittest.TestCase):
    def test_missing_summary_link_is_incomplete(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            library = Path(tmp)
            index = library / "index.md"
            index.write_text("# Paper\n\n## 要約\n未作成。\n", encoding="utf-8")

            ok, message = validate_linked_summary(index, library)

            self.assertFalse(ok)
            self.assertIn("Summaryリンク", message)

    def test_specific_linked_summary_is_complete(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            library = Path(tmp)
            summaries = library / "summaries"
            summaries.mkdir()
            summary = summaries / "paper_summary.md"
            summary.write_text("# Summary\n\n" + "検証済みの要約です。" * 30, encoding="utf-8")
            index = library / "index.md"
            index.write_text(
                "# Paper\n\n- Summary: summaries/paper_summary.md\n\n## 要約\n詳細要約\n",
                encoding="utf-8",
            )

            ok, message = validate_linked_summary(index, library)

            self.assertTrue(ok, message)


if __name__ == "__main__":
    unittest.main()
