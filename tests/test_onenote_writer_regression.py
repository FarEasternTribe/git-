import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class OneNoteWriterRegressionTests(unittest.TestCase):
    def test_text_writer_delegates_to_shared_hardened_writer(self):
        script = (ROOT / "tools" / "create_onenote_text_page.ps1").read_text(
            encoding="utf-8"
        )

        self.assertIn("create_onenote_text_image_page.ps1", script)
        self.assertNotIn("CreateNewPage", script)
        self.assertNotIn("UpdatePageContent", script)

    def test_shared_writer_uses_title_page_and_existing_page_xml(self):
        script = (
            ROOT / "tools" / "create_onenote_text_image_page.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn("CreateNewPage($section.ID, [ref]$pageId, 1)", script)
        self.assertIn("GetPageContent($pageId, [ref]$existingContent, 2)", script)
        self.assertIn("$page = $doc.DocumentElement", script)
        self.assertNotIn("$doc.CreateElement('one', 'Page'", script)

    def test_shared_writer_has_retry_full_readback_and_empty_cleanup(self):
        script = (
            ROOT / "tools" / "create_onenote_text_image_page.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn("$attempt -le 5", script)
        self.assertIn("text read-back mismatch at line", script)
        self.assertIn("Remove-EmptyCreatedPage", script)
        self.assertIn("Refusing to create an empty OneNote page", script)
        self.assertIn("Section match count must be 1", script)

    def test_shared_writer_preserves_typed_visuals_in_source_order(self):
        script = (
            ROOT / "tools" / "create_onenote_text_image_page.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn("IMAGE|FIGURE|EQUATION", script)
        self.assertIn("Figures=$figureCount", script)
        self.assertIn("Equations=$equationCount", script)
        self.assertIn("VisualOrder=", script)
        self.assertIn("visual-kind read-back mismatch", script)


if __name__ == "__main__":
    unittest.main()
