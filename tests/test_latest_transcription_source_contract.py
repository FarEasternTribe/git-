import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class LatestTranscriptionSourceContractTests(unittest.TestCase):
    def test_exporter_selects_date_and_transcription_titles(self):
        script = (ROOT / "tools" / "export_onenote_page_pdf.ps1").read_text(
            encoding="utf-8-sig"
        )

        self.assertIn("[switch]$LatestForDate", script)
        self.assertIn("$transcriptionSuffix", script)
        self.assertIn("FromBase64String", script)
        self.assertIn("lastModifiedTime", script)
        self.assertIn("ResolvedPageTitle=", script)
        self.assertIn("ResolvedPageId=", script)
        self.assertIn("Remove-Item -LiteralPath $absoluteOut -Force", script)

    def test_untouched_automation_output_is_excluded_by_registry(self):
        exporter = (ROOT / "tools" / "export_onenote_page_pdf.ps1").read_text(
            encoding="utf-8-sig"
        )
        writer = (ROOT / "tools" / "create_onenote_text_image_page.ps1").read_text(
            encoding="utf-8-sig"
        )

        self.assertIn("Test-UntouchedAutomationOutput", exporter)
        self.assertIn("ContentHash", exporter)
        self.assertIn(".agent_runtime\\transcription_outputs", writer)
        self.assertIn("ContentHash", writer)
        self.assertIn("AutomationOutputRegistered=", writer)

    def test_date_update_detector_enables_latest_selection(self):
        script = (ROOT / "tools" / "detect_onenote_transcription_updates.ps1").read_text(
            encoding="utf-8-sig"
        )

        self.assertIn("if ($PageTitle -match '^20\\d{6}$')", script)
        self.assertIn("$exportArgs += '-LatestForDate'", script)
        self.assertIn("ResolvedPageId=", script)
        self.assertIn('$resolvedTitle -cne $PageTitle', script)


if __name__ == "__main__":
    unittest.main()
