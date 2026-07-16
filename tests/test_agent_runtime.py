from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import agent_runtime


class AgentRuntimeTests(unittest.TestCase):
    def test_operation_key_normalizes_spacing_and_case(self):
        self.assertEqual(agent_runtime.operation_key("Journal", " A   B "), agent_runtime.operation_key("journal", "a b"))

    def test_begin_and_finish_are_visible_in_status(self):
        with tempfile.TemporaryDirectory() as tmp:
            runtime = Path(tmp)
            with patch.object(agent_runtime, "RUNTIME_DIR", runtime), patch.object(agent_runtime, "STATE_FILE", runtime / "state.json"):
                self.assertEqual(agent_runtime.begin("journal", "", 0, False), 0)
                state = agent_runtime.load_state()
                run_id = state["runs"][0]["run_id"]
                self.assertEqual(agent_runtime.finish(run_id, "succeeded", 0, "ok"), 0)
                self.assertEqual(agent_runtime.load_state()["runs"][0]["status"], "succeeded")

    def test_recent_success_is_deduplicated(self):
        with tempfile.TemporaryDirectory() as tmp:
            runtime = Path(tmp)
            with patch.object(agent_runtime, "RUNTIME_DIR", runtime), patch.object(agent_runtime, "STATE_FILE", runtime / "state.json"):
                agent_runtime.begin("journal", "same", 0, False)
                run_id = agent_runtime.load_state()["runs"][0]["run_id"]
                agent_runtime.finish(run_id, "succeeded", 0, "")
                self.assertEqual(agent_runtime.begin("journal", "same", 300, False), 10)

    def test_restore_rejects_path_traversal(self):
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "bad.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("../escape.txt", "bad")
            self.assertEqual(agent_runtime.restore(str(archive), False, False), 3)


if __name__ == "__main__":
    unittest.main()
