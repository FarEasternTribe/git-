from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET

from summarize_note5 import (
    extract_explicit_todo_items,
    markdown_to_onenote_page_xml,
    select_new_onenote_todos,
)


NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


class OneNoteTodoTagTests(unittest.TestCase):
    def test_fullwidth_todo_items_become_final_onenote_tags(self) -> None:
        text = "＠Todo\n□トリアジンの合成再合成\n□NMRデータ取りに行く"
        items = extract_explicit_todo_items(text)
        xml = markdown_to_onenote_page_xml("日誌", "# 本文\n研究メモ", items)
        root = ET.fromstring(xml.replace("__PAGE_ID__", "page-id"))

        self.assertIsNotNone(root.find("one:TagDef", NS))
        outline_items = root.findall(".//one:Outline//one:OE", NS)
        tagged_items = [node for node in outline_items if node.find("one:Tag", NS) is not None]
        self.assertEqual(["トリアジンの合成再合成", "NMRデータ取りに行く"], [
            node.find("one:T", NS).text for node in tagged_items
        ])
        self.assertEqual(tagged_items, outline_items[-2:])
        self.assertTrue(all(node.find("one:Tag", NS).get("completed") == "false" for node in tagged_items))

    def test_ampersand_todo_marker_from_voice_input_is_supported(self) -> None:
        text = "& Todo\n□次の反応試しておく\n□NMR管から化合物回収\n---"
        self.assertEqual(
            ["次の反応試しておく", "NMR管から化合物回収"],
            extract_explicit_todo_items(text),
        )

    def test_source_history_prevents_corrected_task_from_being_readded(self) -> None:
        original = "脱水化二卜発注"
        new_items, keys = select_new_onenote_todos([original], [])
        self.assertEqual([original], new_items)

        # The user may correct the visible OneNote text later. The unchanged
        # raw source spelling is still known and must not be appended again.
        new_items, updated_keys = select_new_onenote_todos([original, "新しい反応を確認"], keys)
        self.assertEqual(["新しい反応を確認"], new_items)
        self.assertEqual(len(keys) + 1, len(updated_keys))


if __name__ == "__main__":
    unittest.main()
