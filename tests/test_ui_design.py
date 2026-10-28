"""Office UI public behavior: scan results, actions and visual state.

GUI tests are run under Xvfb on Linux and may be skipped in headless CI.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

GUI_AVAILABLE = os.name == "nt" or bool(os.environ.get("DISPLAY"))


class ThemeTests(unittest.TestCase):
    def test_color_palette_distinguishes_primary_and_error(self):
        # Catches a missing centralized palette or incorrect primary action token.
        from excel_sheet_merger.ui_theme import COLORS
        self.assertEqual(COLORS["primary"], "#166D53")
        self.assertEqual(COLORS["background"], "#F6F8F7")
        self.assertNotEqual(COLORS["error"], COLORS["primary"])

    def test_metrics_derive_from_scan_report(self):
        # Catches counting rejected non-CSV source files as mergeable sheets.
        from excel_sheet_merger.classify import scan_folder
        from excel_sheet_merger.gui import summary_metrics
        with tempfile.TemporaryDirectory() as td:
            folder = Path(td)
            for filename in (
                "A사업장_20261028_11시.csv",
                "A사업장_20261028_14시.csv",
                "B사업장_20261028_11시.csv",
                "제외.csv",
            ):
                (folder / filename).write_text("a,b\n001,192.168.0.1\n", encoding="utf-8")
            scan = scan_folder(folder)
            self.assertEqual(summary_metrics(scan), ("4", "1", "1", "1"))


@unittest.skipUnless(GUI_AVAILABLE, "requires Windows desktop or Linux DISPLAY")
class GuiBehaviorTests(unittest.TestCase):
    def setUp(self):
        from excel_sheet_merger.gui import MergeApp
        self.app = MergeApp()
        self.app.update_idletasks()

    def tearDown(self):
        self.app.destroy()

    def test_initial_layout_has_metrics_and_collapsed_log(self):
        # Catches removal of main preview, hidden buttons or default open logs.
        self.assertEqual(self.app.minsize(), (900, 640))
        self.assertEqual(set(self.app.metrics), {"checked", "groups", "singles", "rejected"})
        self.assertEqual(str(self.app.all_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.folder_entry.cget("state")), "readonly")
        self.assertFalse(self.app.log_expanded)
        self.assertEqual(self.app.log_panel.winfo_manager(), "")
        self.assertTrue(self.app.tree.winfo_exists())
        self.assertEqual(str(self.app.tree.cget("selectmode")), "extended")

    def test_small_window_keeps_tree_and_actions_usable(self):
        # A minimum-sized window must show at least two preview rows below headings.
        self.app.geometry("900x640")
        self.app.update()
        self.assertGreaterEqual(self.app.tree.winfo_height(), 100)
        self.assertTrue(self.app.all_button.winfo_ismapped())
        self.assertTrue(self.app.pick_button.winfo_ismapped())

    def test_scan_updates_metric_cards_and_child_selection(self):
        # Catches styling work accidentally breaking Treeview group selection.
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            for name in ("A사업장_20261028_11시.csv", "A사업장_20261028_14시.csv", "B사업장_20261028_11시.csv", "foo.csv"):
                (d / name).write_text("value\n001\n", encoding="utf-8")
            self.app.folder.set(str(d))
            self.app._scan()
            self.assertEqual(tuple(self.app.metrics[name].get() for name in ("checked", "groups", "singles", "rejected")), ("4", "1", "1", "1"))
            self.assertEqual(str(self.app.all_button.cget("state")), "normal")
            parent = next(iter(self.app.group_map))
            child = self.app.tree.get_children(parent)[0]
            self.app.tree.selection_set(child)
            self.assertEqual(self.app._selected_groups(), [self.app.group_map[parent]])

    def test_manual_ui_event_flush_does_not_leave_multiple_poll_timers(self):
        # A manual event drain must replace, not duplicate, the scheduled poll.
        initial = self.app._poll_job
        self.assertIn(initial, self.app.tk.call("after", "info"))
        self.app._consume_events()
        timers = self.app.tk.call("after", "info")
        self.assertNotIn(initial, timers)
        self.assertIn(self.app._poll_job, timers)

    def test_collapsible_log_and_failed_job_expands_it(self):
        # Catches failure details remaining hidden when conversion fails.
        self.app._toggle_log()
        self.assertTrue(self.app.log_expanded)
        self.app._toggle_log()
        self.assertFalse(self.app.log_expanded)
        self.app.busy = True
        self.app.events.put(("log", "실패: 잘못된 CSV"))
        self.app.events.put(("done", 0, 1, 0))
        with patch("excel_sheet_merger.gui.messagebox.showwarning") as popup:
            self.app._consume_events()
        popup.assert_called_once()
        self.assertTrue(self.app.log_expanded)
        self.assertIn("실패 1개", self.app.state_text.get())
        self.assertIn("잘못된 CSV", self.app.log.get("1.0", "end"))


if __name__ == "__main__":
    unittest.main()
