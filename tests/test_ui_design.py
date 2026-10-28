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

    def test_ui_names_both_csv_and_excel_inputs(self):
        self.assertIn('Excel', self.app.title())
        self.assertIn('Excel', self.app.tree.heading('source')['text'])
        # The UI must not mislabel all source inputs as CSV-only.
        self.assertIn('Excel', self.app.state_text.get())

    def test_mixed_input_preview_counts_files_not_only_csv(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            directory = Path(folder)
            (directory / 'A사업장_20261028_11시.csv').write_text('ID\n001\n',encoding='utf-8')
            excel_path = directory / 'A사업장_20261028_14시.xlsx'
            wb = Workbook()
            wb.active['A1'] = 'test'
            wb.save(excel_path)
            wb.close()
            self.app.folder.set(str(directory))
            self.app._scan()
            self.assertEqual(self.app.metrics['groups'].get(), '1')
            self.assertIn('2개 파일', self.app.summary.get())
            group_id = next(iter(self.app.group_map))
            self.assertEqual(len(self.app.tree.get_children(group_id)), 2)

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
        self.assertEqual(str(self.app.tree.cget("selectmode")), "browse")

    def test_small_window_keeps_tree_and_actions_usable(self):
        # A minimum-sized window must show at least two preview rows below headings.
        self.app.geometry("900x640")
        self.app.update()
        self.assertGreaterEqual(self.app.tree.winfo_height(), 100)
        self.assertTrue(self.app.all_button.winfo_ismapped())
        self.assertTrue(self.app.pick_button.winfo_ismapped())

    def test_scan_updates_metrics_and_tree_highlight_is_not_checkmark(self):
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
            self.assertEqual(self.app._selected_groups(), [])
            self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")
            self.app._toggle_group(parent)
            self.assertEqual(self.app._selected_groups(), [self.app.group_map[parent]])


    def _scan_three_groups(self, folder):
        for site in ("A사업장", "B사업장", "C사업장"):
            for hour in ("11시", "14시"):
                (folder / f"{site}_20261028_{hour}.csv").write_text(
                    "ID,IP\n001,192.168.0.1\n", encoding="utf-8"
                )
        self.app.folder.set(str(folder))
        self.app._scan()
        return list(self.app.group_map)

    def test_checking_multiple_groups_updates_preview_and_button(self):
        # A checked set, rather than Treeview row highlighting, must control a merge.
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.assertTrue(self.app.tree.item(ids[0], "text").startswith("☐"))
            self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")
            self.app._toggle_group(ids[0])
            self.app._toggle_group(ids[2])
            self.app.tree.selection_set(ids[1])  # Highlight alone does not check B.
            self.assertEqual(self.app._selected_groups(), [self.app.group_map[ids[0]], self.app.group_map[ids[2]]])
            self.assertTrue(self.app.tree.item(ids[0], "text").startswith("☑"))
            self.assertTrue(self.app.tree.item(ids[2], "text").startswith("☑"))
            self.assertTrue(self.app.tree.item(ids[1], "text").startswith("☐"))
            self.assertIn("2 / 전체 3", self.app.selection_summary.get())
            self.assertEqual(str(self.app.selected_button.cget("state")), "normal")
            self.app._toggle_group(ids[0])
            self.app._toggle_group(ids[2])
            self.assertEqual(self.app._selected_groups(), [])
            self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")

    def test_select_all_clear_and_rescan_resets_group_checkmarks(self):
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.app._set_all_groups(True)
            self.assertEqual(len(self.app._selected_groups()), 3)
            self.assertIn("3 / 전체 3", self.app.selection_summary.get())
            self.app._set_all_groups(False)
            self.assertEqual(self.app._selected_groups(), [])
            self.app._toggle_group(ids[1])
            self.app._scan()  # Same IDs can be regenerated; old selection must not leak.
            self.assertEqual(self.app._selected_groups(), [])
            self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")
            self.assertTrue(all(self.app.tree.item(k, "text").startswith("☐") for k in self.app.group_map))

    def test_busy_disables_checkbox_actions_and_preserves_in_flight_groups(self):
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.app._toggle_group(ids[0])
            self.app.busy = True
            self.app._set_enabled(True)
            self.app._toggle_group(ids[1])
            self.app._set_all_groups(True)
            self.app._scan()
            self.assertEqual(self.app._selected_groups(), [self.app.group_map[ids[0]]])
            self.assertEqual(str(self.app.selected_button.cget("state")), "disabled")
            self.assertEqual(str(self.app.select_all_button.cget("state")), "disabled")
            self.assertEqual(str(self.app.clear_selection_button.cget("state")), "disabled")

    def test_checked_only_merge_writes_only_two_of_three_workbooks(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as td:
            directory = Path(td)
            ids = self._scan_three_groups(directory)
            self.app._toggle_group(ids[0])
            self.app._toggle_group(ids[2])
            with patch("excel_sheet_merger.gui.threading.Thread") as spawn:
                self.app._start_merge(True)
            spawn.assert_called_once()
            selected_groups, output_dir = spawn.call_args.kwargs["args"]
            self.assertEqual([g.site for g in selected_groups], ["A사업장", "C사업장"])
            self.app._worker_merge(selected_groups, output_dir)
            with patch("excel_sheet_merger.gui.messagebox.showinfo"):
                self.app._consume_events()
            for site in ("A사업장", "C사업장"):
                output = directory / "병합결과" / f"{site}_20261028.xlsx"
                wb = load_workbook(output, read_only=True)
                try:
                    self.assertEqual(wb.sheetnames, ["11시", "14시"])
                    self.assertEqual(wb["11시"]["A2"].value, "001")
                finally:
                    wb.close()
            self.assertFalse((directory / "병합결과" / "B사업장_20261028.xlsx").exists())

    def test_completion_message_mentions_renamed_workbooks(self):
        with tempfile.TemporaryDirectory() as td:
            directory = Path(td)
            self._scan_three_groups(directory)
            target = directory / '병합결과'
            target.mkdir()
            (target / 'A사업장_20261028.xlsx').write_bytes(b'PREVIOUS')
            with patch('excel_sheet_merger.gui.threading.Thread'):
                self.app._start_merge(False)
            self.app._worker_merge(list(self.app.report.groups), target)
            with patch('excel_sheet_merger.gui.messagebox.showinfo') as popup:
                self.app._consume_events()
            self.assertIn('이름 변경 1개', self.app.state_text.get())
            self.assertIn('완료 3개', self.app.state_text.get())
            self.assertNotIn('건너뜀', self.app.state_text.get())
            self.assertIn('이름 변경 1개', popup.call_args.args[1])
            self.assertIn('이미 존재', popup.call_args.args[1])
            self.assertTrue(self.app.log_expanded)

    def test_all_merge_ignores_checkbox_filter(self):
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.app._toggle_group(ids[1])
            with patch("excel_sheet_merger.gui.threading.Thread") as spawn:
                self.app._start_merge(False)
            spawn.assert_called_once()
            selected_groups = spawn.call_args.kwargs["args"][0]
            self.assertEqual([g.site for g in selected_groups], ["A사업장", "B사업장", "C사업장"])

    def test_mouse_and_keyboard_toggle_group_without_toggling_child(self):
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.app.update()
            group_box = self.app.tree.bbox(ids[0], column="#0")
            self.assertTrue(group_box)
            self.app.tree.event_generate("<ButtonRelease-1>", x=group_box[0] + 55, y=group_box[1] + 10)
            self.assertEqual([g.site for g in self.app._selected_groups()], ["A사업장"])
            child = self.app.tree.get_children(ids[0])[0]
            child_box = self.app.tree.bbox(child, column="#0")
            self.app.tree.event_generate("<ButtonRelease-1>", x=child_box[0] + 65, y=child_box[1] + 10)
            self.assertEqual([g.site for g in self.app._selected_groups()], ["A사업장"])
            # Hosted Windows CI cannot guarantee that synthetic KeyPress
            # events reach the foreground window. Exercise the actual bound
            # callback with a focused tree item rather than OS keyboard focus.
            self.assertTrue(self.app.tree.bind("<space>"))
            self.app.tree.focus(ids[1])
            self.app._on_tree_space(None)
            self.assertEqual([g.site for g in self.app._selected_groups()], ["A사업장", "B사업장"])
            self.app.tree.focus(child)
            self.app._on_tree_space(None)
            self.assertEqual([g.site for g in self.app._selected_groups()], ["A사업장", "B사업장"])

    def test_expand_arrow_click_does_not_toggle_group_checkbox(self):
        # The Treeview expand/collapse arrow is not the checkbox hit target.
        with tempfile.TemporaryDirectory() as td:
            ids = self._scan_three_groups(Path(td))
            self.app.update()
            x0, y0, width, height = self.app.tree.bbox(ids[0], column="#0")
            arrow_x = next((x for x in range(x0, x0 + min(width, 60))
                            if "indicator" in self.app.tree.identify_element(x, y0 + height // 2).lower()), None)
            if arrow_x is None:
                self.skipTest("theme does not expose tree expander hit region")
            self.app.tree.event_generate("<ButtonRelease-1>", x=arrow_x, y=y0 + height // 2)
            self.assertEqual(self.app._selected_groups(), [])
            self.assertTrue(self.app.tree.item(ids[0], "text").startswith("☐"))

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
