"""Regression tests for UI reliability without changing spreadsheet operations."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

GUI_AVAILABLE = os.name == 'nt' or bool(os.environ.get('DISPLAY'))


@unittest.skipUnless(GUI_AVAILABLE, 'Windows desktop or DISPLAY required')
class GuiStabilityTests(unittest.TestCase):
    def setUp(self):
        from excel_sheet_merger.gui import MergeApp
        self.app = MergeApp()
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.addCleanup(self.app.destroy)

    def _valid_scan(self):
        folder = Path(self.temp.name)
        for hour in ('11시', '14시'):
            (folder / f'A사업장_20261028_{hour}.csv').write_text('ID\n001\n', encoding='utf-8')
        self.app.folder.set(str(folder))
        self.app._scan()
        self.assertIsNotNone(self.app.report)
        return folder

    def test_failed_worker_start_recovers_buttons_and_allows_retry(self):
        folder = self._valid_scan()
        group_id = next(iter(self.app.group_map))
        self.app._toggle_group(group_id)
        with patch('excel_sheet_merger.gui.threading.Thread') as worker, \
                patch('excel_sheet_merger.gui.messagebox.showerror') as popup:
            worker.return_value.start.side_effect = RuntimeError('cannot start new thread')
            self.app._start_merge(True)
        self.assertFalse(self.app.busy, 'failed Thread.start must not strand the UI in busy state')
        self.assertEqual(str(self.app.selected_button.cget('state')), 'normal')
        self.assertEqual(str(self.app.all_button.cget('state')), 'normal')
        self.assertIn('시작 실패', self.app.state_text.get())
        self.assertIn('cannot start new thread', self.app.log.get('1.0', 'end'))
        popup.assert_called_once()
        self.assertEqual([g.site for g in self.app._selected_groups()], ['A사업장'])
        self.assertFalse((folder / '병합결과').exists(), 'starting a failed thread must not write any output')
        with patch('excel_sheet_merger.gui.threading.Thread') as worker:
            self.app._start_merge(True)
            worker.return_value.start.assert_called_once_with()
        self.assertTrue(self.app.busy)
        self.app.busy = False

    def test_worker_start_oserror_is_recovered_without_losing_scan(self):
        self._valid_scan()
        with patch('excel_sheet_merger.gui.threading.Thread', side_effect=OSError(12, 'out of resources')), \
                patch('excel_sheet_merger.gui.messagebox.showerror') as popup:
            self.app._start_merge(False)
        self.assertFalse(self.app.busy)
        self.assertEqual(len(self.app.report.groups), 1)
        self.assertEqual(str(self.app.scan_button.cget('state')), 'normal')
        popup.assert_called_once()

    def test_result_folder_permission_error_is_explained_without_gui_callback_crash(self):
        self._valid_scan()
        with patch('excel_sheet_merger.gui.Path.mkdir',
                   side_effect=PermissionError(13, 'access denied')) as mkdir, \
             patch('excel_sheet_merger.gui.messagebox.showerror') as popup:
            self.app._open_results()
        mkdir.assert_called_once()
        popup.assert_called_once()
        self.assertIn('접근', popup.call_args.args[1])
        self.assertIn('권한', popup.call_args.args[1])
        self.assertIn('병합결과', popup.call_args.args[1])
        self.assertFalse(self.app.busy)

    @unittest.skipUnless(os.name == 'nt', 'Windows Explorer callback is Windows-only')
    def test_result_folder_explorer_error_has_korean_recovery(self):
        folder = self._valid_scan()
        with patch('excel_sheet_merger.gui.os.startfile',
                   side_effect=OSError(53, 'unreachable share')) as opener, \
             patch('excel_sheet_merger.gui.messagebox.showerror') as popup:
            self.app._open_results()
        opener.assert_called_once()
        self.assertTrue((folder / '병합결과').is_dir())
        popup.assert_called_once()
        self.assertIn('결과 폴더', popup.call_args.args[1])

    def test_burst_of_log_events_renders_in_one_text_insert_without_losing_messages(self):
        with patch.object(self.app.log, 'insert', wraps=self.app.log.insert) as writes:
            for i in range(100):
                self.app.events.put(('log', f'burst-{i:03d}'))
            self.app._consume_events()
        self.assertEqual(writes.call_count, 1, 'text widget inserts must be batched for a log burst')
        displayed = self.app.log.get('1.0', 'end')
        self.assertEqual(displayed.count('burst-'), 100)
        self.assertLess(displayed.index('burst-000'), displayed.index('burst-099'))
        self.assertRegex(displayed, r'\[\d{2}:\d{2}:\d{2}\] burst-000')

    def test_long_burst_keeps_every_visible_log_entry_without_pruning(self):
        # GUI memory may increase with full-history retention, but existing
        # browsing semantics must not change in a reliability-only update.
        for i in range(2000):
            self.app.events.put(('log', f'stress-log-{i:04d}'))
        while not self.app.events.empty():
            self.app._consume_events()
        displayed = self.app.log.get('1.0', 'end')
        self.assertEqual(displayed.count('stress-log-'), 2000)
        self.assertLess(displayed.index('stress-log-0000'),
                        displayed.index('stress-log-1999'))

    def test_batch_keeps_order_across_progress_and_done_events(self):
        self.app.busy = True
        self.app.events.put(('log', 'before-progress'))
        self.app.events.put(('progress', 1, 1))
        self.app.events.put(('log', 'after-progress'))
        self.app.events.put(('done', 1, 0, 0))
        with patch('excel_sheet_merger.gui.messagebox.showinfo') as done_popup:
            self.app._consume_events()
        lines = self.app.log.get('1.0', 'end')
        self.assertLess(lines.index('before-progress'), lines.index('after-progress'))
        self.assertLess(lines.index('after-progress'), lines.index('완료 1개'))
        self.assertEqual(self.app.state_text.get(), '완료 1개 / 이름 변경 0개 / 실패 0개')
        self.assertFalse(self.app.busy)
        done_popup.assert_called_once()


if __name__ == '__main__':
    unittest.main()
