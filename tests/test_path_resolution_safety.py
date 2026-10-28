"""Path canonicalization failure must never leave a stale selectable preview."""
from __future__ import annotations

import errno
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

GUI_AVAILABLE = os.name == 'nt' or bool(os.environ.get('DISPLAY'))


@unittest.skipUnless(GUI_AVAILABLE, 'needs a GUI display')
class PathResolutionStabilityTests(unittest.TestCase):
    def setUp(self):
        from excel_sheet_merger.gui import MergeApp
        self.app = MergeApp()
        self.addCleanup(self.app.destroy)
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        for hour in ('11시', '14시'):
            (self.folder / f'A사업장_20261029_{hour}.csv').write_text(
                'ID,IP\n00123,192.0.2.1\n', encoding='utf-8')
        self.app.folder.set(str(self.folder))

    def _scan_ok(self):
        self.app._scan()
        self.assertIsNotNone(self.app.report)
        self.assertEqual(len(self.app.report.groups), 1)

    def test_merge_start_resolution_oserror_clears_preview_and_reports_failure(self):
        self._scan_ok()
        gid = next(iter(self.app.group_map))
        self.app._toggle_group(gid)
        with patch('excel_sheet_merger.gui.Path.resolve',
                   side_effect=OSError(errno.ENETUNREACH, 'network path unavailable')), \
             patch('excel_sheet_merger.gui.threading.Thread') as worker, \
             patch('excel_sheet_merger.gui.messagebox.showerror') as dialog:
            self.app._start_merge(True)
        worker.assert_not_called()
        dialog.assert_called_once()
        self.assertIn('폴더', dialog.call_args.args[1])
        self.assertIn(str(self.folder), dialog.call_args.args[1])
        self.assertIsNone(self.app.report)
        self.assertEqual(self.app.group_map, {})
        self.assertEqual(self.app.checked_group_ids, set())
        self.assertEqual(str(self.app.all_button['state']), 'disabled')
        self.assertEqual(str(self.app.selected_button['state']), 'disabled')
        self.assertFalse(self.app.busy)
        self.assertFalse((self.folder / '병합결과').exists())

    def test_scan_resolution_failure_does_not_commit_half_valid_report(self):
        # Even after the scanner returns, canonicalization can encounter a
        # transient SMB/junction failure. A new scan must remain invalid.
        with patch('excel_sheet_merger.gui.Path.resolve',
                   side_effect=OSError(errno.ENOTCONN, 'share not connected')), \
             patch('excel_sheet_merger.gui.messagebox.showerror') as dialog:
            self.app._scan()
        dialog.assert_called_once()
        self.assertIn('폴더', dialog.call_args.args[1])
        self.assertIsNone(self.app.report)
        self.assertIsNone(self.app._scanned_folder)
        self.assertEqual(self.app.group_map, {})
        self.assertFalse(self.app.checked_group_ids)
        self.assertEqual(str(self.app.all_button['state']), 'disabled')
        self.assertFalse((self.folder / '병합결과').exists())
        self._scan_ok()  # A restored share must still work normally.

    def test_merge_start_resolution_runtime_error_is_recoverable(self):
        self._scan_ok()
        with patch('excel_sheet_merger.gui.Path.resolve',
                   side_effect=RuntimeError('Symlink loop')), \
             patch('excel_sheet_merger.gui.threading.Thread') as worker, \
             patch('excel_sheet_merger.gui.messagebox.showerror') as dialog:
            self.app._start_merge(False)
        worker.assert_not_called()
        dialog.assert_called_once()
        self.assertIn('다시', dialog.call_args.args[1])
        self.assertFalse(self.app.busy)
        self.assertIsNone(self.app.report)


if __name__ == '__main__':
    unittest.main()
