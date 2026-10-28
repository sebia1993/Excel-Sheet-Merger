import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))


class DistributionTests(unittest.TestCase):
    def test_gui_can_be_imported_without_windows_specific_libraries(self):
        from excel_sheet_merger.gui import run_app, summarize_scan
        self.assertTrue(callable(run_app))
        self.assertTrue(callable(summarize_scan))

    def test_merge_worker_uses_explicit_output_folder_not_tk_variables(self):
        import queue
        import tempfile
        from excel_sheet_merger.gui import MergeApp
        from excel_sheet_merger.classify import scan_folder
        from test_merge import FakeExcel
        with tempfile.TemporaryDirectory() as d:
            source = Path(d)
            for time in ('11시', '14시'):
                (source / f'A사업장_20261028_{time}.xlsx').write_text('keep')
            group = scan_folder(source).groups[0]
            app = MergeApp.__new__(MergeApp)
            app.events = queue.Queue()
            class ForbiddenFolder:
                def get(self):
                    raise AssertionError('worker must not access Tk variables')
            app.folder = ForbiddenFolder()
            results = app._process_group(FakeExcel(), group, source / '병합결과', 1, 1, 0, 0, 0, [])
            self.assertEqual(results, (1, 0, 0))

    def test_windows_builder_has_safety_checks_and_real_packager(self):
        workflow = (ROOT / '.github/workflows/windows-exe.yml').read_text(encoding='utf-8')
        self.assertIn('windows-latest', workflow)
        self.assertIn('pyinstaller', workflow.lower())
        self.assertIn('unittest discover', workflow)
        self.assertIn('upload-artifact', workflow)
        self.assertIn('python', workflow.lower())

    def test_main_branch_build_publishes_real_exe_to_github_release(self):
        workflow = (ROOT / '.github/workflows/windows-exe.yml').read_text(encoding='utf-8')
        self.assertIn('contents: write', workflow)
        self.assertIn("refs/heads/main", workflow)
        self.assertIn('gh release create', workflow)
        self.assertIn('ExcelSheetMerger.exe', workflow)
        self.assertIn('SHA256SUMS', workflow)
        self.assertIn('if-no-files-found: error', workflow)
        readme = (ROOT / 'README.md').read_text(encoding='utf-8')
        self.assertIn('releases/latest/download/ExcelSheetMerger.exe', readme)

    def test_readme_explains_excel_dependency_and_exe_build(self):
        readme = (ROOT / 'README.md').read_text(encoding='utf-8')
        self.assertIn('Microsoft Excel', readme)
        self.assertIn('PyInstaller', readme)
        self.assertIn('Windows 11', readme)
        self.assertIn('EXE', readme)


if __name__ == '__main__':
    unittest.main()
