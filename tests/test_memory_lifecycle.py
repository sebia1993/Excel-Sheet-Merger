"""Large source cells should stop consuming memory while their results are copied.

These tests exercise real openpyxl workbooks, not a fake data model; observing
source-cell ownership during transfer catches a future regression to the
old high-peak-memory algorithm.
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

from excel_sheet_merger.merge import _append_openpyxl


class ExcelSourceMemoryLifecycleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / 'A사업장_20261028_11시.xlsx'

    def test_live_source_cell_ownership_decreases_during_copy(self):
        book = Workbook()
        for i in range(1, 401):
            for j in range(1, 11):
                book.active.cell(i, j, f'업무_{i}_{j}')
        book.save(self.source)
        book.close()

        loaded = []
        target_book = Workbook()
        target = target_book.active
        original_cell = target.cell
        calls = 0
        live_at_midpoint = None

        def load_source(*args, **kwargs):
            source = load_workbook(*args, **kwargs)
            loaded.append(source)
            return source

        def tracking_copy(*args, **kwargs):
            nonlocal calls, live_at_midpoint
            calls += 1
            if calls == 2000:
                live_at_midpoint = len(loaded[0].active._cells)
            return original_cell(*args, **kwargs)

        with patch('excel_sheet_merger.merge.load_workbook', side_effect=load_source):
            with patch.object(target, 'cell', side_effect=tracking_copy):
                _append_openpyxl(target, self.source)
        self.assertEqual(calls, 4000)
        self.assertIsNotNone(live_at_midpoint)
        self.assertLessEqual(live_at_midpoint, 2500,
                             'source cells must be released before destination copy finishes')
        self.assertEqual(target['J400'].value, '업무_400_10')
        target_book.close()

    def test_styles_comments_merges_and_sparse_cells_survive_copy(self):
        book = Workbook()
        ws = book.active
        ws['B2'] = '=1+2'
        ws['B2'].font = Font(bold=True, color='008000')
        ws['C3'] = '문서 링크'
        ws['C3'].hyperlink = 'https://example.com/docs'
        ws['C3'].comment = Comment('관리 메모', 'QA')
        ws['Z200'] = '마지막 값'
        ws['D8'] = '헤더'
        ws.merge_cells('D8:F8')
        ws['H10'].fill = PatternFill(fill_type='solid', fgColor='FF00CC')
        ws.column_dimensions['B'].width = 32
        ws.row_dimensions[8].height = 44
        ws.freeze_panes = 'C5'
        ws.auto_filter.ref = 'B2:C3'
        book.save(self.source)
        book.close()
        saved_source = self.source.read_bytes()

        target_book = Workbook()
        _append_openpyxl(target_book.active, self.source)
        output = self.folder / 'copied.xlsx'
        target_book.save(output)
        target_book.close()
        copied = load_workbook(output)
        try:
            sheet = copied.active
            self.assertEqual(sheet['B2'].value, '=1+2')
            self.assertTrue(sheet['B2'].font.bold)
            self.assertEqual(sheet['C3'].hyperlink.target, 'https://example.com/docs')
            self.assertEqual(sheet['C3'].comment.text, '관리 메모')
            self.assertEqual(sheet['Z200'].value, '마지막 값')
            self.assertIn('D8:F8', {str(r) for r in sheet.merged_cells.ranges})
            self.assertEqual(sheet['H10'].fill.fgColor.rgb, '00FF00CC')
            self.assertEqual(sheet.column_dimensions['B'].width, 32)
            self.assertEqual(sheet.row_dimensions[8].height, 44)
            self.assertEqual(sheet.freeze_panes, 'C5')
            self.assertEqual(sheet.auto_filter.ref, 'B2:C3')
        finally:
            copied.close()
        self.assertEqual(self.source.read_bytes(), saved_source,
                         'the on-disk source must remain unmodified')


    def test_blank_comment_cell_survives_excel_sheet_copy(self):
        """A note-only blank cell is user data even though it has no style."""
        from openpyxl.comments import Comment
        from openpyxl import Workbook, load_workbook

        source_book = Workbook()
        ws = source_book.active
        ws['A1'] = '검증 데이터'
        ws['D20'].comment = Comment('점검 완료 후 검토 바랍니다', '담당자')
        self.assertIsNone(ws['D20'].value)
        self.assertFalse(ws['D20'].has_style)
        source_book.save(self.source)
        source_book.close()
        original_bytes = self.source.read_bytes()

        result = Workbook()
        _append_openpyxl(result.active, self.source)
        output = self.folder / 'with-comment.xlsx'
        result.save(output)
        result.close()

        reopened = load_workbook(output)
        try:
            self.assertEqual(reopened.active['A1'].value, '검증 데이터')
            self.assertIsNone(reopened.active['D20'].value)
            self.assertIsNotNone(reopened.active['D20'].comment,
                                 'Note-only blank cells must not silently disappear')
            self.assertEqual(reopened.active['D20'].comment.text,
                             '점검 완료 후 검토 바랍니다')
            self.assertEqual(reopened.active['D20'].comment.author, '담당자')
        finally:
            reopened.close()
        self.assertEqual(self.source.read_bytes(), original_bytes)

    def test_comment_only_cell_survives_public_group_merge_and_renamed_output(self):
        from openpyxl.comments import Comment
        from excel_sheet_merger.classify import scan_folder
        from excel_sheet_merger.merge import merge_group

        source_path = self.folder / 'A사업장_20261029_11시.xlsx'
        excel = Workbook()
        excel.active['A1'] = '시트 원본'
        excel.active['D20'].comment = Comment('값은 없지만 필요한 업무 메모', '검토팀')
        excel.save(source_path)
        excel.close()
        raw = source_path.read_bytes()
        (self.folder / 'A사업장_20261029_14시.csv').write_text(
            'ID,IP\n00123,192.0.2.9\n', encoding='utf-8')
        group = scan_folder(self.folder).groups[0]
        first = merge_group(group, self.folder / '병합결과')
        second = merge_group(group, self.folder / '병합결과')
        self.assertEqual(first.path.name, 'A사업장_20261029.xlsx')
        self.assertEqual(second.path.name, 'A사업장_20261029 (2).xlsx')
        for result in (first, second):
            merged = load_workbook(result.path)
            try:
                self.assertEqual(merged.sheetnames, ['11시', '14시'])
                self.assertIsNone(merged['11시']['D20'].value)
                self.assertEqual(merged['11시']['D20'].comment.text,
                                 '값은 없지만 필요한 업무 메모')
                self.assertEqual(merged['11시']['D20'].comment.author, '검토팀')
                self.assertEqual(merged['14시']['A2'].value, '00123')
            finally:
                merged.close()
        self.assertEqual(source_path.read_bytes(), raw)


if __name__ == '__main__':
    unittest.main()
