"""유효한 보안 문서 진단 범위를 지키면서 사용자에게 도움되는 실패 설명 제공."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from excel_sheet_merger.classify import scan_folder
from excel_sheet_merger.merge import merge_group


class FailureGuidanceTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.dir.cleanup)
        self.root = Path(self.dir.name)
        self.good_csv = self.root / 'A사업장_20261028_11시.csv'
        self.good_csv.write_text('ID,설명\n00123,한글\n', encoding='cp949')

    def _group(self):
        return scan_folder(self.root).groups[0]

    def _failing_guidance(self):
        from excel_sheet_merger.failure_guidance import diagnose_failure
        from excel_sheet_merger.merge import SourceFileError
        try:
            merge_group(self._group(), self.root / '병합결과')
        except SourceFileError as exc:
            return exc, diagnose_failure(exc, self._group())
        self.fail('그룹이 성공해 손상된 파일 재현 테스트를 통과하지 못했습니다')

    def test_cp949_invalid_bytes_reports_specific_file_security_possible_not_certain(self):
        damaged = self.root / 'A사업장_20261028_14시.csv'
        damaged.write_bytes(b'ID\n\xff\x80\xff\x80\n')
        error, result = self._failing_guidance()
        self.assertEqual(error.filename, damaged.name)
        self.assertEqual(result.filename, damaged.name)
        self.assertEqual(result.code, 'CSV_ENCODING')
        self.assertIn('인코딩', result.title)
        self.assertIn('CP949', result.reason)
        self.assertIn('문서보안', result.reason)
        self.assertIn('확정', result.reason)
        self.assertIn('권한', result.action)
        self.assertIn('담당자', result.action)
        self.assertIn('cp949', result.technical.lower())
        self.assertEqual(self.good_csv.read_bytes(), 'ID,설명\n00123,한글\n'.encode('cp949'))
        self.assertEqual(list((self.root / '병합결과').glob('*.xlsx')), [])

    def test_empty_csv_has_correct_filename_and_no_false_protection_diagnosis(self):
        empty = self.root / 'A사업장_20261028_14시.csv'
        empty.write_bytes(b'')
        _, advice = self._failing_guidance()
        self.assertEqual(advice.filename, empty.name)
        self.assertIn('비어', advice.reason)
        self.assertNotIn('암호화', advice.reason)
        self.assertIn('확인', advice.action)

    def test_malformed_csv_has_structure_warning_not_encoding(self):
        broken = self.root / 'A사업장_20261028_14시.csv'
        broken.write_bytes(b'ID,comment\n1,"broken quote\n')
        _, advice = self._failing_guidance()
        self.assertEqual(advice.filename, broken.name)
        self.assertEqual(advice.code, 'CSV_STRUCTURE')
        self.assertIn('CSV', advice.reason)
        self.assertNotIn('암호화', advice.reason)

    def test_invalid_xlsx_reports_excel_open_issue_not_cp949(self):
        excel = self.root / 'A사업장_20261028_14시.xlsx'
        excel.write_bytes(b'not a zip workbook')
        _, advice = self._failing_guidance()
        self.assertEqual(advice.filename, excel.name)
        self.assertEqual(advice.code, 'EXCEL_OPEN')
        self.assertIn('형식', advice.reason)
        self.assertNotIn('CP949', advice.reason)
        self.assertIn('보안', advice.action)

    def test_multiple_sheets_reports_one_sheet_rule(self):
        excel = self.root / 'A사업장_20261028_14시.xlsx'
        book = Workbook()
        book.active['A1'] = 'hello'
        book.create_sheet('another')
        book.save(excel)
        book.close()
        _, advice = self._failing_guidance()
        self.assertEqual(advice.filename, excel.name)
        self.assertEqual(advice.code, 'MULTIPLE_SHEETS')
        self.assertIn('1개', advice.action)

    def test_missing_input_reports_file_not_found(self):
        disappearing = self.root / 'A사업장_20261028_14시.csv'
        disappearing.write_text('ID\n7\n',encoding='utf-8')
        group = self._group()
        disappearing.unlink()
        from excel_sheet_merger.failure_guidance import diagnose_failure
        from excel_sheet_merger.merge import SourceFileError
        with self.assertRaises(SourceFileError) as context:
            merge_group(group, self.root / '병합결과')
        advice = diagnose_failure(context.exception, group)
        self.assertEqual(advice.filename, disappearing.name)
        self.assertEqual(advice.code, 'MISSING_INPUT')
        self.assertIn('존재', advice.action)

    def test_permission_error_has_access_remediation_without_claiming_encryption(self):
        from excel_sheet_merger.failure_guidance import diagnose_failure
        from excel_sheet_merger.merge import SourceFileError
        try:
            try:
                raise PermissionError('permission denied')
            except PermissionError as exc:
                raise SourceFileError('B사업장_20261028_14시.xlsx', exc) from exc
        except SourceFileError as failure:
            advice = diagnose_failure(failure)
        self.assertEqual(advice.filename, 'B사업장_20261028_14시.xlsx')
        self.assertEqual(advice.code, 'PERMISSION')
        self.assertIn('권한', advice.reason)
        self.assertIn('허용', advice.action)
        self.assertNotIn('암호화', advice.reason)

    def test_unexpected_error_does_not_assert_security_encryption(self):
        from excel_sheet_merger.failure_guidance import diagnose_failure
        advice = diagnose_failure(OSError('unexpected disk crash'), None)
        self.assertEqual(advice.code, 'UNKNOWN')
        self.assertIn('unexpected disk crash', advice.technical)
        self.assertNotIn('암호화', advice.reason)


if __name__ == '__main__':
    unittest.main()
