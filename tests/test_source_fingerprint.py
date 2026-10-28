"""Guard against source modifications that preserve filesystem metadata."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from openpyxl import Workbook, load_workbook
from excel_sheet_merger.classify import scan_folder
from excel_sheet_merger.merge import SourceFileError, merge_group, _verify_source
from excel_sheet_merger.failure_guidance import diagnose_failure


class ContentFingerprintTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def csv(self, hour, data):
        p = self.root / f'A사업장_20261028_{hour}.csv'
        p.write_bytes(data)
        return p

    @staticmethod
    def mutate_with_preserved_metadata(path: Path, replacement: bytes, offset: int = 0):
        before = path.stat()
        with path.open('r+b') as output:
            output.seek(offset)
            output.write(replacement)
            output.flush()
            os.fsync(output.fileno())
        os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
        after = path.stat()
        assert (after.st_size, after.st_mtime_ns, after.st_dev, after.st_ino) == (
            before.st_size, before.st_mtime_ns, before.st_dev, before.st_ino
        )

    def test_same_size_same_mtime_mutated_csv_is_not_published(self):
        source = self.csv('11시', b'ID,VALUE\n0001,HELLO\n')
        self.csv('14시', b'ID,VALUE\n0002,WORLD\n')
        report = scan_folder(self.root)
        self.mutate_with_preserved_metadata(source, b'9999', offset=9)
        with self.assertRaisesRegex(SourceFileError, '변경'):
            merge_group(report.groups[0], self.root / '병합결과')
        self.assertFalse(list((self.root / '병합결과').glob('*.xlsx')))

    def test_large_source_mutated_near_middle_with_same_metadata_is_detected(self):
        big = self.csv('11시', b'A' * (9 * 1024 * 1024))
        self.csv('14시', b'ID\n1\n')
        report = scan_folder(self.root)
        source = report.groups[0].inputs[0]
        self.mutate_with_preserved_metadata(big, b'REPLACEMENT', offset=big.stat().st_size // 2)
        with self.assertRaisesRegex(SourceFileError, '변경'):
            _verify_source(source)

    def test_unchanged_mixed_inputs_keep_sheet_names_values_and_formulas(self):
        self.csv('11시', 'IP,코드\n192.0.2.5,00123\n'.encode('utf-8'))
        wb = Workbook()
        wb.active['A1'] = '원본'
        wb.active['B2'] = '=1+2'
        wb.save(self.root / 'A사업장_20261028_14시.xlsx')
        wb.close()
        merged = merge_group(scan_folder(self.root).groups[0], self.root / '병합결과')
        result = load_workbook(merged.path)
        try:
            self.assertEqual(result.sheetnames, ['11시', '14시'])
            self.assertEqual(result['11시']['B2'].value, '00123')
            self.assertEqual(result['14시']['B2'].value, '=1+2')
        finally:
            result.close()

    def test_changed_source_error_explains_rescan_in_korean(self):
        error = SourceFileError('A사업장_20261028_11시.csv', ValueError(
            '파일 분석 이후 원본 파일 내용이 변경되었습니다. 다시 파일 분석을 실행하세요.'))
        advice = diagnose_failure(error)
        self.assertEqual(advice.code, 'SOURCE_CHANGED')
        self.assertEqual(advice.filename, 'A사업장_20261028_11시.csv')
        self.assertIn('변경', advice.reason)
        self.assertIn('파일 분석', advice.action)

    def test_scan_with_unreadable_fingerprint_keeps_group_for_existing_merge_errors(self):
        self.csv('11시', b'col\n1\n')
        self.csv('14시', b'col\n2\n')
        # Corporate file security may permit stat() but reject byte reads.
        # A new fingerprint check must not silently discard an otherwise
        # validly named source at the analysis stage.
        with patch('excel_sheet_merger.classify.file_fingerprint', side_effect=PermissionError('restricted')):
            report = scan_folder(self.root)
        self.assertEqual(len(report.groups), 1)
        self.assertTrue(all(item.content_fingerprint is None for item in report.groups[0].inputs))
        output = merge_group(report.groups[0], self.root / '병합결과')
        self.assertTrue(output.path.is_file())


if __name__ == '__main__':
    unittest.main()
