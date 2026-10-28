"""Excel 프로그램 없이 병합 트랜잭션 동작을 재현하는 테스트."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from excel_sheet_merger.classify import scan_folder
from excel_sheet_merger.merge import merge_group, MergeStatus


class FakeSheet:
    def __init__(self, owner, name):
        self.owner = owner
        self.Name = name

    def Copy(self, *, After):
        dest = After.owner
        dest.sheets.append(FakeSheet(dest, self.Name))

    def Delete(self):
        self.owner.sheets.remove(self)


class FakeSheets:
    def __init__(self, book):
        self.book = book

    @property
    def Count(self):
        return len(self.book.sheets)

    def __call__(self, idx):
        return self.book.sheets[idx - 1]


class FakeBook:
    def __init__(self, names=('Sheet1',), fail_save=False):
        self.sheets = [FakeSheet(self, n) for n in names]
        self.Sheets = FakeSheets(self)
        self.closed = False
        self.fail_save = fail_save

    def SaveAs(self, *, Filename, FileFormat):
        if self.fail_save:
            Path(Filename).write_text('partial')
            raise RuntimeError('시뮬레이션 저장 실패')
        Path(Filename).write_text(json.dumps([s.Name for s in self.sheets], ensure_ascii=False), encoding='utf-8')

    def Close(self, *, SaveChanges):
        self.closed = True


class FakeBooks:
    def __init__(self, sheet_count=1, fail_save=False):
        self.sheet_count = sheet_count
        self.fail_save = fail_save
        self.created = []
        self.opened = []

    def Add(self):
        b = FakeBook(fail_save=self.fail_save)
        self.created.append(b)
        return b

    def Open(self, *, Filename, UpdateLinks, ReadOnly, IgnoreReadOnlyRecommended, AddToMru):
        assert ReadOnly and UpdateLinks == 0 and not AddToMru
        b = FakeBook(names=tuple(f'원본{i+1}' for i in range(self.sheet_count)))
        self.opened.append(b)
        return b


class FakeExcel:
    def __init__(self, **kwargs):
        self.Workbooks = FakeBooks(**kwargs)


class MergeTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        self.folder = Path(self.tempdir.name)
        self.names = ['A사업장_20261028_14시.xlsx','A사업장_20261028_11시.xlsx','A사업장_20261028_17시.xlsx']
        for name in self.names:
            (self.folder / name).write_text('ORIGINAL', encoding='utf-8')
        self.group = scan_folder(self.folder).groups[0]
        self.output = self.folder / '병합결과'

    def test_merges_in_time_order_and_preserves_sources(self):
        excel = FakeExcel()
        result = merge_group(self.group, self.output, excel)
        self.assertEqual(result.status, MergeStatus.CREATED)
        self.assertEqual(json.loads(result.path.read_text(encoding='utf-8')), ['11시','14시','17시'])
        self.assertTrue(all((self.folder / n).read_text(encoding='utf-8') == 'ORIGINAL' for n in self.names))
        self.assertEqual(len(excel.Workbooks.opened), 3)
        self.assertTrue(excel.Workbooks.created[0].closed)
        self.assertTrue(all(b.closed for b in excel.Workbooks.opened))
        self.assertEqual(list(self.output.glob('*.xlsx')), [result.path])

    def test_existing_result_is_never_overwritten(self):
        self.output.mkdir()
        dst = self.output / self.group.output_name
        dst.write_text('DO_NOT_TOUCH', encoding='utf-8')
        result = merge_group(self.group, self.output, FakeExcel())
        self.assertEqual(result.status, MergeStatus.EXISTS)
        self.assertEqual(dst.read_text(encoding='utf-8'), 'DO_NOT_TOUCH')

    def test_multiple_sheets_in_one_source_fails_atomically(self):
        excel = FakeExcel(sheet_count=2)
        with self.assertRaisesRegex(ValueError, '시트.*1개'):
            merge_group(self.group, self.output, excel)
        self.assertFalse((self.output / self.group.output_name).exists())
        self.assertEqual(list(self.output.glob('*.xlsx')), [])
        self.assertTrue(excel.Workbooks.created[0].closed)
        self.assertTrue(all(b.closed for b in excel.Workbooks.opened))

    def test_save_failure_removes_temporary_files(self):
        excel = FakeExcel(fail_save=True)
        with self.assertRaisesRegex(RuntimeError, '저장 실패'):
            merge_group(self.group, self.output, excel)
        self.assertEqual(list(self.output.glob('*.xlsx')), [])
        self.assertFalse((self.output / self.group.output_name).exists())
        self.assertTrue(excel.Workbooks.created[0].closed)


if __name__ == '__main__':
    unittest.main()
