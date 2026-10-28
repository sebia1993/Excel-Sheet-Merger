"""CSV와 Excel 파일의 데이터를 시간별 시트로 병합해 XLSX로 만든다.

CSV의 원본 텍스트 보존 규칙은 그대로 유지한다. 일반 Excel은 셀 자료형과
대표적인 셀 스타일·병합·행/열 크기를 복사한다. Excel COM은 사용하지 않는다.
"""

from __future__ import annotations

import codecs
from copy import copy
import csv
import os
import re
import uuid
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Iterator

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell

from .classify import MergeGroup

MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384
MAX_CELL_UTF16_UNITS = 32_767
_ILLEGAL_EXCEL_TEXT = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class MergeStatus(str, Enum):
    CREATED = "생성 완료"
    EXISTS = "기존 파일 존재로 건너뜀"  # Legacy status: no longer returned for name collisions.
    RENAMED = "기존 파일 존재로 이름 변경 생성"


@dataclass(frozen=True)
class MergeResult:
    path: Path
    status: MergeStatus


def detect_csv_encoding(path: Path) -> str:
    """UTF BOM을 우선하고, 전체 파일의 UTF-8 유효성을 검사한 후 CP949로 대체.

    첫 KB가 영어로만 된 파일이더라도 뒤쪽의 CP949 한글을 놓치지 않는다.
    """
    with path.open("rb") as file:
        bom = file.read(4)
        if bom.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
            raise ValueError(f"UTF-32 CSV는 지원하지 않습니다: {path.name}")
        if bom.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
            return "utf-16"
        if bom.startswith(codecs.BOM_UTF8):
            return "utf-8-sig"
        file.seek(0)
        decoder = codecs.getincrementaldecoder("utf-8")(errors="strict")
        try:
            for chunk in iter(lambda: file.read(65536), b""):
                decoder.decode(chunk)
            decoder.decode(b"", final=True)
            return "utf-8"
        except UnicodeDecodeError:
            return "cp949"


def csv_rows(path: Path) -> Iterator[list[str]]:
    encoding = detect_csv_encoding(path)
    try:
        with path.open("r", encoding=encoding, newline="", errors="strict") as file:
            reader = csv.reader(file, delimiter=",", strict=True)
            for first, row in enumerate(reader):
                # Microsoft Excel이 CSV 첫 행에 구분자를 명시한 파일을 지원한다.
                if first == 0 and row in (["sep=", ""], ["sep=,"]):
                    continue
                yield row
    except (UnicodeError, csv.Error) as exc:
        raise ValueError(f"CSV 데이터 또는 인코딩 오류 ({path.name}): {exc}") from exc


def text_cell(ws, value: str):
    # XML 1.0 파서는 원본 CRLF를 LF+LF로 읽을 수 있어 셀 내 줄바꿈을 LF로 통일한다.
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    if _ILLEGAL_EXCEL_TEXT.search(value):
        raise ValueError("Excel에서 사용할 수 없는 제어 문자가 있습니다")
    if len(value.encode("utf-16-le")) // 2 > MAX_CELL_UTF16_UNITS:
        raise ValueError("Excel 셀 최대 길이(32,767자)를 초과했습니다")
    cell = WriteOnlyCell(ws, value=value)
    # 문자열 '=...'을 수식으로 추론하지 않도록 명시적으로 Excel 텍스트로 고정.
    cell.data_type = "s"
    return cell


def _append_csv(ws, path: Path) -> None:
    """The existing CSV path remains text-only (including formula-looking text)."""
    row_count = 0
    for row in csv_rows(path):
        row_count += 1
        if row_count > MAX_ROWS:
            raise ValueError(f"CSV 행 수가 Excel 한도를 넘었습니다: {path.name}")
        if len(row) > MAX_COLUMNS:
            raise ValueError(f"CSV 열 수가 Excel 한도를 넘었습니다: {path.name}")
        ws.append([text_cell(ws, value) for value in row])
    if not row_count:
        raise ValueError(f"비어 있는 CSV 파일입니다: {path.name}")


def _append_openpyxl(ws, path: Path) -> None:
    """Copy the sole worksheet's ordinary cells and layout without executing Excel."""
    try:
        source_wb = load_workbook(path, read_only=False, data_only=False, keep_links=False)
    except Exception as exc:
        raise ValueError(f"Excel 파일을 열 수 없습니다 ({path.name}): {exc}") from exc
    try:
        if len(source_wb.worksheets) != 1:
            raise ValueError(f"Excel 원본의 시트는 1개여야 합니다: {path.name} (현재 {len(source_wb.worksheets)}개)")
        source = source_wb.worksheets[0]
        if source.max_row > MAX_ROWS or source.max_column > MAX_COLUMNS:
            raise ValueError(f"Excel 행/열 한도를 초과했습니다: {path.name}")
        has_data = False
        for row in source.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                if cell.value is not None:
                    has_data = True
                target = ws.cell(row=cell.row, column=cell.column, value=cell.value)
                # '=' as text stays text; real formulas remain formulas.
                if cell.data_type == 's' and isinstance(cell.value, str):
                    target.data_type = 's'
                if cell.has_style:
                    target.font = copy(cell.font)
                    target.fill = copy(cell.fill)
                    target.border = copy(cell.border)
                    target.alignment = copy(cell.alignment)
                    target.number_format = cell.number_format
                    target.protection = copy(cell.protection)
                if cell.hyperlink is not None:
                    target.hyperlink = copy(cell.hyperlink)
                if cell.comment is not None:
                    target.comment = copy(cell.comment)
        if not has_data:
            raise ValueError(f"비어 있는 Excel 시트입니다: {path.name}")
        for key, dim in source.column_dimensions.items():
            ws.column_dimensions[key] = copy(dim)
        for key, dim in source.row_dimensions.items():
            ws.row_dimensions[key] = copy(dim)
        for merged in source.merged_cells.ranges:
            ws.merge_cells(str(merged))
        ws.freeze_panes = source.freeze_panes
        ws.sheet_view.showGridLines = source.sheet_view.showGridLines
        if source.auto_filter.ref:
            ws.auto_filter.ref = source.auto_filter.ref
    finally:
        source_wb.close()


def _append_xls(ws, path: Path) -> None:
    """Read legacy BIFF .xls values; original .xls cell formatting is not retained."""
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError("구형 XLS 읽기 모듈(xlrd)이 프로그램에 포함되지 않았습니다") from exc
    try:
        book = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ValueError(f"구형 XLS 파일을 열 수 없습니다 ({path.name}): {exc}") from exc
    try:
        if book.nsheets != 1:
            raise ValueError(f"Excel 원본의 시트는 1개여야 합니다: {path.name} (현재 {book.nsheets}개)")
        sheet = book.sheet_by_index(0)
        if sheet.nrows == 0 or sheet.ncols == 0:
            raise ValueError(f"비어 있는 Excel 시트입니다: {path.name}")
        if sheet.nrows > MAX_ROWS or sheet.ncols > MAX_COLUMNS:
            raise ValueError(f"Excel 행/열 한도를 초과했습니다: {path.name}")
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                original = sheet.cell(r, c)
                typ, value = original.ctype, original.value
                if typ in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                if typ == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(value, book.datemode)
                elif typ == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                elif typ == xlrd.XL_CELL_ERROR:
                    value = xlrd.error_text_from_code.get(value, '#N/A')
                target = ws.cell(row=r + 1, column=c + 1, value=value)
                if typ == xlrd.XL_CELL_TEXT and isinstance(value, str):
                    # Legacy Excel formula *results* may be stored as text.
                    target.data_type = 's'
                elif typ == xlrd.XL_CELL_ERROR:
                    target.data_type = 'e'
                elif typ == xlrd.XL_CELL_DATE:
                    target.number_format = 'yyyy-mm-dd hh:mm:ss'
    finally:
        book.release_resources()


def merge_group(group: MergeGroup, output_dir: Path) -> MergeResult:
    """그룹당 한 번만 생성; 읽기 실패 시 최종 파일을 만들지 않는다."""
    if len(group.inputs) < 2:
        raise ValueError("병합 대상 파일이 2개 미만입니다")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    final_path = output_dir / group.output_name
    tmp_path = output_dir / f"_병합중_{uuid.uuid4().hex}.xlsx"
    # CSV-only keeps the original memory-efficient streaming writer.
    csv_only = all(item.path.suffix.casefold() == '.csv' for item in group.inputs)
    wb = Workbook(write_only=csv_only)
    if not csv_only:
        wb.remove(wb.active)  # no default "Sheet" in the output
    try:
        for item in group.inputs:
            ws = wb.create_sheet(title=item.sheet_name)
            extension = item.path.suffix.casefold()
            if extension == '.csv':
                _append_csv(ws, item.path)
            elif extension in ('.xlsx', '.xlsm'):
                _append_openpyxl(ws, item.path)
            elif extension == '.xls':
                _append_xls(ws, item.path)
            else:
                raise ValueError(f"지원하지 않는 Excel 형식입니다: {item.file_name}")
        wb.save(tmp_path)
        # Avoid TOCTOU: the final filename is claimed by an exclusive operation,
        # not by a prior exists() check. Another process may publish first.
        # Windows rename refuses an existing destination; POSIX link does too.
        candidate = final_path
        suffix = 2
        while True:
            try:
                if os.name == "nt":
                    os.rename(tmp_path, candidate)
                else:
                    os.link(tmp_path, candidate)
            except FileExistsError:
                candidate = final_path.with_name(f"{final_path.stem} ({suffix}){final_path.suffix}")
                suffix += 1
            else:
                status = MergeStatus.CREATED if candidate == final_path else MergeStatus.RENAMED
                return MergeResult(candidate, status)
    except Exception:
        # write_only 워크시트의 열린 XML 제너레이터를 닫지 않으면,
        # 실패 후 GC 단계에서 lxml 예외가 stderr에 출력된다.
        if csv_only:
            for ws in wb.worksheets:
                if not ws.closed:
                    try:
                        ws.close()
                    except Exception:
                        pass
        raise
    finally:
        wb.close()
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
