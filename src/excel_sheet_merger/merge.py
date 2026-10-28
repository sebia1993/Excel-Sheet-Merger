"""CSV 파일을 인코딩을 보존해 읽고 시간별 시트로 XLSX에 저장한다.

CSV는 셀 서식과 자료형을 정의하지 않으므로, 값의 변형을 막기 위해 모두
텍스트로 저장한다. Excel 데스크톱 프로그램이나 COM은 사용하지 않는다.
"""

from __future__ import annotations

import codecs
import csv
import os
import re
import uuid
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Iterator

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from .classify import MergeGroup

MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384
MAX_CELL_UTF16_UNITS = 32_767
_ILLEGAL_EXCEL_TEXT = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class MergeStatus(str, Enum):
    CREATED = "생성 완료"
    EXISTS = "기존 파일 존재로 건너뜀"


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
    if _ILLEGAL_EXCEL_TEXT.search(value):
        raise ValueError("Excel에서 사용할 수 없는 제어 문자가 있습니다")
    if len(value.encode("utf-16-le")) // 2 > MAX_CELL_UTF16_UNITS:
        raise ValueError("Excel 셀 최대 길이(32,767자)를 초과했습니다")
    cell = WriteOnlyCell(ws, value=value)
    # 문자열 '=...'을 수식으로 추론하지 않도록 명시적으로 Excel 텍스트로 고정.
    cell.data_type = "s"
    return cell


def merge_group(group: MergeGroup, output_dir: Path) -> MergeResult:
    """그룹당 한 번만 생성; 읽기 실패 시 최종 파일을 만들지 않는다."""
    if len(group.inputs) < 2:
        raise ValueError("병합 대상 CSV가 2개 미만입니다")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    final_path = output_dir / group.output_name
    if final_path.exists():
        return MergeResult(final_path, MergeStatus.EXISTS)
    tmp_path = output_dir / f"_병합중_{uuid.uuid4().hex}.xlsx"
    wb = Workbook(write_only=True)
    try:
        for item in group.inputs:
            ws = wb.create_sheet(title=item.sheet_name)
            row_count = 0
            for row in csv_rows(item.path):
                row_count += 1
                if row_count > MAX_ROWS:
                    raise ValueError(f"CSV 행 수가 Excel 한도를 넘었습니다: {item.file_name}")
                if len(row) > MAX_COLUMNS:
                    raise ValueError(f"CSV 열 수가 Excel 한도를 넘었습니다: {item.file_name}")
                ws.append([text_cell(ws, value) for value in row])
            if not row_count:
                raise ValueError(f"비어 있는 CSV 파일입니다: {item.file_name}")
        wb.save(tmp_path)
        # Windows에서는 os.rename이 기존 파일을 덮어쓰지 않음.
        # POSIX에서는 atomic link를 사용해 no-clobber 동작을 보장.
        if os.name == "nt":
            try:
                os.rename(tmp_path, final_path)
            except FileExistsError:
                return MergeResult(final_path, MergeStatus.EXISTS)
        else:
            try:
                os.link(tmp_path, final_path)
            except FileExistsError:
                return MergeResult(final_path, MergeStatus.EXISTS)
        return MergeResult(final_path, MergeStatus.CREATED)
    except Exception:
        # write_only 워크시트의 열린 XML 제너레이터를 닫지 않으면,
        # 실패 후 GC 단계에서 lxml 예외가 stderr에 출력된다.
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
