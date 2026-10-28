"""Convert file handling exceptions into safe, actionable Korean GUI messages.

A CP949 decoding failure is not proof that the file is encrypted. Both a
protected document and a damaged/incorrectly encoded file can produce it.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from typing import TYPE_CHECKING

from .merge import SourceFileError

if TYPE_CHECKING:
    from .classify import MergeGroup


@dataclass(frozen=True)
class FailureGuidance:
    filename: str
    code: str
    title: str
    reason: str
    action: str
    technical: str

    def one_line(self) -> str:
        return f"{self.filename} — {self.title}"


def _cause_chain(error: Exception) -> tuple[BaseException, ...]:
    """Walk wrapped causes without parsing locale-sensitive exception strings."""
    causes: list[BaseException] = []
    seen: set[int] = set()
    current: BaseException | None = error
    while current is not None and id(current) not in seen:
        causes.append(current)
        seen.add(id(current))
        current = current.__cause__ or current.__context__
    return tuple(causes)


def diagnose_failure(error: Exception, group: MergeGroup | None = None) -> FailureGuidance:
    """Classify the *type* of error, not whether the file really is DRM-encrypted."""
    filename = getattr(error, 'filename', None) if isinstance(error, SourceFileError) else None
    technical = str(error) or type(error).__name__
    if not filename and group is not None:
        filename = next((item.file_name for item in group.inputs if item.file_name in technical), None)
    filename = filename or (group.output_name if group is not None else '전체 작업')
    chain = _cause_chain(error)
    message = technical.lower()
    any_error = lambda kind: any(isinstance(item, kind) for item in chain)

    if any_error(FileNotFoundError):
        return FailureGuidance(filename, 'MISSING_INPUT', '원본 파일을 찾을 수 없음',
                               '파일 분석 이후 원본 파일이 이동·삭제되었거나 접근할 수 없습니다.',
                               '원본 파일이 존재하는지 확인하고 폴더를 다시 분석해 주세요.', technical)
    if any_error(PermissionError):
        return FailureGuidance(filename, 'PERMISSION', '파일 접근 권한 없음',
                               '프로그램이 원본을 열거나 결과를 저장할 권한이 없습니다.',
                               '파일이 다른 프로그램에서 잠겨 있는지 확인하고, 회사에서 허용한 접근 권한을 확인해 주세요.', technical)
    if any_error(UnicodeError) and filename.lower().endswith('.csv'):
        decode = next((item for item in chain if isinstance(item, UnicodeError)), None)
        codec = getattr(decode, 'encoding', '문자 인코딩').upper()
        return FailureGuidance(filename, 'CSV_ENCODING', 'CSV 문자 인코딩 오류',
                               f'파일 데이터를 {codec}로 해석할 수 없습니다. 문서보안 프로그램으로 보호되었거나 '
                               '인코딩/파일 형식이 맞지 않을 수 있습니다. 이 오류만으로 암호화 여부는 확정할 수 없습니다.',
                               '회사 보안 정책을 준수하여 열람 권한이 있는 프로그램에서 정상 CSV/Excel로 내보내고 다시 시도하거나, 문서보안 담당자에게 문의해 주세요.',
                               technical)
    if '파일 분석 이후 원본 파일' in technical or '파일 분석 중 원본 파일' in technical:
        return FailureGuidance(filename, 'SOURCE_CHANGED', '원본 파일이 변경됨',
                               '파일 분석 이후 원본 파일 내용이 변경되어 이전 분석 결과를 사용할 수 없습니다.',
                               '원본 파일 저장이 완료된 후 파일 분석을 다시 실행하고 병합해 주세요.',
                               technical)
    if any_error(csv.Error):
        return FailureGuidance(filename, 'CSV_STRUCTURE', 'CSV 데이터 구조 오류',
                               'CSV의 쉼표·따옴표·행 구성이 올바르지 않거나 데이터가 손상되었습니다.',
                               '원본을 승인된 프로그램에서 확인하고 올바른 CSV 형식으로 다시 저장한 뒤 재시도해 주세요.', technical)
    if '시트는 1개여야' in technical:
        return FailureGuidance(filename, 'MULTIPLE_SHEETS', 'Excel 시트 개수 제한',
                               '이 프로그램은 원본 Excel 파일당 시트 1개만 병합할 수 있습니다.',
                               '원본 Excel에서 병합하려는 시트만 1개 남긴 별도 사본을 준비해 주세요.', technical)
    if '비어 있는 csv' in message or '비어 있는 excel' in message:
        return FailureGuidance(filename, 'EMPTY_INPUT', '원본 파일에 데이터가 없음',
                               '원본 파일이 비어 있거나 병합할 수 있는 행과 셀 값이 없습니다.',
                               '원본 파일 내용을 확인하고 데이터가 있는 파일을 선택해 주세요.', technical)
    if 'csv 행 수' in message or 'excel 행/열' in message or 'excel 셀' in message:
        return FailureGuidance(filename, 'EXCEL_LIMIT', 'Excel 데이터 크기 제한',
                               '행·열 또는 셀의 글자 수가 Excel 허용 범위를 초과했습니다.',
                               '회사 정책을 준수하며 원본 데이터를 여러 파일로 나누어 병합해 주세요.', technical)
    if 'excel 파일을 열 수 없' in message or 'xls 파일을 열 수 없' in message:
        return FailureGuidance(filename, 'EXCEL_OPEN', 'Excel 파일을 읽을 수 없음',
                               'Excel 파일 형식이 올바르지 않거나 손상·암호 보호·회사 문서보안 적용 가능성이 있습니다. 원인은 아직 확정되지 않았습니다.',
                               '회사에서 허용된 프로그램으로 정상적으로 열리는지 확인하고, 접근 권한이 없다면 문서보안 담당자에게 문의해 주세요.', technical)
    if 'utf-32 csv' in message:
        return FailureGuidance(filename, 'UNSUPPORTED_ENCODING', '지원하지 않는 CSV 인코딩',
                               'UTF-32 형식 CSV는 지원하지 않습니다.',
                               '승인된 도구에서 UTF-8 또는 CP949 CSV로 다시 저장해 주세요.', technical)
    return FailureGuidance(filename, 'UNKNOWN', '병합 중 예상하지 못한 오류',
                           '오류 정보만으로 정확한 원인을 특정하지 못했습니다.',
                           '작업 기록에 남은 상세 오류를 확인해 내부 프로그램·문서보안 담당자에게 문의해 주세요.', technical)
