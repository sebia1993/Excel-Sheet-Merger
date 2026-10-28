"""Reject pathological Excel ZIP packages before openpyxl expands them.

Only the ZIP central directory is read. These deliberately conservative
resource limits are not a proof of archive integrity or an anti-virus scan.
Normal Excel ZIP members and workbook contents are never rewritten.
"""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

MAX_ARCHIVE_MEMBERS = 30_000
MAX_EXPANDED_PACKAGE_BYTES = 1024 * 1024 * 1024
MAX_EXPANDED_XML_BYTES = 256 * 1024 * 1024
SUSPICIOUS_EXPANDED_MEMBER_BYTES = 64 * 1024 * 1024
MAX_SUSPICIOUS_COMPRESSION_RATIO = 400


def validate_xlsx_archive(path: Path) -> None:
    """Reject extreme or ambiguous ZIP metadata without decompressing data.

    The in-memory workbook loader can expand and allocate far more memory than
    the compressed XLSX's on-disk size suggests. Files under ordinary resource
    limits proceed into the existing parser without changes to conversion.
    """
    with ZipFile(path, 'r') as archive:
        members = archive.infolist()
        if len(members) > MAX_ARCHIVE_MEMBERS:
            raise ValueError("Excel 압축 파일에 비정상적으로 많은 구성 요소가 있습니다")

        expanded_total = 0
        seen_parts: set[str] = set()
        for member in members:
            name = member.filename
            if name in seen_parts:
                raise ValueError("Excel 압축 파일에 중복된 구성 요소가 있습니다")
            seen_parts.add(name)
            if member.is_dir():
                continue

            expanded_total += member.file_size
            if expanded_total > MAX_EXPANDED_PACKAGE_BYTES:
                raise ValueError("Excel 압축 해제 크기가 비정상적으로 큽니다 (전체 용량 한도)")
            if name.endswith(('.xml', '.rels')) and member.file_size > MAX_EXPANDED_XML_BYTES:
                raise ValueError("Excel 압축 해제 크기가 비정상적으로 큽니다 (XML 용량 한도)")
            if (member.file_size > SUSPICIOUS_EXPANDED_MEMBER_BYTES
                    and member.file_size > member.compress_size * MAX_SUSPICIOUS_COMPRESSION_RATIO):
                raise ValueError("Excel 압축 해제 크기가 비정상적으로 큽니다 (압축 확장 비율)")
