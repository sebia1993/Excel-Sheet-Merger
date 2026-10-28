"""파일명을 기반으로 CSV → XLSX 시트 병합 대상을 분류한다."""

from __future__ import annotations

import re
import hashlib
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path

DATE_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?:[-._]?\d{2}){2}(?!\d)")
TIME_PATTERNS = (
    re.compile(r"(?:오전|오후)?\s*(?:[01]?\d|2[0-3])시(?:[0-5]?\d분)?"),
    re.compile(r"(?:[01]?\d|2[0-3])[-_:][0-5]\d"),
    re.compile(r"(?:[01]\d|2[0-3])[0-5]\d"),
)
SUPPORTED_INPUT_SUFFIXES = {".csv", ".xlsx", ".xlsm", ".xls"}
SOURCE_SUFFIXES = SUPPORTED_INPUT_SUFFIXES | {".xlsb"}
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class SheetInput:
    path: Path
    site: str
    date: str
    hour: int
    minute: int
    sheet_name: str = ""
    snapshot: tuple[int, int, int, int] | None = None
    content_fingerprint: bytes | None = None

    @property
    def file_name(self) -> str:
        return self.path.name


@dataclass(frozen=True)
class MergeGroup:
    site: str
    date: str
    inputs: tuple[SheetInput, ...]
    output_name: str

    @property
    def key(self) -> str:
        return f"{self.site.casefold()}|{self.date}"


@dataclass(frozen=True)
class RejectedFile:
    name: str
    reason: str


@dataclass(frozen=True)
class ScanReport:
    groups: tuple[MergeGroup, ...]
    singles: tuple[MergeGroup, ...]
    rejected: tuple[RejectedFile, ...]
    checked: int


def _bounded(source: str, start: int, end: int) -> bool:
    return ((start == 0 or not source[start - 1].isalnum())
            and (end == len(source) or not source[end].isalnum()))


def _clock(token: str) -> tuple[int, int] | None:
    token = token.strip()
    period = None
    if token.startswith(("오전", "오후")):
        period = token[:2]
        token = token[2:].strip()
    if "시" in token:
        h, _, suffix = token.partition("시")
        hour = int(h)
        minute = int(suffix[:-1]) if suffix.endswith("분") else 0
    elif any(c in token for c in "-:_"):
        h, m = re.split(r"[-:_]", token)
        hour, minute = int(h), int(m)
    else:
        hour, minute = int(token[:2]), int(token[2:])
    if hour > 23 or minute > 59 or (period and not 1 <= hour <= 12):
        return None
    if period == "오전" and hour == 12:
        hour = 0
    elif period == "오후" and hour < 12:
        hour += 12
    return hour, minute


def _site_name(text: str) -> str:
    text = re.sub(r"[\s_]+", "_", text)
    return text.strip("_- .()[]")


def _safe_filename(name: str) -> str:
    name = INVALID_FILENAME_CHARS.sub("_", name).strip(" .")
    if not name:
        name = "사업장"
    if name.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", "COM1", "LPT1"}:
        name = "_" + name
    return name[:150].rstrip(" .") or "사업장"


def clean_sheet_name(name: str) -> str:
    name = INVALID_SHEET_CHARS.sub("-", name).strip(" '\t\r\n") or "시트"
    encoded = name.encode("utf-16-le")[:62]
    if len(encoded) >= 2:
        # 손상된 서로게이트 끝을 만들지 않도록 역순 제거한다.
        while True:
            try:
                return encoded.decode("utf-16-le") or "시트"
            except UnicodeDecodeError:
                encoded = encoded[:-2]
    return "시트"


def unique_sheet_names(names: list[str]) -> list[str]:
    used: set[str] = set()
    result: list[str] = []
    for name in names:
        base = clean_sheet_name(name)
        n = 1
        while True:
            suffix = f" ({n})" if n > 1 else ""
            # 잘림 및 대소문자 중복 방지, Excel 31 UTF-16단위 제한.
            limit = 31 - len(suffix.encode("utf-16-le")) // 2
            proposed = clean_sheet_name(base.encode("utf-16-le")[:limit * 2].decode("utf-16-le", errors="ignore") + suffix)
            if proposed.casefold() not in used:
                used.add(proposed.casefold())
                result.append(proposed)
                break
            n += 1
    return result


def parse_filename(path: Path) -> SheetInput:
    name = path.name
    if name.startswith("~$"):
        raise ValueError("Excel 임시 잠금 파일")
    if path.suffix.casefold() not in SUPPORTED_INPUT_SUFFIXES:
        raise ValueError("지원하지 않는 확장자 (CSV, XLSX, XLSM, XLS만 가능)")
    stem = path.stem
    dates = list(DATE_RE.finditer(stem))
    if len(dates) != 1:
        raise ValueError("날짜를 1개만 인식해야 합니다 (YYYYMMDD 등)")
    date_match = dates[0]
    date = re.sub(r"\D", "", date_match.group())
    try:
        if datetime.strptime(date, "%Y%m%d").strftime("%Y%m%d") != date:
            raise ValueError("날짜 값이 올바르지 않습니다")
    except ValueError as exc:
        raise ValueError("유효하지 않은 날짜") from exc
    rest = stem[:date_match.start()] + "_" + stem[date_match.end():]
    candidates: dict[tuple[int, int], tuple[int, int]] = {}
    for regex in TIME_PATTERNS:
        for match in regex.finditer(rest):
            if not _bounded(rest, match.start(), match.end()):
                continue
            tm = _clock(match.group())
            if tm is not None:
                candidates[(match.start(), match.end())] = tm
    if len(candidates) != 1:
        raise ValueError(f"시간 인식 불가 또는 여러 시간 발견 ({len(candidates)}개)")
    (start, end), (hour, minute) = next(iter(candidates.items()))
    site = _site_name(rest[:start] + "_" + rest[end:])
    if not site or not any(ch.isalpha() for ch in site):
        raise ValueError("사업장명을 인식할 수 없습니다")
    sheet_name = f"{hour:02d}시" + (f"{minute:02d}분" if minute else "")
    return SheetInput(path=path, site=site, date=date, hour=hour, minute=minute, sheet_name=sheet_name)


def file_snapshot(path: Path) -> tuple[int, int, int, int]:
    """Capture size, modification time and file identity for consistency checks."""
    st = path.stat()
    return (st.st_size, st.st_mtime_ns, st.st_dev, st.st_ino)


_FULL_FINGERPRINT_LIMIT = 8 * 1024 * 1024
_SAMPLE_SIZE = 64 * 1024


def file_fingerprint(path: Path, size: int) -> bytes:
    """Fingerprint entire small files or 3 bounded regions of a large file.

    This detects accidental edits retaining size/mtime, but sampled hashing
    is not a cryptographic guarantee for unsampled content above 8 MiB.
    """
    digest = hashlib.blake2b(digest_size=20)
    digest.update(str(size).encode('ascii'))
    with path.open('rb') as stream:
        if size <= _FULL_FINGERPRINT_LIMIT:
            for chunk in iter(lambda: stream.read(128 * 1024), b''):
                digest.update(chunk)
        else:
            offsets = (0, size // 2 - _SAMPLE_SIZE // 2, size - _SAMPLE_SIZE)
            for offset in offsets:
                stream.seek(offset)
                digest.update(offset.to_bytes(8, 'little'))
                digest.update(stream.read(_SAMPLE_SIZE))
    return digest.digest()


def scan_folder(folder: Path) -> ScanReport:
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError(f"존재하지 않는 폴더: {folder}")
    candidates = sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.casefold() in SOURCE_SUFFIXES), key=lambda p: p.name.casefold())
    groups_by_key: dict[tuple[str, str], list[SheetInput]] = {}
    rejected: list[RejectedFile] = []
    for path in candidates:
        try:
            item = parse_filename(path)
            before = file_snapshot(path)
            try:
                fingerprint = file_fingerprint(path, before[0])
            except OSError:
                # An enterprise DLP agent may allow stat(), but block reads.
                # Preserve the established merge-time error guidance instead of
                # silently excluding the file during name-based analysis.
                fingerprint = None
            if file_snapshot(path) != before:
                raise ValueError("파일 분석 중 원본 파일이 변경되었습니다")
            item = replace(item, snapshot=before, content_fingerprint=fingerprint)
        except (ValueError, OSError) as exc:
            rejected.append(RejectedFile(path.name, str(exc)))
            continue
        key = (item.site.casefold(), item.date)
        groups_by_key.setdefault(key, []).append(item)
    all_groups: list[MergeGroup] = []
    for items in groups_by_key.values():
        items.sort(key=lambda i: (i.hour, i.minute, i.file_name.casefold()))
        names = unique_sheet_names([i.sheet_name for i in items])
        assigned = tuple(replace(item, sheet_name=sheet_name) for item, sheet_name in zip(items, names))
        first = assigned[0]
        output_name = _safe_filename(first.site + "_" + first.date) + ".xlsx"
        all_groups.append(MergeGroup(first.site, first.date, assigned, output_name))
    all_groups.sort(key=lambda g: (g.site.casefold(), g.date))
    groups = tuple(g for g in all_groups if len(g.inputs) >= 2)
    singles = tuple(g for g in all_groups if len(g.inputs) == 1)
    # 파일 시스템의 이름 정규화가 결과 파일명을 충돌시키는 경우, 자동 병합을 중단한다.
    counts: dict[str, int] = {}
    for group in groups:
        counts[group.output_name.casefold()] = counts.get(group.output_name.casefold(), 0) + 1
    for name, n in counts.items():
        if n > 1:
            raise ValueError(f"서로 다른 그룹의 결과 파일명이 중복됩니다: {name}")
    return ScanReport(groups, singles, tuple(rejected), len(candidates))
