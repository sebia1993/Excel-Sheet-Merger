"""Per-process isolated scratch folders with conservative crash-residue recovery.

Only removes our own exact-directory format, with an intact ownership marker,
known staged filenames, at least 24h old, and *no live process lock*.
Never attempts to clean pre-v1.4 unknown-ownership `_병합중_*.xlsx` files.
"""
from __future__ import annotations

import os
import re
import shutil
import time
import uuid
from pathlib import Path

MARKER = b'ExcelSheetMerger:scratch:v1\n'
AGE_SECONDS = 24 * 60 * 60
_NAME = re.compile(r'_병합임시_[0-9a-f]{32}\Z')
_ALLOWED = {'owner.lock', 'result.xlsx', 'base.xlsx', 'csv.xlsx', 'log.tmp'}


def _lock(handle, blocking: bool) -> bool:
    """An actual OS lock, released automatically by the kernel on process death."""
    try:
        handle.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK if blocking else msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB))
        return True
    except (OSError, BlockingIOError):
        return False


def _unlock(handle) -> None:
    handle.seek(0)
    if os.name == 'nt':
        import msvcrt
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl
        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def cleanup_stale_workspaces(output_dir: Path) -> None:
    """Best-effort cleanup, never affecting current writes or unowned files."""
    try:
        children = list(output_dir.iterdir())
    except OSError:
        return
    for candidate in children:
        if not _NAME.fullmatch(candidate.name) or candidate.is_symlink() or not candidate.is_dir():
            continue
        marker = candidate / 'owner.lock'
        try:
            if marker.is_symlink() or marker.stat().st_size != len(MARKER):
                continue
            if time.time() - marker.stat().st_mtime < AGE_SECONDS:
                continue
            # Never descend into unknown files or symlinks, even with a matching marker.
            entries = list(candidate.iterdir())
            if any(p.name not in _ALLOWED or p.is_symlink() or not p.is_file() for p in entries):
                continue
            with marker.open('r+b') as handle:
                if handle.read() != MARKER or not _lock(handle, blocking=False):
                    continue
                try:
                    # On Windows we must close the handle before rmtree.
                    pass
                finally:
                    _unlock(handle)
            shutil.rmtree(candidate)
        except (OSError, ValueError):
            # A concurrently running process, locked file, changed directory,
            # readonly path, or third-party data means leave it alone.
            continue


class OwnedScratch:
    """One merge's private, process-locked workspace on the output filesystem."""
    def __init__(self, output_dir: Path):
        self.path = Path(output_dir) / f'_병합임시_{uuid.uuid4().hex}'
        self._handle = None

    def __enter__(self) -> 'OwnedScratch':
        self.path.mkdir(parents=False, exist_ok=False)
        try:
            marker = self.path / 'owner.lock'
            self._handle = marker.open('x+b')
            self._handle.write(MARKER)
            self._handle.flush()
            if not _lock(self._handle, blocking=True):
                raise OSError('임시 작업 파일 잠금 실패')
            return self
        except BaseException:
            if self._handle is not None:
                self._handle.close()
            shutil.rmtree(self.path, ignore_errors=True)
            raise

    def __exit__(self, exc_type, exc_val, traceback) -> None:
        if self._handle is not None:
            try:
                _unlock(self._handle)
            finally:
                self._handle.close()
                self._handle = None
        shutil.rmtree(self.path, ignore_errors=True)
