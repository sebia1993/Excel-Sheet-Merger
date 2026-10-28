"""Opt-in synthetic fingerprint IO benchmark, not a company SMB test.

Run: PYTHONPATH=src python benchmarks/measure_fingerprint.py
Uses a sparse 128MiB file with a few nonzero regions. Reports local timings and
read budgets for legacy three-region sampling, current 17-region sampling,
and full-file hashing. Values are NOT predictive of network share latency.
"""
from __future__ import annotations

import hashlib
import statistics
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

from excel_sheet_merger.classify import file_fingerprint

SIZE = 128 * 1024 * 1024
SAMPLE = 64 * 1024
REPEATS = 5


def _full_hash(path: Path) -> bytes:
    digest = hashlib.blake2b(digest_size=20)
    digest.update(str(SIZE).encode('ascii'))
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(128 * 1024), b''):
            digest.update(chunk)
    return digest.digest()


def _measure(fn) -> float:
    timings = []
    for _ in range(REPEATS):
        start = time.perf_counter()
        result = fn()
        assert len(result) == 20
        timings.append(time.perf_counter() - start)
    return statistics.median(timings)


def main() -> None:
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / 'large.bin'
        with path.open('wb') as stream:
            stream.truncate(SIZE)
            for position in (0, SIZE // 4, SIZE // 2, SIZE * 3 // 4, SIZE - SAMPLE):
                stream.seek(position)
                stream.write(b'A' * SAMPLE)
        with patch('excel_sheet_merger.classify._LARGE_FILE_SAMPLES', 3):
            legacy = _measure(lambda: file_fingerprint(path, SIZE))
        current = _measure(lambda: file_fingerprint(path, SIZE))
        full = _measure(lambda: _full_hash(path))
        print(f'Synthetic file: {SIZE / (1024**2):.0f} MiB, local filesystem')
        print(f'3 windows:  {3 * SAMPLE / 1024:.0f} KiB read, {legacy:.4f}s median')
        print(f'17 windows: {17 * SAMPLE / 1024:.0f} KiB read, {current:.4f}s median')
        print(f'Full read:  {SIZE / (1024**2):.0f} MiB read, {full:.4f}s median')
        print('Network/DRM/SMB costs can differ substantially; use approved samples.')


if __name__ == '__main__':
    main()
