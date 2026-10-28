"""Optional local synthetic comparison. NOT a real SMB/DRM/Windows benchmark.

Usage: PYTHONPATH=src python benchmarks/measure_local_fingerprint.py
"""
from __future__ import annotations

import statistics
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

from excel_sheet_merger.classify import file_fingerprint

SIZES = (17 * 1024**2, 32 * 1024**2, 33 * 1024**2)


def median_seconds(callback) -> float:
    results = []
    for _ in range(4):
        start = time.perf_counter()
        callback()
        results.append(time.perf_counter() - start)
    return statistics.median(results)


def main() -> None:
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / 'local-io-benchmark.bin'
        for size in SIZES:
            with path.open('wb') as fp:
                fp.truncate(size)
            with patch('excel_sheet_merger.classify._is_fixed_local_drive', return_value=False):
                sampled = median_seconds(lambda: file_fingerprint(path, size))
            with patch('excel_sheet_merger.classify._is_fixed_local_drive', return_value=True):
                fixed = median_seconds(lambda: file_fingerprint(path, size))
            fixed_bytes = size if size <= 32 * 1024**2 else 17 * 65536
            print(f'{size // 1024**2:2d} MiB: remote/unknown sampled={sampled:.4f}s '
                  f'(1.0625 MiB read), fixed-local={fixed:.4f}s '
                  f'({fixed_bytes / 1024**2:.4f} MiB read)')
    print('Times depend on host and cache. No corporate SMB or Windows claims.')


if __name__ == '__main__':
    main()
