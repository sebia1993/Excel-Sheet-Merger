# Long Site Name Output Collision: Stability-only Design

## Problem

The current output filename is `_safe_filename(first.site + '_' + first.date) + '.xlsx'`, with a fixed 150-character truncation. For >150-character site names this removes the date and collapses different site/date group output paths. The collision safeguard then aborts the entire scan, even when source filenames themselves are valid.

## Selected behavior

Keep exact filenames for existing ordinary (<=150) site/date combinations. Only if the sanitized/truncated output is missing its date tail, build a deterministic 150-character name with: sanitized site prefix, '~' + 12-character BLAKE2b digest of the entire original site name, and '_YYYYMMDD'. This ensures normal dates stay visible and distinct long sites sharing the same first 150 characters almost certainly remain distinct. The existing collision-detection check remains as defense in depth.

## Safety and scope

Never merge mismatched sites or dates, overwrite an earlier result, remove original input or change sheet names. Do not bypass `scan_folder` rejection of malformed input. Only output paths for previously failing oversized names change. No GUI changes, no additional runtime packages, no 8-hour soak yet.

## Tests

1. A 165-character site with two different dates yields two groups and two distinct, timestamp-suffixed XLSX outputs, rather than raising an output name collision.
2. Two 165-character sites differing only at the far end of the name yield distinct outputs on the same date, without silent group merging. Test resulting workbook contents.
3. Ordinary short site and date output remains byte-for-byte identical to historical naming rule.
