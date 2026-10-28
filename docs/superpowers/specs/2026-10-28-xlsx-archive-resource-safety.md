# Excel Input Archive Resource-Safety Design

Date: 2026-10-28. Scope approved under the user's standing instruction to continue stability improvements without changing existing user features. The 8-hour Windows soak test remains the final stage.

## Problem
A valid .xlsx workbook with worksheet XML expanding beyond 64 MiB from less than 512 KiB on disk was accepted by the existing loader. Normal-mode openpyxl inflates XML before the application can stop excessive memory allocation.

## Alternatives
- **Chosen: metadata-only preflight:** inspect ZIP central-directory entries and suspicious expansion before calling openpyxl. A lightweight, read-only check without new dependencies.
- Post-load cell counting: rejected because memory has already been allocated.
- On-disk compressed file-size limit: rejected because compressed size does not predict expanded memory.

## Safety boundaries
- Reject more than 30,000 ZIP members.
- Reject greater than 1 GiB total advertised expanded bytes.
- Reject any single XML or relationship member over 256 MiB expanded.
- Reject any member over 64 MiB expanded and over 400:1 compression ratio.
- Reject duplicate ZIP entry names (including the content-type manifest).
- Keep ordinary CSV/XLS/XLSX/XLSM handling, styles, formulas, groups, output names, UI and source protection unchanged. A real but extraordinarily large/highly compressible workbook can hit the safety limit; this should be disclosed, not bypassed.
- Do not diagnose actual DRM encryption or malicious intent from compression alone.
- Treat central-directory inspection as a resource guard, not full antivirus or XML-safety proof.

## Acceptance checks
1. Previously accepted >64 MiB expanded worksheet-XML input is refused *before* load_workbook with input untouched and no output.
2. Normal and moderately high-compression workbooks still preserve data/formula/styles.
3. Duplicate ZIP paths, excessive total size, individual XML size and member count are handled cleanly.
4. Failed group leaves existing result and owned temporary workspace intact/clean as appropriate.
5. User sees affected filename and appropriate Korean resource warning; no claim of confirmed malware.
6. Full Linux/Windows tests and packaged Windows EXE smoke verified before new Release. Real company SMB/DRM/AV and 8-hour soak explicitly deferred.
