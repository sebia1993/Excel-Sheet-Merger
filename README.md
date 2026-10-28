# Excel Sheet Merger — CSV → XLSX 시간별 시트 자동 병합

Windows 11용 한국어 GUI 프로그램입니다. 같은 **사업장 + 날짜**의 여러 CSV 파일을 하나의 `.xlsx` 파일로 합치고, 각 CSV의 파일명에 있는 시간을 그대로 **시트 이름**으로 사용합니다.

**Python 설치도, Microsoft Excel 설치도 필요하지 않습니다.** 실행 파일은 Python/Tkinter와 openpyxl을 PyInstaller로 패키징합니다.

## 병합 결과 예시

| 원본 CSV (쉼표로 구분된 값) | 결과 파일 | 시트 이름 |
| --- | --- | --- |
| `A사업장_20261028_11시.csv` | `A사업장_20261028.xlsx` | `11시` |
| `A사업장_20261028_14시.csv` | `A사업장_20261028.xlsx` | `14시` |
| `A사업장_20261028_17시.csv` | `A사업장_20261028.xlsx` | `17시` |

각 시트에는 해당 CSV 파일의 전체 행과 열을 담습니다. 출력 폴더는 선택한 CSV 폴더 아래의 `병합결과`입니다. B사업장 또는 날짜가 다른 데이터는 다른 통합문서로 분류됩니다.

## EXE 사용 방법

1. [최신 Windows EXE 다운로드](https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest/download/ExcelSheetMerger.exe) → **Assets**에서 `ExcelSheetMerger.exe`를 받습니다.
2. Windows 11에서 실행 후 **폴더 선택** → CSV 파일이 담긴 폴더 지정.
3. **파일 분석**으로 그룹, 결과 파일명, `11시` / `14시` / `17시` 시트명, 제외 사유를 미리 확인합니다.
4. **전체 그룹 병합** 또는 **선택 그룹만 병합**을 클릭합니다.
5. `병합결과` 폴더에서 `.xlsx` 파일과 UTF-8 작업 기록을 확인합니다.

## 지원 형식과 제한 사항

- 원본 파일 형식은 **`.csv`만 지원**합니다. `.xlsx`, `.xls`, `.xlsm`, `.xlsb`는 병합 대상에서 제외합니다.
- 날짜: `20261028`, `2026-10-28`, `2026_10_28`, `2026.10.28`.
- 시간: `11시`, `09시30분`, `14-30`, `1430`, `오전9시`, `오후2시` 등.
- 파일명에서 날짜 **하나**와 시간 **하나**를 인식할 수 있어야 합니다. `최종`, `수정` 같은 접미사는 사업장 구분에 포함됩니다. 예기치 않은 자동 병합을 방지하는 정책입니다.
- 사업장명과 날짜가 동일한 파일을 묶고, 시간을 오름차순 정렬합니다. 같은 시간의 파일이 두 개면 `11시`, `11시 (2)`로 지정합니다.
- 한 그룹에 **2개 이상** 파일이 있어야 병합하며, 1개뿐인 파일은 미리보기에 별도 표시합니다.
- CSV 인코딩은 **UTF-8(BOM 유무), CP949/EUC-KR, UTF-16(BOM)** 등을 읽습니다. 쉼표 구분, 쌍따옴표 안의 쉼표/줄바꿈, `sep=,` 지시행을 지원합니다. 손상되거나 다른 구분자를 사용하는 CSV는 제외/오류 처리될 수 있습니다.
- CSV 셀 내부의 CRLF/CR 줄바꿈은 Excel에서 한 번만 줄바꿈되도록 LF로 정규화합니다. 그 외 CSV 값은 **전부 텍스트로 저장**하여 `00123`, IP/MAC 주소, `=SUM(...)` 같은 값이 변환·실행되지 않도록 합니다. 계산이 필요한 숫자 열은 Excel에서 후속 변환하세요.
- Excel 한계: 시트당 최대 **1,048,576행 × 16,384열**, 셀당 최대 **32,767 UTF-16 코드단위**. 한도 초과 시 해당 그룹을 실패 처리합니다.
- 원본 CSV는 절대 변경하지 않습니다. 같은 결과 파일명이 이미 있으면 **덮어쓰지 않고 건너뜁니다**. 임시 파일 작성 후 최종 파일명으로 확정하며 실패 시 정리합니다.
- Excel 설치가 필요하지 않습니다. 단, 결과 파일을 열어 보려면 Excel 또는 다른 `.xlsx` 뷰어가 필요합니다.

## 자동 EXE 빌드 및 Release

`main` 브랜치에 변경이 푸시되면 `.github/workflows/windows-exe.yml`이 Windows 러너에서 **테스트 → PyInstaller EXE 빌드 → PE 형식 검사 → SHA-256 생성 → GitHub Release 업로드**를 수행합니다. 성공 이후 [Releases](https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest)에서 버전별 다운로드를 확인할 수 있습니다. CSV 변환판은 `v1.1.<워크플로 실행 번호>` 태그로 배포됩니다.

개발 PC에 Python 3.12가 설치되어 있다면 `build_windows.bat`로 로컬 빌드할 수 있습니다. 사용자 PC에는 Python을 설치할 필요가 없습니다.

## 개발과 테스트

- Python 3.12+, `openpyxl>=3.1.5`, Tkinter, PyInstaller.
- `python -m pip install -r requirements-build.txt` (개발/빌드 PC만)
- `python -m unittest discover -s tests -v`
- `python app.py`
- 주요 파일: `src/excel_sheet_merger/classify.py`(파일명 분류), `merge.py`(CSV 파싱 및 XLSX 생성), `gui.py`(한국어 GUI).

## 실제 환경 검증 범위

CSV → XLSX 변환은 자동 테스트로 검증합니다. GitHub Actions에서 Windows용 EXE 패키징도 확인할 수 있습니다. **실제 회사 PC/VDI 정책, 백신 탐지, 특수 CSV 인코딩 및 대용량 처리**는 실사용 환경에서 추가 검증이 필요합니다. 코드 서명되지 않은 EXE는 회사 보안정책상 실행 제한될 수 있으니 정책을 준수하세요.
