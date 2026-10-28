# ExcelSheetMerger — Windows 11 엑셀 시트 자동 병합 (Python)

**Python + Superpowers 개발 흐름을 적용한 Windows 11 GUI 프로그램**입니다. 원본 엑셀 파일들을 파일명 기준으로 묶어, 한 파일 안에 시간별 여러 시트를 만듭니다.

## 결과 예시

| 원본 파일 | 병합 결과 | 시트 이름 |
| --- | --- | --- |
| A사업장_20261028_11시.xlsx | A사업장_20261028.xlsx | 11시 |
| A사업장_20261028_14시.xlsx | A사업장_20261028.xlsx | 14시 |
| A사업장_20261028_17시.xlsx | A사업장_20261028.xlsx | 17시 |

B사업장 또는 날짜가 다른 파일은 별도 그룹이 됩니다. **서로 비슷한 이름이어도 자동으로 묶지 않습니다.** 사업장명과 날짜가 동일해야 합니다.

## 최종 사용법 (EXE)

1. [최신 Windows 11 EXE 바로 다운로드](https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest/download/ExcelSheetMerger.exe)에서 `ExcelSheetMerger.exe`를 내려받습니다. **GitHub Actions 빌드 및 Release 생성이 성공한 뒤 링크가 동작합니다.**
2. Windows 11 PC에 **Microsoft Excel 데스크톱 프로그램**이 설치되어 있어야 합니다. Python 설치는 필요 없습니다.
3. EXE를 더블클릭하여 실행합니다.
4. **폴더 선택** → **파일 분석** → 그룹/시트/제외 사유 미리보기.
5. **전체 그룹 병합** 또는 그룹을 선택한 다음 **선택 그룹만 병합**을 누릅니다.
6. 원본 폴더의 `병합결과` 폴더에 `.xlsx`와 작업 기록 `.txt`가 생성됩니다.

## Windows EXE 만드는 방법 — Python 설치 없는 빌드 경로

이 GitHub 저장소는 `main` 브랜치에 코드를 올리면 GitHub Actions의 `Windows 11 Python EXE 빌드 및 릴리즈` 워크플로가 **테스트 → Windows EXE 빌드 → SHA-256 검사값 작성 → GitHub Release 업로드**를 수행합니다. 성공 후 [최신 Release](https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest)에서 EXE 파일을 직접 받을 수 있습니다. Actions의 개별 실행을 열면 `ExcelSheetMerger-Windows11-x64` 아티팩트(ZIP)도 받을 수 있습니다. 버전은 `v1.0.<워크플로 실행 번호>` 형식입니다.

Windows 11에 Python 3.12가 설치된 개발용 PC에서는 `build_windows.bat`를 더블클릭하면 테스트 후 `dist/ExcelSheetMerger.exe`가 만들어집니다. 이때만 Python과 인터넷 연결(필요한 빌드 라이브러리 설치용)이 필요합니다.

**주의:** 이 배포 ZIP 안에 실제 빌드된 EXE가 없다면, 소스코드 ZIP을 EXE라고 간주하면 안 됩니다. Windows 빌드가 성공했는지 GitHub Actions 결과 또는 Windows 빌드 로그로 확인해 주세요.

## 파일 이름 규칙

- 날짜: `20261028`, `2026-10-28`, `2026_10_28`, `2026.10.28`
- 시간: `11시`, `09시30분`, `14-30`, `1430`, `오전9시`, `오후2시`
- 파일 이름에서 날짜 1개, 시간 1개만 명확히 인식되어야 합니다.
- 사업장명/시설명은 날짜·시간 이외의 문자열을 보존하여 판단합니다. 예: `A사업장_본관`과 `A사업장_별관`은 별도 그룹입니다.
- `최종`, `수정`, `본관` 등의 접미사도 분류 기준으로 남습니다. 잘못 합쳐지는 것을 막기 위한 동작입니다.
- `.xlsx`만 실제 병합; `.xls`, `.xlsm`, `.xlsb`와 Excel 잠금 파일(`~$` 시작)은 제외합니다.
- 그룹에 파일이 **2개 이상** 있을 때만 결과 파일을 생성합니다. 단독 파일은 목록에 보류 표시합니다.
- 시트는 시간순으로 정렬되며 동일 시간이면 `11시`, `11시 (2)`처럼 생성합니다.

## 보존 및 안전 정책

- 입력 `.xlsx`를 `ReadOnly=True`, `UpdateLinks=0`으로 열고, Excel 자동화 매크로 실행을 차단합니다.
- **원본 파일은 삭제 또는 덮어쓰기하지 않습니다.** 출력 경로에 같은 결과 파일이 있으면 건너뜁니다.
- 각 원본은 **정확히 1개의 시트**가 있어야 합니다. 하나라도 다르면 해당 그룹을 실패 처리하고 다른 그룹은 계속 진행합니다.
- 임시 파일로 먼저 저장하고 Excel 통합 문서를 닫은 후 최종 파일명으로 확정합니다.
- Excel COM의 시트 복사를 이용해 셀 서식·수식·도형·차트 등을 최대한 유지합니다. 다만 **외부 통합 문서 참조, 데이터 연결, 암호화된 통합 문서 등은 결과 검증이 필요**합니다.
- 네트워크 경로 및 회사 정책에서 Excel 자동화가 제한되어 있으면 작업이 실패할 수 있습니다. 보안 정책을 우회하지 마세요.
- 기업 보안 솔루션은 서명되지 않은 새 EXE를 경고할 수 있습니다. 회사의 정식 배포 승인/코드 서명 절차를 따르세요.

## 개발 정보

- Python 3.12, GUI: Tkinter, Windows Excel 제어: pywin32, 실행 파일 패키징: PyInstaller.
- Windows 전용 COM 모듈은 실행할 때만 동적으로 로드합니다. Linux 개발 환경에서는 파일명 및 트랜잭션 처리에 대한 단위 테스트가 가능합니다.
- 테스트: `python -m unittest discover -s tests -v`
- 직접 Python 실행(개발자용): `python app.py` (윈도우 + pywin32 필요)
- 소스 트리: `src/excel_sheet_merger/classify.py` 파일명 분석, `merge.py` Excel COM, `gui.py` Tkinter, `.github/workflows/windows-exe.yml` 자동 빌드 및 Release 게시.

## 검증 범위

- Python 단위 테스트는 파일명 인식·분류·중복 시트명·기존 결과 덮어쓰기 방지·예외 시 임시 파일 정리 등을 검증합니다.
- Windows GitHub Actions에서 PyInstaller로 실제 EXE를 빌드하고 PE 헤더 및 생성 여부를 검사합니다. Microsoft Excel이 설치된 실제 Windows 11 PC에서의 COM 병합은 **별도 검증이 필요**합니다.
