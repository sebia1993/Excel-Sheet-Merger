# Excel Sheet Merger — CSV·Excel → XLSX 시간별 시트 자동 병합

Windows 11용 미니멀 업무형(화이트·그린) 한국어 GUI 프로그램. 같은 사업장명과 날짜의 CSV·Excel 원본들을 시간순으로 분류하여 한 개의 XLSX 통합 문서로 병합합니다.

**Python 설치가 필요하지 않습니다. Excel 설치가 필요하지 않습니다.** Windows x64 EXE로 제공됩니다.

## 다운로드

최신 Windows EXE:
https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest/download/ExcelSheetMerger.exe

릴리즈 페이지:
https://github.com/sebia1993/Excel-Sheet-Merger/releases/latest

## 사용 방법

1. EXE 실행 후 CSV·Excel 파일이 들어 있는 폴더를 선택.
2. 파일 분석에서 사업장/날짜별 그룹 및 시간별 시트 구성 미리보기.
3. 전체 그룹 병합 또는 선택 그룹 병합 클릭.
4. 원본 폴더 안 병합결과 폴더에서 XLSX 파일과 작업 로그 확인.

## 예제

원본 A사업장_20261028_11시.csv → 시트 11시
원본 A사업장_20261028_14시.xlsx → 시트 14시
원본 A사업장_20261028_17시.xls → 시트 17시

최종 통합 문서: A사업장_20261028.xlsx (시트 11시, 14시, 17시)

CSV끼리, Excel끼리, CSV+Excel 혼합 모두 같은 기준으로 묶습니다. 같은 시각의 서로 다른 파일은 11시, 11시 (2)로 구분합니다. 파일이 한 개뿐인 그룹은 병합 보류합니다.

## 지원 원본

- .csv: UTF-8(BOM 포함), CP949, UTF-16(BOM) 및 쉼표·인용부호·줄바꿈 처리. 모든 셀을 텍스트로 저장하여 00123, IP/MAC, =SUM() 같은 문자열을 보존.
- .xlsx: 숫자·날짜·문자·수식·셀 스타일·병합 셀·행/열 크기·틀고정을 가능한 범위에서 복사.
- .xlsm: .xlsx와 동일하지만 VBA 매크로는 결과 XLSX에 포함되지 않음.
- .xls: xlrd 라이브러리로 데이터·날짜·불리언 중심으로 변환. 원본 셀 서식은 미보존.
- .xlsb 등 미지원 형식은 제외 사유 표시.

**Excel 원본은 파일 하나당 정확히 시트 1개**여야 합니다. 여러 시트가 있는 Excel이나 손상된 파일이 있으면 해당 그룹 병합을 실패 처리하고 작업 기록에 원인을 남깁니다. 다른 정상 그룹은 계속 처리합니다.

## 파일명 규칙과 주의사항

파일명에서 날짜 1개와 시간 1개 인식 필수. 날짜 20261028, 2026-10-28, 2026_10_28, 2026.10.28; 시간 11시, 09시30분, 14-30, 1430, 오후2시 등을 지원합니다. 최종/수정 같은 추가 문구는 사업장 구분에 포함되며, 임의의 유사도 기준으로 합치지 않습니다.

원본 CSV와 Excel은 수정하지 않습니다. 이미 같은 결과 파일이 있으면 덮어쓰지 않고 건너뜁니다. 오류가 나면 불완전한 임시 파일을 삭제합니다. Excel 시트 최대 1,048,576행 × 16,384열.

일반 Excel의 차트·이미지·피벗·조건부 서식·외부 링크·매크로 및 수식 재계산/캐시의 완전한 복제는 보장하지 않습니다. 원본 시트명을 참조하는 수식은 새 시트명과 충돌할 수 있으므로 결과를 확인하세요. Excel 포함 그룹은 메모리 사용량이 증가합니다. 회사 보안 정책상 코드 미서명 EXE 실행이 제한될 수 있습니다.

## 개발과 자동 빌드

Python 3.12, openpyxl, xlrd, Tkinter, ttkbootstrap 및 PyInstaller를 사용합니다.

개발/테스트 명령:
python -m pip install -r requirements-build.txt
python -m unittest discover -s tests -v
python app.py

Windows 빌드: build_windows.bat.
GitHub Actions: .github/workflows/windows-exe.yml — 테스트 → PyInstaller EXE → SHA256SUMS → 버전별 GitHub Release 자동 게시.
혼합 Excel 입력 릴리즈 태그는 v1.3.<GitHub Actions 실행 번호> 입니다.

핵심 파일: src/excel_sheet_merger/classify.py (분류), merge.py (병합), gui.py (화면), ui_theme.py (디자인).

참고 UI: https://github.com/israel-dryer/ttkbootstrap (MIT).
실제 구형 XLS 파일 및 회사 PC/VDI에서의 호환성은 추가 실사용 검증이 필요합니다.
