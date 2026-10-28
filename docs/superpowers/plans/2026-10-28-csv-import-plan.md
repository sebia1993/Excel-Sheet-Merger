# CSV 시트 자동병합 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (or subagent-driven-development) task-by-task with tests before implementation.

**Goal:** Windows 11 EXE가 Excel 없이 CSV를 XLSX의 시간별 시트로 변환하도록 변경한다.

**Architecture:** 기존 `classify.py`에서 CSV만 인식하고, `merge.py`는 CSV 스트리밍+openpyxl write_only 생성으로 대체하며 GUI 작업 스레드는 병합 함수만 사용한다. Windows GitHub Actions가 EXE와 해시 파일을 배포한다.

**Tech Stack:** Python 3.12+, csv, openpyxl 3.1.x, Tkinter, PyInstaller, unittest

**Spec:** `docs/superpowers/specs/2026-10-28-csv-import-design.md`

## Global Constraints
- 원본 파일 수정 금지, 기존 결과 덮어쓰기 금지
- 시간순 시트명 사용, 단독 파일 제외, 한 그룹 실패 시 다른 그룹 계속
- 모든 CSV 필드는 Excel의 텍스트로 기록하고 수식 자동 계산 금지
- Excel 설치/COM 불필요

## Review Focus
- CSV BOM/CP949 환경에서 한글이 깨지지 않는지 테스트한다.
- 쉼표 및 줄바꿈이 포함된 인용 필드가 열과 행을 어긋나게 하지 않는지 테스트한다.
- 앞자리 0과 수식처럼 생긴 데이터가 값 변환 없이 유지되는지 테스트한다.
- CSV 오류가 발생했을 때 최종 파일과 임시 파일을 남기지 않는지 테스트한다.
- 결과 파일이 이미 있으면 다시 실행해도 보존되는지 테스트한다.

### Task 1: 파일명 분류 CSV 지원
**Files:** `tests/test_classify.py`, `src/excel_sheet_merger/classify.py`
- [ ] CSV 분류/시트명과 비CSV 제외 테스트를 먼저 작성하고 실패를 확인
- [ ] 확장자 선택과 안내만 CSV에 맞게 변경
- [ ] 테스트 통과 확인

### Task 2: CSV 데이터 → XLSX 생성기
**Files:** `tests/test_merge.py`, `src/excel_sheet_merger/merge.py`
- [ ] UTF-8 BOM, CP949, quoted newline/comma, 문자열 보존, 덮어쓰기 방지, 실패 원자성 테스트를 먼저 작성하고 실패 확인
- [ ] 인코딩 판별/CSV 파싱/Streaming XLSX/atomic rename 구현
- [ ] 테스트 통과 확인

### Task 3: GUI/배포/문서
**Files:** `tests/test_distribution.py`, `src/excel_sheet_merger/gui.py`, `requirements-build.txt`, `build_windows.bat`, `.github/workflows/windows-exe.yml`, `README.md`, `RELEASE_NOTES.md`
- [ ] GUI 작업 스레드 병합 및 Excel 비의존 테스트를 먼저 작성하고 실패 확인
- [ ] GUI와 사용자 안내 변경
- [ ] Windows 빌드 옵션/의존성 및 Release 설명 변경
- [ ] 전체 테스트 및 컴파일 검사

### Task 4: GitHub 반영 및 릴리즈 확인
- [ ] GitHub connector로 현재 main HEAD 재검증, 안전한 경로로 변경 파일 커밋
- [ ] GitHub Actions 실행 확인 및 최신 Release에 EXE와 SHA-256 아티팩트 업로드 확인
- [ ] 최종 링크와 실사용 검증 제한 사항 보고
