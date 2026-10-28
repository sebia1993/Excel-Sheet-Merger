# Python Excel Sheet Merger for Windows 11 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Python으로 작성한 엑셀 시트 자동병합 GUI와 Windows EXE 빌드 경로를 제공한다.

**Architecture:** Pure Python filename classifier, COM Excel merger, Tkinter GUI; merge runs worker thread; CI uses Windows PyInstaller.

**Tech Stack:** Python 3.12, unittest, tkinter, pywin32, PyInstaller, GitHub Actions.

**Spec:** `docs/superpowers/specs/2026-10-28-python-win11-merge-design.md`

## Global Constraints
- 원본 읽기 전용, 최종 결과 덮어쓰기 금지, 시간순 정렬, 그룹명 엄격 비교.
- COM 보안: 외부 링크 업데이트 금지, 매크로 실행 차단, Excel 자원 해제.
- Windows에서 패키징, Linux에서 Python 단위 테스트; Windows 실행 검증 여부는 정직하게 구분.

## Review Focus
- AM/PM 시계, 00시, 날짜 오류, 동명이인 사업장, 복수 시간 토큰, 중복 시트명, 읽기 전용 파일, Excel COM 실패, 기존 결과 충돌.

### Task 1: 파일명 분류 엔진
**Files:** `src/excel_sheet_merger/classify.py`, `tests/test_classify.py`
**Interfaces:** `parse_filename(path) -> SheetInput`, `scan_folder(path) -> ScanReport`.
- [ ] 분류·시트 중복·누락 테스트 작성 및 RED 확인
- [ ] 구현, GREEN 확인, 전 테스트 재실행
- [ ] 변경 커밋

### Task 2: Excel COM 병합 및 실패 보호
**Files:** `src/excel_sheet_merger/merge.py`, `tests/test_merge.py`
**Interfaces:** `merge_group(group, output_dir, excel) -> Path`, `excel_session()`; GUI에서 호출.
- [ ] 가짜 Excel COM 객체로 병합 순서, 저장 실패, 결과 덮어쓰기 테스트 작성 및 RED 확인
- [ ] 구현, GREEN 확인, 전체 테스트 재실행
- [ ] 변경 커밋

### Task 3: Tkinter GUI / 빌드 / 설명서
**Files:** `src/excel_sheet_merger/gui.py`, `app.py`, `requirements-build.txt`, `build_windows.bat`, `.github/workflows/windows-exe.yml`, `README.md`
**Interfaces:** `run_app()` from `app.py`; GUI scan and merge using earlier modules.
- [ ] GUI 의존 없는 인자·문서·빌드 정적 테스트 작성 및 RED 확인
- [ ] GUI 구현 및 PyInstaller Windows 빌드 구성, 테스트 GREEN
- [ ] 전체 unittest/py_compile 실행, 결과 ZIP 생성, 변경 커밋

## Handoff
사용자의 '개발 진행' 요청을 실행 방식으로 수용하고 샌드박스에서 inline 구현한다. Windows 직접 빌드는 가능한 수단만 수행한다.
