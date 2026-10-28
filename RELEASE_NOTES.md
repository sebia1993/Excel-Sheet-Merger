# Excel Sheet Merger — 긴 사업장명에서 결과 파일명 충돌 방지

- 사업장명이 매우 길면 기존 150자 출력 파일명 제한 때문에 날짜가 잘려 **서로 다른 날짜의 파일도 같은 결과 이름으로 충돌**하여 분석 자체가 중단되던 문제를 수정했습니다.
- 일반적인 짧은 사업장명은 이전과 **완전히 같은 결과 파일명**을 사용합니다. 길이 제한으로 날짜를 포함하지 못하는 경우에만 사업장 이름을 일부 줄이고 12자리 고정 해시·날짜를 포함해 결과 그룹을 구분합니다.
- 서로 다른 날짜 또는 이름 앞부분이 같은 서로 다른 긴 사업장을 합치지 않고, 각 그룹에서 올바른 `11시`·`14시` 시트를 생성하는 테스트를 추가했습니다.
- 기존 CSV/XLSX/XLSM/XLS 입력, 수식·서식·메모, 그룹별 체크박스, 기존 결과 보존, 이름 뒤 `(2)`·`(3)` 추가, 오류 메시지 및 GUI 배치는 변경하지 않았습니다.
- Linux/Xvfb 전체 166개 실행, 162개 통과·Windows 전용 4개 제외. Windows GitHub Actions의 전체 테스트·실제 EXE 검증 결과는 새 Release 기록에서 확인할 수 있습니다.
- 실제 회사 Windows 11/SMB/DRM/백신 환경의 8시간 실사용 시험은 모든 코드 안정성 개선 이후 최종 단계로 보류합니다.

[16차 안정성 재감사](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/superpowers/stability-recheck-2026-10-29-round16.md) · [이전 v1.3.31 변경 이력](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/releases/v1.3.31.md)
