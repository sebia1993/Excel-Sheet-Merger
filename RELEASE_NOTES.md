# Excel Sheet Merger — 메모 전용 빈 셀의 무손실 복사

- 원본 XLSX/XLSM에서 **값과 서식 없이 메모(주석)만 붙은 셀**이 결과 파일에서 누락되던 문제를 수정했습니다. 원본 메모의 텍스트·작성자와 빈 셀 상태를 보존합니다.
- 기존 데이터·수식·스타일·시트명·그룹 선택·CSV 텍스트·중복 출력 번호 부여·오류 조치 안내·원본 파일은 그대로 유지합니다. 일반 빈 셀은 여전히 생성하지 않습니다.
- 실제 XLSX 메모 셀과 CSV를 함께 병합하고 반복 실행하여 메모·CSV 문자열 `00123`·기존 결과 보호·새 `(2)` 파일 생성을 회귀 검증합니다.
- Linux/Xvfb 로컬에서 테스트 157개 실행, 153개 통과 및 Windows 전용 4개 제외를 확인했습니다. Windows Actions 전체 테스트와 EXE 실행 검증 결과는 배포 실행 기록을 참조하세요.
- GUI 로그 3만 건 유지 시 프로세스 RSS 약 7MiB 증가를 측정했습니다. 전체 기록 조회 기능 변경 없이 자동 삭제·축약하지 않았습니다.
- 사내 Windows 11·SMB·DRM·백신 환경의 8시간 실사용 테스트는 아직 진행하지 않았습니다.

[13차 안정성 재감사](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/superpowers/stability-recheck-2026-10-29-round13.md) · [이전 릴리즈 내용 (v1.3.28)](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/releases/v1.3.28.md)
