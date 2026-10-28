# Excel Sheet Merger — 폴더 경로 확인 실패 시 GUI 안정성 개선

- 공유폴더 또는 정션 경로의 연결 상태가 바뀌어 `Path.resolve()`가 오류를 내더라도, 병합 화면의 오래된 분석 결과를 초기화하고 **잘못된 그룹으로 병합을 시작하지 않도록** 보완했습니다.
- 파일 분석 자체는 성공했지만 경로 정규화가 실패하면, 아직 완성되지 않은 분석 결과를 사용하지 않고 **원본 폴더 확인 안내**를 표시합니다.
- 병합 시작 직전에 경로 확인이 실패하면 이전 체크박스 선택과 미리보기를 무효화하고, 문제가 발생한 폴더 경로와 재분석 조치 방법을 안내합니다.
- 기존에 정상 동작하던 CSV·XLSX·XLSM·XLS 입력, 시트 이름·수식·서식, 그룹별 체크박스, 결과 파일명 `(2)`·`(3)` 생성, 원본 보호와 GUI 배치는 유지합니다.
- 네트워크 연결 오류 및 경로 순환 예외를 주입하는 **새 GUI 회귀 테스트 3개**를 추가했습니다. Linux/Xvfb 전체 164개 실행, 160개 통과·Windows 전용 4개 제외. Windows GitHub Actions와 EXE 검증은 Release에 연결된 실행 기록에서 확인하세요.
- 회사 실제 Windows 11·DRM·SMB 환경에서 수행하는 8시간 장시간 안정성 시험은 코드 안정화가 끝난 **최종 단계**로 유보했습니다.

[15차 안정성 재감사](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/superpowers/stability-recheck-2026-10-29-round15.md) · [이전 릴리즈 내용 (v1.3.30)](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/releases/v1.3.30.md)
