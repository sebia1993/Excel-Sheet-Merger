# Excel Sheet Merger — 무관한 고압축 Office XML의 오탐 방지

- 실제로 읽지 않는 Office `customXml/itemN.xml`의 압축률만 과도하다는 이유로 정상 셀 데이터를 가진 XLSX가 거부되는 문제를 수정했습니다. 기존 `openpyxl 3.1.5`에서는 해당 부가 요소를 병합 결과에 복사하지 않습니다.
- 실제 ZIP 테스트에서 정상 데이터가 있는 CSV+XLSX 병합을 확인하고, 과도하게 압축된 워크시트 XML, 임의 부가 XML 경로, XML 크기 제한 초과는 여전히 로딩 전에 차단하도록 회귀 검증했습니다.
- 전체 ZIP 압축 해제 1GiB, 개별 XML 256MiB, ZIP 구성 요소 수 및 중복 검사, 본문 워크시트의 압축률 한도는 변경하지 않았습니다.
- Linux/Xvfb 로컬 자동 테스트 161개 실행, 157개 통과·Windows 전용 4개 제외. Windows 전체 CI 및 실제 EXE 스모크는 본 Release의 GitHub Actions 결과를 참조하세요.
- GUI·그룹 체크박스·파일명 번호 부여·CSV 텍스트·Excel 수식/서식·기존 결과 및 원본 보호는 유지했습니다.
- 회사 Windows 11 PC/SMB/DRM/백신의 8시간 실사용 시험은 아직 수행하지 않았습니다.

[14차 감사](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/superpowers/stability-recheck-2026-10-29-round14.md) · [이전 버전 릴리즈 기록 (v1.3.29)](https://github.com/sebia1993/Excel-Sheet-Merger/blob/main/docs/releases/v1.3.29.md)
