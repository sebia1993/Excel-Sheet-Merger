# Excel Sheet Merger — 안정성 재감사 13차 (2026-10-29)

기준: v1.3.28 (abf56f6dd2726a7765217757d091e377b204a6d6). 사용자 요청: 기존 기능·UI 유지, 지속적 안정성 개선, 8시간 Windows 장시간 실사용 테스트는 마지막 단계.

## 실제로 재현한 데이터 보존 결함

XLSX의 A1에 값이 있고, D20에는 값/서식이 없지만 메모 텍스트·작성자가 있는 입력을 생성했다. openpyxl 로드 후 D20.value=None, has_style=False, comment!=None임을 확인했다. 기존 _append_openpyxl은 값도 스타일도 없다는 이유로 D20을 건너뛰었다. 결과 XLSX에는 D20 메모가 없어졌다.

회귀 테스트를 먼저 작성해 기존 코드에서 `AssertionError: unexpectedly None` 실패를 확인했다. 이후 early-skip 조건에 comment와 hyperlink의 부재를 추가했다. 이 변경은 정상 빈 셀의 스킵, 값·수식·서식·시트명·원본 보호를 유지한다.

## 검증 근거

- 변경 전 Linux/Xvfb 회귀 테스트: 155개 실행, 151개 통과, Windows 전용 4개 제외.
- RED: 새로운 blank comment test는 구버전에서 D20.comment=None으로 실패.
- GREEN: 신규 direct test와 공개 merge_group 테스트를 포함해 157개 실행, 153개 통과, Windows 전용 4개 제외.
- 공개 그룹 병합에서 Excel D20의 주석 텍스트/작성자, CSV 문자열 00123, 시트명 11시·14시, 2회 생성한 결과의 (2) 접미사, 원본 XLSX 바이트 보존을 검증했다.
- 회사 승인 Windows PC의 8시간 시험은 실행하지 않았다. GitHub Windows Actions 전체 회귀 테스트와 EXE 패키징 실검증은 해당 릴리즈의 Actions 기록을 확인한다.

## GUI 로그 메모리

실제 Tk Text에 30,000개의 로그 행을 유지한 Linux/Python 3.13/Xvfb/ttkbootstrap 2.2.6 합성 실험에서 프로세스 RSS 약 127.76MiB에서 134.78MiB로 7.02MiB 늘었다. 기존 요구는 전체 로그를 프로그램에서 확인하는 것이다. 일부 기록 삭제·축약 또는 마지막 N개만 보여주기는 기능 변경에 해당하므로 실시하지 않았다.

## 여전히 남은 위험

- 네트워크 및 32MiB 초과 원본의 17개 표본 범위 밖 변경은 감지 불가능하다. SMB 읽기 비용을 동반하는 변경은 별도 계측 전에는 도입하지 않는다.
- 정상 고압축 Excel의 보수적 ZIP 확장 안전 제한 오탐 가능성은 남는다.
- 차트, 피벗, 외부 링크, 특수 Excel 주석(스레드 대화형)은 완전 복사를 보장하지 않는다.
- GUI 작업 기록 10만 건 이상에서는 Tcl Text 메모리 증가가 지속될 수 있다. 기존 스크롤 조회를 바꾸지 않는 설계가 필요하다.
- Windows 11·SMB·DRM·백신 8시간 실사용 검증은 최종 단계로 유보했다.
