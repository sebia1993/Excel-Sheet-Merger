# Excel Sheet Merger — v1.3 CSV 및 Excel 혼합 병합

기존 CSV 전용 프로그램에 .xlsx·.xlsm·.xls 일반 Excel 원본을 추가했습니다. 미니멀 업무형 GUI는 유지합니다.

## 기능
- 같은 사업장/날짜의 CSV+Excel을 시간순으로 병합해 한 개 XLSX 생성.
- 파일명 시간으로 11시, 14시, 17시 시트 구성; 중복 시간은 (2) 추가.
- .xlsx/.xlsm의 문자·숫자·날짜·수식·대표 서식·병합 셀·틀고정 가능한 범위에서 복사.
- .xls의 셀 값·날짜·불리언을 xlrd로 변환(원본 스타일 미보존).
- CSV의 00123/IP/MAC/수식처럼 보이는 텍스트 보존 및 기존 출력 보호.
- GUI의 CSV 전용 안내 문구를 CSV·Excel 안내로 변경.

## 제한
- 원본 Excel은 파일당 정확히 한 개의 시트 필요. 복수 시트나 손상 파일은 해당 그룹을 실패 기록.
- .xlsb 미지원. .xlsm의 매크로, 차트/이미지/피벗/외부 연결/조건부 서식은 완전 복제되지 않음.
- 원본 시트명을 참조하는 수식은 확인 필요. 대용량 Excel은 추가 RAM 사용 가능.
- Python, Microsoft Excel 별도 설치 없이 Windows 11 EXE 실행 가능. 코드 서명은 제공하지 않음.

## 이용
1. 아래 Assets에서 ExcelSheetMerger.exe 다운로드.
2. CSV·Excel 폴더 선택 → 파일 분석 → 전체/선택 병합.
3. 병합결과 폴더의 XLSX와 작업 로그 확인.

소스: https://github.com/sebia1993/Excel-Sheet-Merger
