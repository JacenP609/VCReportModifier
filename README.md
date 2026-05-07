# VCReportModifier

이 tool은 VectorCAST HTML Report를 수정하고, Test Case 정보를 Excel로 추출하는 Python 자동화 프로그램이다.

주요 기능은 다음과 같다.

원본 Report 폴더를 새 폴더로 복사한다.
복사된 HTML Report를 수정한다.
User Code 섹션 교체
TC 제목을 FunctionName-N 형식으로 변경
Date of Creation / Date of Execution 일괄 변경
Requirements/Notes의 ATG 자동 생성 문구 제거
TOC의 Test Case 항목 재정렬
수정된 HTML Report에서 TC 정보를 추출해서 HTML 파일별 Excel 파일로 저장한다.

필요 패키지는 아래와 같다.

pip install beautifulsoup4 openpyxl

코드 상단 CONFIG 영역에서 아래 경로를 수정해서 사용한다.

ORIGINAL_ROOT: 원본 VectorCAST HTML Report 폴더
NEW_ROOT: 수정된 HTML Report가 저장될 폴더
SWE4_JSON_ROOT: SWE4 FunctionList JSON 폴더
INPUT_ROOT: Excel 추출 대상 HTML 폴더
OUTPUT_XLSX_ROOT: 최종 Excel 출력 루트 폴더
TEST_START_DATE / TEST_END_DATE: 생성 시각 난수 범위 시작/종료

일반적으로 HTML 수정 후 Excel까지 추출하려면 INPUT_ROOT는 NEW_ROOT로 두면 된다.

SWE4_JSON_ROOT 폴더에는 아래 JSON 파일들이 있어야 한다.

HIL_FunctionList.json
FIL_FunctionList.json
Security_FunctionList.json
FTL_FunctionList.json

HTML 파일명 prefix에 따라 사용할 JSON이 자동 선택된다.

HIL_ 로 시작하는 HTML은 HIL_FunctionList.json을 사용한다.
FIL_ 로 시작하는 HTML은 FIL_FunctionList.json을 사용한다.
Security_ 로 시작하는 HTML은 Security_FunctionList.json을 사용한다.
FTL_ 로 시작하는 HTML은 FTL_FunctionList.json을 사용한다.

실행은 아래 명령으로 한다.

python main.py

실행 방식은 main() 함수에서 선택한다.

HTML 수정과 Excel 추출을 둘 다 실행하려면 아래 두 줄을 사용한다.

process_all_html_files()
process_folder_to_excel()

HTML 수정만 실행하려면 process_all_html_files()만 사용한다.

Excel 추출만 실행하려면 process_folder_to_excel()만 사용한다.

출력 결과는 두 가지다.

첫 번째는 NEW_ROOT에 생성되는 수정된 HTML Report 폴더다.
두 번째는 OUTPUT_XLSX_ROOT에 생성되는 Excel 파일들이다.

Excel 출력은 INPUT_ROOT 기준 상대경로를 그대로 유지하며, HTML 파일 1개당 Excel 파일 1개가 생성된다.
예: INPUT_ROOT\\A\\B\\sample.html -> OUTPUT_XLSX_ROOT\\A\\B\\sample.xlsx

Excel 시트에서 `Source HTML`, `Unit Under Test` 컬럼은 제거되어 더 이상 출력되지 않는다.

주의사항은 다음과 같다.

원본 ORIGINAL_ROOT 폴더는 직접 수정하지 않는다.
스크립트는 ORIGINAL_ROOT를 NEW_ROOT로 복사한 뒤, NEW_ROOT 안의 HTML만 수정한다.
OVERWRITE_NEW_ROOT가 True이면 기존 NEW_ROOT 폴더는 삭제 후 다시 생성된다.
HTML 수정 후 Excel을 추출하려면 INPUT_ROOT를 NEW_ROOT로 설정해야 한다.
Date of Creation 생성값은 TEST_START_DATE~TEST_END_DATE 범위에서 결정되며, 08:00~20:00 시간대만 사용된다.
Configuration Data의 `Date/Time of Report Creation`은 같은 Report의 마지막 `Date of Execution`에 몇 분을 더해 자동 설정된다.
