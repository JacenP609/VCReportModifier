import logging

from excel_converter import process_folder_to_excel
from report_modifier import process_all_html_files

CONFIG = {
    "ORIGINAL_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\Security",
    "NEW_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\SecurityFixed",
    "SWE4_JSON_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\SWE4_Json",
    "OUTPUT_XLSX_PATH": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\Security_TCNames.xlsx",
    "OVERWRITE_NEW_ROOT": True,
    "RECURSIVE_SEARCH": True,
    "ENABLE_TIME_LOG": True,
    "TEST_START_DATE": "01 MAY 2026  1:00:00 PM",
    "EXECUTION_START_DATE": "04 MAY 2026  9:00:00 AM",
    "DATE_FORMAT": "%d %b %Y  %I:%M:%S %p",
    "RANDOM_SEED": 20260506,
    "CREATION_MIN_STEP_MINUTES": 10,
    "CREATION_MAX_STEP_MINUTES": 20,
    "EXECUTION_STEP_PER_TC_COUNT": 5,
    "EXECUTION_STEP_MINUTES": 1,
    "JSON_FILE_MAP": {
        "hil": "HIL_FunctionList.json",
        "fil": "FIL_FunctionList.json",
        "security": "Security_FunctionList.json",
        "ftl": "FTL_FunctionList.json",
    },
}

logging.basicConfig(filename="UnitTC.log", filemode="w", level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")


def main():
    process_all_html_files(CONFIG)
    process_folder_to_excel(CONFIG)


if __name__ == "__main__":
    main()
