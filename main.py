import re
from pathlib import Path

from bs4 import BeautifulSoup

from excel_converter import process_folder_to_excel_per_html
from report_modifier import process_all_html_files

# =========================
# Shared Config (single place)
# =========================
CONFIG = {
    "ORIGINAL_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\Security",
    "NEW_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\SecurityFixed",
    "SWE4_JSON_ROOT": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\SWE4_Json",
    "OUTPUT_XLSX_PATH": r"C:\Users\hyper.park\PycharmProjects\VCReportModifier\Security_TCNames.xlsx",
    "OVERWRITE_NEW_ROOT": True,
    "RECURSIVE_SEARCH": True,
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


def clean_text(text: str) -> str:
    return "" if text is None else " ".join(text.split()).strip()


def read_html_soup(html_path: Path):
    try:
        html = html_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        html = html_path.read_text(encoding="cp949", errors="ignore")
    return BeautifulSoup(html, "html.parser")


def safe_name(text: str) -> str:
    text = re.sub(r"\([^)]*\)", "", clean_text(text)).replace("::", "_")
    return re.sub(r"_+", "_", re.sub(r"[^A-Za-z0-9_]+", "_", text)).strip("_")


def extract_func_name(subprogram: str) -> str:
    subprogram = re.sub(r"\([^)]*\)", "", clean_text(subprogram))
    return subprogram.split("::")[-1].strip() if "::" in subprogram else subprogram.strip()


def main():
    process_all_html_files(CONFIG, clean_text, read_html_soup, safe_name, extract_func_name)
    process_folder_to_excel_per_html(CONFIG, clean_text, read_html_soup)


if __name__ == "__main__":
    main()
