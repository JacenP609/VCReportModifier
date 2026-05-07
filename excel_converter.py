import datetime
import logging
import time
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from report_modifier import collect_html_files

HEADER_FILL = PatternFill("solid", start_color="38E3FF")
HEADER_FONT = Font(bold=True, size=12, color="0000FF")
THIN_SIDE = Side(border_style="thin", color="000000")
HEADER_BORDER = Border(top=THIN_SIDE, left=THIN_SIDE, right=THIN_SIDE, bottom=THIN_SIDE)
BODY_BORDER = HEADER_BORDER


def create_excel(output_path: str):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "TCNames"
    ws.cell(row=1, column=1, value="TC-Count")
    ws.cell(row=1, column=2, value="Unit Under Test")
    ws.cell(row=1, column=3, value="Function Name")
    ws.cell(row=1, column=4, value="TestCase Name")
    ws.cell(row=1, column=5, value="Input")
    ws.cell(row=2, column=5, value="Variable")
    ws.cell(row=2, column=6, value="DataType")
    ws.cell(row=2, column=7, value="Value")
    ws.cell(row=1, column=8, value="Output")
    ws.cell(row=2, column=8, value="Variable")
    ws.cell(row=2, column=9, value="DataType")
    ws.cell(row=2, column=10, value="Value")
    for m in ["A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:G1", "H1:J1"]: ws.merge_cells(m)
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=10):
        for c in row:
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.border = HEADER_BORDER; c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A3"; wb.save(output_path)


def process_folder_to_excel(cfg, read_html_soup, get_text, parse_summary_table, parse_test_data_section, append_testcase_to_sheet):
    output_dir = Path(cfg["OUTPUT_XLSX_PATH"]).parent
    output_dir.mkdir(parents=True, exist_ok=True)
    html_files = collect_html_files(cfg["INPUT_ROOT"], cfg["RECURSIVE_SEARCH"])
    for html_path in html_files:
        output_path = output_dir / f"{html_path.stem}_TCNames.xlsx"
        create_excel(str(output_path))
        wb = openpyxl.load_workbook(output_path); ws = wb["TCNames"]
        soup = read_html_soup(html_path)
        main_scroller = soup.find(id="main-scroller")
        if main_scroller:
            tc_count = 0
            for testcase_div in main_scroller.find_all("div", class_="testcase", recursive=False):
                tc_count += 1
                h2 = testcase_div.find("h2")
                tc_name = get_text(h2.find("span") if h2 and h2.find("span") else h2)
                uut, sub = parse_summary_table(testcase_div)
                test_data = parse_test_data_section(testcase_div)
                append_testcase_to_sheet(ws, tc_count, tc_name, uut, sub, test_data.get("Input Test Data", []), test_data.get("Expected Test Data", []))
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=10):
            for c in row: c.border = BODY_BORDER; c.alignment = Alignment(vertical="top", wrap_text=True)
        for col_idx, col_cells in enumerate(ws.iter_cols(), start=1):
            m = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(m + 2, 50)
        wb.save(output_path)
