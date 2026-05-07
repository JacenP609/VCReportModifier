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


def get_text(tag, clean_text) -> str:
    return "" if tag is None else clean_text(tag.get_text(" ", strip=True))


def create_excel(output_path: Path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "TCNames"
    headers = ["Source HTML", "TC-Count", "Unit Under Test", "Function Name", "TestCase Name", "Input", "", "", "Output", "", ""]
    subs = ["", "", "", "", "", "Variable", "DataType", "Value", "Variable", "DataType", "Value"]
    for i, v in enumerate(headers, 1): ws.cell(row=1, column=i, value=v)
    for i, v in enumerate(subs, 1): ws.cell(row=2, column=i, value=v)
    for m in ["A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:H1", "I1:K1"]: ws.merge_cells(m)
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=11):
        for c in row: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.border = HEADER_BORDER; c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A3"; wb.save(output_path)


def parse_html_file_to_sheet(html_path: Path, ws, clean_text, read_html_soup):
    soup = read_html_soup(html_path)
    main = soup.find(id="main-scroller")
    if not main:
        return
    tc_count = 0
    for tc in main.find_all("div", class_="testcase", recursive=False):
        tc_count += 1
        h2 = tc.find("h2")
        tc_name = get_text(h2.find("span") if h2 and h2.find("span") else h2, clean_text)
        r = ws.max_row + 2
        ws.cell(row=r, column=1, value=html_path.name)
        ws.cell(row=r, column=2, value=tc_count)
        ws.cell(row=r, column=5, value=tc_name)


def finalize_sheet(ws):
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=11):
        for c in row: c.border = BODY_BORDER; c.alignment = Alignment(vertical="top", wrap_text=True)
    for col_idx, col_cells in enumerate(ws.iter_cols(), start=1):
        m = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(m + 2, 50)


def process_folder_to_excel_per_html(config, clean_text, read_html_soup):
    output_dir = Path(config["OUTPUT_XLSX_PATH"]).parent
    output_dir.mkdir(parents=True, exist_ok=True)
    html_paths = collect_html_files(config["NEW_ROOT"], config["RECURSIVE_SEARCH"])
    for html_path in html_paths:
        out_path = output_dir / f"{html_path.stem}_TCNames.xlsx"
        create_excel(out_path)
        wb = openpyxl.load_workbook(out_path); ws = wb["TCNames"]
        parse_html_file_to_sheet(html_path, ws, clean_text, read_html_soup)
        finalize_sheet(ws)
        wb.save(out_path)
