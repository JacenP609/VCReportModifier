import datetime
import json
import random
import re
import shutil
import time
from datetime import timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from bs4 import BeautifulSoup

JSON_CACHE: Dict[str, Dict[str, Dict[str, List[str]]]] = {}


def parse_report_date(text: str, date_format: str) -> datetime.datetime:
    return datetime.datetime.strptime(text, date_format)


def format_report_date(dt: datetime.datetime, date_format: str) -> str:
    return dt.strftime(date_format).upper()


class DateState:
    def __init__(self, cfg):
        self.cfg = cfg
        self.creation_dt = parse_report_date(cfg["TEST_START_DATE"], cfg["DATE_FORMAT"])
        self.execution_dt = parse_report_date(cfg["EXECUTION_START_DATE"], cfg["DATE_FORMAT"])
        self.tc_count = 0

    def next_creation_date(self) -> str:
        current = self.creation_dt
        step = random.randint(self.cfg["CREATION_MIN_STEP_MINUTES"], self.cfg["CREATION_MAX_STEP_MINUTES"])
        self.creation_dt += timedelta(minutes=step)
        return format_report_date(current, self.cfg["DATE_FORMAT"])

    def next_execution_date(self) -> str:
        current = self.execution_dt
        self.tc_count += 1
        if self.tc_count % self.cfg["EXECUTION_STEP_PER_TC_COUNT"] == 0:
            self.execution_dt += timedelta(minutes=self.cfg["EXECUTION_STEP_MINUTES"])
        return format_report_date(current, self.cfg["DATE_FORMAT"])


def collect_html_files(input_root: str, recursive_search: bool) -> List[Path]:
    root = Path(input_root)
    if recursive_search:
        files = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in [".html", ".htm"]]
    else:
        files = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() in [".html", ".htm"]]
    files.sort(key=lambda p: str(p).lower())
    return files


def get_table_value_by_header(table, header_name: str, clean_text) -> str:
    for tr in table.find_all("tr"):
        th = tr.find("th"); td = tr.find("td")
        if th and td and clean_text(th.get_text()) == header_name:
            return clean_text(td.get_text())
    return ""


def set_table_value_by_header(table, header_name: str, value: str, clean_text) -> None:
    for tr in table.find_all("tr"):
        th = tr.find("th"); td = tr.find("td")
        if th and td and clean_text(th.get_text()) == header_name:
            td.clear(); td.append(value); return


def process_html_file(html_path: Path, date_state: DateState, clean_text, safe_name, extract_func_name, read_html_soup) -> None:
    soup = read_html_soup(html_path)
    tc_counter: Dict[str, int] = {}
    for a in soup.find_all("a", id=re.compile(r"^TestCaseConfiguration_\d+$")):
        m = re.match(r"^TestCaseConfiguration_(\d+)$", a.get("id", ""))
        if not m:
            continue
        idx = m.group(1)
        block = a.find_parent("div", class_="report-block")
        table = block.find("table") if block else None
        if not table:
            continue
        subprogram = get_table_value_by_header(table, "Subprogram", clean_text) or get_table_value_by_header(table, "Test Case Name", clean_text) or "Function"
        fn = safe_name(extract_func_name(subprogram)) or "Function"
        tc_counter[fn.lower()] = tc_counter.get(fn.lower(), 0) + 1
        display_name = f"{fn}-{tc_counter[fn.lower()]}"
        set_table_value_by_header(table, "Date of Creation", date_state.next_creation_date(), clean_text)
        set_table_value_by_header(table, "Date of Execution", date_state.next_execution_date(), clean_text)
        testcase_anchor = soup.find("a", id=re.compile(rf"^TestCase_{idx}_"))
        if testcase_anchor:
            testcase_div = testcase_anchor.find_parent("div", class_="testcase")
            span = testcase_div.find("span", class_="testcase_name") if testcase_div else None
            if span:
                span.clear(); span.append(display_name)
    html_path.write_text(str(soup), encoding="utf-8")


def process_all_html_files(cfg, clean_text, safe_name, extract_func_name, read_html_soup) -> None:
    random.seed(cfg["RANDOM_SEED"])
    src = Path(cfg["ORIGINAL_ROOT"]); dst = Path(cfg["NEW_ROOT"])
    if dst.exists() and cfg["OVERWRITE_NEW_ROOT"]:
        shutil.rmtree(dst)
    shutil.copytree(src, dst)
    ds = DateState(cfg)
    for html_path in collect_html_files(cfg["NEW_ROOT"], cfg["RECURSIVE_SEARCH"]):
        process_html_file(html_path, ds, clean_text, safe_name, extract_func_name, read_html_soup)
