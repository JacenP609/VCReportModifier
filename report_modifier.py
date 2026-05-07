import datetime
import json
import random
import re
import shutil
from datetime import timedelta
from pathlib import Path
from typing import Callable, Dict, List


class DateState:
    def __init__(self, config: Dict):
        self.config = config
        self.creation_dt = datetime.datetime.strptime(config["TEST_START_DATE"], config["DATE_FORMAT"])
        self.execution_dt = datetime.datetime.strptime(config["EXECUTION_START_DATE"], config["DATE_FORMAT"])
        self.tc_count = 0

    def next_creation_date(self) -> str:
        cur = self.creation_dt
        self.creation_dt += timedelta(minutes=random.randint(self.config["CREATION_MIN_STEP_MINUTES"], self.config["CREATION_MAX_STEP_MINUTES"]))
        return cur.strftime(self.config["DATE_FORMAT"]).upper()

    def next_execution_date(self) -> str:
        cur = self.execution_dt
        self.tc_count += 1
        if self.tc_count % self.config["EXECUTION_STEP_PER_TC_COUNT"] == 0:
            self.execution_dt += timedelta(minutes=self.config["EXECUTION_STEP_MINUTES"])
        return cur.strftime(self.config["DATE_FORMAT"]).upper()


def collect_html_files(input_root: str, recursive: bool) -> List[Path]:
    root = Path(input_root)
    files = [p for p in (root.rglob("*") if recursive else root.iterdir()) if p.is_file() and p.suffix.lower() in [".html", ".htm"]]
    return sorted(files, key=lambda p: str(p).lower())


def get_table_value_by_header(table, header_name: str, clean_text: Callable[[str], str]) -> str:
    for tr in table.find_all("tr"):
        th, td = tr.find("th"), tr.find("td")
        if th and td and clean_text(th.get_text()) == header_name:
            return clean_text(td.get_text())
    return ""


def set_table_value_by_header(table, header_name: str, value: str, clean_text: Callable[[str], str]):
    for tr in table.find_all("tr"):
        th, td = tr.find("th"), tr.find("td")
        if th and td and clean_text(th.get_text()) == header_name:
            td.clear(); td.append(value); return


def process_html_file(html_path: Path, date_state: DateState, clean_text, read_html_soup, safe_name, extract_func_name):
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


def process_all_html_files(config: Dict, clean_text, read_html_soup, safe_name, extract_func_name):
    random.seed(config["RANDOM_SEED"])
    src, dst = Path(config["ORIGINAL_ROOT"]), Path(config["NEW_ROOT"])
    if dst.exists() and config["OVERWRITE_NEW_ROOT"]:
        shutil.rmtree(dst)
    shutil.copytree(src, dst)

    # keep JSON cache loading available for future expansion
    _ = json, config.get("JSON_FILE_MAP"), config.get("SWE4_JSON_ROOT")

    ds = DateState(config)
    for html_path in collect_html_files(config["NEW_ROOT"], config["RECURSIVE_SEARCH"]):
        process_html_file(html_path, ds, clean_text, read_html_soup, safe_name, extract_func_name)
