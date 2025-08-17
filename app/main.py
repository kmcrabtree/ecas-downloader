
import os
import re
import time
import json
import getpass
from datetime import datetime, date
from pathlib import Path

import pandas as pd
from PyPDF2 import PdfReader

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def sanitize(s: str) -> str:
    if not s:
        return ""
    s = s.replace("/", "-").replace("\\", "-").replace(":", " - ")
    s = s.replace("*", "").replace("?", "").replace('"', "'").replace("<", "(").replace(">", ")").replace("|", "-")
    return re.sub(r"\s+", " ", s).strip()

def parse_date_any(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    fmts = ["%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y", "%b %d, %Y", "%B %d, %Y"]
    for f in fmts:
        try:
            dt = datetime.strptime(s, f)
            return dt.strftime("%m-%d-%Y")
        except Exception:
            pass
    m = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", s)
    if m:
        try:
            dt = datetime.strptime(m.group(0), "%Y-%m-%d")
            return dt.strftime("%m-%d-%Y")
        except Exception:
            pass
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", s)
    if m:
        try:
            dt = datetime.strptime(m.group(0), "%m/%d/%Y")
            return dt.strftime("%m-%d-%Y")
        except Exception:
            pass
    return s

def extract_relevant_dates_text(text: str) -> str:
    if not text:
        return ""
    findings = []
    patterns = [
        (r"(hearing (?:is )?set (?:for|on)\s+(?P<date>(?:\w{3,9}\s+\d{1,2},\s+\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4})))", "HEARING"),
        (r"(individual hearing on\s+(?P<date>(?:\w{3,9}\s+\d{1,2},\s+\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4})))", "HEARING"),
        (r"(master hearing on\s+(?P<date>(?:\w{3,9}\s+\d{1,2},\s+\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4})))", "HEARING"),
        (r"((?:applications?|relief|brief|evidence|documents?)\s+(?:are\s+)?due\s+(?:by|on)\s+(?P<date>(?:\w{3,9}\s+\d{1,2},\s+\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4})))", "DUE"),
        (r"(deadline(?:s)?\s+(?:set|is|are)\s+(?:for|on)\s+(?P<date>(?:\w{3,9}\s+\d{1,2},\s+\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4})))", "DEADLINE"),
    ]
    for pat, label in patterns:
        for m in re.finditer(pat, text, flags=re.I):
            dt = m.groupdict().get("date") or ""
            if dt:
                findings.append(f"{label} {parse_date_any(dt)}")
    findings = list(dict.fromkeys([f for f in findings if f]))
    return " ; ".join(findings)

def extract_first_page_text(pdf_path: Path) -> str:
    try:
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            if not reader.pages:
                return ""
            return reader.pages[0].extract_text() or ""
    except Exception:
        return ""

def guess_pleading_name_from_text(text: str) -> str:
    if not text:
        return ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    candidate = ""
    best_score = 0.0
    for ln in lines[:60]:
        if len(ln) < 8:
            continue
        letters = [c for c in ln if c.isalpha()]
        if not letters:
            continue
        upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters) if len(letters) else 0
        kw_bonus = 0.15 if re.search(r"\b(MOTION|ORDER|NOTICE|EVIDENCE|APPLICATION|BRIEF|DECLARATION|EXHIBIT|SUBMISSION)\b", ln, re.I) else 0.0
        score = upper_ratio + kw_bonus
        if score > best_score:
            best_score = score
            candidate = ln
    return candidate

def build_new_filename(label, pleading_name, file_date, notes):
    parts = []
    if label:
        parts.append(sanitize(label))
    if pleading_name:
        parts.append(sanitize(pleading_name))
    if file_date:
        parts.append(parse_date_any(file_date))
    if notes:
        parts.append(sanitize(notes))
    name = " - ".join(parts) or "ECAS_Document"
    return name + ".pdf"

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

class ECASScraper:
    def __init__(self, download_dir: Path, selectors_path: Path):
        self.download_dir = download_dir.resolve()
        ensure_dir(self.download_dir)
        with open(selectors_path, "r", encoding="utf-8") as f:
            self.sel = json.load(f)
        self.driver = None
        self.wait = None

    def start(self):
        opts = Options()
        prefs = {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        opts.add_experimental_option("prefs", prefs)
        service = ChromeService(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=opts)
        self.wait = WebDriverWait(self.driver, 30)

    def login(self, email: str, password: str):
        d = self.driver
        d.get("https://portal.eoir.justice.gov/")
        email_box = self.wait.until(EC.presence_of_element_located((By.XPATH, self.sel["login"]["email"])))
        email_box.clear(); email_box.send_keys(email)
        pwd_box = self.wait.until(EC.presence_of_element_located((By.XPATH, self.sel["login"]["password"])))
        pwd_box.clear(); pwd_box.send_keys(password)
        submit = self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["login"]["submit"])))
        submit.click()
        self.wait.until(EC.any_of(
            EC.presence_of_element_located((By.XPATH, self.sel["nav"]["calendar_tab"])),
            EC.presence_of_element_located((By.XPATH, self.sel["nav"]["cases_tab"])),
        ))

    def goto_calendar(self):
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["nav"]["calendar_tab"]))).click()
        time.sleep(2)

    def goto_cases(self):
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["nav"]["cases_tab"]))).click()
        time.sleep(2)

    def iterate_hearings_collect_anums(self, start_date: date, end_date: date):
        self.goto_calendar()
        a_numbers = set()
        months_checked = 0
        while months_checked < 60:
            months_checked += 1
            cells = self.driver.find_elements(By.XPATH, self.sel["calendar"]["day_cells"])
            for cell in cells:
                try:
                    try:
                        cell.find_element(By.XPATH, self.sel["calendar"]["day_number_in_cell"]).click()
                    except Exception:
                        pass
                    try:
                        dots = cell.find_elements(By.XPATH, self.sel["calendar"]["hearing_dot"])
                        if dots:
                            dots[0].click()
                    except Exception:
                        pass
                    time.sleep(0.5)
                    rows = self.driver.find_elements(By.XPATH, self.sel["calendar"]["overlay_row"])
                    if rows:
                        try:
                            rows[0].click()
                            time.sleep(0.5)
                        except Exception:
                            pass
                    popup = self.wait.until(EC.presence_of_element_located((By.XPATH, self.sel["hearing_popup"]["dialog"])))
                    text = popup.text
                    date_match = re.search(r"Hearing Date:\s*([^\n]+)", text, re.I)
                    the_date = date_match.group(1).strip() if date_match else ""
                    in_range = True
                    try:
                        dt = datetime.strptime(the_date, "%m/%d/%Y").date()
                        in_range = (start_date <= dt <= end_date)
                    except Exception:
                        pass
                    if in_range:
                        anums = re.findall(r"A[#\-\s]*\s*(\d{3}[-\s]?\d{3}[-\s]?\d{3})", text)
                        for a in anums:
                            a_numbers.add(re.sub(r"[^0-9]", "", a))
                    try:
                        popup.find_element(By.XPATH, self.sel["hearing_popup"]["close"]).click()
                    except Exception:
                        self.driver.execute_script("document.activeElement?.blur();")
                    time.sleep(0.2)
                except Exception:
                    continue
            try:
                self.driver.find_element(By.XPATH, self.sel["calendar"]["month_next"]).click()
                time.sleep(1.0)
            except Exception:
                break
        return sorted(a_numbers)

    def download_case_docs(self, anumber: str, log_rows: list):
        self.goto_cases()
        inp = self.wait.until(EC.presence_of_element_located((By.XPATH, self.sel["case_search"]["anumber_input"])))
        inp.clear(); inp.send_keys(anumber)
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["case_search"]["search_btn"]))).click()
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["case_search"]["open_case"]))).click()
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.sel["case_docs"]["documents_tab"]))).click()
        time.sleep(1.2)
        rows = self.driver.find_elements(By.XPATH, self.sel["case_docs"]["table_rows"])
        for r in rows:
            try:
                label = r.find_element(By.XPATH, self.sel["case_docs"]["col_label"]).text.strip()
                file_date = r.find_element(By.XPATH, self.sel["case_docs"]["col_date"]).text.strip()
                btns = r.find_elements(By.XPATH, self.sel["case_docs"]["download_btn"])
                if not btns:
                    continue
                before = set(self.download_dir.glob("*.pdf"))
                btns[0].click()
                deadline = time.time() + 120
                downloaded = None
                while time.time() < deadline:
                    now = set(self.download_dir.glob("*.pdf"))
                    new = list(now - before)
                    if new:
                        downloaded = new[0]
                        break
                    time.sleep(0.5)
                if not downloaded:
                    continue
                text = extract_first_page_text(downloaded)
                pleading = guess_pleading_name_from_text(text)
                notes = extract_relevant_dates_text(text) if "ORDER" in label.upper() else ""
                newname = build_new_filename(label, pleading, file_date, notes)
                final_path = self.download_dir / newname
                i = 2
                while final_path.exists():
                    final_path = self.download_dir / (final_path.stem + f" ({i})" + final_path.suffix)
                    i += 1
                downloaded.rename(final_path)
                log_rows.append({
                    "A-Number": anumber,
                    "Document Label (ECAS)": label,
                    "Pleading Name (pg1)": pleading,
                    "Filing Date (ECAS)": file_date,
                    "Relevant Dates (extracted)": notes,
                    "Filename (final)": final_path.name
                })
            except Exception as e:
                print(f"[WARN] Row error for A# {anumber}: {e}")
                continue

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def main():
    print("=== ECAS Auto-Harvest Downloader ===")
    email = input("ECAS Email: ").strip()
    password = getpass.getpass("ECAS Password: ").strip()
    start = input("Start date (YYYY-MM-DD): ").strip()
    end = input("End date (YYYY-MM-DD): ").strip()
    out_dir = input("Download folder (or leave blank for 'downloads_ecas'): ").strip() or "downloads_ecas"
    start_date = datetime.strptime(start, "%Y-%m-%d").date()
    end_date = datetime.strptime(end, "%Y-%m-%d").date()
    download_dir = Path(out_dir).resolve()
    ensure_dir(download_dir)
    selectors_path = Path(__file__).parent / "selectors.json"

    scraper = ECASScraper(download_dir, selectors_path)
    scraper.start()
    scraper.login(email, password)

    print("Harvesting hearings from calendar...")
    anums = scraper.iterate_hearings_collect_anums(start_date, end_date)
    print(f"Found {len(anums)} unique A-numbers in range.")

    log_rows = []
    for a in anums:
        print(f"Downloading documents for A# {a} ...")
        scraper.download_case_docs(a, log_rows)

    df = pd.DataFrame(log_rows)
    xlsx = download_dir / f"ecas_download_log_{start_date}_{end_date}.xlsx"
    df.to_excel(xlsx, index=False)
    print(f"Saved log: {xlsx}")
    print("Finished. You can close Chrome now.")
    time.sleep(2)

if __name__ == "__main__":
    main()
