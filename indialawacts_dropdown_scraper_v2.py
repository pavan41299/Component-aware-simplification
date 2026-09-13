#!/usr/bin/env python3
"""
Expand IndiaLawActs section dropdowns, extract one row per section, and audit
scraped section counts against expected counts.

Output:
    indialawacts_sections_dropdown_audit.xlsx

Primary columns in the workbook:
    1. Act
    2. Section Number
    3. Bracket Text
    4. Section
    5. Clause
    6. Description

Also writes:
    - Count_Audit sheet
    - Sources sheet
    - Failures sheet

Recommended install:
    pip install playwright openpyxl beautifulsoup4 requests
    playwright install chromium

Notes:
- The script first tries a normal HTML fetch.
- If that fails or parses too few sections, it falls back to Playwright and
  explicitly expands accordion / dropdown content on the act page.
- HMA has an on-site count inconsistency (29 vs 30 on different site blocks),
  so the audit accepts either 29 or 30.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except Exception:
    PLAYWRIGHT_AVAILABLE = False
    PlaywrightTimeoutError = Exception  # type: ignore
    sync_playwright = None  # type: ignore


OUTPUT_XLSX = "indialawacts_new_sections_dropdown_audit.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36"
)


@dataclass(frozen=True)
class ActConfig:
    act_name: str
    url: str
    expected_counts: Tuple[int, ...]
    note: str = ""


ACTS: Sequence[ActConfig] = (
    ActConfig("Indian Penal Code (IPC) 1860", "https://indialawacts.in/ipc.php", (511,)),
    ActConfig("Bharatiya Nyaya Sanhita (BNS) 2023", "https://indialawacts.in/bns.php", (358,)),
    ActConfig("Code of Criminal Procedure (CrPC) 1973", "https://indialawacts.in/crpc.php", (484,)),
    ActConfig("Bharatiya Nagarik Suraksha Sanhita (BNSS) 2023", "https://indialawacts.in/bnss.php", (531,)),
    ActConfig("Indian Evidence Act 1872", "https://indialawacts.in/evidence.php", (167,)),
    ActConfig("Bharatiya Sakshya Adhiniyam (BSA) 2023", "https://indialawacts.in/bsa.php", (170,)),
    ActConfig("Code of Civil Procedure (CPC) 1908", "https://indialawacts.in/cpc.php", (158,)),
    ActConfig("Negotiable Instruments Act, 1881", "https://indialawacts.in/ni.php", (147,), "Count based on the site section listing."),
    ActConfig("Indian Divorce Act, 1869", "https://indialawacts.in/ida.php", (62,)),
    ActConfig("Hindu Marriage Act, 1955", "https://indialawacts.in/hma.php", (29, 30), "Site shows 29 in one block and 30 in another."),
    ActConfig("Companies Act, 2013", "https://indialawacts.in/companies.php", (470,)),
    ActConfig("Right to Information Act, 2005", "https://indialawacts.in/rti.php", (31,), "Count based on the site section listing."),
    ActConfig("Motor Vehicles Act, 1988", "https://indialawacts.in/mva.php", (217,)),
    ActConfig("Wildlife (Protection) Act, 1972", "https://indialawacts.in/wildlife.php", (66,)),
    ActConfig("Environment (Protection) Act, 1986", "https://indialawacts.in/epa.php", (26,)),
    ActConfig("Biological Diversity Act, 2002", "https://indialawacts.in/bda.php", (65,)),
    ActConfig("Forest (Conservation) Act, 1980", "https://indialawacts.in/fca.php", (5,)),
)


SECTION_RE = re.compile(
    r"^(?:#+\s*)?Section\s+([A-Za-z0-9()./\-]+)\s*[:.\-–—]?\s*(.*?)\s*$",
    re.IGNORECASE,
)
CHAPTER_RE = re.compile(r"^(?:#+\s*)?Chapter\s+[A-Z0-9IVXLCDM]+\b", re.IGNORECASE)
CLAUSE_LABEL_RE = re.compile(r"^Clause\s*:?\s*(.*)$", re.IGNORECASE)
DESCRIPTION_LABEL_RE = re.compile(
    r"^(Explanation|Explanations|Exception|Exceptions|First Exception|Second Exception|Third Exception|Fourth Exception|Fifth Exception|Illustration|Illustrations|Comment|Comments|Notes?)\b\s*[:.\-–—]?\s*(.*)$",
    re.IGNORECASE,
)
AMENDMENT_NOTE_RE = re.compile(r"^(Ins\.|Subs\.|Omitted by|Inserted by|Substituted by|Repealed by|State Amendments?)\b", re.IGNORECASE)
FOOTER_RE = re.compile(
    r"^(Disclaimer|Privacy Policy|Terms & Conditions|Comprehensive Legal Resource|©\s*2025|Verify with:)\b",
    re.IGNORECASE,
)


@dataclass
class SectionRecord:
    act_name: str
    section_no: str
    section_title: str
    clause: str
    description: str
    url: str

    @property
    def bracket_text(self) -> str:
        title = self.section_title.strip()
        if title.startswith("(") and title.endswith(")") and len(title) >= 2:
            return title[1:-1].strip()
        return title

    @property
    def section_display(self) -> str:
        title = f": {self.section_title}" if self.section_title else ""
        return f"Section {self.section_no}{title}"

    @property
    def combined_section_display(self) -> str:
        return f"{self.act_name} - {self.section_display}"


class ExtractionError(RuntimeError):
    pass


# -----------------------------
# Fetch helpers
# -----------------------------

def fetch_html_requests(url: str, timeout: int = 45, retries: int = 3) -> str:
    last_error: Optional[Exception] = None
    session = requests.Session()
    headers = {"User-Agent": USER_AGENT}
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, headers=headers, timeout=timeout)
            resp.raise_for_status()
            resp.encoding = resp.encoding or "utf-8"
            return resp.text
        except Exception as exc:  # pragma: no cover - network dependent
            last_error = exc
            time.sleep(min(2 * attempt, 5))
    raise ExtractionError(f"requests fetch failed for {url}: {last_error}")


EXPAND_SELECTORS: Sequence[str] = (
    "button[aria-expanded='false']",
    "button.accordion-button",
    "[data-bs-toggle='collapse']",
    "[data-toggle='collapse']",
    "summary",
    "button:has-text('Read more')",
    "button:has-text('Show more')",
    "a:has-text('Read more')",
)


def fetch_text_playwright(url: str, timeout_ms: int = 60000, headless: bool = True) -> str:
    if not PLAYWRIGHT_AVAILABLE:
        raise ExtractionError(
            "Playwright is not installed. Run: pip install playwright && playwright install chromium"
        )

    with sync_playwright() as p:  # pragma: no cover - browser dependent
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page(user_agent=USER_AGENT, viewport={"width": 1440, "height": 2000})
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        try:
            page.wait_for_load_state("networkidle", timeout=timeout_ms)
        except PlaywrightTimeoutError:
            pass

        # Expand accordions / dropdowns multiple passes while scrolling.
        for _ in range(8):
            try:
                page.mouse.wheel(0, 4000)
            except Exception:
                pass
            for selector in EXPAND_SELECTORS:
                try:
                    locator = page.locator(selector)
                    count = min(locator.count(), 5000)
                    for i in range(count):
                        try:
                            item = locator.nth(i)
                            item.scroll_into_view_if_needed(timeout=1500)
                            item.click(timeout=1200, force=True)
                            page.wait_for_timeout(40)
                        except Exception:
                            continue
                except Exception:
                    continue
            try:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                page.wait_for_timeout(250)
            except Exception:
                pass

        # Force open <details> blocks if present.
        try:
            page.evaluate(
                """
                () => {
                  document.querySelectorAll('details').forEach(d => d.open = true);
                  document.querySelectorAll('[aria-expanded="false"]').forEach(el => el.click());
                }
                """
            )
        except Exception:
            pass

        page.wait_for_timeout(700)
        text = page.locator("body").inner_text(timeout=timeout_ms)
        browser.close()
        return text


# -----------------------------
# Parse helpers
# -----------------------------

def normalize_text(text: str) -> List[str]:
    text = text.replace("\xa0", " ")
    text = text.replace("\r", "\n")
    out: List[str] = []
    for raw in text.splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if line:
            out.append(line)
    return out


def html_to_text_lines(html: str) -> List[str]:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "footer"]):
        tag.decompose()
    return normalize_text(soup.get_text("\n"))


def strip_footer(lines: List[str]) -> List[str]:
    for i, line in enumerate(lines):
        if FOOTER_RE.match(line):
            return lines[:i]
    return lines


def start_at_first_section(lines: List[str]) -> List[str]:
    for i, line in enumerate(lines):
        if SECTION_RE.match(line):
            return lines[i:]
    return lines


def section_blocks(lines: List[str]) -> List[Tuple[str, str, List[str]]]:
    blocks: List[Tuple[str, str, List[str]]] = []
    current_no: Optional[str] = None
    current_title: str = ""
    body: List[str] = []

    def flush() -> None:
        nonlocal current_no, current_title, body
        if current_no is not None:
            blocks.append((current_no, current_title.strip(), body[:]))
        current_no = None
        current_title = ""
        body = []

    i = 0
    while i < len(lines):
        line = lines[i]
        if CHAPTER_RE.match(line):
            i += 1
            continue
        m = SECTION_RE.match(line)
        if m:
            flush()
            current_no = m.group(1).strip()
            current_title = (m.group(2) or "").strip()
            # Occasionally a heading title spills to the next line.
            if not current_title and i + 1 < len(lines):
                nxt = lines[i + 1]
                if not SECTION_RE.match(nxt) and not CHAPTER_RE.match(nxt):
                    if not CLAUSE_LABEL_RE.match(nxt) and not DESCRIPTION_LABEL_RE.match(nxt):
                        current_title = nxt
                        i += 1
            i += 1
            continue
        if current_no is not None:
            body.append(line)
        i += 1

    flush()
    return blocks


def collapse_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def join_paragraphs(lines: Iterable[str]) -> str:
    cleaned = [collapse_spaces(x) for x in lines if collapse_spaces(x)]
    return "\n".join(cleaned).strip()


def parse_section_body(body_lines: List[str]) -> Tuple[str, str]:
    clause_lines: List[str] = []
    description_parts: List[str] = []

    current_mode: Optional[str] = None
    current_label: Optional[str] = None
    current_buffer: List[str] = []

    def flush_desc_block() -> None:
        nonlocal current_label, current_buffer
        if current_label and current_buffer:
            block = join_paragraphs(current_buffer)
            if block:
                description_parts.append(f"{current_label}: {block}")
        current_label = None
        current_buffer = []

    def flush_clause_buffer() -> None:
        nonlocal clause_lines
        if current_buffer:
            clause_lines.extend(current_buffer)

    i = 0
    while i < len(body_lines):
        line = body_lines[i].strip()
        i += 1
        if not line:
            continue
        if FOOTER_RE.match(line) or CHAPTER_RE.match(line) or SECTION_RE.match(line):
            break
        if AMENDMENT_NOTE_RE.match(line):
            continue
        if line in {"[Omitted]", "Omitted"}:
            if not clause_lines:
                clause_lines.append(line)
            continue

        cm = CLAUSE_LABEL_RE.match(line)
        if cm:
            if current_mode == "description":
                flush_desc_block()
            elif current_mode == "clause":
                flush_clause_buffer()
                current_buffer = []
            current_mode = "clause"
            current_buffer = []
            first = cm.group(1).strip()
            if first:
                current_buffer.append(first)
            continue

        dm = DESCRIPTION_LABEL_RE.match(line)
        if dm:
            if current_mode == "clause":
                flush_clause_buffer()
                current_buffer = []
            elif current_mode == "description":
                flush_desc_block()
            current_mode = "description"
            current_label = collapse_spaces(dm.group(1))
            current_buffer = []
            first = (dm.group(2) or "").strip()
            if first:
                current_buffer.append(first)
            continue

        low = line.lower()
        if low.startswith(("source:", "related sections", "view all", "download ")):
            continue

        if current_mode is None:
            current_mode = "clause"
            current_buffer = []
            current_buffer.append(line)
            continue

        current_buffer.append(line)

    if current_mode == "clause":
        flush_clause_buffer()
    elif current_mode == "description":
        flush_desc_block()

    clause = join_paragraphs(clause_lines)
    description = "\n\n".join(p for p in description_parts if p.strip()).strip()
    return clause, description


def parse_records(act: ActConfig, text: str) -> List[SectionRecord]:
    lines = strip_footer(start_at_first_section(normalize_text(text)))
    blocks = section_blocks(lines)
    records: List[SectionRecord] = []
    for sec_no, sec_title, body in blocks:
        clause, description = parse_section_body(body)
        records.append(
            SectionRecord(
                act_name=act.act_name,
                section_no=sec_no,
                section_title=sec_title,
                clause=clause,
                description=description,
                url=act.url,
            )
        )
    return records


# -----------------------------
# Workbook helpers
# -----------------------------

def short_sheet_title(act_name: str) -> str:
    replacements = {
        "Indian Penal Code (IPC) 1860": "IPC",
        "Bharatiya Nyaya Sanhita (BNS) 2023": "BNS",
        "Code of Criminal Procedure (CrPC) 1973": "CrPC",
        "Bharatiya Nagarik Suraksha Sanhita (BNSS) 2023": "BNSS",
        "Indian Evidence Act 1872": "IEA",
        "Bharatiya Sakshya Adhiniyam (BSA) 2023": "BSA",
        "Code of Civil Procedure (CPC) 1908": "CPC",
        "Negotiable Instruments Act, 1881": "NI_Act",
        "Indian Divorce Act, 1869": "IDA",
        "Hindu Marriage Act, 1955": "HMA",
        "Companies Act, 2013": "Companies",
        "Right to Information Act, 2005": "RTI",
        "Motor Vehicles Act, 1988": "MVA",
        "Wildlife (Protection) Act, 1972": "Wildlife",
        "Environment (Protection) Act, 1986": "EPA",
        "Biological Diversity Act, 2002": "BDA",
        "Forest (Conservation) Act, 1980": "FCA",
    }
    return replacements.get(act_name, act_name[:31])


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN_GRAY = Side(style="thin", color="D9D9D9")
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
STATUS_OK_FILL = PatternFill("solid", fgColor="E2F0D9")
STATUS_BAD_FILL = PatternFill("solid", fgColor="FCE4D6")
STATUS_WARN_FILL = PatternFill("solid", fgColor="FFF2CC")


def style_header(ws, row: int = 1) -> None:
    for cell in ws[row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_TOP
        cell.border = Border(bottom=THIN_GRAY)


def style_data(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.alignment = WRAP_TOP


def create_workbook(
    all_records: List[SectionRecord],
    per_act_records: Dict[str, List[SectionRecord]],
    audit_rows: List[Tuple[str, str, int, int, str, str]],
    failures: List[Tuple[str, str, str]],
) -> Workbook:
    wb = Workbook()

    # All acts sheet
    ws = wb.active
    ws.title = "All_Acts"
    ws.append(["Act", "Section Number", "Bracket Text", "Section", "Clause", "Description"])
    for rec in all_records:
        ws.append([
            rec.act_name,
            rec.section_no,
            rec.bracket_text,
            rec.section_display,
            rec.clause,
            rec.description,
        ])
    style_header(ws)
    style_data(ws, 2, ws.max_row, 1, 6)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 95
    ws.column_dimensions["F"].width = 120

    # Per-act sheets
    for act_name, records in per_act_records.items():
        aws = wb.create_sheet(short_sheet_title(act_name))
        aws.append(["Act", "Section Number", "Bracket Text", "Section", "Clause", "Description"])
        for rec in records:
            aws.append([
                rec.act_name,
                rec.section_no,
                rec.bracket_text,
                rec.section_display,
                rec.clause,
                rec.description,
            ])
        style_header(aws)
        style_data(aws, 2, aws.max_row, 1, 6)
        aws.freeze_panes = "A2"
        aws.auto_filter.ref = f"A1:F{aws.max_row}"
        aws.column_dimensions["A"].width = 42
        aws.column_dimensions["B"].width = 16
        aws.column_dimensions["C"].width = 48
        aws.column_dimensions["D"].width = 44
        aws.column_dimensions["E"].width = 95
        aws.column_dimensions["F"].width = 120

    # Count audit sheet
    aud = wb.create_sheet("Count_Audit")
    aud.append(["Act", "Expected Count(s)", "Scraped Sections", "Unique Sections", "Status", "Note"])
    for row in audit_rows:
        aud.append(list(row))
    style_header(aud)
    style_data(aud, 2, aud.max_row, 1, 6)
    aud.freeze_panes = "A2"
    aud.auto_filter.ref = f"A1:F{aud.max_row}"
    aud.column_dimensions["A"].width = 42
    aud.column_dimensions["B"].width = 18
    aud.column_dimensions["C"].width = 16
    aud.column_dimensions["D"].width = 16
    aud.column_dimensions["E"].width = 14
    aud.column_dimensions["F"].width = 60
    for row in range(2, aud.max_row + 1):
        status_cell = aud[f"E{row}"]
        if status_cell.value == "MATCH":
            status_cell.fill = STATUS_OK_FILL
        elif status_cell.value == "WARN":
            status_cell.fill = STATUS_WARN_FILL
        else:
            status_cell.fill = STATUS_BAD_FILL

    # Sources
    src = wb.create_sheet("Sources")
    src.append(["Act", "URL", "Expected Count(s)", "Notes"])
    for act in ACTS:
        src.append([
            act.act_name,
            act.url,
            "/".join(str(x) for x in act.expected_counts),
            act.note,
        ])
    style_header(src)
    style_data(src, 2, src.max_row, 1, 4)
    src.freeze_panes = "A2"
    src.auto_filter.ref = f"A1:D{src.max_row}"
    src.column_dimensions["A"].width = 42
    src.column_dimensions["B"].width = 44
    src.column_dimensions["C"].width = 18
    src.column_dimensions["D"].width = 70

    # Failures
    fail = wb.create_sheet("Failures")
    fail.append(["Act", "URL", "Error"])
    for item in failures:
        fail.append(list(item))
    style_header(fail)
    style_data(fail, 2, fail.max_row, 1, 3)
    fail.freeze_panes = "A2"
    fail.auto_filter.ref = f"A1:C{fail.max_row}"
    fail.column_dimensions["A"].width = 42
    fail.column_dimensions["B"].width = 44
    fail.column_dimensions["C"].width = 100

    return wb


# -----------------------------
# Main workflow
# -----------------------------

def try_extract_with_requests(act: ActConfig) -> List[SectionRecord]:
    html = fetch_html_requests(act.url)
    lines = html_to_text_lines(html)
    text = "\n".join(lines)
    return parse_records(act, text)


def extract_act(act: ActConfig, timeout_ms: int, headless: bool) -> List[SectionRecord]:
    # Try static HTML first.
    req_error: Optional[Exception] = None
    try:
        records = try_extract_with_requests(act)
        unique_sections = len({r.section_no for r in records})
        if unique_sections >= min(act.expected_counts):
            return records
    except Exception as exc:
        req_error = exc

    # Browser fallback.
    browser_error: Optional[Exception] = None
    try:
        text = fetch_text_playwright(act.url, timeout_ms=timeout_ms, headless=headless)
        records = parse_records(act, text)
        if records:
            return records
    except Exception as exc:
        browser_error = exc

    raise ExtractionError(
        f"Could not extract {act.act_name}. requests_error={req_error}; browser_error={browser_error}"
    )


def build_audit_row(act: ActConfig, records: List[SectionRecord]) -> Tuple[str, str, int, int, str, str]:
    scraped_sections = len(records)
    unique_sections = len({rec.section_no for rec in records})
    expected_str = "/".join(str(x) for x in act.expected_counts)
    if unique_sections in act.expected_counts:
        status = "MATCH"
    elif unique_sections == 0:
        status = "FAIL"
    else:
        status = "FAIL"
    note = act.note
    return (act.act_name, expected_str, scraped_sections, unique_sections, status, note)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape IndiaLawActs dropdown text into an Excel workbook.")
    parser.add_argument("--output", default=OUTPUT_XLSX, help="Output workbook path")
    parser.add_argument("--timeout-ms", type=int, default=60000, help="Per-page browser timeout in milliseconds")
    parser.add_argument("--headed", action="store_true", help="Run browser in headed mode")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    all_records: List[SectionRecord] = []
    per_act_records: Dict[str, List[SectionRecord]] = {}
    failures: List[Tuple[str, str, str]] = []
    audit_rows: List[Tuple[str, str, int, int, str, str]] = []

    print(f"Output workbook: {args.output}")
    for act in ACTS:
        print(f"\n[START] {act.act_name} -> {act.url}")
        try:
            records = extract_act(act, timeout_ms=args.timeout_ms, headless=not args.headed)
            # Deduplicate by section number, keeping the first parsed instance.
            deduped: List[SectionRecord] = []
            seen = set()
            for rec in records:
                if rec.section_no in seen:
                    continue
                seen.add(rec.section_no)
                deduped.append(rec)
            records = deduped
            per_act_records[act.act_name] = records
            all_records.extend(records)
            audit_row = build_audit_row(act, records)
            audit_rows.append(audit_row)
            print(f"[OK] scraped={len(records)} unique_sections={audit_row[3]} expected={audit_row[1]} status={audit_row[4]}")
        except Exception as exc:
            failures.append((act.act_name, act.url, str(exc)))
            audit_rows.append((act.act_name, "/".join(map(str, act.expected_counts)), 0, 0, "FAIL", f"{act.note} {exc}".strip()))
            print(f"[FAIL] {exc}")

    wb = create_workbook(all_records, per_act_records, audit_rows, failures)
    wb.save(args.output)

    matches = sum(1 for row in audit_rows if row[4] == "MATCH")
    fails = sum(1 for row in audit_rows if row[4] == "FAIL")
    print("\nDone.")
    print(f"Total section rows written: {len(all_records)}")
    print(f"Acts matched expected count: {matches}/{len(ACTS)}")
    print(f"Acts failed: {fails}")
    print(f"Workbook saved to: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
