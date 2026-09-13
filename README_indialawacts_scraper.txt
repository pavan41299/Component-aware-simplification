IndiaLawActs dropdown scraper
============================

Purpose
-------
This script opens each Act page on https://indialawacts.in/, expands the
section dropdown / accordion content, extracts one row per section, and writes
an Excel workbook with these columns:

1. Section
2. Clause
3. Description

It also builds a Count_Audit sheet so you can verify that the number of scraped
sections matches the expected count for each Act.

Files
-----
- indialawacts_dropdown_scraper.py
- Output workbook after running: indialawacts_sections_dropdown_audit.xlsx

What the workbook contains
--------------------------
- All_Acts            : all Acts combined
- One sheet per Act   : IPC, BNS, CrPC, BNSS, etc.
- Count_Audit         : expected vs scraped section counts
- Sources             : act URLs and expected counts
- Failures            : pages that failed / timed out

Expected counts used for audit
------------------------------
- IPC 1860: 511
- BNS 2023: 358
- CrPC 1973: 484
- BNSS 2023: 531
- Indian Evidence Act 1872: 167
- BSA 2023: 170
- CPC 1908: 158
- Negotiable Instruments Act 1881: 147
- Indian Divorce Act 1869: 62
- Hindu Marriage Act 1955: 29 or 30
  The site shows an inconsistency for HMA on different site blocks, so the
  audit accepts either 29 or 30.
- Companies Act 2013: 470
- RTI Act 2005: 31
- Motor Vehicles Act 1988: 217
- Wildlife (Protection) Act 1972: 66
- Environment (Protection) Act 1986: 26
- Biological Diversity Act 2002: 65
- Forest (Conservation) Act 1980: 5

Install
-------
pip install playwright openpyxl beautifulsoup4 requests
playwright install chromium

Run
---
python indialawacts_dropdown_scraper.py

Optional
--------
Run browser in visible mode:
python indialawacts_dropdown_scraper.py --headed

Change browser timeout:
python indialawacts_dropdown_scraper.py --timeout-ms 90000

How it works
------------
1. Tries a normal HTML fetch first.
2. If the section parse looks incomplete, it falls back to Playwright.
3. In Playwright mode it scrolls through the page and clicks common collapse /
   accordion / details controls to expand all section dropdown content.
4. It parses each section into:
   - Section   : e.g. Section 10A: Dissolution of marriage by mutual consent
   - Clause    : main clause text
   - Description: explanation / exception / illustration / note text combined
5. It deduplicates by section number per Act.
6. It writes a Count_Audit sheet comparing the scraped unique section count to
   the expected count.

Notes
-----
- If some pages time out, rerun with a larger timeout.
- If your environment blocks the normal requests fetch, the browser fallback is
  the important path.
- The script is designed so the section count audit is based on one row per
  section, not one row per clause/exception fragment.
