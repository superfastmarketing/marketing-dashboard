"""
Log in to LP, open the Report Generator, and print all form field names/IDs.
Run this once to map the exact selectors needed.
"""
import os, time
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv(Path(__file__).parent.parent / ".env")
LP_URL  = os.getenv("LP_URL").rstrip("/")
LP_USER = os.getenv("LP_USERNAME")
LP_PASS = os.getenv("LP_PASSWORD")

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, slow_mo=300)
    context = browser.new_context()
    page = context.new_page()

    # Login
    page.goto(f"{LP_URL}/Login.aspx", wait_until="networkidle")
    page.wait_for_selector('#txtUserName', state="visible", timeout=15000)
    page.locator('#txtUserName').fill(LP_USER)
    page.locator('#txtPassword').fill(LP_PASS)
    page.locator('#btnLogin').click()
    page.wait_for_load_state("networkidle", timeout=15000)
    if "Start.html" not in page.url:
        page.evaluate("document.querySelector('form').submit()")
        page.wait_for_load_state("networkidle", timeout=10000)
    print(f"Logged in: {page.url}")

    # Open report generator
    page.goto(f"{LP_URL}/ReportCtrl.html?BC=Reports|ReportGenerator", wait_until="networkidle")
    time.sleep(2)
    print(f"Report page: {page.url}\n")

    # Select group and wait for ReportName to populate
    print("Selecting Marketing group...")
    page.select_option('#ReportGroup', label='Marketing')
    time.sleep(3)

    # Print what's in ReportName now
    opts = page.locator('#ReportName').locator('option').all()
    print("ReportName options after selecting Marketing:")
    for o in opts:
        print(f"  value={o.get_attribute('value')!s:<6} label={o.inner_text()}")

    print("\nSelecting Marketing Sub-Source Report...")
    page.select_option('#ReportName', label='Marketing Sub-Source Report')
    time.sleep(3)

    print("\n=== VISIBLE SELECTS AFTER REPORT SELECTED ===")
    for s in page.locator("select").all():
        if not s.is_visible():
            continue
        id_  = s.get_attribute("id") or s.get_attribute("name") or ""
        opts = s.locator("option").all_text_contents()
        val  = s.input_value() if s.count() else ""
        print(f"  id={id_:<20} current={val:<25} options={opts[:8]}")

    # Check all TextBox IDs specifically
    print("\n=== TEXTBOX FIELDS (TextBox1-8) ===")
    for i in range(1, 9):
        el = page.locator(f'#TextBox{i}')
        if el.count() > 0:
            visible = el.is_visible()
            val = el.input_value() if visible else "(hidden)"
            ph  = el.get_attribute("placeholder") or ""
            print(f"  TextBox{i}  visible={visible}  placeholder={ph!s:<20}  value={val}")

    # Check format dropdown options now
    print("\n=== FORMAT OPTIONS ===")
    fmt_opts = page.locator('#rFormat option').all_text_contents()
    print(f"  {fmt_opts}")

    # Check all visible inputs
    print("\n=== ALL VISIBLE INPUTS ===")
    for inp in page.locator("input").all():
        if not inp.is_visible():
            continue
        id_  = inp.get_attribute("id") or inp.get_attribute("name") or ""
        type_ = inp.get_attribute("type") or "text"
        ph   = inp.get_attribute("placeholder") or ""
        val  = inp.input_value()
        if id_ not in ("chatTextBox", "txtFooterSubscribe", "srch_phone", "srch_lastname",
                       "srch_prospectid", "srch_jobnumber", "srch_apptdate"):
            print(f"  id={id_:<20} type={type_:<10} placeholder={ph:<20} value={val}")

    input("\nPress Enter to close...")
    browser.close()
