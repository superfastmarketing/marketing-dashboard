"""
Download all 5 LeadPerfection reports for 4 date periods (20 files total).
Saves Excel files to ../reports/<period>/ folders.

Usage:
    python download_reports.py
"""

import os, sys, time
from datetime import date, timedelta
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv(Path(__file__).parent.parent / ".env")

LP_URL  = os.getenv("LP_URL", "https://p5wfa.leadperfection.com").rstrip("/")
LP_USER = os.getenv("LP_USERNAME")
LP_PASS = os.getenv("LP_PASSWORD")

REPORTS_DIR = Path(__file__).parent.parent / "reports"
REPORTS_DIR.mkdir(exist_ok=True)


# ── Date range helpers ────────────────────────────────────────────────────────

def date_ranges():
    today = date.today()

    # Prior week = Sun–Sat of the previous week
    days_since_sunday = (today.weekday() + 1) % 7  # Sun=0 … Sat=6
    last_sat  = today - timedelta(days=days_since_sunday + 1)
    last_sun  = last_sat - timedelta(days=6)

    # Prior month
    first_of_month = today.replace(day=1)
    pm_end   = first_of_month - timedelta(days=1)
    pm_start = pm_end.replace(day=1)

    return {
        "prior_week":  (last_sun,             last_sat),
        "prior_month": (pm_start,             pm_end),
        "mtd":         (today.replace(day=1), today),
        "ytd":         (today.replace(month=1, day=1), today),
    }


def fmt(d: date) -> str:
    return d.strftime("%m/%d/%Y")


# ── Report definitions ────────────────────────────────────────────────────────

REPORTS = [
    {
        "name":   "marketing_sub_source",
        "group":  "Marketing",
        "report": "Marketing Sub-Source Report",
        "source": "ALL",
        "market": "ALL",
    },
    {
        "name":   "appt_by_setter",
        "group":  "Call Center",
        "report": "Appointment Statistics by Setter",
        "source": "--All Sources--",
        "market": "--ALL--",
    },
    {
        "name":   "appt_by_product",
        "group":  "Call Center",
        "report": "Appointment Statistics by Product",
        "market": "--ALL--",
    },
    {
        "name":   "appt_by_subsource",
        "group":  "Call Center",
        "report": "Appointment Statistics by Sub-Source",
        "source": "--All Sources--",
        "market": "--ALL--",
    },
    {
        "name":   "appt_by_source",
        "group":  "Call Center",
        "report": "Appointment Statistics by Source",
        "source": "--All Sources--",
    },
    {
        "name":   "dispo_distribution",
        "group":  "Call Center",
        "report": "Dispo Distribution",
        "source": "--All Sources--",
        "market": "--ALL--",
    },
    {
        "name":   "appt_by_promoter",
        "group":  "Call Center",
        "report": "Appointment Statistics by Promoter",
        "source": "--All Sources--",
        "market": "--ALL--",
    },
]


# ── Playwright helpers ────────────────────────────────────────────────────────

def login(page):
    print("Logging in to LeadPerfection...")
    page.goto(f"{LP_URL}/Login.aspx", wait_until="networkidle")
    page.wait_for_selector('#txtUserName', state="visible", timeout=15000)
    page.locator('#txtUserName').fill(LP_USER)
    page.locator('#txtPassword').fill(LP_PASS)
    page.locator('#btnLogin').click()
    page.wait_for_load_state("networkidle", timeout=15000)

    print(f"  After login, URL: {page.url}")

    # LP sometimes shows a "Enter the code." verification step on Login.aspx
    # before redirecting to Start.html.  Handle it here.
    for _attempt in range(5):
        if "Start.html" in page.url:
            break
        if "Login.aspx" not in page.url:
            # Some other intermediate page — just wait
            page.wait_for_load_state("networkidle", timeout=10000)
            continue

        # Still on Login.aspx — check whether it's the code-verification step
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(page.content(), "html.parser")
        err_el = soup.find(id="lblLoginError") or soup.find(class_="alert")
        err_txt = err_el.get_text(strip=True) if err_el else ""

        if "enter the code" in err_txt.lower():
            print(f"  Verification code page detected, submitting pre-filled code...")
            # The code is already pre-filled; just click the submit button
            btns = page.locator(
                'input[type="submit"]:not(#btnFooterSubscribe), '
                'button[type="submit"]:not(#btnFooterSubscribe)'
            ).all()
            for btn in btns:
                if btn.is_visible():
                    btn.click()
                    break
            else:
                page.evaluate("document.querySelector('form').submit()")
            page.wait_for_load_state("networkidle", timeout=15000)
            print(f"  After code submit, URL: {page.url}")
        else:
            # Genuine login failure
            print(f"  LP error message: {err_txt or 'none found'}")
            page.screenshot(path=str(Path(__file__).parent.parent / "login_error.png"))
            print("  Screenshot saved: login_error.png")
            raise RuntimeError("Login failed")
    else:
        raise RuntimeError("Login: stuck in redirect loop after 5 attempts")

    print(f"  Logged in successfully.")


def open_report_generator(page):
    page.goto(f"{LP_URL}/ReportCtrl.html?BC=Reports|ReportGenerator", wait_until="networkidle")
    time.sleep(1)


def select_opt(page, label_text, value, optional=False):
    """Select a dropdown option by the label text next to it."""
    try:
        sel = page.locator(f'xpath=//label[contains(text(),"{label_text}")]/following::select[1]')
        if sel.count() == 0:
            sel = page.locator(f'select:near(:text("{label_text}"))').first
        sel.select_option(label=value)
    except Exception as e:
        if not optional:
            print(f"    Warning: could not set '{label_text}' to '{value}': {e}")


def set_date(page, field_id, date_str):
    """Set a calendar-picker date field via JS and trigger change events."""
    page.evaluate(f"""
        var el = document.querySelector('#{field_id}');
        if (el) {{
            el.value = '{date_str}';
            el.dispatchEvent(new Event('input',  {{bubbles:true}}));
            el.dispatchEvent(new Event('change', {{bubbles:true}}));
        }}
    """)


def download_one(page, context, rdef, period, start_dt, end_dt, out_dir):
    fname = f"{rdef['name']}.xlsx"
    out_path = out_dir / fname

    open_report_generator(page)

    # Group
    page.select_option('#ReportGroup', label=rdef["group"])
    time.sleep(2)

    # Report name (dropdown refreshes after group changes)
    page.select_option('#ReportName', label=rdef["report"])
    time.sleep(2)

    # Set dates via JS to bypass calendar picker
    # Find the two visible TextBox inputs (Start/End Date)
    textbox_ids = page.evaluate("""
        Array.from(document.querySelectorAll('input[id^="TextBox"]'))
             .filter(el => el.offsetParent !== null)
             .map(el => el.id)
    """)
    print(f"    Date field IDs: {textbox_ids}")
    if len(textbox_ids) >= 2:
        set_date(page, textbox_ids[0], fmt(start_dt))
        set_date(page, textbox_ids[1], fmt(end_dt))
    else:
        print(f"    WARNING: Expected 2 date fields, found {len(textbox_ids)}")

    # Optional filters — DropDown1=Source, DropDown2=Market
    # Set these BEFORE format so any AJAX reload doesn't reset the format.
    for dd_id, key in [('#DropDown1', 'source'), ('#DropDown2', 'market')]:
        if key not in rdef:
            continue
        try:
            opts = page.locator(f'{dd_id} option').all_text_contents()
            target = rdef[key]
            match = next((o for o in opts if o.strip().upper() == target.strip('-').upper() or o == target), None)
            if match:
                page.select_option(dd_id, label=match)
                time.sleep(1)
        except Exception:
            pass

    # Format = Excel Spreadsheet — set LAST so dropdown AJAX can't reset it
    fmt_opts = page.locator('#rFormat option').all_text_contents()
    excel_label = next((o for o in fmt_opts if 'excel' in o.lower()), None)
    if excel_label:
        page.select_option('#rFormat', label=excel_label)
        print(f"    Format set to: {excel_label}")
    else:
        print(f"    WARNING: No Excel option in format dropdown: {fmt_opts}")
    time.sleep(0.5)

    # Verify dates and format before clicking Go
    v1 = page.locator('#TextBox1').input_value()
    v2 = page.locator('#TextBox2').input_value()
    cur_fmt = page.locator('#rFormat').input_value()
    print(f"    Dates: {v1} to {v2}  |  Format value: {cur_fmt}")

    # Close any popup windows left over from a previous download so LP is
    # forced to open a fresh one (LP reuses a named window which prevents
    # Playwright from seeing the download event on subsequent reports).
    for p in context.pages:
        if p != page:
            try:
                p.close()
            except Exception:
                pass

    # Click Go and capture the download
    with page.expect_download(timeout=90000) as dl_info:
        page.locator('button:has-text("Go"), input[value="Go"]').first.click()
    dl_info.value.save_as(out_path)
    time.sleep(0.5)
    print(f"    Saved {fname}")


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    ranges = date_ranges()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=False,
            slow_mo=300,
            args=[
                "--disable-save-password-bubble",
                "--disable-features=PasswordManagerEnabled",
                "--no-default-browser-check",
                "--password-store=basic",
            ]
        )
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        login(page)

        for period, (start_dt, end_dt) in ranges.items():
            out_dir = REPORTS_DIR / period
            out_dir.mkdir(exist_ok=True)
            print(f"\n=== {period.upper()} ({fmt(start_dt)} - {fmt(end_dt)}) ===")
            for rdef in REPORTS:
                try:
                    download_one(page, context, rdef, period, start_dt, end_dt, out_dir)
                except Exception as e:
                    print(f"    ERROR {rdef['name']}: {e}")
                    # Reset page state so next report starts clean
                    try:
                        page.goto(f"{LP_URL}/Start.html", wait_until="networkidle", timeout=10000)
                    except Exception:
                        pass

        browser.close()
    print("\nAll downloads complete.")


if __name__ == "__main__":
    run()
