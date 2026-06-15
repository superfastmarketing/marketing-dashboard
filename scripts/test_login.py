"""Quick test: log in to LeadPerfection and print page title + URL."""
import os
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv(Path(__file__).parent.parent / ".env")

LP_URL  = os.getenv("LP_URL", "https://p5wfa.leadperfection.com")
LP_USER = os.getenv("LP_USERNAME")
LP_PASS = os.getenv("LP_PASSWORD")

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False)  # visible so we can see what happens
    page = browser.new_page()

    print(f"Navigating to {LP_URL} ...")
    page.goto(LP_URL, wait_until="domcontentloaded")
    print(f"  Title: {page.title()}")
    print(f"  URL:   {page.url}")

    # Print all input fields on the page
    inputs = page.locator("input").all()
    print(f"\nFound {len(inputs)} input fields:")
    for inp in inputs:
        name = inp.get_attribute("name") or ""
        id_  = inp.get_attribute("id") or ""
        type_ = inp.get_attribute("type") or "text"
        print(f"  type={type_:<10} name={name:<25} id={id_}")

    # Try to fill login
    try:
        page.locator('input[type="text"]').first.fill(LP_USER)
        page.locator('input[type="password"]').first.fill(LP_PASS)
        print("\nFilled credentials. Submitting...")
        page.locator('input[type="submit"], button[type="submit"]').first.click()
        page.wait_for_load_state("networkidle", timeout=10000)
        print(f"  After login — Title: {page.title()}")
        print(f"  After login — URL:   {page.url}")
    except Exception as e:
        print(f"Login attempt error: {e}")

    input("\nPress Enter to close browser...")
    browser.close()
