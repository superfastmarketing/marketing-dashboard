"""Print all input fields on the LP login page to find exact IDs."""
import os
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv(Path(__file__).parent.parent / ".env")
LP_URL = os.getenv("LP_URL").rstrip("/")

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto(f"{LP_URL}/Login.aspx", wait_until="networkidle")
    import time; time.sleep(2)

    print("=== LOGIN PAGE INPUTS ===")
    for inp in page.locator("input").all():
        id_   = inp.get_attribute("id") or ""
        name  = inp.get_attribute("name") or ""
        type_ = inp.get_attribute("type") or "text"
        ph    = inp.get_attribute("placeholder") or ""
        vis   = inp.is_visible()
        print(f"  id={id_:<25} name={name:<20} type={type_:<12} placeholder={ph:<20} visible={vis}")

    input("\nPress Enter to close...")
    browser.close()
