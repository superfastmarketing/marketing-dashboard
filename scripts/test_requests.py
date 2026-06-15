"""
Map LeadPerfection login form fields + attempt login via requests.
"""
import os, re
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")

LP_URL  = os.getenv("LP_URL", "https://p5wfa.leadperfection.com")
LP_USER = os.getenv("LP_USERNAME")
LP_PASS = os.getenv("LP_PASSWORD")

s = requests.Session()
s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/125.0"

login_url = LP_URL.rstrip("/") + "/Login.aspx"
resp = s.get(login_url, timeout=15)
soup = BeautifulSoup(resp.text, "html.parser")

# Build payload from hidden fields only, then add credentials manually
payload = {}
for inp in soup.find_all("input", {"type": "hidden"}):
    name = inp.get("name")
    if name:
        payload[name] = inp.get("value", "")

# Add login credentials and submit button
payload["username"] = LP_USER
payload["password"] = LP_PASS
payload["btnLogin"]  = " Log In"

print(f"Posting to: {login_url}")
print(f"Payload keys: {list(payload.keys())}")

resp2 = s.post(login_url, data=payload, timeout=15, allow_redirects=True)
print(f"Status: {resp2.status_code}  Final URL: {resp2.url}")

soup2 = BeautifulSoup(resp2.text, "html.parser")
title = soup2.title.string if soup2.title else "N/A"
print(f"Page title: {title}")

# Check for error message
err = soup2.find(id="lblLoginError")
if err:
    print(f"Login error: {err.get_text(strip=True)}")

if resp2.url != login_url:
    print("LOGIN SUCCEEDED (redirected away from login page)")
elif "dashboard" in resp2.text.lower() or "logout" in resp2.text.lower():
    print("LOGIN SUCCEEDED (dashboard content found)")
else:
    print("LOGIN FAILED")
