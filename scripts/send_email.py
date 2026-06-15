"""
send_email.py — Email a Google Drive link to the marketing dashboard.

Sends a link rather than an attachment to avoid antivirus flags on HTML files.
The dashboard is uploaded to Drive first (upload_to_drive.py), then recipients
click the link to open it directly in their browser.

Required .env keys:
  SMTP_FROM          — sending address (must match the Gmail account)
  SMTP_APP_PASSWORD  — 16-char App Password from myaccount.google.com/apppasswords
  SMTP_TO            — comma-separated recipient list
  GDRIVE_FILE_ID     — Google Drive file ID (set during Drive upload setup)
"""

import os
import smtplib
from datetime import date
from email.message import EmailMessage
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).parent.parent
load_dotenv(ROOT / ".env")

SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_FROM     = os.getenv("SMTP_FROM", "")
SMTP_PASS     = os.getenv("SMTP_APP_PASSWORD", "")
SMTP_TO       = [a.strip() for a in os.getenv("SMTP_TO", "").split(",") if a.strip()]
GDRIVE_FILE_ID = os.getenv("GDRIVE_FILE_ID", "")


def _drive_link():
    if GDRIVE_FILE_ID:
        return f"https://drive.google.com/file/d/{GDRIVE_FILE_ID}/view"
    return None


def _build_body(link):
    today = date.today().strftime("%B %d, %Y")

    plain = (
        f"SuperFast Kitchen & Bath — Marketing Dashboard ({today})\n\n"
        "Your marketing dashboard has been updated with today's data from LeadPerfection.\n\n"
        f"View dashboard: {link}\n\n"
        "The dashboard opens directly in your browser — no download required.\n"
        "Use Chrome or Edge for the best experience.\n\n"
        "Data sources: Marketing Sub-Source Report (raw leads) + "
        "Call Center Appointment Statistics (set / demo / sales)."
    )

    html = f"""
<html><body style="font-family:Arial,sans-serif;color:#222;max-width:600px;margin:0 auto;padding:0;">
<div style="background:linear-gradient(135deg,#1a2e4a,#2d5a8e);color:white;padding:28px 32px;border-radius:8px 8px 0 0;">
  <h1 style="margin:0;font-size:1.35rem;font-weight:700;">SuperFast Kitchen &amp; Bath</h1>
  <p style="margin:6px 0 0;opacity:0.75;font-size:0.88rem;">Marketing Dashboard &nbsp;·&nbsp; {today}</p>
</div>
<div style="background:#f9fafb;padding:28px 32px;border:1px solid #e5e7eb;border-top:none;">
  <p style="margin:0 0 20px;font-size:0.95rem;">
    Your marketing dashboard has been updated with today's data from LeadPerfection.
  </p>
  <a href="{link}"
     style="display:inline-block;background:#2d5a8e;color:white;text-decoration:none;
            padding:13px 28px;border-radius:6px;font-weight:700;font-size:0.95rem;
            letter-spacing:0.3px;">
    View Dashboard &rarr;
  </a>
  <p style="margin:20px 0 0;font-size:0.82rem;color:#6b7280;">
    Opens directly in your browser &mdash; no download required.<br>
    Use Chrome or Edge for the best experience.
  </p>
</div>
<div style="background:#f3f4f6;padding:16px 32px;border:1px solid #e5e7eb;border-top:none;
            border-radius:0 0 8px 8px;">
  <p style="margin:0;font-size:0.75rem;color:#9ca3af;">
    Data: Marketing Sub-Source Report (raw leads) &nbsp;+&nbsp;
    Call Center Appointment Statistics (set / demo / sales).<br>
    Generated automatically by the LP automation pipeline.
  </p>
</div>
</body></html>
"""
    return plain, html


def send():
    if not SMTP_FROM or not SMTP_PASS:
        print("  SMTP_FROM or SMTP_APP_PASSWORD not set in .env — skipping email.")
        return
    if not SMTP_TO:
        print("  SMTP_TO not set in .env — skipping email.")
        return

    link = _drive_link()
    if not link:
        print("  GDRIVE_FILE_ID not set in .env — skipping email.")
        return

    today = date.today().strftime("%B %d, %Y")
    plain, html_body = _build_body(link)

    msg = EmailMessage()
    msg["Subject"] = f"Marketing Dashboard — {today}"
    msg["From"]    = SMTP_FROM
    msg["To"]      = ", ".join(SMTP_TO)
    msg.set_content(plain)
    msg.add_alternative(html_body, subtype="html")

    print(f"Sending dashboard email to {len(SMTP_TO)} recipients...")
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_FROM, SMTP_PASS)
        server.send_message(msg)
    print(f"  Email sent to: {', '.join(SMTP_TO)}")
    print(f"  Dashboard link: {link}")


if __name__ == "__main__":
    send()
