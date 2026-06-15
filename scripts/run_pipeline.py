"""
run_pipeline.py — Full daily automation pipeline.

Steps:
  1. Download all 20 LP reports (5 reports × 4 date periods)
  2. Build Marketing_Dashboard.html from the downloaded data
  3. Upload to Google Drive
  4. Email to distribution list

Run manually:
  python scripts/run_pipeline.py

Scheduled automatically by Windows Task Scheduler every day at 8:00 AM.
Log output is written to logs/pipeline_YYYY-MM-DD.log in the project root.
"""

import sys
import traceback
from datetime import date
from pathlib import Path

ROOT = Path(__file__).parent.parent
LOG_DIR = ROOT / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / f"pipeline_{date.today().isoformat()}.log"


class Tee:
    """Write to both stdout and a log file simultaneously."""
    def __init__(self, path):
        self.file = open(path, "w", encoding="utf-8")
        self.stdout = sys.stdout

    def write(self, data):
        self.stdout.write(data)
        self.file.write(data)
        self.file.flush()

    def flush(self):
        self.stdout.flush()
        self.file.flush()

    def close(self):
        self.file.close()


def run_step(name, fn):
    print(f"\n{'='*60}")
    print(f"STEP: {name}")
    print('='*60)
    try:
        fn()
        print(f"[OK] {name}")
        return True
    except Exception:
        print(f"[FAILED] {name}")
        traceback.print_exc()
        return False


def main():
    tee = Tee(LOG_FILE)
    sys.stdout = tee

    try:
        print(f"Pipeline started: {date.today().isoformat()}")

        # Step 1 — Download reports from LeadPerfection
        def download():
            sys.path.insert(0, str(Path(__file__).parent))
            import download_reports
            download_reports.run()

        ok1 = run_step("Download LP reports", download)
        if not ok1:
            print("\nAborting: download failed. Check login or LP availability.")
            sys.exit(1)

        # Step 2 — Build dashboard HTML
        def build():
            import build_dashboard
            build_dashboard.run()

        ok2 = run_step("Build dashboard HTML", build)
        if not ok2:
            print("\nAborting: dashboard build failed.")
            sys.exit(1)

        # Step 3 — Upload to Google Drive
        def drive():
            import upload_to_drive
            upload_to_drive.upload()

        run_step("Upload to Google Drive", drive)  # non-fatal if Drive fails

        # Step 4 — Send email
        def email():
            import send_email
            send_email.send()

        run_step("Send email", email)  # non-fatal if email fails

        print(f"\n{'='*60}")
        print(f"Pipeline complete: {date.today().isoformat()}")
        print(f"Log saved to: {LOG_FILE}")

    finally:
        sys.stdout = tee.stdout
        tee.close()


if __name__ == "__main__":
    main()
