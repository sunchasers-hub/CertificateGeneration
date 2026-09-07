"""
main.py
Orchestrates the full pipeline:

  1. Read participants.xlsx
  2. For each row:
       - Time > 0  -> fill certificate template, render PDF, send email
                       (with PDF attached) using email_with_time.txt
       - Time == 0 -> send email (no attachment) using email_without_time.txt
  3. Write a run log (output/run_log.csv) recording success/failure per row.

Run with:  python main.py
Set DRY_RUN=false in your .env once you've verified everything looks right.
"""

import csv
import sys
from datetime import datetime

import config
from excel_reader import read_participants, ExcelValidationError
from certificate_generator import generate_certificate_pdf
from email_sender import build_email, send_email


def run():
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        participants = list(read_participants())
    except (FileNotFoundError, ExcelValidationError) as e:
        print(f"ERROR reading Excel file: {e}")
        sys.exit(1)

    if not participants:
        print("No valid rows found in the Excel file. Nothing to do.")
        return

    log_rows = []
    sent, failed, skipped_zero_time_cert = 0, 0, 0

    for p in participants:
        status = "OK"
        detail = ""
        pdf_path = None
        try:
            if p.has_time:
                pdf_path = generate_certificate_pdf(p)
            msg = build_email(p, pdf_path)
            send_email(msg)
            sent += 1
        except Exception as e:
            status = "FAILED"
            detail = str(e)
            failed += 1
            print(f"[Row {p.row_number}] FAILED for {p.name} <{p.email}>: {e}")

        log_rows.append({
            "row": p.row_number,
            "name": p.name,
            "email": p.email,
            "time": p.time_display,
            "certificate_generated": bool(pdf_path),
            "status": status,
            "detail": detail,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        })

    with config.LOG_CSV.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(log_rows[0].keys()))
        writer.writeheader()
        writer.writerows(log_rows)

    print(f"\nDone. {sent} email(s) processed OK, {failed} failed.")
    print(f"Log written to: {config.LOG_CSV}")
    if config.DRY_RUN:
        print("NOTE: DRY_RUN is enabled — no emails were actually sent. "
              "Set DRY_RUN=false in .env to send for real.")


if __name__ == "__main__":
    run()
