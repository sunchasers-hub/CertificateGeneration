"""
email_sender.py
Loads the two email body templates, fills placeholders, and sends the
message (with or without a certificate PDF attached) over SMTP.
"""

import smtplib
from email.message import EmailMessage
from pathlib import Path

import config
from excel_reader import Participant


def _load_and_fill_template(template_path: Path, participant: Participant) -> str:
    if not template_path.exists():
        raise FileNotFoundError(f"Email template not found: {template_path}")

    text = template_path.read_text(encoding="utf-8")
    replacements = {
        "{{Name}}": participant.name,
        "{{Distance}}": participant.distance,
        "{{Time}}": participant.time_display,
        "{{Date}}": participant.date_display,
    }
    for token, value in replacements.items():
        text = text.replace(token, value)
    return text


def build_email(participant: Participant, pdf_path: Path = None) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = f"{config.SMTP_FROM_NAME} <{config.SMTP_USERNAME}>"
    msg["To"] = participant.email

    if participant.has_time:
        msg["Subject"] = config.EMAIL_SUBJECT_WITH_TIME
        body = _load_and_fill_template(config.EMAIL_WITH_TIME_TXT, participant)
    else:
        msg["Subject"] = config.EMAIL_SUBJECT_WITHOUT_TIME
        body = _load_and_fill_template(config.EMAIL_WITHOUT_TIME_TXT, participant)

    msg.set_content(body)

    if participant.has_time and pdf_path is not None:
        msg.add_attachment(
            pdf_path.read_bytes(),
            maintype="application",
            subtype="pdf",
            filename=f"{participant.name}_certificate.pdf",
        )

    return msg


def send_email(msg: EmailMessage):
    if config.DRY_RUN:
        print(f"[DRY RUN] Would send email to {msg['To']} — subject: {msg['Subject']}")
        return

    if not config.SMTP_USERNAME or not config.SMTP_PASSWORD:
        raise RuntimeError("SMTP_USERNAME / SMTP_PASSWORD are not set. Check your .env file.")

    with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT) as server:
        server.starttls()
        server.login(config.SMTP_USERNAME, config.SMTP_PASSWORD)
        server.send_message(msg)
