# Certificate & Email Automation App

Reads a participants Excel sheet, generates a personalized PDF certificate
for anyone with a recorded finish time, and emails everyone — certificate
attached for finishers, a plain message for non-finishers.

## How it works

```
participants.xlsx ──► excel_reader.py ──► for each row:
                                              │
                              Time > 0 ───────┼─────── Time == 0
                                 │                          │
                     certificate_generator.py         (skip certificate)
                     (fill .pptx → convert PDF)              │
                                 │                            │
                              email_sender.py (attach PDF)  email_sender.py
                                 │                            │
                                 └────────► SMTP send ◄───────┘
```

## 1. Setup

```bash
pip install -r requirements.txt
```

You also need **LibreOffice** installed on the machine that runs this
(used headlessly to convert the filled `.pptx` into a `.pdf`):
- macOS: `brew install --cask libreoffice`
- Ubuntu/Debian: `sudo apt install libreoffice`
- Windows: install LibreOffice, then set `SOFFICE_BIN` in `.env` to the
  full path of `soffice.exe`.

Copy the env template and fill in your SMTP details:

```bash
cp .env.example .env
```

## 2. Provide your input files

- **Excel** — place at `data/participants.xlsx` (or set `EXCEL_PATH`).
  Required columns, exact header names: `Name`, `Distance`, `Time`, `Date`, `email_id`.
  A `Time` of `0`, blank, or non-numeric text (e.g. "DNF") is treated as "no time recorded".

- **Certificate template** — place at `templates/certificate_template.pptx`
  (or set `PPTX_TEMPLATE_PATH`). Design it however you like in PowerPoint,
  and anywhere you want data inserted, type these exact placeholder tokens
  as text on the slide:

  ```
  {{Name}}   {{Distance}}   {{Time}}   {{Date}}
  ```

  Double curly braces are used deliberately so the app never accidentally
  matches a stray occurrence of the word "Time" elsewhere in your
  certificate's decorative text.

- **Email bodies** — `templates/email_with_time.txt` and
  `templates/email_without_time.txt` are provided as starting samples;
  edit the wording to taste. They support the same `{{Name}}`,
  `{{Distance}}`, `{{Time}}`, `{{Date}}` placeholders.

## 3. Test run (no emails actually sent)

`DRY_RUN=true` is the default in `.env.example`. With it on, the app still
generates all the PDFs and prints what it *would* send, so you can check
`output/pdfs/` and `output/run_log.csv` before going live.

```bash
python main.py
```

## 4. Go live

Set `DRY_RUN=false` in `.env`, then run again:

```bash
python main.py
```

## Notes & things you may want to adjust

- **Gmail users**: you'll need an [App Password](https://myaccount.google.com/apppasswords)
  rather than your normal password (Gmail blocks plain SMTP login otherwise).
- **Large participant lists**: SMTP servers often rate-limit outgoing mail;
  if you're sending hundreds of emails, consider adding a short delay
  between sends in `main.py`, or switching to a transactional email
  provider (SendGrid, SES, Mailgun) via their SMTP relay.
- **Re-runs**: the app doesn't currently track "already sent" state across
  runs — if you re-run it, it will resend everyone. If you need
  resume-safety, the `output/run_log.csv` file has everything needed to
  filter out already-successful rows before the next run.
- **Column names**: if your Excel uses different header text, update the
  `COL_*` constants in `config.py` rather than renaming your spreadsheet.
