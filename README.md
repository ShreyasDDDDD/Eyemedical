# Eyemedical Reminder System

This project converts a weekly eye-medication worksheet into normalized reminder events, stores them, and sends reminders 5 minutes before each dose.

## What it does

- Reads an `.xlsx` worksheet grid.
- Detects:
  - weekday headers (`Monday` ... `Sunday`)
  - left/right/both eye columns
  - time rows
  - medicine names in cells
- Normalizes the sheet into row-by-row reminder records.
- Stores reminders in SQLite with weekly recurrence semantics.
- Every minute (configurable), checks if a medicine is due in 5 minutes.
- Sends reminder messages (currently printed to stdout, easy to replace with email/SMS/Telegram).
- Detects worksheet changes via file hash and auto-refreshes the reminder list.
- Prevents duplicates so the same reminder is not sent twice for the same due date.

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

### Parse/update reminders once

```bash
python reminder_system.py /path/to/worksheet.xlsx --db reminders.db --run-once
```

### Run live reminder loop (every 60s)

```bash
python reminder_system.py /path/to/worksheet.xlsx --db reminders.db --tick 60
```

Optional sheet selection:

```bash
python reminder_system.py /path/to/worksheet.xlsx --sheet "Updated April 26"
```

## Reminder message format

Example output:

```text
Reminder: In 5 minutes, take Dorzolamide in your right eye at 6:00 AM.
```

## Notes

- The loop uses UTC by default (`datetime.utcnow()`). If you want local timezone reminders, adjust `now` handling in `send_due_reminders`.
- To send via external channels, replace `print_sender` with integrations (SMTP, Twilio, Telegram Bot API, WhatsApp API, etc.).
