# Modern Mailer

A PySide6 desktop application for sending daily report emails and weekly invoice emails from one UI. The current implementation is designed to run on both Windows and macOS.

## Requirements

- Python 3.9 or higher
- PySide6
- reportlab
- tzdata

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Configuration

The app reads optional values from `.env` in the app folder. If `.env` is missing, built-in defaults are used.

Example:

```env
YOUR_NAME=Your Full Name
SENDER_EMAIL=your.email@example.com
SENDER_PASSWORD="your_app_password_here"

TO_EMAIL=your-recipient@example.com
CC_EMAIL=your-cc@example.com
INVOICE_TO_EMAILS=finance@example.com,manager@example.com
INVOICE_CC_EMAILS=

SMTP_SERVERS=smtp.gmail.com,smtp.mail.yahoo.com,smtp-mail.outlook.com
SMTP_PORT=587

INVOICE_HOURLY_RATE=7.50
INVOICE_PAYMENT_METHOD=Payoneer
INVOICE_PAYMENT_EMAIL=payoneer@example.com
INVOICE_PAYMENT_TERMS=Please make payment within 30 days of the invoice date.

INVOICE_ADDRESS=123 Main St
INVOICE_CITY_STATE_ZIP=Portland, OR 97204
INVOICE_PHONE=503-555-0100
```

Notes:

- If a password contains `#` or other special characters, keep it quoted.
- `SMTP_SERVERS` is a comma-separated list used to populate the SMTP dropdown.
- Form edits are also persisted to `config.json` for future launches.

## Running the App

### macOS

Use the bundled launcher so the local virtual environment is used:

```bash
./run_mac.sh
```

### Windows

Double-click `run.bat`, or run:

```bat
run.bat
```

The batch file prefers `.venv\Scripts\python.exe`, then falls back to `py -3`, then `python`.

### Direct Python launch

```bash
python main.py
```

## UI Notes

The app has two modes:

- `Daily Report` for the end-of-day email workflow
- `Invoice Sender` for generating a PDF invoice and emailing it as an attachment

The invoice date fields use a Qt calendar dialog, which works on both Windows and macOS. Daily hours are generated automatically from the selected week range and summed into `Hours Worked`.

Invoice sender includes two PDF templates:

- `Standard Invoice` for the current single-line invoice layout
- `Client Tracking Invoice` for a timesheet-style layout that prints each daily time entry with hours and amount

## Building the App

### Build only

```bash
python -m PyInstaller --noconfirm --windowed --name "PITC-VA-Mailer" --hidden-import tzdata --collect-all PySide6 --collect-all shiboken6 main.py
```

On Windows this produces a `.exe` in `dist/`.

On macOS this produces a `.app` bundle in `dist/`.

### Build and package for sharing

```bash
bash ./copy_dist_to_pitc_mailer.sh
```

This script:

- builds a platform-appropriate app bundle
- creates `../PITC-mailer`
- copies the generated `.exe` on Windows or `.app` bundle on macOS

## Release Checklist

1. Run `python -m py_compile main.py`.
2. Run `bash ./copy_dist_to_pitc_mailer.sh`.
3. Verify the packaged app opens on the target platform.
4. Verify the SMTP dropdown, Daily Report mode, Invoice Sender mode, and date picker all work.
5. Confirm `.env` values are loaded correctly.

## Gmail App Password

Google blocks direct username/password sign-ins for many accounts. Use an App Password:

1. Open your Google Account security settings.
2. Enable 2-Step Verification.
3. Open `App Passwords`.
4. Generate an app password for Mail.
5. Use that value in the app password field or `.env`.

## Defaults

Daily report defaults:

- To: `sid@learncodinganywhere.com`
- CC: `jack@learncodinganywhere.com`
- Subject: `{Your Name} Daily Report of {Date}`

Invoice defaults:

- To: `finances@learncodinganywhere.com`, `sid@learncodinganywhere.com`
- Payment Method: `Payoneer`
- Hourly Rate: `$7.50`
- Payment Terms: `Please make payment within 30 days of the invoice date.`
- Template: `Standard Invoice`

## Report Questions

1. What course are you on?
2. Which step of the course are you on?
3. What non-course related tasks did you complete today?
4. Do you have any slows or issues that you need assistance with?
5. Do you have any questions?
6. Is there anything else you would like to communicate?

## Tips

- Use `Preview Email` and `Preview Invoice` before sending.
- The time in the header updates automatically.
- The invoice flow writes the PDF first, then attaches it to the email.
