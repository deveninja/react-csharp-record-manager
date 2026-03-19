# Training Mailer

A Python desktop application for sending your End-of-Day report email and your weekly training invoice from the same desktop app.

---

## Requirements

- **Python 3.9 or higher** — download from [python.org](https://www.python.org/downloads/)
- `tkinter` — bundled with Python on Windows and macOS (no extra install needed)
- `reportlab` — used to generate the invoice PDF attachment
- `tkcalendar` — used for the invoice date dropdown calendars

---

## Setup

### 1 — Clone / copy the folder

Place the `eod-desktop-app` folder anywhere you like.

### 2 — Install dependencies

```bash
pip install -r requirements.txt
```

### 3 — (Optional) Configure via .env

Copy `.env-sample` to `.env`, then edit the values you want to customize.

Example:

```env
# User Information
YOUR_NAME=Your Full Name
SENDER_EMAIL=your.email@example.com
SENDER_PASSWORD="your_app_password_here"

# Email recipients
TO_EMAIL=your-recipient@example.com
CC_EMAIL=your-cc@example.com
INVOICE_TO_EMAILS=finance@example.com,manager@example.com
INVOICE_CC_EMAILS=

# SMTP Server settings
SMTP_SERVERS=smtp.gmail.com,smtp.mail.yahoo.com,smtp-mail.outlook.com
SMTP_PORT=587

# Invoice defaults
INVOICE_HOURLY_RATE=7.50
INVOICE_PAYMENT_METHOD=Payoneer
INVOICE_PAYMENT_EMAIL=payoneer@example.com
INVOICE_PAYMENT_TERMS=Please make payment within 30 days of the invoice date.

# Invoice Address Defaults
INVOICE_ADDRESS=123 Main St
INVOICE_CITY_STATE_ZIP=Portland, OR 97204
INVOICE_PHONE=503-555-0100
```

Notes:

- If your password contains `#` or other special characters, keep it wrapped in quotes.
- `SMTP_SERVERS` is a comma-separated list used to populate the SMTP server dropdown in the app.
- The SMTP port defaults to `587` but remains editable in the UI.

If `.env` is not present or is incomplete, the app falls back to built-in defaults.

---

## Running the App

### Run from Python

```bash
python main.py
```

### Run the packaged app

After building:

1. Open the packaged output folder.
2. Copy `.env-sample` and rename the copy to `.env`.
3. Edit `.env` with your own values.
4. Run the exe.

The app opens with two modes:

- `Daily Report` for the end-of-day email workflow
- `Invoice Sender` for generating a PDF invoice and emailing it as an attachment

---

## First-time configuration

The app can prefill these fields from `.env` if values are provided:

- `Your Name`
- `Sender Email`
- `Password`
- `SMTP Server`
- `SMTP Port`
- `Payment Method`
- `Payment Email`
- `Address`
- `City/State/Zip`
- `Phone Number`

Values you edit in the form are also stored in `config.json` for future launches.

---

## Build the App

### Build only

```bash
python -m PyInstaller --onefile --windowed --name "PITC-VA-Mailer" --hidden-import tzdata main.py
```

This creates the exe in `dist/`.

### Build and package for sharing

```bash
bash ./copy_dist_to_pitc_mailer.sh
```

This script:

- builds `PITC-VA-Mailer.exe`
- creates `PITC-mailer` in the parent `react-csharp-record-manager` folder
- copies only:
  - `PITC-VA-Mailer.exe`
  - `.env-sample`

Final packaged output:

- `../PITC-mailer/PITC-VA-Mailer.exe`
- `../PITC-mailer/.env-sample`

After packaging, copy `../PITC-mailer/.env-sample` to `../PITC-mailer/.env` and edit the `.env` file before sharing or running the app.

### Release checklist

Before sharing the app:

1. Update `.env-sample` with the latest placeholders and SMTP server list.
2. Run `python -m py_compile main.py` to confirm the app has no syntax errors.
3. Run `bash ./copy_dist_to_pitc_mailer.sh` to build and package the release.
4. Open `../PITC-mailer/PITC-VA-Mailer.exe` and verify:

- the app opens successfully
- the SMTP dropdown shows the expected server list
- Daily Report mode loads correctly
- Invoice Sender mode loads correctly

5. Confirm `../PITC-mailer` contains only:

- `PITC-VA-Mailer.exe`
- `.env-sample`

6. Copy `.env-sample` to `.env`, update the values, and verify the app starts with the new `.env` file.
7. Share the `PITC-mailer` folder.

When the app opens, check the **Settings** section at the top:

| Field        | What to enter                                                  |
| ------------ | -------------------------------------------------------------- |
| Your Name    | Your full name (used in the subject line)                      |
| Sender Email | The email address you will send **from**                       |
| Password     | Your email password or App Password (see Gmail note below)     |
| SMTP Server  | Choose from the dropdown populated by `SMTP_SERVERS` in `.env` |
| SMTP Port    | Usually `587` for TLS (default)                                |

---

## Gmail — App Password (required)

Google blocks "less secure app" sign-ins. You must create an **App Password**:

1. Go to your Google Account → **Security**
2. Make sure **2-Step Verification** is turned on
3. Under 2-Step Verification, scroll to **App Passwords**
4. Create a new App Password (select app: _Mail_, device: _Windows Computer_)
5. Copy the 16-character password and paste it into the **Password** field in the app

---

## Email details

|               |                                        |
| ------------- | -------------------------------------- |
| **To**        | sid@learncodinganywhere.com            |
| **CC**        | jack@learncodinganywhere.com           |
| **Subject**   | `{Your Name} Daily Report of {Date}`   |
| **Date/Time** | Always shown in Eastern Time (EST/EDT) |

### Invoice sender defaults

The invoice mode reads the included `Developer Training Invoice.docx` instructions and uses these defaults unless overridden in `.env`:

- `To`: `finances@learncodinganywhere.com`, `sid@learncodinganywhere.com`
- `Payment Method`: `Payoneer`
- `Payment Email`: Your sender email (or set via `INVOICE_PAYMENT_EMAIL` in `.env`)
- `Hourly Rate`: `$7.50`
- `Payment Terms`: `Please make payment within 30 days of the invoice date.`

The invoice mode can:

- Generate the `Notes` line automatically from the date range, total hours, hourly rate, and total due
- Append anything you type into the notes box as extra context
- Generate a PDF invoice in the required training format
- Save the PDF anywhere you choose
- Email the PDF as an attachment from the same app

#### Configuring Different Payment Methods or Emails

Each invoice can use a different payment method or payment email. Edit these in the Invoice Sender tab, or set defaults in `.env`:

```env
INVOICE_PAYMENT_METHOD=Payoneer  # Or: Bank Transfer, Wise, etc.
INVOICE_PAYMENT_EMAIL=payoneer@example.com  # Or your payment account email
```

If `INVOICE_PAYMENT_EMAIL` is empty in `.env`, the app defaults to your sender email for each invoice.

The auto-generated note follows this format:

```text
Total number of hours worked for 2 weeks is 80 at $7.50/hour, which comes up to $600.00.
```

---

## Report questions

1. What course are you on?
2. Which step of the course are you on?
3. What non-course related tasks did you complete today?
4. Do you have any slows or issues that you need assistance with?
5. Do you have any questions?
6. Is there anything else you would like to communicate?

---

## Tips

- Click **Preview Email** to review your message before sending.
- Click **Clear Answers** to reset the answer fields.
- The date/time in the header refreshes automatically every minute.
- The invoice send flow creates the PDF first, then attaches that saved PDF to the email.
