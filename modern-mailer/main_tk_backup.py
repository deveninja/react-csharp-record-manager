from __future__ import annotations

import importlib
import json
import os
import re
import smtplib
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from tkinter import filedialog, messagebox, scrolledtext, ttk

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    from tkcalendar import Calendar, DateEntry
except Exception:
    Calendar = None
    DateEntry = None

try:
    from zoneinfo import ZoneInfo

    EASTERN_TZ = ZoneInfo("America/New_York")
except Exception:
    try:
        ZoneInfo = importlib.import_module("backports.zoneinfo").ZoneInfo
        EASTERN_TZ = ZoneInfo("America/New_York")
    except Exception:
        EASTERN_TZ = None


# Determine the app directory (handles both PyInstaller exe and normal Python execution)
if getattr(sys, "frozen", False):
    # Running as PyInstaller exe
    APP_DIR = os.path.dirname(sys.executable)
else:
    # Running as normal Python script
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "config.json")
ENV_FILE = os.path.join(APP_DIR, ".env")
GENERATED_DIR = os.path.join(APP_DIR, "generated-invoices")


def load_env_file(path: str):
    if not os.path.exists(path):
        return

    try:
        with open(path, "r", encoding="utf-8") as handle:
            for raw_line in handle:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except Exception:
        pass


# Load .env file from APP_DIR, then try current working directory as fallback
load_env_file(ENV_FILE)
# If not found in APP_DIR, try current working directory
if not os.path.exists(ENV_FILE):
    load_env_file(os.path.join(os.getcwd(), ".env"))


DAILY_TO_EMAIL = os.getenv("TO_EMAIL", "sid@learncodinganywhere.com")
DAILY_CC_EMAIL = os.getenv("CC_EMAIL", "jack@learncodinganywhere.com")
INVOICE_TO_EMAILS = os.getenv(
    "INVOICE_TO_EMAILS",
    "finances@learncodinganywhere.com,sid@learncodinganywhere.com",
)
INVOICE_CC_EMAILS = os.getenv("INVOICE_CC_EMAILS", "sid@learncodinganywhere.com")

# Parse SMTP servers from .env (comma-separated list)
SMTP_SERVERS_STR = os.getenv("SMTP_SERVERS", "smtp.gmail.com,smtp.mail.yahoo.com,smtp-mail.outlook.com")
SMTP_SERVERS_LIST = [s.strip() for s in SMTP_SERVERS_STR.split(",") if s.strip()]
DEFAULT_SMTP_SERVER = SMTP_SERVERS_LIST[0] if SMTP_SERVERS_LIST else "smtp.gmail.com"
DEFAULT_SMTP_PORT = os.getenv("SMTP_PORT", "587")
DEFAULT_HOURLY_RATE = os.getenv("INVOICE_HOURLY_RATE", "7.50")
DEFAULT_PAYMENT_METHOD = os.getenv("INVOICE_PAYMENT_METHOD", "Payoneer")
DEFAULT_PAYMENT_EMAIL = os.getenv("INVOICE_PAYMENT_EMAIL", "")
DEFAULT_PAYMENT_TERMS = os.getenv(
    "INVOICE_PAYMENT_TERMS",
    "Please make payment within 30 days of the invoice date.",
)

# User Information Defaults
DEFAULT_NAME = os.getenv("YOUR_NAME", "")
DEFAULT_SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
DEFAULT_SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "")

# Invoice Address Defaults
DEFAULT_ADDRESS = os.getenv("INVOICE_ADDRESS", "")
DEFAULT_CITY_STATE_ZIP = os.getenv("INVOICE_CITY_STATE_ZIP", "")
DEFAULT_PHONE = os.getenv("INVOICE_PHONE", "")

COMPANY_LINES = [
    "The Tech Academy",
    "310 SW 4th Ave Suite 200",
    "Portland, Oregon 97204",
]

QUESTIONS = [
    {
        "key": "q1",
        "label": "1. What course are you on?",
        "hint": "Example: Basics of English Course",
    },
    {
        "key": "q2",
        "label": "2. Which step of the course are you on?",
        "hint": "Example: Step 92 at 52% training progress.",
    },
    {
        "key": "q3",
        "label": "3. What non-course related tasks did you complete today?",
        "hint": "Example: None. or Called students in the relinquished pipeline for one hour.",
    },
    {
        "key": "q4",
        "label": "4. Do you have any slows or issues that you need assistance with?",
        "hint": "Example: None. or The video on step 15 didn't play.",
    },
    {
        "key": "q5",
        "label": "5. Do you have any questions?",
        "hint": "Example: I'm all good for now. or What credentials should I use to access my training courses?",
    },
    {
        "key": "q6",
        "label": "6. Is there anything else you would like to communicate?",
        "hint": "Example: I really enjoyed learning about the company today and I look forward to working with you!",
    },
]


def _now_eastern() -> datetime:
    if EASTERN_TZ:
        return datetime.now(EASTERN_TZ)
    return datetime.now()


def get_report_date() -> str:
    return _now_eastern().strftime("%B %d, %Y")


def get_display_datetime() -> str:
    now = _now_eastern()
    tz_label = now.strftime("%Z") if EASTERN_TZ else "Local Time"
    return now.strftime(f"%A, %B %d, %Y  %I:%M %p  ({tz_label})")


def get_invoice_date() -> str:
    return _now_eastern().strftime("%m/%d/%Y")


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except Exception:
            pass
    return {}


def save_config(data: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
    except Exception:
        pass


def parse_email_list(value: str) -> list[str]:
    return [item.strip() for item in re.split(r"[;,]", value or "") if item.strip()]


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    cleaned = cleaned.strip(".-")
    return cleaned or "invoice"


def currency(value: float) -> str:
    return f"${value:,.2f}"


def htmlize_multiline(value: str) -> str:
    return escape(value).replace("\n", "<br>")


def _format_quantity(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _parse_user_date(value: str) -> datetime | None:
    raw = (value or "").strip()
    if not raw:
        return None

    formats = ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y"]
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def _date_value(value: str) -> date:
    parsed = _parse_user_date(value)
    if parsed:
        return parsed.date()
    return _now_eastern().date()


def build_invoice_notes(week_start: str, week_end: str, hours: float, rate: float, total_due: float, extra_notes: str = "") -> str:
    start_date = _parse_user_date(week_start)
    end_date = _parse_user_date(week_end)

    if start_date and end_date and end_date >= start_date:
        inclusive_days = (end_date - start_date).days + 1
        if inclusive_days % 7 == 0:
            week_count = inclusive_days // 7
            period_label = f"{week_count} week" if week_count == 1 else f"{week_count} weeks"
        else:
            period_label = f"the period from {week_start} to {week_end}"
    else:
        period_label = f"the period from {week_start} to {week_end}"

    base_note = (
        f"Total number of hours worked for {period_label} is {_format_quantity(hours)} "
        f"at {currency(rate)}/hour, which comes up to {currency(total_due)}."
    )

    extra_notes = extra_notes.strip()
    if extra_notes:
        return f"{base_note}\n\nAdditional Notes: {extra_notes}"
    return base_note


class MailerApp:
    CLR_BG = "#f3f4f6"
    CLR_CARD = "#ffffff"
    CLR_HEADER = "#16324f"
    CLR_ACCENT = "#0f766e"
    CLR_LABEL = "#0f172a"
    CLR_HINT = "#64748b"
    CLR_MUTED = "#475569"
    CLR_WARN = "#92400e"
    CLR_WARN_BG = "#fef3c7"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PITC Mailer")
        self.root.geometry("860x940")
        self.root.minsize(700, 760)
        self.root.configure(bg=self.CLR_BG)

        self.config = load_config()
        self.answer_widgets: dict[str, scrolledtext.ScrolledText] = {}
        self._daily_hours_vars: dict[str, tk.StringVar] = {}
        self._saved_daily_hours: dict[str, str] = {}
        self._recomputing_hours = False
        self._show_password = False
        self._mode_var = tk.StringVar(value="daily")

        self._build_ui()
        self._load_saved_config()
        self._update_mode_view()
        self._schedule_date_refresh()
        self._update_invoice_total()

    def _build_ui(self):
        self._build_header()

        outer = tk.Frame(self.root, bg=self.CLR_BG)
        outer.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(outer, bg=self.CLR_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._inner = tk.Frame(self._canvas, bg=self.CLR_BG)
        self._frame_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind(
            "<Configure>",
            lambda _event: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.bind(
            "<Configure>",
            lambda event: self._canvas.itemconfig(self._frame_id, width=event.width),
        )
        self._canvas.configure(yscrollcommand=vscroll.set)
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind_all("<Button-4>", self._on_mousewheel)
        self._canvas.bind_all("<Button-5>", self._on_mousewheel)

        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_settings(self._inner)
        self._build_mode_selector(self._inner)
        self._daily_frame = tk.Frame(self._inner, bg=self.CLR_BG)
        self._invoice_frame = tk.Frame(self._inner, bg=self.CLR_BG)
        self._build_daily_mode(self._daily_frame)
        self._build_invoice_mode(self._invoice_frame)
        tk.Frame(self._inner, bg=self.CLR_BG, height=18).pack()

        self._status_var = tk.StringVar(value="Ready. Use Daily Report or Invoice Sender from the mode switch above.")
        tk.Label(
            self.root,
            textvariable=self._status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#e2e8f0",
            fg=self.CLR_MUTED,
            font=("Helvetica", 9),
            padx=6,
        ).pack(side=tk.BOTTOM, fill=tk.X)

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=self.CLR_HEADER, pady=14)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr,
            text="PITC Mailer",
            font=("Helvetica", 18, "bold"),
            bg=self.CLR_HEADER,
            fg="white",
        ).pack()
        self._header_subtitle = tk.Label(
            hdr,
            text="",
            font=("Helvetica", 9),
            bg=self.CLR_HEADER,
            fg="#cbd5e1",
        )
        self._header_subtitle.pack()

    def _card(self, parent, padx=16, pady=6) -> tk.Frame:
        wrapper = tk.Frame(parent, bg=self.CLR_BG)
        wrapper.pack(fill=tk.X, padx=padx, pady=pady)
        card = tk.Frame(wrapper, bg=self.CLR_CARD, bd=1, relief=tk.GROOVE)
        card.pack(fill=tk.X)
        return card

    def _section_title(self, parent, text: str):
        tk.Label(
            parent,
            text=text,
            font=("Helvetica", 11, "bold"),
            bg=self.CLR_BG,
            fg=self.CLR_HEADER,
        ).pack(anchor=tk.W, padx=16, pady=(10, 2))

    def _labeled_entry(self, parent, label_text: str, width: int = 17, show: str = ""):
        row = tk.Frame(parent, bg=self.CLR_CARD)
        row.pack(fill=tk.X, padx=12, pady=4)
        tk.Label(
            row,
            text=label_text,
            width=width,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        var = tk.StringVar()
        entry = tk.Entry(
            row,
            textvariable=var,
            show=show,
            font=("Helvetica", 10),
            fg="#0f172a",
            bg="#ffffff",
            insertbackground="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#0f766e",
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        return var, entry

    def _double_entry_row(self, parent, left_label: str, right_label: str, left_width: int = 17):
        row = tk.Frame(parent, bg=self.CLR_CARD)
        row.pack(fill=tk.X, padx=12, pady=4)

        tk.Label(
            row,
            text=left_label,
            width=left_width,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        left_var = tk.StringVar()
        left_entry = tk.Entry(
            row,
            textvariable=left_var,
            font=("Helvetica", 10),
            fg="#0f172a",
            bg="#ffffff",
            insertbackground="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#0f766e",
        )
        left_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            row,
            text=right_label,
            width=12,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT, padx=(10, 0))
        right_var = tk.StringVar()
        right_entry = tk.Entry(
            row,
            textvariable=right_var,
            font=("Helvetica", 10),
            fg="#0f172a",
            bg="#ffffff",
            insertbackground="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#0f766e",
        )
        right_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        return left_var, left_entry, right_var, right_entry

    def _double_date_row(self, parent, left_label: str, right_label: str, left_width: int = 17):
        row = tk.Frame(parent, bg=self.CLR_CARD)
        row.pack(fill=tk.X, padx=12, pady=4)

        tk.Label(
            row,
            text=left_label,
            width=left_width,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        left_var = tk.StringVar()
        left_entry = self._create_date_input(row, left_var)
        left_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            row,
            text=right_label,
            width=12,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT, padx=(10, 0))
        right_var = tk.StringVar()
        right_entry = self._create_date_input(row, right_var)
        right_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        return left_var, left_entry, right_var, right_entry

    def _create_date_input(self, parent, variable: tk.StringVar):
        if Calendar is not None:
            wrapper = tk.Frame(parent, bg=self.CLR_CARD)
            entry = tk.Entry(
                wrapper,
                textvariable=variable,
                font=("Helvetica", 10),
                fg="#0f172a",
                bg="#ffffff",
                insertbackground="#0f172a",
                relief=tk.FLAT,
                highlightthickness=1,
                highlightbackground="#cbd5e1",
                highlightcolor="#0f766e",
                width=16,
            )
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

            ttk.Button(
                wrapper,
                text="v",
                width=2,
                command=lambda: self._open_calendar_picker(variable),
            ).pack(side=tk.LEFT, padx=(4, 0))
            return wrapper

        entry = tk.Entry(
            parent,
            textvariable=variable,
            font=("Helvetica", 10),
            fg="#0f172a",
            bg="#ffffff",
            insertbackground="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#0f766e",
        )
        return entry

    def _open_calendar_picker(self, variable: tk.StringVar):
        if Calendar is None:
            return

        parsed = _parse_user_date(variable.get().strip())
        initial = parsed.date() if parsed else _now_eastern().date()

        picker = tk.Toplevel(self.root)
        picker.title("Select Date")
        picker.configure(bg=self.CLR_CARD)
        picker.resizable(False, False)
        picker.transient(self.root)
        picker.grab_set()

        calendar = Calendar(
            picker,
            selectmode="day",
            year=initial.year,
            month=initial.month,
            day=initial.day,
            date_pattern="mm/dd/yyyy",
            background=self.CLR_HEADER,
            foreground="white",
            headersbackground=self.CLR_HEADER,
            headersforeground="white",
            selectbackground=self.CLR_ACCENT,
            selectforeground="white",
            normalbackground="#ffffff",
            normalforeground="#0f172a",
            weekendbackground="#f8fafc",
            weekendforeground="#0f172a",
        )
        calendar.pack(padx=10, pady=10)

        buttons = tk.Frame(picker, bg=self.CLR_CARD)
        buttons.pack(fill=tk.X, padx=10, pady=(0, 10))

        def apply_date():
            variable.set(calendar.get_date())
            picker.destroy()

        ttk.Button(buttons, text="Cancel", command=picker.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Select", command=apply_date).pack(side=tk.RIGHT, padx=(0, 6))

    def _build_settings(self, parent):
        self._section_title(parent, "Shared Email Settings")
        card = self._card(parent, pady=6)

        date_row = tk.Frame(card, bg=self.CLR_CARD)
        date_row.pack(fill=tk.X, padx=12, pady=6)
        tk.Label(
            date_row,
            text="Current Time:",
            width=17,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        self._date_label = tk.Label(
            date_row,
            text=get_display_datetime(),
            bg=self.CLR_WARN_BG,
            fg=self.CLR_WARN,
            font=("Helvetica", 10, "bold"),
            padx=8,
            pady=2,
        )
        self._date_label.pack(side=tk.LEFT)

        self._name_var, _ = self._labeled_entry(card, "Your Name:")
        self._email_var, _ = self._labeled_entry(card, "Sender Email:")

        pw_row = tk.Frame(card, bg=self.CLR_CARD)
        pw_row.pack(fill=tk.X, padx=12, pady=4)
        tk.Label(
            pw_row,
            text="Password:",
            width=17,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        self._password_var = tk.StringVar()
        self._password_entry = tk.Entry(
            pw_row,
            textvariable=self._password_var,
            show="*",
            font=("Helvetica", 10),
            fg="#0f172a",
            bg="#ffffff",
            insertbackground="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#0f766e",
        )
        self._password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(pw_row, text="Show", width=6, command=self._toggle_password).pack(side=tk.LEFT, padx=(4, 0))

        tk.Label(
            card,
            text="  Gmail users: use an App Password via Google Account > Security > 2-Step Verification > App Passwords",
            bg=self.CLR_CARD,
            fg=self.CLR_HINT,
            font=("Helvetica", 8, "italic"),
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))

        # SMTP Server Dropdown
        smtp_row = tk.Frame(card, bg=self.CLR_CARD)
        smtp_row.pack(fill=tk.X, padx=12, pady=4)
        tk.Label(
            smtp_row,
            text="SMTP Server:",
            width=17,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        self._smtp_var = tk.StringVar(value=DEFAULT_SMTP_SERVER)
        self._smtp_dropdown = ttk.Combobox(
            smtp_row,
            textvariable=self._smtp_var,
            values=SMTP_SERVERS_LIST,
            font=("Helvetica", 10),
            state="readonly",
        )
        self._smtp_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._port_var, _ = self._labeled_entry(card, "SMTP Port:")

        tk.Frame(card, bg=self.CLR_CARD, height=4).pack()

    def _build_mode_selector(self, parent):
        self._section_title(parent, "Mode")
        card = self._card(parent, pady=6)
        row = tk.Frame(card, bg=self.CLR_CARD)
        row.pack(fill=tk.X, padx=12, pady=10)

        ttk.Radiobutton(
            row,
            text="Daily Report",
            variable=self._mode_var,
            value="daily",
            command=self._update_mode_view,
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Radiobutton(
            row,
            text="Invoice Sender",
            variable=self._mode_var,
            value="invoice",
            command=self._update_mode_view,
        ).pack(side=tk.LEFT)

    def _build_daily_mode(self, parent):
        self._section_title(parent, "Daily Report")
        for question in QUESTIONS:
            card = self._card(parent, pady=4)

            tk.Label(
                card,
                text=question["label"],
                font=("Helvetica", 10, "bold"),
                bg=self.CLR_CARD,
                fg=self.CLR_LABEL,
                wraplength=720,
                justify=tk.LEFT,
                anchor=tk.W,
            ).pack(anchor=tk.W, padx=12, pady=(8, 0))

            tk.Label(
                card,
                text=question["hint"],
                font=("Helvetica", 8, "italic"),
                bg=self.CLR_CARD,
                fg=self.CLR_HINT,
                wraplength=720,
                justify=tk.LEFT,
                anchor=tk.W,
            ).pack(anchor=tk.W, padx=12, pady=(1, 4))

            text_widget = scrolledtext.ScrolledText(
                card,
                height=3,
                font=("Helvetica", 10),
                wrap=tk.WORD,
                relief=tk.FLAT,
                bd=0,
                highlightthickness=1,
                highlightbackground="#cbd5e1",
                highlightcolor=self.CLR_ACCENT,
                bg="#f8fafc",
                fg="#0f172a",
                insertbackground="#0f172a",
            )
            text_widget.pack(fill=tk.X, padx=12, pady=(0, 10))
            self.answer_widgets[question["key"]] = text_widget

        button_row = tk.Frame(parent, bg=self.CLR_BG)
        button_row.pack(fill=tk.X, padx=16, pady=8)
        ttk.Button(button_row, text="Preview Email", command=self._preview_daily_email).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Clear Answers", command=self._clear_answers).pack(side=tk.LEFT, padx=4)
        self._daily_send_button = ttk.Button(button_row, text="Send Daily Report", command=self._send_daily_email_thread)
        self._daily_send_button.pack(side=tk.RIGHT, padx=4)

    def _build_invoice_mode(self, parent):
        self._section_title(parent, "Invoice Sender")

        dest_card = self._card(parent, pady=4)
        tk.Label(
            dest_card,
            text="Invoice email recipients are configurable via .env and default to the training instructions.",
            bg=self.CLR_CARD,
            fg=self.CLR_HINT,
            font=("Helvetica", 9),
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(10, 2))
        self._invoice_dest_label = tk.Label(
            dest_card,
            text="",
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 9, "bold"),
            anchor=tk.NW,
            justify=tk.LEFT,
            wraplength=0,
        )
        self._invoice_dest_label.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        details_card = self._card(parent, pady=4)
        self._invoice_date_var, _, self._invoice_number_var, _ = self._double_entry_row(
            details_card,
            "Invoice Date:",
            "Invoice Number:",
        )
        self._week_start_var, self._week_start_entry, self._week_end_var, self._week_end_entry = self._double_date_row(
            details_card,
            "Week Start:",
            "Week End:",
        )
        self._hours_var, self._hours_entry, self._rate_var, _ = self._double_entry_row(
            details_card,
            "Hours Worked:",
            "Hourly Rate:",
        )
        self._hours_entry.configure(state="readonly", readonlybackground="#ffffff")

        tk.Label(
            details_card,
            text="Daily Hours (generated from week range)",
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10, "bold"),
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(6, 2))
        tk.Label(
            details_card,
            text="Enter hours per day. The app sums these automatically into Hours Worked.",
            bg=self.CLR_CARD,
            fg=self.CLR_HINT,
            font=("Helvetica", 8, "italic"),
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        self._daily_hours_frame = tk.Frame(details_card, bg=self.CLR_CARD)
        self._daily_hours_frame.pack(fill=tk.X, padx=12, pady=(0, 6))

        self._week_start_var.trace_add("write", lambda *_args: self._refresh_daily_hours_inputs())
        self._week_end_var.trace_add("write", lambda *_args: self._refresh_daily_hours_inputs())
        self._hours_var.trace_add("write", lambda *_args: self._update_invoice_total())
        self._rate_var.trace_add("write", lambda *_args: self._update_invoice_total())

        if DateEntry is None:
            tk.Label(
                details_card,
                text="Install tkcalendar to enable calendar dropdowns for week start and week end.",
                bg=self.CLR_CARD,
                fg=self.CLR_HINT,
                font=("Helvetica", 8, "italic"),
                anchor=tk.W,
            ).pack(fill=tk.X, padx=12, pady=(0, 4))

        total_row = tk.Frame(details_card, bg=self.CLR_CARD)
        total_row.pack(fill=tk.X, padx=12, pady=4)
        tk.Label(
            total_row,
            text="Total Due:",
            width=17,
            anchor=tk.W,
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            font=("Helvetica", 10),
        ).pack(side=tk.LEFT)
        self._invoice_total_var = tk.StringVar(value="$0.00")
        tk.Label(
            total_row,
            textvariable=self._invoice_total_var,
            bg="#ecfeff",
            fg="#155e75",
            font=("Helvetica", 10, "bold"),
            padx=8,
            pady=3,
        ).pack(side=tk.LEFT)

        self._payment_method_var, _ = self._labeled_entry(details_card, "Payment Method:")
        self._payment_email_var, _ = self._labeled_entry(details_card, "Payment Email:")

        tk.Label(
            details_card,
            text="The PDF will use the training invoice format: services description, amount due, payment terms, payment method, and notes.",
            bg=self.CLR_CARD,
            fg=self.CLR_HINT,
            font=("Helvetica", 8, "italic"),
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(0, 10))

        contact_card = self._card(parent, pady=4)
        self._address_var, _ = self._labeled_entry(contact_card, "Address:")
        self._city_state_zip_var, _ = self._labeled_entry(contact_card, "City/State/Zip:")
        self._phone_var, _ = self._labeled_entry(contact_card, "Phone Number:")

        notes_card = self._card(parent, pady=4)
        tk.Label(
            notes_card,
            text="Additional Notes (Optional)",
            font=("Helvetica", 10, "bold"),
            bg=self.CLR_CARD,
            fg=self.CLR_LABEL,
            anchor=tk.W,
        ).pack(fill=tk.X, padx=12, pady=(10, 2))
        tk.Label(
            notes_card,
            text="The invoice note is generated automatically from the dates, hours, and rate. Use this box only for extra context that should be appended.",
            font=("Helvetica", 8, "italic"),
            bg=self.CLR_CARD,
            fg=self.CLR_HINT,
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        self._invoice_notes = scrolledtext.ScrolledText(
            notes_card,
            height=6,
            font=("Helvetica", 10),
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor=self.CLR_ACCENT,
            bg="#f8fafc",
            fg="#0f172a",
            insertbackground="#0f172a",
        )
        self._invoice_notes.pack(fill=tk.X, padx=12, pady=(0, 10))

        button_row = tk.Frame(parent, bg=self.CLR_BG)
        button_row.pack(fill=tk.X, padx=16, pady=8)
        ttk.Button(button_row, text="Preview Invoice", command=self._preview_invoice).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Save PDF", command=self._save_invoice_pdf).pack(side=tk.LEFT, padx=4)
        self._invoice_send_button = ttk.Button(button_row, text="Send Invoice", command=self._send_invoice_thread)
        self._invoice_send_button.pack(side=tk.RIGHT, padx=4)

    def _load_saved_config(self):
        self._name_var.set(self.config.get("name", DEFAULT_NAME))
        self._email_var.set(self.config.get("email", DEFAULT_SENDER_EMAIL))
        self._password_var.set(DEFAULT_SENDER_PASSWORD)
        self._smtp_var.set(self.config.get("smtp_server", DEFAULT_SMTP_SERVER))
        self._port_var.set(self.config.get("smtp_port", DEFAULT_SMTP_PORT))

        self._invoice_date_var.set(self.config.get("invoice_date", get_invoice_date()))
        self._invoice_number_var.set(self.config.get("invoice_number", "01"))
        self._week_start_var.set(self.config.get("week_start", ""))
        self._week_end_var.set(self.config.get("week_end", ""))
        self._hours_var.set(self.config.get("hours_worked", "40"))
        self._saved_daily_hours = self.config.get("daily_hours", {})
        self._rate_var.set(self.config.get("hourly_rate", DEFAULT_HOURLY_RATE))
        self._payment_method_var.set(self.config.get("payment_method", DEFAULT_PAYMENT_METHOD))
        self._payment_email_var.set(self.config.get("payment_email", DEFAULT_PAYMENT_EMAIL or self.config.get("email", "")))
        self._address_var.set(self.config.get("address", DEFAULT_ADDRESS))
        self._city_state_zip_var.set(self.config.get("city_state_zip", DEFAULT_CITY_STATE_ZIP))
        self._phone_var.set(self.config.get("phone", DEFAULT_PHONE))
        self._invoice_notes.insert("1.0", self.config.get("invoice_notes", ""))
        self._sync_date_widgets()
        self._refresh_daily_hours_inputs()

    def _sync_date_widgets(self):
        return

    def _save_form_config(self):
        save_config(
            {
                "name": self._name_var.get().strip(),
                "email": self._email_var.get().strip(),
                "smtp_server": self._smtp_var.get().strip(),
                "smtp_port": self._port_var.get().strip(),
                "invoice_date": self._invoice_date_var.get().strip(),
                "invoice_number": self._invoice_number_var.get().strip(),
                "week_start": self._week_start_var.get().strip(),
                "week_end": self._week_end_var.get().strip(),
                "hours_worked": self._hours_var.get().strip(),
                "daily_hours": {key: var.get().strip() for key, var in self._daily_hours_vars.items()},
                "hourly_rate": self._rate_var.get().strip(),
                "payment_method": self._payment_method_var.get().strip(),
                "payment_email": self._payment_email_var.get().strip(),
                "address": self._address_var.get().strip(),
                "city_state_zip": self._city_state_zip_var.get().strip(),
                "phone": self._phone_var.get().strip(),
                "invoice_notes": self._invoice_notes.get("1.0", tk.END).strip(),
            }
        )

    def _toggle_password(self):
        self._show_password = not self._show_password
        self._password_entry.configure(show="" if self._show_password else "*")

    def _schedule_date_refresh(self):
        self._date_label.configure(text=get_display_datetime())
        self.root.after(60_000, self._schedule_date_refresh)

    def _on_mousewheel(self, event):
        if getattr(event, "num", None) == 4:
            self._canvas.yview_scroll(-1, "units")
            return

        if getattr(event, "num", None) == 5:
            self._canvas.yview_scroll(1, "units")
            return

        delta = getattr(event, "delta", 0)
        if delta == 0:
            return

        if sys.platform == "darwin":
            step = -1 if delta > 0 else 1
        else:
            step = int(-1 * (delta / 120))
            if step == 0:
                step = -1 if delta > 0 else 1

        self._canvas.yview_scroll(step, "units")

    def _update_mode_view(self):
        self._daily_frame.pack_forget()
        self._invoice_frame.pack_forget()

        if self._mode_var.get() == "daily":
            self._daily_frame.pack(fill=tk.X)
            subtitle = f"Daily Report recipients: To: {DAILY_TO_EMAIL} | CC: {DAILY_CC_EMAIL or 'None'}"
        else:
            self._invoice_frame.pack(fill=tk.X)
            invoice_to = ", ".join(parse_email_list(INVOICE_TO_EMAILS)) or "None"
            invoice_cc = ", ".join(parse_email_list(INVOICE_CC_EMAILS)) or "None"
            subtitle = f"Invoice recipients: To: {invoice_to} | CC: {invoice_cc}"

        self._header_subtitle.configure(text=subtitle)
        self._invoice_dest_label.configure(
            text=(
                f"To: {', '.join(parse_email_list(INVOICE_TO_EMAILS)) or 'None'}\n"
                f"CC: {', '.join(parse_email_list(INVOICE_CC_EMAILS)) or 'None'}"
            )
        )
        self.root.after_idle(lambda: self._canvas.configure(scrollregion=self._canvas.bbox("all")))

    def _validate_shared_fields(self, require_password: bool) -> bool:
        checks = [
            (self._name_var.get().strip(), "Please enter your name."),
            (self._email_var.get().strip(), "Please enter your sender email address."),
            (self._smtp_var.get().strip(), "Please enter the SMTP server address."),
            (self._port_var.get().strip(), "Please enter the SMTP port."),
        ]
        if require_password:
            checks.insert(2, (self._password_var.get().strip(), "Please enter your email password."))

        for value, message in checks:
            if not value:
                messagebox.showerror("Missing Information", message)
                return False

        try:
            port = int(self._port_var.get().strip())
            if not (1 <= port <= 65535):
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Port", "SMTP port must be a number between 1 and 65535.")
            return False

        return True

    def _get_daily_answers(self) -> dict:
        return {key: widget.get("1.0", tk.END).strip() for key, widget in self.answer_widgets.items()}

    def _build_daily_subject(self) -> str:
        name = self._name_var.get().strip() or "(Your Name)"
        return f"{name} Daily Report of {get_report_date()}"

    def _build_daily_body(self) -> tuple[str, str]:
        answers = self._get_daily_answers()

        plain_lines = []
        for question in QUESTIONS:
            plain_lines.append(question["label"])
            plain_lines.append(answers.get(question["key"], ""))
            plain_lines.append("")
        plain_text = "\n".join(plain_lines).strip()

        html_lines = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            '  <meta charset="utf-8">',
            '  <meta name="viewport" content="width=device-width, initial-scale=1">',
            "  <style>",
            "    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #243041; max-width: 680px; margin: 0 auto; padding: 24px; background: #f8fafc; }",
            "    .header { background: linear-gradient(135deg, #0f766e 0%, #164e63 100%); color: white; padding: 28px; border-radius: 14px; margin-bottom: 24px; }",
            "    .header h1 { margin: 0; font-size: 26px; }",
            "    .header p { margin: 6px 0 0; opacity: 0.92; }",
            "    .card { background: #ffffff; border: 1px solid #dbe4ee; border-radius: 12px; padding: 18px; margin-bottom: 14px; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04); }",
            "    .card h3 { margin: 0 0 10px; font-size: 14px; color: #0f172a; }",
            "    .card p { margin: 0; white-space: pre-wrap; color: #334155; }",
            "  </style>",
            "</head>",
            "<body>",
            f"  <div class='header'><h1>{escape(self._name_var.get().strip() or 'Daily Report')}</h1><p>{escape(get_report_date())}</p></div>",
        ]

        for question in QUESTIONS:
            label = escape(question["label"])
            answer = answers.get(question["key"], "").strip()
            answer_html = htmlize_multiline(answer) if answer else "<em>No response</em>"
            html_lines.append(f"  <div class='card'><h3>{label}</h3><p>{answer_html}</p></div>")

        html_lines.extend(["</body>", "</html>"])
        return plain_text, "\n".join(html_lines)

    def _validate_daily(self) -> bool:
        if not self._validate_shared_fields(require_password=True):
            return False

        answers = self._get_daily_answers()
        for question in QUESTIONS:
            if not answers.get(question["key"]):
                messagebox.showerror("Missing Answer", f"Please answer:\n\n{question['label']}")
                return False
        return True

    def _preview_daily_email(self):
        subject = self._build_daily_subject()
        plain_text, _html_text = self._build_daily_body()

        win = tk.Toplevel(self.root)
        win.title("Daily Report Preview")
        win.geometry("680x620")
        win.configure(bg=self.CLR_BG)

        info = tk.Frame(win, bg=self.CLR_CARD, bd=1, relief=tk.GROOVE)
        info.pack(fill=tk.X, padx=12, pady=(12, 4))
        for label, value in [("To:", DAILY_TO_EMAIL), ("CC:", DAILY_CC_EMAIL), ("Subject:", subject)]:
            row = tk.Frame(info, bg=self.CLR_CARD)
            row.pack(fill=tk.X, padx=10, pady=2)
            tk.Label(row, text=label, width=9, anchor=tk.W, bg=self.CLR_CARD, font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
            tk.Label(row, text=value or "None", bg=self.CLR_CARD, font=("Helvetica", 10), fg="#334155").pack(side=tk.LEFT)

        body = scrolledtext.ScrolledText(win, font=("Helvetica", 10), wrap=tk.WORD, padx=12, pady=10, bg="#f8fafc", relief=tk.FLAT, bd=0)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        body.insert("1.0", plain_text)
        body.configure(state=tk.DISABLED)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 10))

    def _clear_answers(self):
        if messagebox.askyesno("Clear Answers", "This will erase all daily report answers. Continue?"):
            for widget in self.answer_widgets.values():
                widget.delete("1.0", tk.END)
            self._status_var.set("Daily report answers cleared.")

    def _build_email_message(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        plain_body: str,
        html_body: str,
        attachments: list[tuple[str, bytes, str]] | None = None,
    ) -> MIMEMultipart:
        sender = self._email_var.get().strip()
        sender_name = self._name_var.get().strip()

        message = MIMEMultipart("mixed")
        message["From"] = f"{sender_name} <{sender}>"
        message["To"] = ", ".join(to_list)
        if cc_list:
            message["CC"] = ", ".join(cc_list)
        message["Subject"] = subject

        alternative = MIMEMultipart("alternative")
        alternative.attach(MIMEText(plain_body, "plain", "utf-8"))
        alternative.attach(MIMEText(html_body, "html", "utf-8"))
        message.attach(alternative)

        for file_name, payload, mime_subtype in attachments or []:
            attachment = MIMEApplication(payload, _subtype=mime_subtype)
            attachment.add_header("Content-Disposition", "attachment", filename=file_name)
            message.attach(attachment)

        return message

    def _send_message(self, message: MIMEMultipart, recipients: list[str]):
        sender = self._email_var.get().strip()
        password = self._password_var.get().strip()
        smtp_server = self._smtp_var.get().strip()
        smtp_port = int(self._port_var.get().strip())

        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender, password)
            server.sendmail(sender, recipients, message.as_string())

    def _send_daily_email_thread(self):
        if not self._validate_daily():
            return
        self._save_form_config()
        self._daily_send_button.configure(state=tk.DISABLED)
        self._status_var.set("Sending daily report...")
        threading.Thread(target=self._do_send_daily, daemon=True).start()

    def _do_send_daily(self):
        try:
            to_list = parse_email_list(DAILY_TO_EMAIL)
            cc_list = parse_email_list(DAILY_CC_EMAIL)
            plain_body, html_body = self._build_daily_body()
            message = self._build_email_message(
                to_list,
                cc_list,
                self._build_daily_subject(),
                plain_body,
                html_body,
            )
            self._send_message(message, to_list + cc_list)
            self.root.after(0, self._on_daily_send_success)
        except smtplib.SMTPAuthenticationError:
            self.root.after(0, lambda: self._on_send_error("Authentication failed. Gmail users should use an App Password."))
        except Exception as exc:
            self.root.after(0, lambda: self._on_send_error(str(exc)))

    def _on_daily_send_success(self):
        self._daily_send_button.configure(state=tk.NORMAL)
        self._status_var.set("Daily report sent successfully.")
        messagebox.showinfo(
            "Report Sent",
            f"Your daily report was sent successfully.\n\nTo: {DAILY_TO_EMAIL}\nCC: {DAILY_CC_EMAIL or 'None'}",
        )

    def _update_invoice_total(self):
        try:
            hours = float(self._hours_var.get().strip() or "0")
            rate = float(self._rate_var.get().strip() or "0")
            total = hours * rate
        except ValueError:
            total = 0.0
        self._invoice_total_var.set(currency(total))

    def _refresh_daily_hours_inputs(self):
        if not hasattr(self, "_daily_hours_frame"):
            return

        # Keep any in-memory edits before rebuilding the dynamic day rows.
        if self._daily_hours_vars:
            self._saved_daily_hours.update({key: var.get().strip() for key, var in self._daily_hours_vars.items()})

        for child in self._daily_hours_frame.winfo_children():
            child.destroy()

        self._daily_hours_vars = {}

        start_dt = _parse_user_date(self._week_start_var.get().strip())
        end_dt = _parse_user_date(self._week_end_var.get().strip())
        if not start_dt or not end_dt or end_dt < start_dt:
            tk.Label(
                self._daily_hours_frame,
                text="Set valid Week Start and Week End dates to generate day-hour inputs.",
                bg=self.CLR_CARD,
                fg=self.CLR_HINT,
                font=("Helvetica", 9),
                anchor=tk.W,
            ).pack(fill=tk.X, pady=(0, 2))
            self._hours_var.set("0")
            return

        day_count = (end_dt.date() - start_dt.date()).days + 1
        if day_count > 31:
            tk.Label(
                self._daily_hours_frame,
                text="Date range is too large. Please use a range up to 31 days.",
                bg=self.CLR_CARD,
                fg=self.CLR_WARN,
                font=("Helvetica", 9),
                anchor=tk.W,
            ).pack(fill=tk.X, pady=(0, 2))
            self._hours_var.set("0")
            return

        for offset in range(day_count):
            current_date = start_dt.date().fromordinal(start_dt.date().toordinal() + offset)
            date_key = current_date.strftime("%Y-%m-%d")
            var = tk.StringVar(value=str(self._saved_daily_hours.get(date_key, "")))
            var.trace_add("write", lambda *_args: self._recompute_hours_from_days())
            self._daily_hours_vars[date_key] = var

            cell = tk.Frame(self._daily_hours_frame, bg=self.CLR_CARD)
            row_index = offset // 7
            col_index = offset % 7
            cell.grid(row=row_index, column=col_index, padx=4, pady=2, sticky="w")

            tk.Label(
                cell,
                text=current_date.strftime("%m/%d"),
                anchor=tk.CENTER,
                bg=self.CLR_CARD,
                fg=self.CLR_LABEL,
                font=("Helvetica", 9, "bold"),
            ).pack()

            entry = tk.Entry(
                cell,
                textvariable=var,
                width=5,
                font=("Helvetica", 10),
                fg="#0f172a",
                bg="#ffffff",
                insertbackground="#0f172a",
                relief=tk.FLAT,
                highlightthickness=1,
                highlightbackground="#cbd5e1",
                highlightcolor="#0f766e",
                justify=tk.CENTER,
            )
            entry.pack(pady=(2, 0))

        self._recompute_hours_from_days()

    def _recompute_hours_from_days(self):
        if self._recomputing_hours:
            return

        total_hours = 0.0
        for var in self._daily_hours_vars.values():
            value = var.get().strip()
            if not value:
                continue
            try:
                total_hours += float(value)
            except ValueError:
                continue

        self._recomputing_hours = True
        self._hours_var.set(_format_quantity(total_hours))
        self._recomputing_hours = False

    def _collect_invoice_data(self, require_password: bool = False) -> dict | None:
        if not self._validate_shared_fields(require_password=require_password):
            return None

        required_values = [
            (self._invoice_date_var.get().strip(), "Please enter the invoice date."),
            (self._invoice_number_var.get().strip(), "Please enter the invoice number."),
            (self._week_start_var.get().strip(), "Please enter the week start date."),
            (self._week_end_var.get().strip(), "Please enter the week end date."),
            (self._address_var.get().strip(), "Please enter your address."),
            (self._city_state_zip_var.get().strip(), "Please enter your city/state/zip."),
            (self._phone_var.get().strip(), "Please enter your phone number."),
        ]
        for value, message in required_values:
            if not value:
                messagebox.showerror("Missing Information", message)
                return None

        payment_method = self._payment_method_var.get().strip() or DEFAULT_PAYMENT_METHOD
        payment_email = self._payment_email_var.get().strip() or self._email_var.get().strip()
        if not payment_email:
            messagebox.showerror("Missing Information", "Please enter the payment email address.")
            return None

        try:
            hours = float(self._hours_var.get().strip())
            rate = float(self._rate_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Number", "Hours worked and hourly rate must both be numeric values.")
            return None

        if hours <= 0 or rate <= 0:
            messagebox.showerror("Invalid Number", "Hours worked and hourly rate must both be greater than zero.")
            return None

        total_due = hours * rate
        description = f"Virtual Assistant Training for the week of {self._week_start_var.get().strip()} to {self._week_end_var.get().strip()}"
        extra_notes = self._invoice_notes.get("1.0", tk.END).strip()
        notes = build_invoice_notes(
            self._week_start_var.get().strip(),
            self._week_end_var.get().strip(),
            hours,
            rate,
            total_due,
            extra_notes,
        )

        data = {
            "invoice_date": self._invoice_date_var.get().strip(),
            "invoice_number": self._invoice_number_var.get().strip(),
            "name": self._name_var.get().strip(),
            "address": self._address_var.get().strip(),
            "city_state_zip": self._city_state_zip_var.get().strip(),
            "email": self._email_var.get().strip(),
            "phone": self._phone_var.get().strip(),
            "week_start": self._week_start_var.get().strip(),
            "week_end": self._week_end_var.get().strip(),
            "hours": hours,
            "rate": rate,
            "total_due": total_due,
            "description": description,
            "payment_terms": DEFAULT_PAYMENT_TERMS,
            "payment_method": payment_method,
            "payment_email": payment_email,
            "notes": notes,
        }
        return data

    def _build_invoice_subject(self, invoice_data: dict) -> str:
        return (
            f"Invoice {invoice_data['invoice_number']} - "
            f"{invoice_data['name']} - {invoice_data['week_start']} to {invoice_data['week_end']}"
        )

    def _build_invoice_email_body(self, invoice_data: dict) -> tuple[str, str]:
        plain_text = (
            "Hello,\n\n"
            f"Attached is my invoice for Virtual Assistant Training for the week of {invoice_data['week_start']} to {invoice_data['week_end']}.\n\n"
            f"Invoice Number: {invoice_data['invoice_number']}\n"
            f"Total Due: {currency(invoice_data['total_due'])}\n"
            f"Payment Method: {invoice_data['payment_method']} ({invoice_data['payment_email']})\n\n"
            "Thank you."
        )

        html_body = "".join(
            [
                "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>",
                "<style>",
                "body{font-family:'Segoe UI',Arial,sans-serif;background:#f8fafc;color:#243041;padding:24px;}",
                ".card{max-width:680px;margin:0 auto;background:#fff;border:1px solid #dbe4ee;border-radius:14px;overflow:hidden;}",
                ".hero{background:linear-gradient(135deg,#16324f 0%,#0f766e 100%);color:#fff;padding:24px 28px;}",
                ".hero h1{margin:0;font-size:24px;}",
                ".content{padding:24px 28px;}",
                ".meta{background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:16px;margin:18px 0;}",
                ".meta-row{margin:6px 0;color:#334155;}",
                "</style></head><body>",
                "<div class='card'>",
                "<div class='hero'><h1>Invoice Attached</h1></div>",
                "<div class='content'>",
                f"<p>Hello,</p><p>Attached is my invoice for Virtual Assistant Training for the week of <strong>{escape(invoice_data['week_start'])}</strong> to <strong>{escape(invoice_data['week_end'])}</strong>.</p>",
                "<div class='meta'>",
                f"<div class='meta-row'><strong>Invoice Number:</strong> {escape(invoice_data['invoice_number'])}</div>",
                f"<div class='meta-row'><strong>Total Due:</strong> {escape(currency(invoice_data['total_due']))}</div>",
                f"<div class='meta-row'><strong>Payment Method:</strong> {escape(invoice_data['payment_method'])} ({escape(invoice_data['payment_email'])})</div>",
                "</div>",
                "<p>Thank you.</p>",
                "</div></div></body></html>",
            ]
        )
        return plain_text, html_body

    def _invoice_preview_text(self, invoice_data: dict) -> str:
        return (
            f"Invoice Date: {invoice_data['invoice_date']}\n"
            f"Invoice Number: {invoice_data['invoice_number']}\n\n"
            f"From:\n{invoice_data['name']}\n{invoice_data['address']}\n{invoice_data['city_state_zip']}\n"
            f"{invoice_data['email']}\n{invoice_data['phone']}\n\n"
            f"To:\n" + "\n".join(COMPANY_LINES) + "\n\n"
            f"Description of Services Provided:\n{invoice_data['description']}\n\n"
            f"Hours Worked: {invoice_data['hours']:.2f}\n"
            f"Hourly Rate: {currency(invoice_data['rate'])}\n"
            f"Total Due: {currency(invoice_data['total_due'])}\n\n"
            f"Payment Terms: {invoice_data['payment_terms']}\n"
            f"Payment Method: {invoice_data['payment_method']} ({invoice_data['payment_email']})\n\n"
            f"Notes:\n{invoice_data['notes']}\n"
        )

    def _build_default_invoice_path(self, invoice_data: dict) -> str:
        os.makedirs(GENERATED_DIR, exist_ok=True)
        filename = sanitize_filename(
            f"invoice-{invoice_data['invoice_number']}-{invoice_data['week_end']}-{invoice_data['name']}"
        )
        return os.path.join(GENERATED_DIR, f"{filename}.pdf")

    def _write_invoice_pdf(self, invoice_data: dict, output_path: str):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        doc = SimpleDocTemplate(
            output_path,
            pagesize=LETTER,
            leftMargin=0.7 * inch,
            rightMargin=0.7 * inch,
            topMargin=0.6 * inch,
            bottomMargin=0.6 * inch,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "InvoiceTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=28,
            textColor=colors.HexColor("#16324f"),
            spaceAfter=12,
        )
        heading_style = ParagraphStyle(
            "SectionHeading",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#0f172a"),
            spaceAfter=4,
        )
        body_style = ParagraphStyle(
            "BodyCopy",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#334155"),
        )

        story = []
        story.append(Paragraph("INVOICE", title_style))

        meta_table = Table(
            [
                [Paragraph("<b>Invoice Date</b>", body_style), Paragraph(escape(invoice_data["invoice_date"]), body_style)],
                [Paragraph("<b>Invoice Number</b>", body_style), Paragraph(escape(invoice_data["invoice_number"]), body_style)],
            ],
            colWidths=[1.7 * inch, 1.9 * inch],
        )
        meta_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eff6ff")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#bfdbfe")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bfdbfe")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(meta_table)
        story.append(Spacer(1, 14))

        from_text = "<br/>".join(
            [
                f"<b>{escape(invoice_data['name'])}</b>",
                escape(invoice_data["address"]),
                escape(invoice_data["city_state_zip"]),
                escape(invoice_data["email"]),
                escape(invoice_data["phone"]),
            ]
        )
        to_text = "<br/>".join([f"<b>{escape(line)}</b>" if index == 0 else escape(line) for index, line in enumerate(COMPANY_LINES)])

        address_table = Table(
            [
                [Paragraph("From", heading_style), Paragraph("To", heading_style)],
                [Paragraph(from_text, body_style), Paragraph(to_text, body_style)],
            ],
            colWidths=[3.05 * inch, 3.05 * inch],
        )
        address_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f8fafc")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#dbe4ee")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dbe4ee")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(address_table)
        story.append(Spacer(1, 16))

        description_table = Table(
            [
                ["Description of Services Provided", "Hours", "Rate", "Amount"],
                [
                    Paragraph(escape(invoice_data["description"]), body_style),
                    Paragraph(f"{invoice_data['hours']:.2f}", body_style),
                    Paragraph(currency(invoice_data["rate"]), body_style),
                    Paragraph(currency(invoice_data["total_due"]), body_style),
                ],
            ],
            colWidths=[3.45 * inch, 0.8 * inch, 0.8 * inch, 1.05 * inch],
        )
        description_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16324f")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#cbd5e1")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(description_table)
        story.append(Spacer(1, 14))

        totals_table = Table(
            [[Paragraph("<b>Total Due</b>", body_style), Paragraph(f"<b>{escape(currency(invoice_data['total_due']))}</b>", body_style)]],
            colWidths=[1.5 * inch, 1.2 * inch],
            hAlign="RIGHT",
        )
        totals_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ecfeff")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#99f6e4")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(totals_table)
        story.append(Spacer(1, 16))

        story.append(Paragraph("Payment Terms", heading_style))
        story.append(Paragraph(escape(invoice_data["payment_terms"]), body_style))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Payment Method", heading_style))
        story.append(Paragraph(escape(f"{invoice_data['payment_method']} ({invoice_data['payment_email']})"), body_style))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Notes", heading_style))
        story.append(Paragraph(htmlize_multiline(invoice_data["notes"]), body_style))
        story.append(Spacer(1, 16))
        story.append(Paragraph("Thank you for your business!", heading_style))

        doc.build(story)

    def _preview_invoice(self):
        invoice_data = self._collect_invoice_data(require_password=False)
        if not invoice_data:
            return
        self._save_form_config()

        win = tk.Toplevel(self.root)
        win.title("Invoice Preview")
        win.geometry("700x620")
        win.configure(bg=self.CLR_BG)

        body = scrolledtext.ScrolledText(win, font=("Helvetica", 10), wrap=tk.WORD, padx=12, pady=10, bg="#f8fafc", relief=tk.FLAT, bd=0)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        body.insert("1.0", self._invoice_preview_text(invoice_data))
        body.configure(state=tk.DISABLED)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 12))

    def _save_invoice_pdf(self):
        invoice_data = self._collect_invoice_data(require_password=False)
        if not invoice_data:
            return
        self._save_form_config()

        default_path = self._build_default_invoice_path(invoice_data)
        output_path = filedialog.asksaveasfilename(
            title="Save Invoice PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialdir=os.path.dirname(default_path),
            initialfile=os.path.basename(default_path),
        )
        if not output_path:
            return

        try:
            self._write_invoice_pdf(invoice_data, output_path)
            self._status_var.set(f"Invoice PDF saved to {output_path}")
            messagebox.showinfo("PDF Saved", f"Invoice PDF saved successfully.\n\n{output_path}")
        except Exception as exc:
            messagebox.showerror("PDF Error", f"Could not create the invoice PDF.\n\n{exc}")

    def _send_invoice_thread(self):
        invoice_data = self._collect_invoice_data(require_password=True)
        if not invoice_data:
            return
        self._save_form_config()
        self._invoice_send_button.configure(state=tk.DISABLED)
        self._status_var.set("Generating invoice PDF and sending email...")
        threading.Thread(target=self._do_send_invoice, args=(invoice_data,), daemon=True).start()

    def _do_send_invoice(self, invoice_data: dict):
        try:
            pdf_path = self._build_default_invoice_path(invoice_data)
            self._write_invoice_pdf(invoice_data, pdf_path)
            with open(pdf_path, "rb") as handle:
                pdf_bytes = handle.read()

            to_list = parse_email_list(INVOICE_TO_EMAILS)
            cc_list = parse_email_list(INVOICE_CC_EMAILS)
            plain_body, html_body = self._build_invoice_email_body(invoice_data)
            message = self._build_email_message(
                to_list,
                cc_list,
                self._build_invoice_subject(invoice_data),
                plain_body,
                html_body,
                attachments=[(os.path.basename(pdf_path), pdf_bytes, "pdf")],
            )
            self._send_message(message, to_list + cc_list)
            self.root.after(0, lambda: self._on_invoice_send_success(pdf_path, to_list, cc_list))
        except smtplib.SMTPAuthenticationError:
            self.root.after(0, lambda: self._on_send_error("Authentication failed. Gmail users should use an App Password."))
        except Exception as exc:
            self.root.after(0, lambda: self._on_send_error(str(exc)))

    def _on_invoice_send_success(self, pdf_path: str, to_list: list[str], cc_list: list[str]):
        self._invoice_send_button.configure(state=tk.NORMAL)
        self._status_var.set("Invoice sent successfully.")
        messagebox.showinfo(
            "Invoice Sent",
            "Your invoice was sent successfully.\n\n"
            f"To: {', '.join(to_list) or 'None'}\n"
            f"CC: {', '.join(cc_list) or 'None'}\n\n"
            f"PDF: {pdf_path}",
        )

    def _on_send_error(self, message: str):
        self._daily_send_button.configure(state=tk.NORMAL)
        self._invoice_send_button.configure(state=tk.NORMAL)
        self._status_var.set("Send failed. See the error message for details.")
        messagebox.showerror("Send Failed", f"Could not send email.\n\n{message}")


def main():
    root = tk.Tk()
    MailerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
