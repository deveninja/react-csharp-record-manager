from __future__ import annotations

import json
import os
import re
import smtplib
import sys
import threading
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtGui import QDoubleValidator
from PySide6.QtWidgets import (
    QApplication,
    QCalendarWidget,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    from zoneinfo import ZoneInfo

    EASTERN_TZ = ZoneInfo("America/New_York")
except Exception:
    EASTERN_TZ = None

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "config.json")
ENV_FILE = os.path.join(APP_DIR, ".env")
GENERATED_DIR = os.path.join(APP_DIR, "generated-invoices")
INVOICE_TEMPLATE_STANDARD = "Standard Invoice"
INVOICE_TEMPLATE_CLIENT_TRACKING = "Client Tracking Invoice"


def load_env_file(path: str):
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except Exception:
        pass


load_env_file(ENV_FILE)
if not os.path.exists(ENV_FILE):
    load_env_file(os.path.join(os.getcwd(), ".env"))

DAILY_TO_EMAIL = os.getenv("TO_EMAIL", "sid@learncodinganywhere.com")
DAILY_CC_EMAIL = os.getenv("CC_EMAIL", "jack@learncodinganywhere.com")
INVOICE_TO_EMAILS = os.getenv("INVOICE_TO_EMAILS", "finances@learncodinganywhere.com,sid@learncodinganywhere.com")
INVOICE_CC_EMAILS = os.getenv("INVOICE_CC_EMAILS", "sid@learncodinganywhere.com")
SMTP_SERVERS_LIST = [s.strip() for s in os.getenv("SMTP_SERVERS", "smtp.gmail.com,smtp.mail.yahoo.com,smtp-mail.outlook.com").split(",") if s.strip()]
DEFAULT_SMTP_SERVER = SMTP_SERVERS_LIST[0] if SMTP_SERVERS_LIST else "smtp.gmail.com"
DEFAULT_SMTP_PORT = os.getenv("SMTP_PORT", "587")
DEFAULT_HOURLY_RATE = os.getenv("INVOICE_HOURLY_RATE", "7.50")
DEFAULT_PAYMENT_METHOD = os.getenv("INVOICE_PAYMENT_METHOD", "Payoneer")
DEFAULT_PAYMENT_EMAIL = os.getenv("INVOICE_PAYMENT_EMAIL", "")
DEFAULT_PAYMENT_TERMS = os.getenv("INVOICE_PAYMENT_TERMS", "Please make payment within 30 days of the invoice date.")
DEFAULT_NAME = os.getenv("YOUR_NAME", "")
DEFAULT_SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
DEFAULT_SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "")
DEFAULT_ADDRESS = os.getenv("INVOICE_ADDRESS", "")
DEFAULT_CITY_STATE_ZIP = os.getenv("INVOICE_CITY_STATE_ZIP", "")
DEFAULT_PHONE = os.getenv("INVOICE_PHONE", "")

COMPANY_LINES = [
    "The Tech Academy",
    "310 SW 4th Ave Suite 200",
    "Portland, Oregon 97204",
]

QUESTIONS = [
    {"key": "q1", "label": "1. What course are you on?", "hint": "Example: Basics of English Course"},
    {"key": "q2", "label": "2. Which step of the course are you on?", "hint": "Example: Step 92 at 52% training progress."},
    {"key": "q3", "label": "3. What non-course related tasks did you complete today?", "hint": "Example: None. or Called students in the relinquished pipeline for one hour."},
    {"key": "q4", "label": "4. Do you have any slows or issues that you need assistance with?", "hint": "Example: None. or The video on step 15 didn't play."},
    {"key": "q5", "label": "5. Do you have any questions?", "hint": "Example: I'm all good for now."},
    {"key": "q6", "label": "6. Is there anything else you would like to communicate?", "hint": "Example: I enjoyed learning today."},
]


def _now_eastern() -> datetime:
    return datetime.now(EASTERN_TZ) if EASTERN_TZ else datetime.now()


def get_report_date() -> str:
    return _now_eastern().strftime("%B %d, %Y")


def get_display_datetime() -> str:
    now = _now_eastern()
    tz = now.strftime("%Z") if EASTERN_TZ else "Local Time"
    return now.strftime(f"%A, %B %d, %Y  %I:%M %p  ({tz})")


def get_invoice_date() -> str:
    return _now_eastern().strftime("%m/%d/%Y")


def _parse_user_date(value: str) -> datetime | None:
    raw = (value or "").strip()
    if not raw:
        return None
    for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y"]:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def parse_email_list(value: str) -> list[str]:
    return [item.strip() for item in re.split(r"[;,]", value or "") if item.strip()]


def currency(value: float) -> str:
    return f"${value:,.2f}"


def htmlize_multiline(value: str) -> str:
    return escape(value).replace("\n", "<br>")


def _format_quantity(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    cleaned = cleaned.strip(".-")
    return cleaned or "invoice"


def build_invoice_notes(week_start: str, week_end: str, hours: float, rate: float, total_due: float, extra_notes: str = "") -> str:
    base_note = (
        f"Total number of hours worked for the period from {week_start} to {week_end} is {_format_quantity(hours)} "
        f"at {currency(rate)}/hour, which comes up to {currency(total_due)}."
    )
    extra_notes = extra_notes.strip()
    return f"{base_note}\n\nAdditional Notes: {extra_notes}" if extra_notes else base_note


def invoice_template_options() -> list[str]:
    return [INVOICE_TEMPLATE_STANDARD, INVOICE_TEMPLATE_CLIENT_TRACKING]


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(data: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


class CalendarDialog(QDialog):
    def __init__(self, parent: QWidget, initial: datetime | None):
        super().__init__(parent)
        self.setWindowTitle("Select Date")
        self.setModal(True)
        self.resize(360, 320)
        self.setStyleSheet(
            """
            QDialog { background: #f8fafc; }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background: #16324f;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }
            QCalendarWidget QToolButton {
                color: #ffffff;
                font-weight: 700;
                background: transparent;
                border: none;
                padding: 6px 10px;
            }
            QCalendarWidget QMenu {
                background: #ffffff;
                color: #0f172a;
            }
            QCalendarWidget QSpinBox {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 2px 6px;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: #0f172a;
                background: #ffffff;
                selection-background-color: #0f766e;
                selection-color: #ffffff;
            }
            """
        )
        layout = QVBoxLayout(self)
        self.calendar = QCalendarWidget(self)
        self.calendar.setGridVisible(False)
        d = initial.date() if initial else _now_eastern().date()
        self.calendar.setSelectedDate(QDate(d.year, d.month, d.day))
        layout.addWidget(self.calendar)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_date(self) -> str:
        d = self.calendar.selectedDate()
        return f"{d.month():02d}/{d.day():02d}/{d.year():04d}"


class MailerWindow(QMainWindow):
    daily_done = Signal(bool, str)
    invoice_done = Signal(bool, str, str, str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PITC Mailer")
        self.resize(1020, 920)

        self.config = load_config()
        self.answer_widgets: dict[str, QTextEdit] = {}
        self.daily_hours_edits: dict[str, QLineEdit] = {}
        self.daily_hours_errors: dict[str, QLabel] = {}
        self.saved_daily_hours: dict[str, str] = self.config.get("daily_hours", {})
        self.recomputing_hours = False

        self.daily_done.connect(self._on_daily_done)
        self.invoice_done.connect(self._on_invoice_done)

        self._build_ui()
        self._apply_style()
        self._load_saved_config()
        self._refresh_daily_hours_inputs()
        self._update_invoice_total()

    def _apply_style(self):
        self.setStyleSheet(
            """
            QWidget { background: #f4f7fb; color: #0f172a; font-size: 13px; }
            QGroupBox { border: 1px solid #d7e1ee; border-radius: 10px; margin-top: 10px; padding: 14px 12px 10px 12px; background: #ffffff; font-weight: 700; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 6px; color: #16324f; }
            QLineEdit, QTextEdit, QComboBox { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px 8px; selection-background-color: #0f766e; }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus { border: 1px solid #0f766e; }
            QPushButton { background: #16324f; color: #ffffff; border: none; border-radius: 8px; padding: 7px 12px; font-weight: 600; }
            QPushButton:hover { background: #1d4369; }
            QPushButton:disabled { background: #94a3b8; }
            QPushButton[datePicker='true'] { background: #ecfeff; color: #155e75; border: 1px solid #99f6e4; padding: 6px 10px; }
            QPushButton[datePicker='true']:hover { background: #cffafe; }
            QLabel[muted='true'] { color: #64748b; }
            QLabel[error='true'] { color: #b91c1c; font-size: 11px; }
            """
        )

    def _build_ui(self):
        root_host = QWidget(self)
        self.setCentralWidget(root_host)
        root = QVBoxLayout(root_host)

        title = QLabel("Modern Mailer")
        title.setStyleSheet("font-size: 24px; font-weight: 800; color: #16324f;")
        subtitle = QLabel("Daily Report + Invoice Sender")
        subtitle.setProperty("muted", True)
        root.addWidget(title)
        root.addWidget(subtitle)

        shared_box = QGroupBox("Shared Email Settings")
        shared = QGridLayout(shared_box)

        self.time_label = QLabel(get_display_datetime())
        self.name_edit = QLineEdit()
        self.email_edit = QLineEdit()
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.smtp_combo = QComboBox()
        self.smtp_combo.addItems(SMTP_SERVERS_LIST)
        self.port_edit = QLineEdit()

        shared.addWidget(QLabel("Current Time:"), 0, 0)
        shared.addWidget(self.time_label, 0, 1, 1, 3)
        shared.addWidget(QLabel("Your Name:"), 1, 0)
        shared.addWidget(self.name_edit, 1, 1)
        shared.addWidget(QLabel("Sender Email:"), 1, 2)
        shared.addWidget(self.email_edit, 1, 3)
        shared.addWidget(QLabel("Password:"), 2, 0)
        shared.addWidget(self.password_edit, 2, 1)
        btn_pw = QPushButton("Show")
        btn_pw.clicked.connect(self._toggle_password)
        shared.addWidget(btn_pw, 2, 2)
        shared.addWidget(QLabel("SMTP Server:"), 3, 0)
        shared.addWidget(self.smtp_combo, 3, 1)
        shared.addWidget(QLabel("SMTP Port:"), 3, 2)
        shared.addWidget(self.port_edit, 3, 3)

        root.addWidget(shared_box)

        mode_box = QGroupBox("Mode")
        mode = QHBoxLayout(mode_box)
        self.daily_radio = QRadioButton("Daily Report")
        self.invoice_radio = QRadioButton("Invoice Sender")
        self.daily_radio.setChecked(True)
        self.daily_radio.toggled.connect(self._update_mode_view)
        self.invoice_radio.toggled.connect(self._update_mode_view)
        mode.addWidget(self.daily_radio)
        mode.addWidget(self.invoice_radio)
        mode.addStretch()
        root.addWidget(mode_box)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll.setWidget(self.scroll_content)
        root.addWidget(self.scroll, 1)

        self.daily_panel = self._build_daily_panel()
        self.invoice_panel = self._build_invoice_panel()
        self.scroll_layout.addWidget(self.daily_panel)
        self.scroll_layout.addWidget(self.invoice_panel)
        self.scroll_layout.addStretch()

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("background: #e2e8f0; color:#334155; padding: 8px; border-radius: 6px;")
        root.addWidget(self.status_label)

    def _build_daily_panel(self) -> QWidget:
        panel = QGroupBox("Daily Report")
        layout = QVBoxLayout(panel)
        for q in QUESTIONS:
            box = QWidget()
            row = QVBoxLayout(box)
            title = QLabel(q["label"])
            title.setStyleSheet("font-weight: 700;")
            hint = QLabel(q["hint"])
            hint.setProperty("muted", True)
            hint.setWordWrap(True)
            txt = QTextEdit()
            txt.setFixedHeight(80)
            row.addWidget(title)
            row.addWidget(hint)
            row.addWidget(txt)
            layout.addWidget(box)
            self.answer_widgets[q["key"]] = txt

        action = QHBoxLayout()
        btn_preview = QPushButton("Preview Email")
        btn_preview.clicked.connect(self._preview_daily_email)
        btn_clear = QPushButton("Clear Answers")
        btn_clear.clicked.connect(self._clear_answers)
        self.btn_send_daily = QPushButton("Send Daily Report")
        self.btn_send_daily.clicked.connect(self._send_daily)
        action.addWidget(btn_preview)
        action.addWidget(btn_clear)
        action.addStretch()
        action.addWidget(self.btn_send_daily)
        layout.addLayout(action)
        return panel

    def _create_date_button(self, target: QLineEdit) -> QPushButton:
        button = QPushButton("Pick date")
        button.setProperty("datePicker", True)
        button.setCursor(Qt.PointingHandCursor)
        button.clicked.connect(lambda: self._pick_date(target))
        return button

    def _build_invoice_panel(self) -> QWidget:
        panel = QGroupBox("Invoice Sender")
        layout = QVBoxLayout(panel)

        self.invoice_dest_label = QLabel("")
        self.invoice_dest_label.setWordWrap(True)
        layout.addWidget(self.invoice_dest_label)

        g = QGridLayout()
        self.invoice_date_edit = QLineEdit()
        self.invoice_number_edit = QLineEdit()
        self.week_start_edit = QLineEdit()
        self.week_end_edit = QLineEdit()
        self.hours_edit = QLineEdit()
        self.hours_edit.setReadOnly(True)
        self.rate_edit = QLineEdit()
        self.rate_edit.setValidator(QDoubleValidator(0.0, 100000.0, 2))

        g.addWidget(QLabel("Invoice Date:"), 0, 0)
        g.addWidget(self.invoice_date_edit, 0, 1)
        b1 = self._create_date_button(self.invoice_date_edit)
        g.addWidget(b1, 0, 2)
        g.addWidget(QLabel("Invoice Number:"), 0, 3)
        g.addWidget(self.invoice_number_edit, 0, 4)

        g.addWidget(QLabel("Week Start:"), 1, 0)
        g.addWidget(self.week_start_edit, 1, 1)
        b2 = self._create_date_button(self.week_start_edit)
        g.addWidget(b2, 1, 2)
        g.addWidget(QLabel("Week End:"), 1, 3)
        g.addWidget(self.week_end_edit, 1, 4)
        b3 = self._create_date_button(self.week_end_edit)
        g.addWidget(b3, 1, 5)

        g.addWidget(QLabel("Hours Worked:"), 2, 0)
        g.addWidget(self.hours_edit, 2, 1)
        g.addWidget(QLabel("Hourly Rate:"), 2, 3)
        g.addWidget(self.rate_edit, 2, 4)
        g.addWidget(QLabel("Template:"), 3, 0)
        self.invoice_template_combo = QComboBox()
        self.invoice_template_combo.addItems(invoice_template_options())
        g.addWidget(self.invoice_template_combo, 3, 1, 1, 2)
        layout.addLayout(g)

        day_box = QGroupBox("Daily Hours")
        day_layout = QVBoxLayout(day_box)
        hint = QLabel("Generated from Week Start to Week End. Use values from 0 to 24.")
        hint.setProperty("muted", True)
        day_layout.addWidget(hint)
        self.day_grid = QGridLayout()
        day_layout.addLayout(self.day_grid)
        self.day_invalid_msg = QLabel("")
        self.day_invalid_msg.setProperty("error", True)
        day_layout.addWidget(self.day_invalid_msg)
        layout.addWidget(day_box)

        total = QHBoxLayout()
        total.addWidget(QLabel("Total Due:"))
        self.total_due_label = QLabel("$0.00")
        self.total_due_label.setStyleSheet("font-weight: 800; color: #155e75;")
        total.addWidget(self.total_due_label)
        total.addStretch()
        layout.addLayout(total)

        c = QGridLayout()
        self.payment_method_edit = QLineEdit()
        self.payment_email_edit = QLineEdit()
        self.address_edit = QLineEdit()
        self.city_state_zip_edit = QLineEdit()
        self.phone_edit = QLineEdit()

        c.addWidget(QLabel("Payment Method:"), 0, 0)
        c.addWidget(self.payment_method_edit, 0, 1)
        c.addWidget(QLabel("Payment Email:"), 0, 2)
        c.addWidget(self.payment_email_edit, 0, 3)
        c.addWidget(QLabel("Address:"), 1, 0)
        c.addWidget(self.address_edit, 1, 1, 1, 3)
        c.addWidget(QLabel("City/State/Zip:"), 2, 0)
        c.addWidget(self.city_state_zip_edit, 2, 1)
        c.addWidget(QLabel("Phone:"), 2, 2)
        c.addWidget(self.phone_edit, 2, 3)
        layout.addLayout(c)

        layout.addWidget(QLabel("Additional Notes (Optional)"))
        self.notes_edit = QTextEdit()
        self.notes_edit.setFixedHeight(120)
        layout.addWidget(self.notes_edit)

        row = QHBoxLayout()
        btn_preview = QPushButton("Preview Invoice")
        btn_preview.clicked.connect(self._preview_invoice)
        btn_save = QPushButton("Save PDF")
        btn_save.clicked.connect(self._save_pdf)
        self.btn_send_invoice = QPushButton("Send Invoice")
        self.btn_send_invoice.clicked.connect(self._send_invoice)
        row.addWidget(btn_preview)
        row.addWidget(btn_save)
        row.addStretch()
        row.addWidget(self.btn_send_invoice)
        layout.addLayout(row)

        self.week_start_edit.textChanged.connect(self._refresh_daily_hours_inputs)
        self.week_end_edit.textChanged.connect(self._refresh_daily_hours_inputs)
        self.rate_edit.textChanged.connect(self._update_invoice_total)
        return panel

    def _pick_date(self, target: QLineEdit):
        dlg = CalendarDialog(self, _parse_user_date(target.text().strip()))
        if dlg.exec() == QDialog.Accepted:
            target.setText(dlg.selected_date())

    def _toggle_password(self):
        self.password_edit.setEchoMode(QLineEdit.Normal if self.password_edit.echoMode() == QLineEdit.Password else QLineEdit.Password)

    def _update_mode_view(self):
        daily = self.daily_radio.isChecked()
        self.daily_panel.setVisible(daily)
        self.invoice_panel.setVisible(not daily)
        if daily:
            msg = f"Daily Report recipients: To: {DAILY_TO_EMAIL} | CC: {DAILY_CC_EMAIL or 'None'}"
        else:
            to = ", ".join(parse_email_list(INVOICE_TO_EMAILS)) or "None"
            cc = ", ".join(parse_email_list(INVOICE_CC_EMAILS)) or "None"
            msg = f"Invoice recipients: To: {to} | CC: {cc}"
        self.status_label.setText(msg)
        self.invoice_dest_label.setText(
            f"To: {', '.join(parse_email_list(INVOICE_TO_EMAILS)) or 'None'}\n"
            f"CC: {', '.join(parse_email_list(INVOICE_CC_EMAILS)) or 'None'}"
        )

    def _load_saved_config(self):
        cfg = self.config
        self.name_edit.setText(cfg.get("name", DEFAULT_NAME))
        self.email_edit.setText(cfg.get("email", DEFAULT_SENDER_EMAIL))
        self.password_edit.setText(DEFAULT_SENDER_PASSWORD)
        idx = self.smtp_combo.findText(cfg.get("smtp_server", DEFAULT_SMTP_SERVER))
        if idx >= 0:
            self.smtp_combo.setCurrentIndex(idx)
        self.port_edit.setText(cfg.get("smtp_port", DEFAULT_SMTP_PORT))

        self.invoice_date_edit.setText(cfg.get("invoice_date", get_invoice_date()))
        self.invoice_number_edit.setText(cfg.get("invoice_number", "01"))
        self.week_start_edit.setText(cfg.get("week_start", ""))
        self.week_end_edit.setText(cfg.get("week_end", ""))
        self.hours_edit.setText(cfg.get("hours_worked", "0"))
        template_idx = self.invoice_template_combo.findText(cfg.get("invoice_template", INVOICE_TEMPLATE_STANDARD))
        if template_idx >= 0:
            self.invoice_template_combo.setCurrentIndex(template_idx)
        self.rate_edit.setText(cfg.get("hourly_rate", DEFAULT_HOURLY_RATE))
        self.payment_method_edit.setText(cfg.get("payment_method", DEFAULT_PAYMENT_METHOD))
        self.payment_email_edit.setText(cfg.get("payment_email", DEFAULT_PAYMENT_EMAIL or cfg.get("email", "")))
        self.address_edit.setText(cfg.get("address", DEFAULT_ADDRESS))
        self.city_state_zip_edit.setText(cfg.get("city_state_zip", DEFAULT_CITY_STATE_ZIP))
        self.phone_edit.setText(cfg.get("phone", DEFAULT_PHONE))
        self.notes_edit.setPlainText(cfg.get("invoice_notes", ""))
        self.saved_daily_hours = cfg.get("daily_hours", {})
        self._update_mode_view()

    def _save_form_config(self):
        save_config(
            {
                "name": self.name_edit.text().strip(),
                "email": self.email_edit.text().strip(),
                "smtp_server": self.smtp_combo.currentText().strip(),
                "smtp_port": self.port_edit.text().strip(),
                "invoice_date": self.invoice_date_edit.text().strip(),
                "invoice_number": self.invoice_number_edit.text().strip(),
                "week_start": self.week_start_edit.text().strip(),
                "week_end": self.week_end_edit.text().strip(),
                "hours_worked": self.hours_edit.text().strip(),
                "invoice_template": self.invoice_template_combo.currentText().strip(),
                "daily_hours": {k: e.text().strip() for k, e in self.daily_hours_edits.items()},
                "hourly_rate": self.rate_edit.text().strip(),
                "payment_method": self.payment_method_edit.text().strip(),
                "payment_email": self.payment_email_edit.text().strip(),
                "address": self.address_edit.text().strip(),
                "city_state_zip": self.city_state_zip_edit.text().strip(),
                "phone": self.phone_edit.text().strip(),
                "invoice_notes": self.notes_edit.toPlainText().strip(),
            }
        )

    def _clear_day_grid(self):
        while self.day_grid.count():
            item = self.day_grid.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _refresh_daily_hours_inputs(self):
        self._clear_day_grid()
        self.day_invalid_msg.clear()

        if self.daily_hours_edits:
            self.saved_daily_hours.update({k: e.text().strip() for k, e in self.daily_hours_edits.items()})
        self.daily_hours_edits = {}
        self.daily_hours_errors = {}

        start_dt = _parse_user_date(self.week_start_edit.text().strip())
        end_dt = _parse_user_date(self.week_end_edit.text().strip())
        if not start_dt or not end_dt or end_dt < start_dt:
            self.hours_edit.setText("0")
            self._update_invoice_total()
            return

        day_count = (end_dt.date() - start_dt.date()).days + 1
        if day_count > 31:
            self.day_invalid_msg.setText("Date range is too large. Use up to 31 days.")
            self.hours_edit.setText("0")
            self._update_invoice_total()
            return

        for i in range(day_count):
            d = start_dt.date() + timedelta(days=i)
            key = d.strftime("%Y-%m-%d")
            col = i % 7
            base_row = (i // 7) * 3

            lbl = QLabel(d.strftime("%m/%d"))
            lbl.setAlignment(Qt.AlignCenter)
            edit = QLineEdit(self.saved_daily_hours.get(key, ""))
            edit.setAlignment(Qt.AlignCenter)
            edit.setFixedWidth(72)
            edit.setValidator(QDoubleValidator(0.0, 100.0, 2))
            err = QLabel("")
            err.setProperty("error", True)
            err.setAlignment(Qt.AlignCenter)

            edit.textChanged.connect(lambda _text, day_key=key: self._validate_day(day_key))
            edit.textChanged.connect(self._recompute_hours)

            self.day_grid.addWidget(lbl, base_row, col)
            self.day_grid.addWidget(edit, base_row + 1, col)
            self.day_grid.addWidget(err, base_row + 2, col)

            self.daily_hours_edits[key] = edit
            self.daily_hours_errors[key] = err

        self._recompute_hours()

    def _validate_day(self, day_key: str):
        edit = self.daily_hours_edits.get(day_key)
        err = self.daily_hours_errors.get(day_key)
        if edit is None or err is None:
            return

        raw = edit.text().strip()
        if not raw:
            err.setText("")
            edit.setStyleSheet("")
            return

        ok = True
        try:
            v = float(raw)
            if v < 0 or v > 24:
                ok = False
        except ValueError:
            ok = False

        if ok:
            err.setText("")
            edit.setStyleSheet("")
        else:
            err.setText("0-24")
            edit.setStyleSheet("border: 1px solid #b91c1c;")

    def _recompute_hours(self):
        if self.recomputing_hours:
            return

        total = 0.0
        invalid = []
        for key, edit in self.daily_hours_edits.items():
            raw = edit.text().strip()
            if not raw:
                continue
            try:
                v = float(raw)
                if 0 <= v <= 24:
                    total += v
                else:
                    invalid.append(key)
            except ValueError:
                invalid.append(key)

        self.recomputing_hours = True
        self.hours_edit.setText(_format_quantity(total))
        self.recomputing_hours = False

        if invalid:
            short = [datetime.strptime(k, "%Y-%m-%d").strftime("%m/%d") for k in invalid[:4]]
            suffix = "..." if len(invalid) > 4 else ""
            self.day_invalid_msg.setText(f"Invalid hours on: {', '.join(short)}{suffix}")
        else:
            self.day_invalid_msg.setText("")

        self._update_invoice_total()

    def _update_invoice_total(self):
        try:
            hours = float(self.hours_edit.text().strip() or "0")
            rate = float(self.rate_edit.text().strip() or "0")
            total = hours * rate
        except ValueError:
            total = 0.0
        self.total_due_label.setText(currency(total))

    def _daily_time_rows(self, rate: float) -> list[dict]:
        rows = []
        for day_key in sorted(self.daily_hours_edits.keys()):
            raw_hours = self.daily_hours_edits[day_key].text().strip()
            try:
                hours = float(raw_hours) if raw_hours else 0.0
            except ValueError:
                hours = 0.0

            current = datetime.strptime(day_key, "%Y-%m-%d")
            rows.append(
                {
                    "date": current.strftime("%m/%d/%Y"),
                    "day": current.strftime("%A"),
                    "hours": hours,
                    "amount": hours * rate,
                }
            )
        return rows

    def _validate_shared(self, require_password: bool) -> bool:
        checks = [
            (self.name_edit.text().strip(), "Please enter your name."),
            (self.email_edit.text().strip(), "Please enter your sender email address."),
            (self.smtp_combo.currentText().strip(), "Please enter the SMTP server address."),
            (self.port_edit.text().strip(), "Please enter the SMTP port."),
        ]
        if require_password:
            checks.insert(2, (self.password_edit.text().strip(), "Please enter your email password."))
        for value, msg in checks:
            if not value:
                QMessageBox.critical(self, "Missing Information", msg)
                return False
        try:
            port = int(self.port_edit.text().strip())
            if not (1 <= port <= 65535):
                raise ValueError
        except ValueError:
            QMessageBox.critical(self, "Invalid Port", "SMTP port must be a number between 1 and 65535.")
            return False
        return True

    def _daily_answers(self) -> dict:
        return {k: w.toPlainText().strip() for k, w in self.answer_widgets.items()}

    def _build_daily_subject(self) -> str:
        name = self.name_edit.text().strip() or "(Your Name)"
        return f"{name} Daily Report of {get_report_date()}"

    def _build_daily_body(self) -> tuple[str, str]:
        answers = self._daily_answers()
        plain_lines = []
        for q in QUESTIONS:
            plain_lines.append(q["label"])
            plain_lines.append(answers.get(q["key"], ""))
            plain_lines.append("")
        plain_text = "\n".join(plain_lines).strip()

        html = [
            "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>",
            "<style>body{font-family:'Segoe UI',Arial,sans-serif;line-height:1.6;color:#243041;max-width:680px;margin:0 auto;padding:24px;background:#f8fafc;}.header{background:linear-gradient(135deg,#0f766e 0%,#164e63 100%);color:white;padding:28px;border-radius:14px;margin-bottom:24px;}.card{background:#fff;border:1px solid #dbe4ee;border-radius:12px;padding:18px;margin-bottom:14px;}</style>",
            "</head><body>",
            f"<div class='header'><h1>{escape(self.name_edit.text().strip() or 'Daily Report')}</h1><p>{escape(get_report_date())}</p></div>",
        ]
        for q in QUESTIONS:
            ans = answers.get(q["key"], "").strip()
            ans_html = htmlize_multiline(ans) if ans else "<em>No response</em>"
            html.append(f"<div class='card'><h3>{escape(q['label'])}</h3><p>{ans_html}</p></div>")
        html.append("</body></html>")
        return plain_text, "\n".join(html)

    def _preview_daily_email(self):
        subject = self._build_daily_subject()
        plain, _ = self._build_daily_body()
        dlg = QDialog(self)
        dlg.setWindowTitle("Daily Report Preview")
        dlg.resize(720, 640)
        v = QVBoxLayout(dlg)
        v.addWidget(QLabel(f"To: {DAILY_TO_EMAIL}"))
        v.addWidget(QLabel(f"CC: {DAILY_CC_EMAIL or 'None'}"))
        v.addWidget(QLabel(f"Subject: {subject}"))
        body = QTextEdit()
        body.setReadOnly(True)
        body.setPlainText(plain)
        v.addWidget(body)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.button(QDialogButtonBox.Close).clicked.connect(dlg.reject)
        v.addWidget(buttons)
        dlg.exec()

    def _clear_answers(self):
        if QMessageBox.question(self, "Clear Answers", "This will erase all daily report answers. Continue?") != QMessageBox.Yes:
            return
        for w in self.answer_widgets.values():
            w.clear()
        self.status_label.setText("Daily report answers cleared.")

    def _build_email_message(self, to_list, cc_list, subject, plain_body, html_body, attachments=None):
        sender = self.email_edit.text().strip()
        sender_name = self.name_edit.text().strip()
        msg = MIMEMultipart("mixed")
        msg["From"] = f"{sender_name} <{sender}>"
        msg["To"] = ", ".join(to_list)
        if cc_list:
            msg["CC"] = ", ".join(cc_list)
        msg["Subject"] = subject
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain_body, "plain", "utf-8"))
        alt.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alt)
        for fname, payload, subtype in attachments or []:
            a = MIMEApplication(payload, _subtype=subtype)
            a.add_header("Content-Disposition", "attachment", filename=fname)
            msg.attach(a)
        return msg

    def _send_message(self, msg, recipients):
        sender = self.email_edit.text().strip()
        password = self.password_edit.text().strip()
        smtp_server = self.smtp_combo.currentText().strip()
        smtp_port = int(self.port_edit.text().strip())
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, password)
            s.sendmail(sender, recipients, msg.as_string())

    def _send_daily(self):
        if not self._validate_shared(require_password=True):
            return
        answers = self._daily_answers()
        for q in QUESTIONS:
            if not answers.get(q["key"]):
                QMessageBox.critical(self, "Missing Answer", f"Please answer:\n\n{q['label']}")
                return
        self._save_form_config()
        self.btn_send_daily.setEnabled(False)
        self.status_label.setText("Sending daily report...")
        threading.Thread(target=self._do_send_daily, daemon=True).start()

    def _do_send_daily(self):
        try:
            to_list = parse_email_list(DAILY_TO_EMAIL)
            cc_list = parse_email_list(DAILY_CC_EMAIL)
            plain, html = self._build_daily_body()
            msg = self._build_email_message(to_list, cc_list, self._build_daily_subject(), plain, html)
            self._send_message(msg, to_list + cc_list)
            self.daily_done.emit(True, "")
        except smtplib.SMTPAuthenticationError:
            self.daily_done.emit(False, "Authentication failed. Gmail users should use an App Password.")
        except Exception as exc:
            self.daily_done.emit(False, str(exc))

    def _on_daily_done(self, ok: bool, message: str):
        self.btn_send_daily.setEnabled(True)
        if ok:
            self.status_label.setText("Daily report sent successfully.")
            QMessageBox.information(self, "Report Sent", f"Your daily report was sent successfully.\n\nTo: {DAILY_TO_EMAIL}\nCC: {DAILY_CC_EMAIL or 'None'}")
        else:
            self.status_label.setText("Send failed. See the error message for details.")
            QMessageBox.critical(self, "Send Failed", f"Could not send email.\n\n{message}")

    def _collect_invoice(self, require_password=False):
        if not self._validate_shared(require_password=require_password):
            return None
        required = [
            (self.invoice_date_edit.text().strip(), "Please enter the invoice date."),
            (self.invoice_number_edit.text().strip(), "Please enter the invoice number."),
            (self.week_start_edit.text().strip(), "Please enter the week start date."),
            (self.week_end_edit.text().strip(), "Please enter the week end date."),
            (self.address_edit.text().strip(), "Please enter your address."),
            (self.city_state_zip_edit.text().strip(), "Please enter your city/state/zip."),
            (self.phone_edit.text().strip(), "Please enter your phone number."),
        ]
        for value, msg in required:
            if not value:
                QMessageBox.critical(self, "Missing Information", msg)
                return None

        try:
            hours = float(self.hours_edit.text().strip())
            rate = float(self.rate_edit.text().strip())
        except ValueError:
            QMessageBox.critical(self, "Invalid Number", "Hours worked and hourly rate must both be numeric values.")
            return None
        if hours <= 0 or rate <= 0:
            QMessageBox.critical(self, "Invalid Number", "Hours worked and hourly rate must both be greater than zero.")
            return None

        payment_method = self.payment_method_edit.text().strip() or DEFAULT_PAYMENT_METHOD
        payment_email = self.payment_email_edit.text().strip() or self.email_edit.text().strip()
        if not payment_email:
            QMessageBox.critical(self, "Missing Information", "Please enter the payment email address.")
            return None

        ws = self.week_start_edit.text().strip()
        we = self.week_end_edit.text().strip()
        total = hours * rate
        notes = build_invoice_notes(ws, we, hours, rate, total, self.notes_edit.toPlainText().strip())

        return {
            "invoice_date": self.invoice_date_edit.text().strip(),
            "invoice_number": self.invoice_number_edit.text().strip(),
            "invoice_template": self.invoice_template_combo.currentText().strip() or INVOICE_TEMPLATE_STANDARD,
            "name": self.name_edit.text().strip(),
            "address": self.address_edit.text().strip(),
            "city_state_zip": self.city_state_zip_edit.text().strip(),
            "email": self.email_edit.text().strip(),
            "phone": self.phone_edit.text().strip(),
            "week_start": ws,
            "week_end": we,
            "hours": hours,
            "rate": rate,
            "total_due": total,
            "description": f"Virtual Assistant Training for the week of {ws} to {we}",
            "payment_terms": DEFAULT_PAYMENT_TERMS,
            "payment_method": payment_method,
            "payment_email": payment_email,
            "notes": notes,
            "daily_time_rows": self._daily_time_rows(rate),
        }

    def _invoice_subject(self, d):
        return f"Invoice {d['invoice_number']} - {d['name']} - {d['week_start']} to {d['week_end']}"

    def _invoice_body(self, d):
        plain = (
            "Hello,\n\n"
            f"Attached is my invoice for Virtual Assistant Training for the week of {d['week_start']} to {d['week_end']}.\n\n"
            f"Invoice Number: {d['invoice_number']}\n"
            f"Total Due: {currency(d['total_due'])}\n"
            f"Payment Method: {d['payment_method']} ({d['payment_email']})\n\n"
            "Thank you."
        )
        html = (
            "<!DOCTYPE html><html><body style='font-family:Segoe UI,Arial,sans-serif;background:#f8fafc;padding:24px;'>"
            "<div style='max-width:680px;margin:0 auto;background:#fff;border:1px solid #dbe4ee;border-radius:14px;overflow:hidden'>"
            "<div style='background:linear-gradient(135deg,#16324f 0%,#0f766e 100%);color:#fff;padding:24px 28px;'><h1 style='margin:0'>Invoice Attached</h1></div>"
            "<div style='padding:24px 28px'>"
            f"<p>Hello,</p><p>Attached is my invoice for Virtual Assistant Training for the week of <strong>{escape(d['week_start'])}</strong> to <strong>{escape(d['week_end'])}</strong>.</p>"
            f"<p><strong>Invoice Number:</strong> {escape(d['invoice_number'])}<br><strong>Total Due:</strong> {escape(currency(d['total_due']))}<br><strong>Payment Method:</strong> {escape(d['payment_method'])} ({escape(d['payment_email'])})</p>"
            "<p>Thank you.</p></div></div></body></html>"
        )
        return plain, html

    def _invoice_preview_text(self, d):
        tracking_lines = []
        if d["invoice_template"] == INVOICE_TEMPLATE_CLIENT_TRACKING:
            tracking_lines.append("Client Tracking")
            tracking_lines.append("Date | Day | Hours | Amount")
            for row in d["daily_time_rows"]:
                tracking_lines.append(
                    f"{row['date']} | {row['day']} | {_format_quantity(row['hours'])} | {currency(row['amount'])}"
                )
            tracking_lines.append("")

        return (
            f"Template: {d['invoice_template']}\n"
            f"Invoice Date: {d['invoice_date']}\n"
            f"Invoice Number: {d['invoice_number']}\n\n"
            f"From:\n{d['name']}\n{d['address']}\n{d['city_state_zip']}\n{d['email']}\n{d['phone']}\n\n"
            + "To:\n" + "\n".join(COMPANY_LINES) + "\n\n"
            + f"Description of Services Provided:\n{d['description']}\n\n"
            + f"Hours Worked: {d['hours']:.2f}\nHourly Rate: {currency(d['rate'])}\nTotal Due: {currency(d['total_due'])}\n\n"
            + ("\n".join(tracking_lines) if tracking_lines else "")
            + f"Payment Terms: {d['payment_terms']}\n"
            + f"Payment Method: {d['payment_method']} ({d['payment_email']})\n\n"
            + f"Notes:\n{d['notes']}\n"
        )

    def _default_pdf_path(self, d):
        os.makedirs(GENERATED_DIR, exist_ok=True)
        template_suffix = "client-tracking" if d["invoice_template"] == INVOICE_TEMPLATE_CLIENT_TRACKING else "standard"
        fname = sanitize_filename(f"invoice-{d['invoice_number']}-{d['week_end']}-{template_suffix}-{d['name']}")
        return os.path.join(GENERATED_DIR, f"{fname}.pdf")

    def _write_pdf(self, d, output_path):
        if d["invoice_template"] == INVOICE_TEMPLATE_CLIENT_TRACKING:
            self._write_client_tracking_pdf(d, output_path)
            return

        self._write_standard_pdf(d, output_path)

    def _write_standard_pdf(self, d, output_path):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc = SimpleDocTemplate(output_path, pagesize=LETTER, leftMargin=0.7 * inch, rightMargin=0.7 * inch, topMargin=0.6 * inch, bottomMargin=0.6 * inch)
        styles = getSampleStyleSheet()
        title = ParagraphStyle("TitleX", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=24, leading=28, textColor=colors.HexColor("#16324f"), spaceAfter=12)
        heading = ParagraphStyle("HeadingX", parent=styles["Heading3"], fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=colors.HexColor("#0f172a"), spaceAfter=4)
        body = ParagraphStyle("BodyX", parent=styles["BodyText"], fontName="Helvetica", fontSize=10, leading=14, textColor=colors.HexColor("#334155"))

        story = [Paragraph("INVOICE", title)]
        meta = Table([
            [Paragraph("<b>Invoice Date</b>", body), Paragraph(escape(d["invoice_date"]), body)],
            [Paragraph("<b>Invoice Number</b>", body), Paragraph(escape(d["invoice_number"]), body)],
        ], colWidths=[1.7 * inch, 1.9 * inch])
        meta.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eff6ff")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#bfdbfe")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bfdbfe")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story += [meta, Spacer(1, 14)]

        from_text = "<br/>".join([f"<b>{escape(d['name'])}</b>", escape(d["address"]), escape(d["city_state_zip"]), escape(d["email"]), escape(d["phone"])])
        to_text = "<br/>".join([f"<b>{escape(line)}</b>" if i == 0 else escape(line) for i, line in enumerate(COMPANY_LINES)])
        addr = Table([[Paragraph("From", heading), Paragraph("To", heading)], [Paragraph(from_text, body), Paragraph(to_text, body)]], colWidths=[3.05 * inch, 3.05 * inch])
        addr.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f8fafc")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#dbe4ee")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dbe4ee")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [addr, Spacer(1, 16)]

        desc = Table([
            ["Description of Services Provided", "Hours", "Rate", "Amount"],
            [Paragraph(escape(d["description"]), body), Paragraph(f"{d['hours']:.2f}", body), Paragraph(currency(d["rate"]), body), Paragraph(currency(d["total_due"]), body)],
        ], colWidths=[3.45 * inch, 0.8 * inch, 0.8 * inch, 1.05 * inch])
        desc.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16324f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [desc, Spacer(1, 14)]

        total = Table([[Paragraph("<b>Total Due</b>", body), Paragraph(f"<b>{escape(currency(d['total_due']))}</b>", body)]], colWidths=[1.5 * inch, 1.2 * inch], hAlign="RIGHT")
        total.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ecfeff")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#99f6e4")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [total, Spacer(1, 16)]

        story.append(Paragraph("Payment Terms", heading))
        story.append(Paragraph(escape(d["payment_terms"]), body))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Payment Method", heading))
        story.append(Paragraph(escape(f"{d['payment_method']} ({d['payment_email']})"), body))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Notes", heading))
        story.append(Paragraph(htmlize_multiline(d["notes"]), body))
        story.append(Spacer(1, 16))
        story.append(Paragraph("Thank you for your business!", heading))
        doc.build(story)

    def _write_client_tracking_pdf(self, d, output_path):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc = SimpleDocTemplate(output_path, pagesize=LETTER, leftMargin=0.55 * inch, rightMargin=0.55 * inch, topMargin=0.55 * inch, bottomMargin=0.55 * inch)
        styles = getSampleStyleSheet()
        title = ParagraphStyle("TrackingTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=colors.HexColor("#16324f"), spaceAfter=10)
        heading = ParagraphStyle("TrackingHeading", parent=styles["Heading3"], fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=colors.HexColor("#0f172a"), spaceAfter=4)
        body = ParagraphStyle("TrackingBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=13, textColor=colors.HexColor("#334155"))

        story = [Paragraph("CLIENT TRACKING INVOICE", title)]
        meta = Table(
            [
                [Paragraph("<b>Invoice Date</b>", body), Paragraph(escape(d["invoice_date"]), body), Paragraph("<b>Invoice Number</b>", body), Paragraph(escape(d["invoice_number"]), body)],
                [Paragraph("<b>Billing Period</b>", body), Paragraph(escape(f"{d['week_start']} to {d['week_end']}"), body), Paragraph("<b>Template</b>", body), Paragraph(escape(d["invoice_template"]), body)],
            ],
            colWidths=[1.15 * inch, 1.85 * inch, 1.15 * inch, 2.0 * inch],
        )
        meta.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eff6ff")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#bfdbfe")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bfdbfe")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story += [meta, Spacer(1, 12)]

        from_text = "<br/>".join([f"<b>{escape(d['name'])}</b>", escape(d["address"]), escape(d["city_state_zip"]), escape(d["email"]), escape(d["phone"])])
        to_text = "<br/>".join([f"<b>{escape(line)}</b>" if i == 0 else escape(line) for i, line in enumerate(COMPANY_LINES)])
        addr = Table([[Paragraph("From", heading), Paragraph("To", heading)], [Paragraph(from_text, body), Paragraph(to_text, body)]], colWidths=[3.1 * inch, 3.1 * inch])
        addr.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f8fafc")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#dbe4ee")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dbe4ee")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [addr, Spacer(1, 12)]

        summary = Table(
            [
                [Paragraph("<b>Description</b>", body), Paragraph("<b>Total Hours</b>", body), Paragraph("<b>Rate</b>", body), Paragraph("<b>Total Due</b>", body)],
                [Paragraph(escape(d["description"]), body), Paragraph(_format_quantity(d["hours"]), body), Paragraph(currency(d["rate"]), body), Paragraph(currency(d["total_due"]), body)],
            ],
            colWidths=[3.4 * inch, 0.9 * inch, 0.8 * inch, 1.15 * inch],
        )
        summary.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16324f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [summary, Spacer(1, 14)]

        story.append(Paragraph("Client Time Tracking", heading))
        tracking_rows = [["Date", "Day", "Hours", "Rate", "Amount"]]
        for row in d["daily_time_rows"]:
            tracking_rows.append([
                Paragraph(escape(row["date"]), body),
                Paragraph(escape(row["day"]), body),
                Paragraph(_format_quantity(row["hours"]), body),
                Paragraph(currency(d["rate"]), body),
                Paragraph(currency(row["amount"]), body),
            ])

        tracking = Table(tracking_rows, colWidths=[1.05 * inch, 1.55 * inch, 0.8 * inch, 0.9 * inch, 1.05 * inch], repeatRows=1)
        tracking.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story += [tracking, Spacer(1, 14)]

        total = Table([[Paragraph("<b>Total Due</b>", body), Paragraph(f"<b>{escape(currency(d['total_due']))}</b>", body)]], colWidths=[1.5 * inch, 1.2 * inch], hAlign="RIGHT")
        total.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ecfeff")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#99f6e4")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story += [total, Spacer(1, 14)]

        story.append(Paragraph("Payment Terms", heading))
        story.append(Paragraph(escape(d["payment_terms"]), body))
        story.append(Spacer(1, 8))
        story.append(Paragraph("Payment Method", heading))
        story.append(Paragraph(escape(f"{d['payment_method']} ({d['payment_email']})"), body))
        story.append(Spacer(1, 8))
        story.append(Paragraph("Notes", heading))
        story.append(Paragraph(htmlize_multiline(d["notes"]), body))
        doc.build(story)

    def _preview_invoice(self):
        d = self._collect_invoice(False)
        if not d:
            return
        self._save_form_config()
        dlg = QDialog(self)
        dlg.setWindowTitle("Invoice Preview")
        dlg.resize(760, 700)
        v = QVBoxLayout(dlg)
        body = QTextEdit()
        body.setReadOnly(True)
        body.setPlainText(self._invoice_preview_text(d))
        v.addWidget(body)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.button(QDialogButtonBox.Close).clicked.connect(dlg.reject)
        v.addWidget(buttons)
        dlg.exec()

    def _save_pdf(self):
        d = self._collect_invoice(False)
        if not d:
            return
        self._save_form_config()
        default = self._default_pdf_path(d)
        out, _ = QFileDialog.getSaveFileName(self, "Save Invoice PDF", default, "PDF files (*.pdf)")
        if not out:
            return
        try:
            self._write_pdf(d, out)
            self.status_label.setText(f"Invoice PDF saved to {out}")
            QMessageBox.information(self, "PDF Saved", f"Invoice PDF saved successfully.\n\n{out}")
        except Exception as exc:
            QMessageBox.critical(self, "PDF Error", f"Could not create the invoice PDF.\n\n{exc}")

    def _send_invoice(self):
        d = self._collect_invoice(True)
        if not d:
            return
        self._save_form_config()
        self.btn_send_invoice.setEnabled(False)
        self.status_label.setText("Generating invoice PDF and sending email...")
        threading.Thread(target=self._do_send_invoice, args=(d,), daemon=True).start()

    def _do_send_invoice(self, d):
        try:
            pdf_path = self._default_pdf_path(d)
            self._write_pdf(d, pdf_path)
            with open(pdf_path, "rb") as f:
                payload = f.read()

            to_list = parse_email_list(INVOICE_TO_EMAILS)
            cc_list = parse_email_list(INVOICE_CC_EMAILS)
            plain, html = self._invoice_body(d)
            msg = self._build_email_message(
                to_list,
                cc_list,
                self._invoice_subject(d),
                plain,
                html,
                attachments=[(os.path.basename(pdf_path), payload, "pdf")],
            )
            self._send_message(msg, to_list + cc_list)
            self.invoice_done.emit(True, pdf_path, ", ".join(to_list) or "None", ", ".join(cc_list) or "None")
        except smtplib.SMTPAuthenticationError:
            self.invoice_done.emit(False, "Authentication failed. Gmail users should use an App Password.", "", "")
        except Exception as exc:
            self.invoice_done.emit(False, str(exc), "", "")

    def _on_invoice_done(self, ok, a, b, c):
        self.btn_send_invoice.setEnabled(True)
        if ok:
            self.status_label.setText("Invoice sent successfully.")
            QMessageBox.information(self, "Invoice Sent", f"Your invoice was sent successfully.\n\nTo: {b}\nCC: {c}\n\nPDF: {a}")
        else:
            self.status_label.setText("Send failed. See the error message for details.")
            QMessageBox.critical(self, "Send Failed", f"Could not send email.\n\n{a}")


def main():
    app = QApplication(sys.argv)
    win = MailerWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
