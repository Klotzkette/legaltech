"""
Word-Gliederungs-Retter – PyQt6 GUI

Design: identical colour palette and layout to Tom's Super Simple PDF Anonymizer.
Soft blue-teal tones, Swiss-style minimalism, Arial font, no bold text.
"""

import os
import subprocess
import sys
import traceback
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QDialog,
    QLineEdit,
    QGroupBox,
    QFormLayout,
    QFrame,
    QSizePolicy,
    QProgressBar,
)
from PyQt6.QtCore import (
    Qt,
    QThread,
    pyqtSignal,
    QSettings,
    QTimer,
    QPropertyAnimation,
    QEasingCurve,
)
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QPalette, QColor, QMouseEvent
from PyQt6.QtWidgets import QGraphicsDropShadowEffect

try:
    from word_processor import process_document, SUPPORTED_EXTENSIONS
    from ai_engine import AIEngine
except ImportError as _imp_err:
    _import_error = _imp_err
    SUPPORTED_EXTENSIONS = {".doc", ".docx"}

    class AIEngine:  # type: ignore
        def __init__(self, api_key):
            self.api_key = api_key
else:
    _import_error = None

# ---------------------------------------------------------------------------
# Processing modes
# ---------------------------------------------------------------------------

MODE_DIRECT = "direct"           # clean output, no track changes
MODE_TRACK  = "track_changes"    # output with Word revision marks

# ---------------------------------------------------------------------------
# Settings helpers
# ---------------------------------------------------------------------------

SETTINGS_ORG = "toms_super_simple_word_gliederungsretter"
SETTINGS_APP = "toms_super_simple_word_gliederungsretter"


def _settings() -> QSettings:
    return QSettings(SETTINGS_ORG, SETTINGS_APP)


def save_api_key(key: str):
    _settings().setValue("api_key", key)


def load_api_key() -> str:
    return _settings().value("api_key", "")


def save_output_dir(path: str):
    _settings().setValue("output_dir", path)


def load_output_dir() -> str:
    return _settings().value("output_dir", "")


def save_mode(mode: str):
    _settings().setValue("processing_mode", mode)


def load_mode() -> str:
    return _settings().value("processing_mode", MODE_DIRECT)


# ---------------------------------------------------------------------------
# Colour palette  (identical to PDF Anonymizer)
# ---------------------------------------------------------------------------

BG_DARK        = "#EEF2F9"
BG_CARD        = "#F7F9FC"
BG_SURFACE     = "#E2E9F3"
BG_HOVER       = "#D4DEEE"

ACCENT         = "#2563EB"
ACCENT_HOVER   = "#1D4ED8"
ACCENT_SOFT    = "#6B9AE8"
ACCENT_GLOW    = "#93B5F5"

BORDER         = "#C8D6EA"
BORDER_FOCUS   = "#3B7DD8"

TEXT_PRIMARY   = "#0F172A"
TEXT_SECONDARY = "#475569"
TEXT_MUTED     = "#94A3B8"

SUCCESS        = "#16A34A"
SUCCESS_BG     = "#F0FDF4"
SUCCESS_BORDER = "#86EFAC"

ERROR          = "#DC2626"
ERROR_BG       = "#FEF2F2"
ERROR_BORDER   = "#FCA5A5"

# ---------------------------------------------------------------------------
# Stylesheet  (identical structure to PDF Anonymizer)
# ---------------------------------------------------------------------------

STYLESHEET = f"""
* {{
    font-family: "SF Pro Display", "Segoe UI", "Helvetica Neue", "Arial", sans-serif;
    font-weight: normal;
}}
QMainWindow {{
    background-color: {BG_DARK};
}}
QWidget#centralWidget {{
    background-color: {BG_DARK};
}}
QLabel {{
    color: {TEXT_SECONDARY};
    font-size: 13px;
    background: transparent;
}}
QLabel#titleLabel {{
    color: {TEXT_PRIMARY};
    font-size: 22px;
    letter-spacing: -0.3px;
}}
QLabel#titleAccent {{
    color: {ACCENT};
    font-size: 22px;
    letter-spacing: -0.3px;
}}
QLabel#subtitleLabel {{
    color: {TEXT_MUTED};
    font-size: 12px;
    line-height: 1.6;
}}
QLabel#dropIcon {{
    font-size: 48px;
    background: transparent;
}}
QLabel#dropLabel {{
    color: {TEXT_PRIMARY};
    font-size: 15px;
}}
QLabel#dropHint {{
    color: {TEXT_MUTED};
    font-size: 12px;
}}
QLabel#fileLabel {{
    color: {TEXT_SECONDARY};
    font-size: 12px;
    padding: 4px 0px;
}}
QLabel#stepLabel {{
    color: {ACCENT};
    font-size: 12px;
}}
QLabel#modelPill {{
    color: {ACCENT};
    background-color: {BG_CARD};
    border: 1px solid {ACCENT_GLOW};
    border-radius: 14px;
    padding: 4px 12px;
    font-size: 11px;
}}
QPushButton {{
    background-color: {ACCENT};
    color: #FFFFFF;
    border: none;
    border-radius: 16px;
    padding: 10px 28px;
    font-size: 13px;
}}
QPushButton:hover {{
    background-color: {ACCENT_HOVER};
}}
QPushButton:pressed {{
    background-color: {ACCENT_SOFT};
}}
QPushButton:disabled {{
    background-color: {BG_SURFACE};
    color: {TEXT_MUTED};
    border-radius: 16px;
}}
QPushButton#settingsBtn {{
    background-color: transparent;
    color: {TEXT_MUTED};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 6px 16px;
    font-size: 11px;
}}
QPushButton#settingsBtn:hover {{
    color: {ACCENT};
    background-color: {BG_CARD};
    border-color: {ACCENT_GLOW};
}}
QPushButton#selectBtn {{
    background-color: {BG_CARD};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 9px 22px;
    font-size: 13px;
}}
QPushButton#selectBtn:hover {{
    background-color: {BG_HOVER};
    border-color: {ACCENT_SOFT};
}}
QPushButton#selectBtn:disabled {{
    color: {TEXT_MUTED};
    border-color: {BORDER};
    background-color: {BG_SURFACE};
}}
QPushButton#openFolderBtn {{
    background-color: {SUCCESS_BG};
    color: {SUCCESS};
    border: 1px solid {SUCCESS_BORDER};
    border-radius: 14px;
    padding: 6px 16px;
    font-size: 12px;
}}
QPushButton#openFolderBtn:hover {{
    background-color: #DCFCE7;
    border-color: {SUCCESS};
}}
QProgressBar {{
    border: none;
    border-radius: 10px;
    text-align: center;
    color: #FFFFFF;
    background-color: {BG_SURFACE};
    min-height: 20px;
    max-height: 20px;
    font-size: 10px;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {ACCENT}, stop:0.5 {ACCENT_SOFT}, stop:1 {ACCENT});
    border-radius: 10px;
}}
QLineEdit {{
    background-color: #FFFFFF;
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 12px;
    padding: 10px 14px;
    font-size: 13px;
}}
QLineEdit:focus {{
    border-color: {ACCENT};
}}
QLineEdit[valid="true"] {{
    border-color: {SUCCESS};
}}
QGroupBox {{
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 16px;
    margin-top: 14px;
    padding: 22px 18px 14px 18px;
    font-size: 13px;
    background-color: {BG_CARD};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 18px;
    padding: 0 10px;
    color: {TEXT_SECONDARY};
    font-size: 12px;
}}
QDialog {{
    background-color: {BG_DARK};
}}
QStatusBar {{
    background-color: {BG_CARD};
    color: {TEXT_MUTED};
    font-size: 10px;
    border-top: 1px solid {BORDER};
    padding: 3px 12px;
}}
QFrame#dropZone {{
    background-color: {BG_CARD};
    border: 2px dashed {BORDER};
    border-radius: 24px;
}}
QFrame#dropZone:hover {{
    border-color: {ACCENT_SOFT};
    background-color: #FFFFFF;
}}
QFrame#dropZone[dragOver="true"] {{
    border-color: {ACCENT};
    border-style: solid;
    border-width: 2px;
    background-color: #EBF4FF;
}}
QFrame#dropZone[processing="true"] {{
    border-color: {ACCENT_GLOW};
    border-style: solid;
    background-color: {BG_CARD};
}}
QFrame#dropZone[success="true"] {{
    border-color: {SUCCESS_BORDER};
    border-style: solid;
    background-color: {SUCCESS_BG};
}}
QFrame#dropZone[error="true"] {{
    border-color: {ERROR_BORDER};
    border-style: solid;
    background-color: {ERROR_BG};
}}
QScrollBar:vertical {{
    background-color: transparent;
    width: 6px;
    margin: 4px 2px;
}}
QScrollBar::handle:vertical {{
    background-color: {BG_HOVER};
    border-radius: 3px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background-color: {ACCENT_SOFT};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QToolTip {{
    background-color: {TEXT_PRIMARY};
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 11px;
}}
"""

# ---------------------------------------------------------------------------
# Drop zone
# ---------------------------------------------------------------------------

_ACCEPTED_EXT = tuple(SUPPORTED_EXTENSIONS)


class DropZone(QFrame):
    """Drag-and-drop zone accepting .doc and .docx files."""

    file_dropped = pyqtSignal(str)
    clicked = pyqtSignal()

    STATE_IDLE       = "idle"
    STATE_PROCESSING = "processing"
    STATE_SUCCESS    = "success"
    STATE_ERROR      = "error"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        self.setMinimumHeight(220)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._state = self.STATE_IDLE

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(32)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 22))
        self.setGraphicsEffect(shadow)
        self._shadow = shadow

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(8)

        self.icon_label = QLabel()
        self.icon_label.setObjectName("dropIcon")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        self.primary_label = QLabel()
        self.primary_label.setObjectName("dropLabel")
        self.primary_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.primary_label)

        self.secondary_label = QLabel()
        self.secondary_label.setObjectName("dropHint")
        self.secondary_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.secondary_label)

        self.step_label = QLabel()
        self.step_label.setObjectName("stepLabel")
        self.step_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.step_label.setVisible(False)
        layout.addWidget(self.step_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(20)
        self.progress_bar.setMinimumWidth(320)
        self.progress_bar.setMaximumWidth(400)
        self.progress_bar.setFormat("%p %")
        layout.addWidget(self.progress_bar, alignment=Qt.AlignmentFlag.AlignCenter)

        # Idle shadow pulse
        self._pulse_anim = QPropertyAnimation(shadow, b"blurRadius")
        self._pulse_anim.setDuration(2400)
        self._pulse_anim.setStartValue(24)
        self._pulse_anim.setEndValue(40)
        self._pulse_anim.setEasingCurve(QEasingCurve.Type.InOutSine)
        self._pulse_anim.setLoopCount(-1)

        self.set_state(self.STATE_IDLE)

    def set_state(self, state: str, detail: str = ""):
        self._state = state
        for prop in ("dragOver", "processing", "success", "error"):
            self.setProperty(prop, False)
        self._pulse_anim.stop()
        self.primary_label.setStyleSheet("")

        if state == self.STATE_IDLE:
            self.icon_label.setText("\u2B06")
            self.primary_label.setText("Word-Dokument hier ablegen")
            self.secondary_label.setText("DOC oder DOCX \u2013 oder klicken zum Ausw\u00e4hlen")
            self.secondary_label.setVisible(True)
            self.step_label.setVisible(False)
            self.progress_bar.setVisible(False)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            self.setAcceptDrops(True)
            self._shadow.setColor(QColor(37, 99, 235, 18))
            self._pulse_anim.start()

        elif state == self.STATE_PROCESSING:
            self.setProperty("processing", True)
            self._shadow.setColor(QColor(37, 99, 235, 30))
            self._shadow.setBlurRadius(28)
            self.icon_label.setText("\u2699")
            self.primary_label.setText("Wird verarbeitet \u2026")
            self.secondary_label.setVisible(False)
            self.step_label.setVisible(True)
            self.step_label.setText(detail or "Initialisiere \u2026")
            self.progress_bar.setVisible(True)
            self.setCursor(Qt.CursorShape.WaitCursor)
            self.setAcceptDrops(False)

        elif state == self.STATE_SUCCESS:
            self.setProperty("success", True)
            self._shadow.setColor(QColor(22, 163, 74, 30))
            self._shadow.setBlurRadius(32)
            self.icon_label.setText("\u2713")
            self.primary_label.setText("Gliederung standardisiert")
            self.primary_label.setStyleSheet(f"color: {SUCCESS}; font-size: 15px;")
            self.secondary_label.setText(detail or "")
            self.secondary_label.setVisible(bool(detail))
            self.step_label.setVisible(False)
            self.progress_bar.setVisible(False)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            self.setAcceptDrops(True)

        elif state == self.STATE_ERROR:
            self.setProperty("error", True)
            self._shadow.setColor(QColor(220, 38, 38, 25))
            self._shadow.setBlurRadius(28)
            self.icon_label.setText("\u2717")
            self.primary_label.setText("Fehler aufgetreten")
            self.primary_label.setStyleSheet(f"color: {ERROR}; font-size: 15px;")
            self.secondary_label.setText(
                "Klicken oder neue Datei ablegen, um es erneut zu versuchen"
            )
            self.secondary_label.setVisible(True)
            self.step_label.setVisible(False)
            self.progress_bar.setVisible(False)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            self.setAcceptDrops(True)

        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def set_progress(self, value: int):
        self.progress_bar.setValue(value)

    def set_step(self, text: str):
        self.step_label.setText(text)

    # ── Events ──────────────────────────────────────────────────────────────

    def mousePressEvent(self, event: QMouseEvent):
        if self._state != self.STATE_PROCESSING:
            self.clicked.emit()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if Path(url.toLocalFile()).suffix.lower() in _ACCEPTED_EXT:
                    event.acceptProposedAction()
                    self.setProperty("dragOver", True)
                    self.style().unpolish(self)
                    self.style().polish(self)
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event: QDropEvent):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if Path(path).suffix.lower() in _ACCEPTED_EXT:
                self.file_dropped.emit(path)
                return


# ---------------------------------------------------------------------------
# Worker thread
# ---------------------------------------------------------------------------

class ProcessWorker(QThread):
    progress = pyqtSignal(int)
    step     = pyqtSignal(str)
    finished_ok  = pyqtSignal(str)   # output path
    finished_err = pyqtSignal(str)   # error message

    def __init__(
        self,
        input_path: str,
        output_path: str,
        track_changes: bool,
        api_key: str,
    ):
        super().__init__()
        self.input_path   = input_path
        self.output_path  = output_path
        self.track_changes = track_changes
        self.api_key       = api_key

    def run(self):
        try:
            ai_engine = AIEngine(self.api_key) if self.api_key else None

            step_values = {
                "Schritt 1/4 – Datei vorbereiten …":        10,
                "Schritt 2/4 – Überschriften analysieren …": 35,
                "Schritt 3/4 – Gliederung standardisieren …": 70,
                "Schritt 4/4 – Nummerierung einrichten …":   90,
            }

            def on_progress(msg: str):
                self.step.emit(msg)
                self.progress.emit(step_values.get(msg, 50))

            process_document(
                input_path=self.input_path,
                output_path=self.output_path,
                track_changes=self.track_changes,
                ai_engine=ai_engine,
                progress_callback=on_progress,
            )
            self.progress.emit(100)
            self.finished_ok.emit(self.output_path)

        except Exception as e:
            self.finished_err.emit(f"{e}\n\n{traceback.format_exc()}")


# ---------------------------------------------------------------------------
# Settings dialog
# ---------------------------------------------------------------------------

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Einstellungen")
        self.setMinimumWidth(480)
        self.setStyleSheet(STYLESHEET)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(32, 28, 32, 24)

        header = QLabel("\u2699  Einstellungen")
        header.setStyleSheet(f"color: {TEXT_PRIMARY}; font-size: 18px; letter-spacing: -0.3px;")
        layout.addWidget(header)

        desc = QLabel(
            "OpenAI API-Key f\u00fcr KI-gest\u00fctzte \u00dcberschriften-Erkennung (optional).\n"
            "Ohne Key wird die Erkennung \u00fcber Formatierung und Stilnamen durchgef\u00fchrt."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        layout.addWidget(desc)

        layout.addSpacing(8)

        model_label = QLabel("\u2728  Modell: GPT-5.2")
        model_label.setStyleSheet(
            f"color: {ACCENT}; font-size: 12px; "
            f"background-color: #EBF4FF; "
            f"border: 1px solid {ACCENT_GLOW}; "
            f"border-radius: 10px; padding: 8px 14px;"
        )
        layout.addWidget(model_label)

        layout.addSpacing(4)

        keys_group = QGroupBox("API-Key (OpenAI)")
        keys_layout = QFormLayout(keys_group)
        keys_layout.setSpacing(10)

        self.key_field = QLineEdit(load_api_key())
        self.key_field.setPlaceholderText("sk-…  (leer lassen für Offline-Modus)")
        self.key_field.setEchoMode(QLineEdit.EchoMode.Password)
        self.key_field.setMinimumHeight(38)
        self.key_field.textChanged.connect(self._on_key_changed)
        keys_layout.addRow("OpenAI:", self.key_field)
        layout.addWidget(keys_group)
        layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Abbrechen")
        cancel_btn.setObjectName("selectBtn")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("\u2713  Speichern")
        save_btn.setMinimumWidth(130)
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

    def _on_key_changed(self):
        has_value = bool(self.key_field.text().strip())
        self.key_field.setProperty("valid", has_value)
        self.key_field.style().unpolish(self.key_field)
        self.key_field.style().polish(self.key_field)

    def _save(self):
        save_api_key(self.key_field.text().strip())
        self.accept()


# ---------------------------------------------------------------------------
# Mode selection dialog  (shown after file selection)
# ---------------------------------------------------------------------------

_MODE_OPTIONS = [
    (
        MODE_DIRECT,
        "Direkt",
        "\u2588\u2588",   # solid blocks icon
        "Alle \u00dcberschriften werden direkt\n"
        "auf 1. / 1.1 / 1.1.1 \u2026 umgestellt.",
    ),
    (
        MODE_TRACK,
        "\u00c4nderungsmodus",
        "\u21C4",         # arrows icon
        "Die Umstellung wird als \u00c4nderungsverfolgung\n"
        "in Word angezeigt (Track Changes).",
    ),
]


class ModeSelectionDialog(QDialog):
    """Shown after file selection; user picks Direct or Track Changes."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Verarbeitungsmodus")
        self.setFixedWidth(480)
        self.setStyleSheet(STYLESHEET)
        self.selected_mode: str | None = None

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(32, 28, 32, 24)

        header = QLabel("Verarbeitungsmodus")
        header.setStyleSheet(
            f"color: {TEXT_PRIMARY}; font-size: 17px; letter-spacing: -0.2px;"
        )
        layout.addWidget(header)

        sub = QLabel("W\u00e4hlen Sie, wie die Gliederung gespeichert werden soll:")
        sub.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        layout.addWidget(sub)
        layout.addSpacing(8)

        saved_mode = load_mode()

        for mode_key, title, icon, desc in _MODE_OPTIONS:
            is_saved = mode_key == saved_mode
            card = QFrame()
            card.setCursor(Qt.CursorShape.PointingHandCursor)
            border_col = ACCENT if is_saved else BORDER
            bg_col = "#EBF4FF" if is_saved else BG_CARD
            bw = "2" if is_saved else "1"
            card.setStyleSheet(f"""
                QFrame {{
                    background-color: {bg_col};
                    border: {bw}px solid {border_col};
                    border-radius: 14px;
                    padding: 14px 18px;
                }}
                QFrame:hover {{
                    border-color: {ACCENT};
                    background-color: #EBF4FF;
                }}
            """)

            card_shadow = QGraphicsDropShadowEffect(card)
            card_shadow.setBlurRadius(16)
            card_shadow.setOffset(0, 2)
            card_shadow.setColor(QColor(0, 0, 0, 15))
            card.setGraphicsEffect(card_shadow)

            card_layout = QHBoxLayout(card)
            card_layout.setSpacing(14)
            card_layout.setContentsMargins(0, 0, 0, 0)

            icon_label = QLabel(icon)
            icon_label.setStyleSheet(
                f"color: {ACCENT}; font-size: 20px; border: none; background: transparent;"
            )
            icon_label.setFixedWidth(32)
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            card_layout.addWidget(icon_label)

            text_layout = QVBoxLayout()
            text_layout.setSpacing(2)

            title_label = QLabel(title)
            title_label.setStyleSheet(
                f"color: {TEXT_PRIMARY}; font-size: 14px; border: none; background: transparent;"
            )
            text_layout.addWidget(title_label)

            desc_label = QLabel(desc)
            desc_label.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 11px; border: none; background: transparent;"
            )
            desc_label.setWordWrap(True)
            text_layout.addWidget(desc_label)

            card_layout.addLayout(text_layout, stretch=1)
            card.mousePressEvent = lambda _event, m=mode_key: self._select(m)
            layout.addWidget(card)

        layout.addSpacing(4)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        cancel_btn = QPushButton("Abbrechen")
        cancel_btn.setObjectName("selectBtn")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def _select(self, mode: str):
        self.selected_mode = mode
        save_mode(mode)
        self.accept()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    _file_counter = 0

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tom's Super Simple Word-Gliederungs-Retter")
        self.setMinimumSize(620, 520)
        self.resize(700, 600)
        self.setStyleSheet(STYLESHEET)

        self.worker: ProcessWorker | None = None
        self.current_file: str | None = None
        self._last_output: str | None = None
        self._selected_mode: str = MODE_DIRECT

        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(32, 24, 32, 16)
        main_layout.setSpacing(0)

        # ── Header ──────────────────────────────────────────────────────────
        header = QHBoxLayout()
        header.setSpacing(8)

        title_layout = QHBoxLayout()
        title_layout.setSpacing(0)
        t1 = QLabel("Tom\u2019s Super Simple ")
        t1.setObjectName("titleLabel")
        title_layout.addWidget(t1)
        t2 = QLabel("Word-Gliederungs-Retter")
        t2.setObjectName("titleAccent")
        title_layout.addWidget(t2)
        title_layout.addStretch()
        header.addLayout(title_layout, stretch=1)

        self.model_pill = QLabel("GPT-5.2")
        self.model_pill.setObjectName("modelPill")
        self._update_model_pill()
        header.addWidget(self.model_pill, alignment=Qt.AlignmentFlag.AlignVCenter)

        header.addSpacing(6)

        settings_btn = QPushButton("\u2699  Einstellungen")
        settings_btn.setObjectName("settingsBtn")
        settings_btn.clicked.connect(self.open_settings)
        header.addWidget(settings_btn, alignment=Qt.AlignmentFlag.AlignVCenter)

        main_layout.addLayout(header)
        main_layout.addSpacing(2)

        # Subtitle
        subtitle = QLabel(
            "Automatische Gliederungsstandardisierung f\u00fcr Word-Dokumente  \u00b7  "
            "DOC und DOCX  \u00b7  1. / 1.1 / 1.1.1 \u2026 mit automatischer Fortsetzung"
        )
        subtitle.setObjectName("subtitleLabel")
        subtitle.setWordWrap(True)
        main_layout.addWidget(subtitle)
        main_layout.addSpacing(14)

        # ── Drop zone ────────────────────────────────────────────────────────
        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self.on_file_selected)
        self.drop_zone.clicked.connect(self.browse_file)
        main_layout.addWidget(self.drop_zone, stretch=1)
        main_layout.addSpacing(12)

        # ── Bottom bar ───────────────────────────────────────────────────────
        bottom = QHBoxLayout()
        bottom.setSpacing(10)

        self.file_label = QLabel("")
        self.file_label.setObjectName("fileLabel")
        self.file_label.setWordWrap(True)
        bottom.addWidget(self.file_label, stretch=1)

        self.open_folder_btn = QPushButton("\u2197  Ordner \u00f6ffnen")
        self.open_folder_btn.setObjectName("openFolderBtn")
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_output_folder)
        bottom.addWidget(self.open_folder_btn)

        self.select_btn = QPushButton("Datei ausw\u00e4hlen")
        self.select_btn.setObjectName("selectBtn")
        self.select_btn.clicked.connect(self.browse_file)
        bottom.addWidget(self.select_btn)

        main_layout.addLayout(bottom)
        main_layout.addSpacing(2)

        # ── Status bar ───────────────────────────────────────────────────────
        self._update_statusbar_idle()

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _update_model_pill(self):
        has_key = bool(load_api_key())
        if has_key:
            self.model_pill.setText("GPT-5.2  \u00b7  KI aktiv")
            self.model_pill.setStyleSheet("")
        else:
            self.model_pill.setText("Offline-Modus")
            self.model_pill.setStyleSheet(
                f"color: {TEXT_SECONDARY}; background-color: {BG_SURFACE}; "
                f"border: 1px solid {BORDER}; "
                f"border-radius: 14px; padding: 4px 12px; font-size: 11px;"
            )

    def _update_statusbar_idle(self):
        has_key = bool(load_api_key())
        if has_key:
            self.statusBar().showMessage(
                "Bereit  \u00b7  Word-Datei ablegen oder ausw\u00e4hlen  \u00b7  v1.0"
            )
        else:
            self.statusBar().showMessage(
                "Offline-Modus (kein API-Key)  \u00b7  "
                "Erkennung via Stilnamen & Formatierung  \u00b7  v1.0"
            )

    @classmethod
    def _output_filename(cls, mode: str) -> str:
        from datetime import date
        cls._file_counter += 1
        today = date.today().strftime("%Y%m%d")
        suffix = "Gliederung" if mode == MODE_DIRECT else "Gliederung_Aenderungsmodus"
        return f"{today}_Dokument_{suffix}_{cls._file_counter:03d}.docx"

    def _set_processing(self, active: bool):
        self.select_btn.setEnabled(not active)

    def _open_output_folder(self):
        if self._last_output:
            self._open_path(os.path.dirname(self._last_output))

    @staticmethod
    def _open_path(path: str):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass

    # ── Slots ────────────────────────────────────────────────────────────────

    def open_settings(self):
        dlg = SettingsDialog(self)
        if dlg.exec():
            self._update_model_pill()
            self._update_statusbar_idle()

    def browse_file(self):
        if self.worker and self.worker.isRunning():
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Word-Datei ausw\u00e4hlen",
            "",
            "Word-Dokumente (*.docx *.doc);;Alle Dateien (*)",
        )
        if path:
            self.on_file_selected(path)

    def on_file_selected(self, path: str):
        if self.worker and self.worker.isRunning():
            return
        self.current_file = path
        self._last_output = None
        self.open_folder_btn.setVisible(False)
        self.file_label.setText(os.path.basename(path))
        self.statusBar().showMessage(f"Geladen: {os.path.basename(path)}")
        self._start_processing()

    def _start_processing(self):
        if not self.current_file:
            return

        # Mode selection
        mode_dlg = ModeSelectionDialog(self)
        if not mode_dlg.exec():
            self.drop_zone.set_state(DropZone.STATE_IDLE)
            return
        mode = mode_dlg.selected_mode
        self._selected_mode = mode

        # Output location
        default_dir = load_output_dir() or os.path.dirname(self.current_file)
        default_name = self._output_filename(mode)
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Ausgabe speichern unter",
            os.path.join(default_dir, default_name),
            "Word-Dokumente (*.docx)",
        )
        if not output_path:
            self.drop_zone.set_state(DropZone.STATE_IDLE)
            return
        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        save_output_dir(os.path.dirname(output_path))

        # Launch processing
        self._set_processing(True)
        self.drop_zone.set_state(DropZone.STATE_PROCESSING, "Initialisiere \u2026")

        api_key = load_api_key()
        track_changes = mode == MODE_TRACK

        self.worker = ProcessWorker(
            input_path=self.current_file,
            output_path=output_path,
            track_changes=track_changes,
            api_key=api_key,
        )
        self.worker.progress.connect(self.drop_zone.set_progress)
        self.worker.step.connect(self.drop_zone.set_step)
        self.worker.step.connect(lambda s: self.statusBar().showMessage(s))
        self.worker.finished_ok.connect(self.on_success)
        self.worker.finished_err.connect(self.on_error)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    def on_success(self, output_path: str):
        self._set_processing(False)
        self._last_output = output_path
        out_name = os.path.basename(output_path)
        mode_label = (
            "direkt gespeichert" if self._selected_mode == MODE_DIRECT
            else "mit \u00c4nderungsmodus gespeichert"
        )
        detail = f"Gliederung standardisiert \u2013 {mode_label}  \u2192  {out_name}"
        self.drop_zone.set_state(DropZone.STATE_SUCCESS, detail)
        self.file_label.setText(out_name)
        self.open_folder_btn.setVisible(True)
        self.statusBar().showMessage(f"Gespeichert: {output_path}")
        self._open_path(output_path)
        QTimer.singleShot(8000, self._reset_to_idle)

    def on_error(self, msg: str):
        self._set_processing(False)
        self.drop_zone.set_state(DropZone.STATE_ERROR)
        self.statusBar().showMessage("Fehler bei der Verarbeitung")
        QMessageBox.critical(
            self,
            "Fehler",
            f"Bei der Verarbeitung ist ein Fehler aufgetreten:\n\n{msg}",
        )
        QTimer.singleShot(500, self._reset_to_idle)

    def _reset_to_idle(self):
        if self.worker and self.worker.isRunning():
            return
        if self.drop_zone._state in (DropZone.STATE_SUCCESS, DropZone.STATE_ERROR):
            self.drop_zone.set_state(DropZone.STATE_IDLE)
            self.file_label.setText("")
            self.current_file = None
            self._update_statusbar_idle()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _check_dependencies() -> str | None:
    missing = []
    try:
        import docx  # noqa: F401
    except ImportError:
        missing.append("python-docx  (pip install python-docx)")
    try:
        from lxml import etree  # noqa: F401
    except ImportError:
        missing.append("lxml  (pip install lxml)")
    try:
        import openai  # noqa: F401
    except ImportError:
        missing.append("openai  (pip install openai)  – optional, für KI-Modus")
    if _import_error is not None and not missing:
        missing.append(str(_import_error))
    hard = [m for m in missing if "optional" not in m]
    if hard:
        return "Fehlende Pflicht-Pakete:\n\n" + "\n".join(f"  •  {m}" for m in hard)
    return None


def run_app():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    dep_err = _check_dependencies()
    if dep_err:
        QMessageBox.critical(
            None,
            "Fehlende Pakete",
            f"{dep_err}\n\nBitte installieren:\n  pip install -r requirements.txt",
        )
        sys.exit(1)

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor(BG_DARK))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.Base,            QColor(BG_CARD))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor(BG_SURFACE))
    palette.setColor(QPalette.ColorRole.Text,            QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.Button,          QColor(BG_SURFACE))
    palette.setColor(QPalette.ColorRole.ButtonText,      QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor(ACCENT))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(TEXT_MUTED))
    app.setPalette(palette)

    font = QFont("Arial")
    font.setPointSize(10)
    font.setWeight(QFont.Weight.Normal)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
