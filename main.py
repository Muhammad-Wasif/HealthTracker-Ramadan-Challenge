
import sys
import os
import datetime
import logging
import shutil
import csv
from typing import Optional, Tuple, Dict, Any

import bcrypt
import numpy as np
import pandas as pd
from openpyxl import Workbook

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QFormLayout, QLabel, QPushButton, QLineEdit,
    QSpinBox, QDoubleSpinBox, QCheckBox, QScrollArea, QFrame,
    QDialog, QTabWidget, QMessageBox, QSizePolicy, QStackedWidget,
    QComboBox, QFileDialog, QGroupBox, QSplitter
)
from PySide6.QtCore import (
    Qt, Signal, QObject, QTimer, QPropertyAnimation, QEasingCurve,
    QSize, QThread, QDate
)
from PySide6.QtGui import (
    QFont, QFontDatabase, QPalette, QColor, QIcon, QPixmap,
    QLinearGradient, QPainter
)

import matplotlib
matplotlib.use("QtAgg")
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.animation as animation

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("HealthTracker")

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_backup")

USERS_FILE       = os.path.join(DATA_DIR, "users.xlsx")
PROFILES_FILE    = os.path.join(DATA_DIR, "profiles.xlsx")
DAILY_INDEX_FILE = os.path.join(DATA_DIR, "daily_index.xlsx")
AVOID_LOGS_FILE  = os.path.join(DATA_DIR, "avoid_logs.xlsx")
HEALTH_LOGS_FILE = os.path.join(DATA_DIR, "health_logs.xlsx")
WEIGHT_HIST_FILE = os.path.join(DATA_DIR, "weight_history.xlsx")

DARK_THEME = """
QWidget { background-color: #ffffff; color: #000000; font-family: 'Segoe UI', sans-serif; font-size: 13px; }
QMainWindow { background-color: #ffffff; }
QGroupBox { border: 1px solid #000000; border-radius: 8px; margin-top: 12px; padding: 8px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; color: #0000ff; font-weight: bold; }
QPushButton { background-color: #0000ff; color: #ffffff; border-radius: 6px; padding: 8px 16px; font-weight: bold; border: none; }
QPushButton:hover { background-color: #00008b; }
QPushButton:pressed { background-color: #0000cd; }
QPushButton#danger { background-color: #ff0000; }
QPushButton#danger:hover { background-color: #8b0000; }
QPushButton#success { background-color: #0000ff; color: #ffffff; }
QPushButton#success:hover { background-color: #00008b; }
QPushButton#secondary { background-color: #ffd700; color: #000000; }
QPushButton#secondary:hover { background-color: #b8860b; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { background-color: #ffffff; border: 1px solid #000000; border-radius: 6px; padding: 6px 10px; color: #000000; }
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus { border: 1px solid #0000ff; }
QCheckBox { spacing: 8px; }
QCheckBox::indicator { width: 18px; height: 18px; border-radius: 4px; border: 2px solid #000000; background: #ffffff; }
QCheckBox::indicator:checked { background-color: #0000ff; border-color: #0000ff; }
QScrollArea { border: none; }
QLabel#title { font-size: 28px; font-weight: bold; color: #0000ff; }
QLabel#subtitle { font-size: 14px; color: #000000; }
QLabel#bmi_normal { color: #0000ff; font-weight: bold; font-size: 15px; }
QLabel#bmi_underweight { color: #0000ff; font-weight: bold; font-size: 15px; }
QLabel#bmi_overweight { color: #ffd700; font-weight: bold; font-size: 15px; }
QLabel#bmi_obese { color: #ff0000; font-weight: bold; font-size: 15px; }
QTabWidget::pane { border: 1px solid #000000; border-radius: 6px; }
QTabBar::tab { background: #ffffff; color: #000000; padding: 8px 16px; border-radius: 4px; margin: 2px; border: 1px solid #000000; }
QTabBar::tab:selected { background: #0000ff; color: #ffffff; }
QFrame#card { background-color: #ffffff; border: 1px solid #000000; border-radius: 10px; padding: 12px; }
"""

LIGHT_THEME = DARK_THEME

# ─────────────────────────────────────────────
# DateManager
# ─────────────────────────────────────────────
class DateManager:
    @staticmethod
    def get_today() -> datetime.date:
        return datetime.date.today()

    @staticmethod
    def get_or_create_day_number(user_id: str, date: datetime.date, df: pd.DataFrame) -> Tuple[int, pd.DataFrame]:
        date_str = str(date)
        user_rows = df[df["user_id"] == user_id]
        existing = user_rows[user_rows["date"] == date_str]
        if not existing.empty:
            return int(existing.iloc[0]["day_number"]), df
        # New date: find max day_number for this user
        if user_rows.empty:
            day_number = 1
        else:
            day_number = int(user_rows["day_number"].max()) + 1
        new_row = pd.DataFrame([{"user_id": user_id, "date": date_str, "day_number": day_number}])
        df = pd.concat([df, new_row], ignore_index=True)
        return day_number, df


# ─────────────────────────────────────────────
# BMIManager
# ─────────────────────────────────────────────
class BMIManager:
    @staticmethod
    def calculate(weight_kg: float, height_cm: float) -> float:
        if height_cm <= 0:
            return 0.0
        h = height_cm / 100.0
        return round(weight_kg / (h * h), 2)

    @staticmethod
    def categorize(bmi: float) -> str:
        if bmi < 18.5:
            return "Underweight"
        elif bmi < 25.0:
            return "Normal"
        elif bmi < 30.0:
            return "Overweight"
        else:
            return "Obese"

    @staticmethod
    def color_object_name(category: str) -> str:
        mapping = {
            "Underweight": "bmi_underweight",
            "Normal": "bmi_normal",
            "Overweight": "bmi_overweight",
            "Obese": "bmi_obese",
        }
        return mapping.get(category, "bmi_normal")


# ─────────────────────────────────────────────
# ExcelUserManager
# ─────────────────────────────────────────────
class ExcelUserManager:
    COLUMNS = ["user_id", "username", "hashed_password", "created_date"]

    def __init__(self):
        os.makedirs(DATA_DIR, exist_ok=True)
        self._ensure_file()

    def _ensure_file(self):
        if not os.path.exists(USERS_FILE):
            df = pd.DataFrame(columns=self.COLUMNS)
            df.to_excel(USERS_FILE, sheet_name="Users", index=False)
            logger.info("Created users.xlsx")

    def _load(self) -> pd.DataFrame:
        try:
            return pd.read_excel(USERS_FILE, sheet_name="Users", dtype=str)
        except Exception:
            return pd.DataFrame(columns=self.COLUMNS)

    def _save(self, df: pd.DataFrame):
        df.to_excel(USERS_FILE, sheet_name="Users", index=False)

    def register(self, username: str, password: str) -> Optional[str]:
        df = self._load()
        if not df[df["username"] == username].empty:
            return None  # username taken
        user_id = f"u{len(df)+1:04d}"
        hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        new_row = pd.DataFrame([{
            "user_id": user_id,
            "username": username,
            "hashed_password": hashed,
            "created_date": str(datetime.date.today())
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        self._save(df)
        logger.info(f"Registered user: {username} ({user_id})")
        return user_id

    def authenticate(self, username: str, password: str) -> Optional[str]:
        df = self._load()
        row = df[df["username"] == username]
        if row.empty:
            return None
        stored_hash = row.iloc[0]["hashed_password"]
        if bcrypt.checkpw(password.encode(), stored_hash.encode()):
            return str(row.iloc[0]["user_id"])
        return None

    def username_exists(self, username: str) -> bool:
        df = self._load()
        return not df[df["username"] == username].empty


# ─────────────────────────────────────────────
# ExcelLogManager
# ─────────────────────────────────────────────
class ExcelLogManager:
    PROFILE_COLS   = ["user_id","age","gender","height_cm","current_weight_kg","bmi","bmi_category","last_updated"]
    INDEX_COLS     = ["user_id","date","day_number"]
    AVOID_COLS     = ["user_id","date","day_number","sugar_servings","sweet_drinks_count","caffeine_cups",
                      "fried_food","excess_oil","overeating","late_screen"]
    HEALTH_COLS    = ["user_id","date","day_number","water_glasses","walking_minutes","steps","sleep_hours",
                      "protein_grams","vegetables","fruit","exercise_done","portion_control"]
    WEIGHT_COLS    = ["user_id","date","weight","bmi"]

    def __init__(self):
        os.makedirs(DATA_DIR, exist_ok=True)
        self._ensure_all_files()

    def _ensure_all_files(self):
        specs = [
            (PROFILES_FILE,    "Profiles",     self.PROFILE_COLS),
            (DAILY_INDEX_FILE, "DailyIndex",   self.INDEX_COLS),
            (AVOID_LOGS_FILE,  "AvoidLogs",    self.AVOID_COLS),
            (HEALTH_LOGS_FILE, "HealthLogs",   self.HEALTH_COLS),
            (WEIGHT_HIST_FILE, "WeightHistory",self.WEIGHT_COLS),
        ]
        for path, sheet, cols in specs:
            if not os.path.exists(path):
                pd.DataFrame(columns=cols).to_excel(path, sheet_name=sheet, index=False)
                logger.info(f"Created {os.path.basename(path)}")

    def _load(self, path: str, sheet: str, cols: list) -> pd.DataFrame:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            for c in cols:
                if c not in df.columns:
                    df[c] = None
            return df
        except Exception:
            return pd.DataFrame(columns=cols)

    def _save(self, df: pd.DataFrame, path: str, sheet: str):
        df.to_excel(path, sheet_name=sheet, index=False)

    # ── Profile ──────────────────────────────
    def get_profile(self, user_id: str) -> Optional[Dict]:
        df = self._load(PROFILES_FILE, "Profiles", self.PROFILE_COLS)
        row = df[df["user_id"] == user_id]
        if row.empty:
            return None
        return row.iloc[0].to_dict()

    def create_profile(self, user_id: str, age: int, gender: str, height_cm: float, weight_kg: float):
        df = self._load(PROFILES_FILE, "Profiles", self.PROFILE_COLS)
        bmi = BMIManager.calculate(weight_kg, height_cm)
        cat = BMIManager.categorize(bmi)
        new_row = pd.DataFrame([{
            "user_id": user_id, "age": age, "gender": gender,
            "height_cm": height_cm, "current_weight_kg": weight_kg,
            "bmi": bmi, "bmi_category": cat,
            "last_updated": str(datetime.date.today())
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        self._save(df, PROFILES_FILE, "Profiles")

    def update_profile(self, user_id: str, **kwargs):
        df = self._load(PROFILES_FILE, "Profiles", self.PROFILE_COLS)
        idx = df[df["user_id"] == user_id].index
        if idx.empty:
            return
        for k, v in kwargs.items():
            df.loc[idx, k] = v
        # Recalculate BMI if weight or height changed
        if "current_weight_kg" in kwargs or "height_cm" in kwargs:
            w = float(df.loc[idx[0], "current_weight_kg"])
            h = float(df.loc[idx[0], "height_cm"])
            bmi = BMIManager.calculate(w, h)
            df.loc[idx, "bmi"] = bmi
            df.loc[idx, "bmi_category"] = BMIManager.categorize(bmi)
        df.loc[idx, "last_updated"] = str(datetime.date.today())
        self._save(df, PROFILES_FILE, "Profiles")

    # ── Daily Index + Log Rows ────────────────
    def ensure_day_exists(self, user_id: str, date: datetime.date) -> int:
        date_str = str(date)
        # DailyIndex
        df_idx = self._load(DAILY_INDEX_FILE, "DailyIndex", self.INDEX_COLS)
        day_number, df_idx = DateManager.get_or_create_day_number(user_id, date, df_idx)
        self._save(df_idx, DAILY_INDEX_FILE, "DailyIndex")

        # AvoidLogs
        df_avoid = self._load(AVOID_LOGS_FILE, "AvoidLogs", self.AVOID_COLS)
        if df_avoid[(df_avoid["user_id"]==user_id) & (df_avoid["date"]==date_str)].empty:
            blank = {c: 0 for c in self.AVOID_COLS}
            blank.update({"user_id": user_id, "date": date_str, "day_number": day_number})
            df_avoid = pd.concat([df_avoid, pd.DataFrame([blank])], ignore_index=True)
            self._save(df_avoid, AVOID_LOGS_FILE, "AvoidLogs")

        # HealthLogs
        df_health = self._load(HEALTH_LOGS_FILE, "HealthLogs", self.HEALTH_COLS)
        if df_health[(df_health["user_id"]==user_id) & (df_health["date"]==date_str)].empty:
            blank = {c: 0 for c in self.HEALTH_COLS}
            blank.update({"user_id": user_id, "date": date_str, "day_number": day_number})
            df_health = pd.concat([df_health, pd.DataFrame([blank])], ignore_index=True)
            self._save(df_health, HEALTH_LOGS_FILE, "HealthLogs")

        return day_number

    def load_logs(self, user_id: str, date: datetime.date) -> Tuple[Dict, Dict]:
        date_str = str(date)
        df_avoid  = self._load(AVOID_LOGS_FILE,  "AvoidLogs",  self.AVOID_COLS)
        df_health = self._load(HEALTH_LOGS_FILE, "HealthLogs", self.HEALTH_COLS)
        avoid_row  = df_avoid[(df_avoid["user_id"]==user_id)  & (df_avoid["date"]==date_str)]
        health_row = df_health[(df_health["user_id"]==user_id) & (df_health["date"]==date_str)]
        avoid_data  = avoid_row.iloc[0].to_dict()  if not avoid_row.empty  else {}
        health_data = health_row.iloc[0].to_dict() if not health_row.empty else {}
        return avoid_data, health_data

    def save_logs(self, user_id: str, date: datetime.date, avoid_data: Dict, health_data: Dict):
        date_str = str(date)
        # Avoid
        df_avoid = self._load(AVOID_LOGS_FILE, "AvoidLogs", self.AVOID_COLS)
        idx = df_avoid[(df_avoid["user_id"]==user_id) & (df_avoid["date"]==date_str)].index
        for k, v in avoid_data.items():
            if k in self.AVOID_COLS:
                df_avoid.loc[idx, k] = v
        self._save(df_avoid, AVOID_LOGS_FILE, "AvoidLogs")
        # Health
        df_health = self._load(HEALTH_LOGS_FILE, "HealthLogs", self.HEALTH_COLS)
        idx = df_health[(df_health["user_id"]==user_id) & (df_health["date"]==date_str)].index
        for k, v in health_data.items():
            if k in self.HEALTH_COLS:
                df_health.loc[idx, k] = v
        self._save(df_health, HEALTH_LOGS_FILE, "HealthLogs")
        logger.info(f"Saved logs for {user_id} on {date_str}")

    def add_weight_history(self, user_id: str, date: datetime.date, weight: float, bmi: float):
        df = self._load(WEIGHT_HIST_FILE, "WeightHistory", self.WEIGHT_COLS)
        date_str = str(date)
        existing = df[(df["user_id"]==user_id) & (df["date"]==date_str)]
        if existing.empty:
            new_row = pd.DataFrame([{"user_id": user_id, "date": date_str, "weight": weight, "bmi": bmi}])
            df = pd.concat([df, new_row], ignore_index=True)
        else:
            idx = existing.index
            df.loc[idx, "weight"] = weight
            df.loc[idx, "bmi"] = bmi
        self._save(df, WEIGHT_HIST_FILE, "WeightHistory")

    def get_history(self, user_id: str) -> Dict[str, pd.DataFrame]:
        df_avoid  = self._load(AVOID_LOGS_FILE,  "AvoidLogs",  self.AVOID_COLS)
        df_health = self._load(HEALTH_LOGS_FILE, "HealthLogs", self.HEALTH_COLS)
        df_weight = self._load(WEIGHT_HIST_FILE, "WeightHistory", self.WEIGHT_COLS)
        return {
            "avoid":  df_avoid[df_avoid["user_id"]==user_id].copy(),
            "health": df_health[df_health["user_id"]==user_id].copy(),
            "weight": df_weight[df_weight["user_id"]==user_id].copy(),
        }

    def export_csv(self, user_id: str, path: str):
        history = self.get_history(user_id)
        merged = pd.merge(history["health"], history["avoid"], on=["user_id","date","day_number"], how="outer")
        merged = pd.merge(merged, history["weight"], on=["user_id","date"], how="outer")
        merged.to_csv(path, index=False)
        logger.info(f"Exported CSV to {path}")

    def backup_all(self):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        for fname in ["users.xlsx","profiles.xlsx","daily_index.xlsx",
                      "avoid_logs.xlsx","health_logs.xlsx","weight_history.xlsx"]:
            src = os.path.join(DATA_DIR, fname)
            if os.path.exists(src):
                dst = os.path.join(BACKUP_DIR, f"{ts}_{fname}")
                shutil.copy2(src, dst)
        logger.info("Backup completed")


# ─────────────────────────────────────────────
# ProfileManager
# ─────────────────────────────────────────────
class ProfileManager:
    def __init__(self, log_manager: ExcelLogManager):
        self.lm = log_manager

    def get_profile(self, user_id: str) -> Optional[Dict]:
        return self.lm.get_profile(user_id)

    def update_weight_height(self, user_id: str, weight: float, height: float):
        self.lm.update_profile(user_id, current_weight_kg=weight, height_cm=height)
        profile = self.lm.get_profile(user_id)
        if profile:
            bmi = float(profile["bmi"])
            self.lm.add_weight_history(user_id, DateManager.get_today(), weight, bmi)


# ─────────────────────────────────────────────
# ChartManager
# ─────────────────────────────────────────────
class ChartManager:
    def __init__(self, dark_mode: bool = True):
        self.dark_mode = dark_mode

    def _style(self, fig: Figure, ax):
        if self.dark_mode:
            fig.patch.set_facecolor("#ffffff")
            ax.set_facecolor("#ffffff")
            ax.tick_params(colors="#000000")
            ax.xaxis.label.set_color("#000000")
            ax.yaxis.label.set_color("#000000")
            ax.title.set_color("#0000ff")
            for spine in ax.spines.values():
                spine.set_edgecolor("#000000")
        else:
            fig.patch.set_facecolor("#ffffff")
            ax.set_facecolor("#ffffff")
            ax.tick_params(colors="#000000")
            ax.xaxis.label.set_color("#000000")
            ax.yaxis.label.set_color("#000000")
            ax.title.set_color("#0000ff")
            for spine in ax.spines.values():
                spine.set_edgecolor("#000000")

    def create_charts_dialog(self, parent, history: Dict[str, pd.DataFrame], dark_mode: bool) -> QDialog:
        self.dark_mode = dark_mode
        dlg = QDialog(parent)
        dlg.setWindowTitle("Health Charts")
        dlg.resize(900, 650)
        layout = QVBoxLayout(dlg)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        charts = [
            ("Sugar",   history["avoid"],  "date", "sugar_servings",   "#ff0000"),
            ("Water",   history["health"], "date", "water_glasses",    "#0000ff"),
            ("Walking", history["health"], "date", "walking_minutes",  "#ffd700"),
            ("Weight",  history["weight"], "date", "weight",           "#0000ff"),
            ("BMI",     history["weight"], "date", "bmi",              "#ffd700"),
        ]

        for title, df, xcol, ycol, color in charts:
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            fig = Figure(figsize=(8, 4), tight_layout=True)
            ax = fig.add_subplot(111)
            self._style(fig, ax)

            if not df.empty and xcol in df.columns and ycol in df.columns:
                df_sorted = df.dropna(subset=[xcol, ycol]).sort_values(xcol)
                x = list(range(len(df_sorted)))
                y = pd.to_numeric(df_sorted[ycol], errors="coerce").fillna(0).tolist()
                labels = df_sorted[xcol].astype(str).tolist()

                line, = ax.plot([], [], color=color, linewidth=2, marker="o", markersize=5)
                ax.set_xlim(-0.5, max(len(x)-1, 1) + 0.5)
                ax.set_ylim(0, max(max(y)*1.2, 1))
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=8)
                ax.set_xlabel("Date")
                ax.set_ylabel(ycol.replace("_", " ").title())
                ax.set_title(f"{title} Over Time")

                def animate(frame, line=line, x=x, y=y):
                    end = min(frame+1, len(x))
                    line.set_data(x[:end], y[:end])
                    return line,

                ani = animation.FuncAnimation(fig, animate, frames=len(x)+1,
                                              interval=80, blit=True, repeat=False)
                canvas = FigureCanvas(fig)
                canvas._ani = ani  # keep reference
            else:
                ax.text(0.5, 0.5, "No data yet", transform=ax.transAxes,
                        ha="center", va="center", color="#a6adc8", fontsize=14)
                canvas = FigureCanvas(fig)

            tab_layout.addWidget(canvas)
            tabs.addTab(tab, title)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
        return dlg


# ─────────────────────────────────────────────
# AuthUI
# ─────────────────────────────────────────────
class AuthUI(QWidget):
    login_success = Signal(str)  # emits user_id

    def __init__(self, user_manager: ExcelUserManager, log_manager: ExcelLogManager):
        super().__init__()
        self.um = user_manager
        self.lm = log_manager
        self.stack = QStackedWidget()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stack)
        self._build_intro()
        self._build_login()
        self._build_register()
        self.stack.setCurrentIndex(0)

    def _build_intro(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setAlignment(Qt.AlignCenter)
        v.setSpacing(20)

        icon_lbl = QLabel("🌙")
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setStyleSheet("font-size: 64px;")

        title = QLabel("Ramadan Health Tracker")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)

        desc = QLabel(
            "Track your daily habits, water intake, diet, and exercise\n"
            "throughout Ramadan. Stay consistent and healthy!"
        )
        desc.setObjectName("subtitle")
        desc.setAlignment(Qt.AlignCenter)
        desc.setWordWrap(True)

        btn = QPushButton("Continue →")
        btn.setFixedWidth(200)
        btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))

        v.addStretch()
        v.addWidget(icon_lbl)
        v.addWidget(title)
        v.addWidget(desc)
        v.addSpacing(20)
        v.addWidget(btn, alignment=Qt.AlignCenter)
        v.addStretch()
        self.stack.addWidget(page)

    def _build_login(self):
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setObjectName("card")
        card.setFixedWidth(380)
        form_layout = QVBoxLayout(card)
        form_layout.setSpacing(12)

        title = QLabel("Welcome Back 👋")
        title.setObjectName("title")
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        form_layout.addWidget(title)

        self.login_user = QLineEdit()
        self.login_user.setPlaceholderText("Username")
        self.login_pass = QLineEdit()
        self.login_pass.setPlaceholderText("Password")
        self.login_pass.setEchoMode(QLineEdit.Password)
        self.login_err = QLabel("")
        self.login_err.setStyleSheet("color: #f38ba8;")

        btn_login = QPushButton("Login")
        btn_login.setObjectName("success")
        btn_login.clicked.connect(self._do_login)
        self.login_pass.returnPressed.connect(self._do_login)

        btn_to_reg = QPushButton("New user? Register")
        btn_to_reg.setObjectName("secondary")
        btn_to_reg.clicked.connect(lambda: self.stack.setCurrentIndex(2))

        btn_back = QPushButton("← Back")
        btn_back.setObjectName("secondary")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(0))

        for w in [self.login_user, self.login_pass, self.login_err, btn_login, btn_to_reg, btn_back]:
            form_layout.addWidget(w)

        outer.addStretch()
        outer.addWidget(card, alignment=Qt.AlignCenter)
        outer.addStretch()
        self.stack.addWidget(page)

    def _build_register(self):
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setObjectName("card")
        card.setFixedWidth(420)
        fl = QFormLayout(card)
        fl.setSpacing(10)
        fl.setLabelAlignment(Qt.AlignRight)

        title = QLabel("Create Account ✨")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        fl.addRow(title)

        self.reg_user   = QLineEdit(); self.reg_user.setPlaceholderText("Choose a username")
        self.reg_pass   = QLineEdit(); self.reg_pass.setPlaceholderText("Password"); self.reg_pass.setEchoMode(QLineEdit.Password)
        self.reg_pass2  = QLineEdit(); self.reg_pass2.setPlaceholderText("Confirm password"); self.reg_pass2.setEchoMode(QLineEdit.Password)
        self.reg_age    = QSpinBox(); self.reg_age.setRange(5, 120); self.reg_age.setValue(25)
        self.reg_gender = QComboBox(); self.reg_gender.addItems(["Male", "Female", "Other"])
        self.reg_height = QDoubleSpinBox(); self.reg_height.setRange(50, 250); self.reg_height.setValue(170); self.reg_height.setSuffix(" cm")
        self.reg_weight = QDoubleSpinBox(); self.reg_weight.setRange(10, 500); self.reg_weight.setValue(70); self.reg_weight.setSuffix(" kg")
        self.reg_err    = QLabel(""); self.reg_err.setStyleSheet("color: #f38ba8;"); self.reg_err.setWordWrap(True)

        fl.addRow("Username:", self.reg_user)
        fl.addRow("Password:", self.reg_pass)
        fl.addRow("Confirm:", self.reg_pass2)
        fl.addRow("Age:", self.reg_age)
        fl.addRow("Gender:", self.reg_gender)
        fl.addRow("Height:", self.reg_height)
        fl.addRow("Weight:", self.reg_weight)
        fl.addRow(self.reg_err)

        btn_reg = QPushButton("Register & Continue")
        btn_reg.setObjectName("success")
        btn_reg.clicked.connect(self._do_register)

        btn_to_login = QPushButton("Already have account? Login")
        btn_to_login.setObjectName("secondary")
        btn_to_login.clicked.connect(lambda: self.stack.setCurrentIndex(1))

        fl.addRow(btn_reg)
        fl.addRow(btn_to_login)

        outer.addStretch()
        outer.addWidget(card, alignment=Qt.AlignCenter)
        outer.addStretch()
        self.stack.addWidget(page)

    def _do_login(self):
        username = self.login_user.text().strip()
        password = self.login_pass.text()
        if not username or not password:
            self.login_err.setText("Please fill in all fields.")
            return
        uid = self.um.authenticate(username, password)
        if uid:
            self.login_err.setText("")
            self.login_success.emit(uid)
        else:
            self.login_err.setText("Invalid username or password.")

    def _do_register(self):
        username = self.reg_user.text().strip()
        password = self.reg_pass.text()
        confirm  = self.reg_pass2.text()
        if not username or not password:
            self.reg_err.setText("Username and password are required.")
            return
        if password != confirm:
            self.reg_err.setText("Passwords do not match.")
            return
        if len(password) < 6:
            self.reg_err.setText("Password must be at least 6 characters.")
            return
        if self.um.username_exists(username):
            self.reg_err.setText("Username already taken.")
            return
        uid = self.um.register(username, password)
        if uid:
            self.lm.create_profile(
                uid,
                age=self.reg_age.value(),
                gender=self.reg_gender.currentText(),
                height_cm=self.reg_height.value(),
                weight_kg=self.reg_weight.value()
            )
            self.reg_err.setText("")
            self.login_success.emit(uid)
        else:
            self.reg_err.setText("Registration failed. Try again.")


# ─────────────────────────────────────────────
# DashboardUI
# ─────────────────────────────────────────────
class DashboardUI(QWidget):
    logout_requested = Signal()
    save_requested   = Signal(dict, dict)
    weight_updated   = Signal(float, float)

    def __init__(self, log_manager: ExcelLogManager, profile_manager: ProfileManager):
        super().__init__()
        self.lm = log_manager
        self.pm = profile_manager
        self.user_id: Optional[str] = None
        self.today = DateManager.get_today()
        self._build_ui()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # ── Top bar ──────────────────────────
        top_bar = QHBoxLayout()
        self.lbl_greeting = QLabel("Welcome!")
        self.lbl_greeting.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.lbl_date = QLabel(str(self.today))
        self.lbl_date.setObjectName("subtitle")
        top_bar.addWidget(self.lbl_greeting)
        top_bar.addStretch()
        top_bar.addWidget(self.lbl_date)
        main_layout.addLayout(top_bar)

        # ── Scroll area ──────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(14)
        scroll.setWidget(content)
        main_layout.addWidget(scroll)

        # Profile panel
        content_layout.addWidget(self._build_profile_panel())
        # Avoid table
        content_layout.addWidget(self._build_avoid_table())
        # Health table
        content_layout.addWidget(self._build_health_table())
        content_layout.addStretch()

        # ── Control buttons ──────────────────
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_save    = QPushButton("💾 Save")
        self.btn_save.setObjectName("success")
        self.btn_history = QPushButton("📋 History")
        self.btn_history.setObjectName("secondary")
        self.btn_charts  = QPushButton("📊 Charts")
        self.btn_charts.setObjectName("secondary")
        self.btn_export  = QPushButton("📤 Export CSV")
        self.btn_export.setObjectName("secondary")
        self.btn_logout  = QPushButton("🚪 Logout")
        self.btn_logout.setObjectName("danger")

        for btn in [self.btn_save, self.btn_history, self.btn_charts, self.btn_export, self.btn_logout]:
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            btn_row.addWidget(btn)

        main_layout.addLayout(btn_row)

    def _build_profile_panel(self) -> QGroupBox:
        grp = QGroupBox("👤 Profile")
        layout = QHBoxLayout(grp)
        layout.setSpacing(20)

        # Info labels
        info = QVBoxLayout()
        self.lbl_weight   = QLabel("Weight: —")
        self.lbl_height   = QLabel("Height: —")
        self.lbl_bmi      = QLabel("BMI: —")
        self.lbl_bmi_cat  = QLabel("—")
        self.lbl_bmi_cat.setObjectName("bmi_normal")
        self.lbl_day      = QLabel("Day: —")
        for lbl in [self.lbl_weight, self.lbl_height, self.lbl_bmi, self.lbl_bmi_cat, self.lbl_day]:
            info.addWidget(lbl)
        layout.addLayout(info)

        layout.addStretch()

        # Update form
        update_form = QFormLayout()
        self.spin_weight = QDoubleSpinBox()
        self.spin_weight.setRange(10, 500)
        self.spin_weight.setSuffix(" kg")
        self.spin_height = QDoubleSpinBox()
        self.spin_height.setRange(50, 250)
        self.spin_height.setSuffix(" cm")
        self.btn_update_profile = QPushButton("Update BMI")
        self.btn_update_profile.clicked.connect(self._on_update_profile)
        update_form.addRow("New Weight:", self.spin_weight)
        update_form.addRow("New Height:", self.spin_height)
        update_form.addRow(self.btn_update_profile)
        layout.addLayout(update_form)

        return grp

    def _build_avoid_table(self) -> QGroupBox:
        grp = QGroupBox("🚫 Things to Avoid Today")
        grid = QGridLayout(grp)
        grid.setSpacing(10)

        self.avoid_widgets: Dict[str, Any] = {}

        numeric_fields = [
            ("sugar_servings",    "Sugar Servings",      0, 50),
            ("sweet_drinks_count","Sweet Drinks",         0, 20),
            ("caffeine_cups",     "Caffeine Cups",        0, 20),
        ]
        bool_fields = [
            ("fried_food",  "Fried Food"),
            ("excess_oil",  "Excess Oil"),
            ("overeating",  "Overeating"),
            ("late_screen", "Late Screen Time"),
        ]

        row = 0
        for key, label, mn, mx in numeric_fields:
            lbl = QLabel(label + ":")
            spin = QSpinBox()
            spin.setRange(mn, mx)
            grid.addWidget(lbl,  row, 0)
            grid.addWidget(spin, row, 1)
            self.avoid_widgets[key] = spin
            row += 1

        for key, label in bool_fields:
            cb = QCheckBox(label)
            grid.addWidget(cb, row, 0, 1, 2)
            self.avoid_widgets[key] = cb
            row += 1

        return grp

    def _build_health_table(self) -> QGroupBox:
        grp = QGroupBox("✅ Healthy Habits Today")
        grid = QGridLayout(grp)
        grid.setSpacing(10)

        self.health_widgets: Dict[str, Any] = {}

        numeric_fields = [
            ("water_glasses",   "Water Glasses",    0, 30),
            ("walking_minutes", "Walking (min)",    0, 300),
            ("steps",           "Steps",            0, 100000),
            ("sleep_hours",     "Sleep Hours",      0, 24),
            ("protein_grams",   "Protein (g)",      0, 500),
        ]
        bool_fields = [
            ("vegetables",     "Vegetables"),
            ("fruit",          "Fruit"),
            ("exercise_done",  "Exercise Done"),
            ("portion_control","Portion Control"),
        ]

        row = 0
        for key, label, mn, mx in numeric_fields:
            lbl = QLabel(label + ":")
            spin = QSpinBox()
            spin.setRange(mn, mx)
            grid.addWidget(lbl,  row, 0)
            grid.addWidget(spin, row, 1)
            self.health_widgets[key] = spin
            row += 1

        for key, label in bool_fields:
            cb = QCheckBox(label)
            grid.addWidget(cb, row, 0, 1, 2)
            self.health_widgets[key] = cb
            row += 1

        return grp

    def load_user(self, user_id: str):
        self.user_id = user_id
        self.today = DateManager.get_today()
        self.lbl_date.setText(str(self.today))

        profile = self.pm.get_profile(user_id)
        if profile:
            w = float(profile.get("current_weight_kg", 0))
            h = float(profile.get("height_cm", 0))
            bmi = float(profile.get("bmi", 0))
            cat = str(profile.get("bmi_category", "Normal"))
            self.lbl_weight.setText(f"Weight: {w} kg")
            self.lbl_height.setText(f"Height: {h} cm")
            self.lbl_bmi.setText(f"BMI: {bmi}")
            self.lbl_bmi_cat.setText(cat)
            self.lbl_bmi_cat.setObjectName(BMIManager.color_object_name(cat))
            self.lbl_bmi_cat.style().unpolish(self.lbl_bmi_cat)
            self.lbl_bmi_cat.style().polish(self.lbl_bmi_cat)
            self.spin_weight.setValue(w)
            self.spin_height.setValue(h)
            self.lbl_greeting.setText(f"Welcome back! 🌙")

        day_number = self.lm.ensure_day_exists(user_id, self.today)
        self.lbl_day.setText(f"Day: {day_number}")

        avoid_data, health_data = self.lm.load_logs(user_id, self.today)
        self._populate_avoid(avoid_data)
        self._populate_health(health_data)

    def _populate_avoid(self, data: Dict):
        for key, widget in self.avoid_widgets.items():
            val = data.get(key, 0)
            if isinstance(widget, QCheckBox):
                widget.setChecked(bool(int(val)) if val else False)
            else:
                widget.setValue(int(val) if val else 0)

    def _populate_health(self, data: Dict):
        for key, widget in self.health_widgets.items():
            val = data.get(key, 0)
            if isinstance(widget, QCheckBox):
                widget.setChecked(bool(int(val)) if val else False)
            else:
                widget.setValue(int(val) if val else 0)

    def collect_avoid(self) -> Dict:
        result = {}
        for key, widget in self.avoid_widgets.items():
            if isinstance(widget, QCheckBox):
                result[key] = 1 if widget.isChecked() else 0
            else:
                result[key] = widget.value()
        return result

    def collect_health(self) -> Dict:
        result = {}
        for key, widget in self.health_widgets.items():
            if isinstance(widget, QCheckBox):
                result[key] = 1 if widget.isChecked() else 0
            else:
                result[key] = widget.value()
        return result

    def _on_update_profile(self):
        w = self.spin_weight.value()
        h = self.spin_height.value()
        self.weight_updated.emit(w, h)


# ─────────────────────────────────────────────
# AppController
# ─────────────────────────────────────────────
class AppController(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ramadan Health Tracker 🌙")
        self.resize(900, 700)
        self.setMinimumSize(600, 500)

        self.dark_mode = True
        self.current_user_id: Optional[str] = None

        # Managers
        self.user_manager    = ExcelUserManager()
        self.log_manager     = ExcelLogManager()
        self.profile_manager = ProfileManager(self.log_manager)
        self.chart_manager   = ChartManager(dark_mode=True)

        # Central stack
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # Auth UI
        self.auth_ui = AuthUI(self.user_manager, self.log_manager)
        self.auth_ui.login_success.connect(self._on_login)
        self.stack.addWidget(self.auth_ui)  # index 0

        # Dashboard UI
        self.dashboard_ui = DashboardUI(self.log_manager, self.profile_manager)
        self.dashboard_ui.logout_requested.connect(self._on_logout)
        self.dashboard_ui.weight_updated.connect(self._on_weight_updated)
        self.dashboard_ui.btn_save.clicked.connect(self._on_save)
        self.dashboard_ui.btn_charts.clicked.connect(self._on_show_charts)
        self.dashboard_ui.btn_export.clicked.connect(self._on_export_csv)
        self.dashboard_ui.btn_logout.clicked.connect(self._on_logout)
        self.dashboard_ui.btn_history.clicked.connect(self._on_show_history)
        self.stack.addWidget(self.dashboard_ui)  # index 1

        # Theme toggle in menu bar
        menu = self.menuBar()
        view_menu = menu.addMenu("View")
        self.theme_action = view_menu.addAction("Switch to Light Theme")
        self.theme_action.triggered.connect(self._toggle_theme)

        self._apply_theme()

    def _apply_theme(self):
        qss = DARK_THEME if self.dark_mode else LIGHT_THEME
        self.setStyleSheet(qss)
        self.chart_manager.dark_mode = self.dark_mode
        label = "Switch to Light Theme" if self.dark_mode else "Switch to Dark Theme"
        self.theme_action.setText(label)

    def _toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self._apply_theme()

    def _on_login(self, user_id: str):
        self.current_user_id = user_id
        self.dashboard_ui.load_user(user_id)
        self.stack.setCurrentIndex(1)

    def _on_logout(self):
        self.current_user_id = None
        self.stack.setCurrentIndex(0)

    def _on_save(self):
        if not self.current_user_id:
            return
        avoid_data  = self.dashboard_ui.collect_avoid()
        health_data = self.dashboard_ui.collect_health()
        self.log_manager.save_logs(self.current_user_id, self.dashboard_ui.today, avoid_data, health_data)
        QMessageBox.information(self, "Saved", "✅ Daily log saved successfully!")

    def _on_weight_updated(self, weight: float, height: float):
        if not self.current_user_id:
            return
        self.profile_manager.update_weight_height(self.current_user_id, weight, height)
        self.dashboard_ui.load_user(self.current_user_id)
        QMessageBox.information(self, "Updated", f"✅ Weight & BMI updated!\nWeight: {weight} kg | Height: {height} cm")

    def _on_show_charts(self):
        if not self.current_user_id:
            return
        history = self.log_manager.get_history(self.current_user_id)
        dlg = self.chart_manager.create_charts_dialog(self, history, self.dark_mode)
        dlg.exec()

    def _on_export_csv(self):
        if not self.current_user_id:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export CSV", "", "CSV Files (*.csv)")
        if path:
            self.log_manager.export_csv(self.current_user_id, path)
            QMessageBox.information(self, "Exported", f"✅ Data exported to:\n{path}")

    def _on_show_history(self):
        if not self.current_user_id:
            return
        history = self.log_manager.get_history(self.current_user_id)
        dlg = QDialog(self)
        dlg.setWindowTitle("Health History")
        dlg.resize(800, 500)
        layout = QVBoxLayout(dlg)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        for name, df in [("Health Logs", history["health"]),
                         ("Avoid Logs",  history["avoid"]),
                         ("Weight History", history["weight"])]:
            tab = QWidget()
            tl = QVBoxLayout(tab)
            if df.empty:
                tl.addWidget(QLabel("No data yet."))
            else:
                from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
                table = QTableWidget(len(df), len(df.columns))
                table.setHorizontalHeaderLabels(list(df.columns))
                table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                for r, (_, row) in enumerate(df.iterrows()):
                    for c, val in enumerate(row):
                        table.setItem(r, c, QTableWidgetItem(str(val)))
                tl.addWidget(table)
            tabs.addTab(tab, name)

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignRight)
        dlg.exec()

    def closeEvent(self, event):
        self.log_manager.backup_all()
        logger.info("Application closed. Backup done.")
        super().closeEvent(event)


# ─────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────
def main():
    # DPI awareness
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")

    app = QApplication(sys.argv)
    app.setApplicationName("Ramadan Health Tracker")
    app.setOrganizationName("HealthTracker")

    # Font
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    controller = AppController()
    controller.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
