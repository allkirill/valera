import sys, os, io, re, requests, warnings, traceback, time, shutil, logging, gc
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Tuple
from urllib.parse import urlparse, quote
from collections import deque, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from html import escape

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QCheckBox, QComboBox, QMessageBox, QGridLayout,
    QToolButton, QFileDialog, QGroupBox, QFrame, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QHeaderView, QDialog, QMenu,
    QScrollArea, QSizePolicy
)
from PySide6.QtCore import QThread, Signal, QSettings, Qt, QRect, QTimer, QMutex, QWaitCondition, QStandardPaths, QUrl
from PySide6.QtGui import QPixmap, QPainter, QColor, QBrush, QPen, QPalette, QDesktopServices, QKeySequence, QShortcut, QIcon, QLinearGradient

from PIL import Image as PILImage, ImageOps
import openpyxl
from openpyxl.styles import PatternFill
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# ============================================================
# LOGGING & THEME
# ============================================================
LOGGER = logging.getLogger("Valera")
LOGGER.setLevel(logging.DEBUG)

def setup_logging():
    try: app_dir = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    except: app_dir = os.path.expanduser("~")
    log_dir = os.path.join(app_dir, "ValeraSoft", "logs")
    os.makedirs(log_dir, exist_ok=True)
    handler = RotatingFileHandler(os.path.join(log_dir, "valera.log"), maxBytes=5*1024*1024, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    if not LOGGER.handlers: LOGGER.addHandler(handler)

def apply_light_theme(app):
    app.setStyle("Fusion")
    p = QPalette()
    p.setColor(QPalette.Window, QColor(240, 240, 240)); p.setColor(QPalette.WindowText, Qt.black)
    p.setColor(QPalette.Base, QColor(255, 255, 255)); p.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    p.setColor(QPalette.ToolTipBase, QColor(255, 255, 220)); p.setColor(QPalette.ToolTipText, Qt.black)
    p.setColor(QPalette.Text, Qt.black); p.setColor(QPalette.Button, QColor(240, 240, 240))
    p.setColor(QPalette.ButtonText, Qt.black); p.setColor(QPalette.BrightText, Qt.red)
    p.setColor(QPalette.Link, QColor(42, 130, 218)); p.setColor(QPalette.Highlight, QColor(42, 130, 218))
    p.setColor(QPalette.HighlightedText, Qt.white); app.setPalette(p)

# ============================================================
# DATA STRUCTURES
# ============================================================
@dataclass
class ProcessingPreset:
    name: str; min_px: int; max_px: int; max_upscale_pct: int; align: str; fmt: str
    center_square: bool; padding_pct: int; remove_meta: bool = False; replace_transparent: bool = True

PRESETS = {
    "Пользовательский": ProcessingPreset("Пользовательский", 0, 4000, 50, "height", "jpg", True, 10, False, True),
    "santehnica.ru": ProcessingPreset("santehnica.ru", 1080, 4000, 100, "height", "jpg", True, 10, True, True),
    "A4": ProcessingPreset("A4", 2480, 3508, 20, "height", "png", False, 0, True, False),
}

@dataclass
class AppSettings:
    source: str = ""; source_type: str = "Excel"; out_dir: str = ""
    article_col: str = "A"; url_from: str = "B"; url_to: str = "P"
    rename_mode: str = "article"; folder_sort: str = "В порядке Excel"
    min_px: int = 0; max_px: int = 4000; max_upscale_pct: int = 50; align: str = "height"
    fmt: str = "jpg"; center_square: bool = True; padding_pct: int = 10
    replace_transparent: bool = True; remove_meta: bool = False; report_copy: bool = False
    clean_raw: bool = False; ssl_verify: bool = True; preset_name: str = "santehnica.ru"
    pdf_always_square: bool = True; selected_rejected_files: List[str] = field(default_factory=list)
    white_threshold: int = 255; watermark_path: str = ""

# ============================================================
# VALERA PIXMAP CREATOR
# ============================================================
def create_valera_pixmap(size=50):
    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    valera_path = os.path.join(exe_dir, "valera.png")
    if os.path.exists(valera_path):
        pix = QPixmap(valera_path)
        if not pix.isNull(): return pix.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    pix = QPixmap(size, int(size * 1.1)); pix.fill(Qt.transparent)
    p = QPainter(pix); p.setRenderHint(QPainter.Antialiasing); s = size / 50.0
    p.setBrush(QBrush(QColor(255, 193, 7))); p.setPen(QPen(QColor(230, 160, 0), max(1, s)))
    p.drawRoundedRect(int(10*s), int(1*s), int(30*s), int(12*s), int(4*s), int(4*s)); p.drawRect(int(7*s), int(10*s), int(36*s), int(4*s))
    p.setBrush(QBrush(QColor(255, 224, 189))); p.setPen(QPen(QColor(210, 180, 150), max(1, s)))
    p.drawEllipse(int(14*s), int(12*s), int(22*s), int(20*s))
    p.setBrush(QBrush(QColor(50, 50, 50))); p.setPen(Qt.NoPen)
    p.drawEllipse(int(19*s), int(20*s), int(3*s), int(3*s)); p.drawEllipse(int(28*s), int(20*s), int(3*s), int(3*s))
    p.setBrush(QBrush(QColor(255, 255, 255)))
    p.drawEllipse(int(20*s), int(20*s), max(1, int(1.5*s)), max(1, int(1.5*s))); p.drawEllipse(int(29*s), int(20*s), max(1, int(1.5*s)), max(1, int(1.5*s)))
    p.setPen(QPen(QColor(180, 100, 80), max(1, int(1.5*s)))); p.setBrush(Qt.NoBrush); p.drawArc(int(20*s), int(25*s), int(10*s), int(6*s), 0, -180*16)
    p.setBrush(QBrush(QColor(76, 163, 224))); p.setPen(QPen(QColor(50, 130, 190), max(1, s)))
    p.drawRoundedRect(int(16*s), int(32*s), int(18*s), int(15*s), int(2*s), int(2*s))
    p.setBrush(QBrush(QColor(76, 163, 224))); p.drawRect(int(8*s), int(34*s), int(8*s), int(5*s)); p.drawRect(int(34*s), int(34*s), int(8*s), int(5*s))
    p.setBrush(QBrush(QColor(255, 224, 189))); p.setPen(QPen(QColor(210, 180, 150), max(1, s)))
    p.drawEllipse(int(6*s), int(33*s), int(6*s), int(6*s)); p.drawEllipse(int(38*s), int(33*s), int(6*s), int(6*s))
    p.setBrush(QBrush(QColor(90, 90, 90))); p.setPen(QPen(QColor(60, 60, 60), max(1, s)))
    p.drawRect(int(17*s), int(46*s), int(7*s), int(7*s)); p.drawRect(int(26*s), int(46*s), int(7*s), int(7*s))
    p.setBrush(QBrush(QColor(70, 50, 30))); p.setPen(Qt.NoPen)
    p.drawRoundedRect(int(15*s), int(51*s), int(9*s), int(4*s), int(2*s), int(2*s)); p.drawRoundedRect(int(26*s), int(51*s), int(9*s), int(4*s), int(2*s), int(2*s))
    p.end(); return pix

# ============================================================
# UTILS
# ============================================================
def col_letter_to_index(letter):
    col = 0
    for ch in letter.upper(): col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1

def generate_columns(count=120):
    res = []
    for i in range(count):
        idx = i + 1; chars = []
        while idx > 0: idx, rem = divmod(idx - 1, 26); chars.append(chr(rem + ord("A")))
        res.append("".join(reversed(chars)))
    return res

def is_excel_locked(fp):
    if not os.path.exists(fp): return False
    try: open(fp, "rb+").close(); return False
    except: return True

URL_REGEX = re.compile(r'(https?://[^\s,;]+|www\.[^\s,;]+)')
def extract_urls(val):
    if not val: return []
    urls = URL_REGEX.findall(str(val).strip()); cleaned = []
    for u in urls:
        u = u.rstrip('.,;:')
        if u.startswith('www.'): u = 'https://' + u
        if u.startswith('http'): cleaned.append(u)
    return cleaned

# ============================================================
# UNDO MANAGER
# ============================================================
class UndoManager:
    def __init__(self, max_steps=15): self.history = deque(maxlen=max_steps); self.temp_dir = None
    def init_temp(self, path): self.temp_dir = path; os.makedirs(self.temp_dir, exist_ok=True)
    def backup_for_edit(self, filepath):
        if not self.temp_dir: return None
        ts = time.strftime("%Y%m%d_%H%M%S") + f"_{os.urandom(2).hex()}"; backup = os.path.join(self.temp_dir, f"{ts}_{os.path.basename(filepath)}")
        try: shutil.copy2(filepath, backup); return backup
        except Exception as e: LOGGER.error(f"Undo backup edit failed: {e}"); return None
    def backup_for_delete(self, filepath):
        if not self.temp_dir: return None
        ts = time.strftime("%Y%m%d_%H%M%S") + f"_{os.urandom(2).hex()}"; backup = os.path.join(self.temp_dir, f"{ts}_{os.path.basename(filepath)}")
        try: os.makedirs(os.path.dirname(backup), exist_ok=True); shutil.move(filepath, backup); return backup
        except Exception as e: LOGGER.error(f"Undo backup delete failed: {e}"); return None
    def push_batch(self, operations):
        if operations: self.history.append(operations)
    def undo(self):
        if not self.history: return False, "Нет действий для отмены"
        batch = self.history.pop()
        try:
            for op in reversed(batch):
                action, orig, backup = op
                if action == "edit": shutil.copy2(backup, orig)
                elif action == "delete": os.makedirs(os.path.dirname(orig), exist_ok=True); shutil.move(backup, orig)
            return True, f"Отменено {len(batch)} действий"
        except Exception as e: return False, f"Ошибка при отмене: {e}"
    def clear(self):
        self.history.clear()
        if self.temp_dir and os.path.exists(self.temp_dir):
            try: shutil.rmtree(self.temp_dir)
            except: pass

# ============================================================
# IMAGE PROCESSOR
# ============================================================
class ImageProcessor:
    def classify_white_background(self, img, threshold=255):
        img_rgba = img.convert("RGBA") if img.mode != "RGBA" else img.copy()
        img_small = img_rgba.copy(); img_small.thumbnail((100, 100)); w, h = img_small.size
        if w < 5 or h < 5: return "interior"
        px = img_small.load(); is_pure = [[False]*w for _ in range(h)]; is_near = [[False]*w for _ in range(h)]
        near_thr = max(threshold - 15, 0)
        for y in range(h):
            for x in range(w):
                r, g, b, a = px[x, y]
                if a == 0 or (a == 255 and r >= threshold and g >= threshold and b >= threshold): is_pure[y][x] = True
                elif a == 255 and r >= near_thr and g >= near_thr and b >= near_thr: is_near[y][x] = True
        border = set()
        for x in range(w): border.update([(x, 0), (x, h-1)])
        for y in range(1, h-1): border.update([(0, y), (w-1, y)])
        vis = [[False]*w for _ in range(h)]; pure = 0
        for bx, by in border:
            if is_pure[by][bx] and not vis[by][bx]:
                q = deque([(bx, by)]); vis[by][bx] = True
                while q:
                    cx, cy = q.popleft(); pure += 1
                    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
                        nx, ny = cx+dx, cy+dy
                        if 0<=nx<w and 0<=ny<h and not vis[ny][nx] and is_pure[ny][nx]: vis[ny][nx]=True; q.append((nx,ny))
        total = w * h
        if total > 0 and (pure/total) >= 0.12: return "white_bg"
        vis2 = [[False]*w for _ in range(h)]; near = 0
        for bx, by in border:
            if (is_pure[by][bx] or is_near[by][bx]) and not vis2[by][bx]:
                q = deque([(bx, by)]); vis2[by][bx] = True
                while q:
                    cx, cy = q.popleft(); near += 1
                    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
                        nx, ny = cx+dx, cy+dy
                        if 0<=nx<w and 0<=ny<h and not vis2[ny][nx] and (is_pure[ny][nx] or is_near[ny][nx]): vis2[ny][nx]=True; q.append((nx,ny))
        if total > 0 and (near/total) >= 0.12: return "questionable"
        if total > 0 and 0.03 <= (pure/total) < 0.12: return "questionable"
        return "interior"

    def _crop_to_content(self, img, threshold=255):
        img_rgba = img.convert("RGBA") if img.mode != "RGBA" else img.copy()
        w, h = img_rgba.size
        if w == 0 or h == 0: return None
        try: data = list(img_rgba.get_flattened_data())
        except AttributeError: data = list(img_rgba.getdata())
        mask_data = bytearray(w*h)
        for i, (r, g, b, a) in enumerate(data):
            if a == 0 or (a == 255 and r >= threshold and g >= threshold and b >= threshold): mask_data[i] = 0
            else: mask_data[i] = 255
        mask = PILImage.frombytes("L", (w, h), bytes(mask_data)); bbox = mask.getbbox()
        if bbox and (bbox[2]-bbox[0])>0 and (bbox[3]-bbox[1])>0: return img.crop(bbox)
        return None

    def _center_in_square(self, img, padding_pct):
        w, h = img.size
        if w == 0 or h == 0: return img
        max_side = max(w, h); pad = int(max_side * (padding_pct / 100.0)); cs = max_side + 2*pad
        bg = PILImage.new("RGBA", (cs, cs), (255, 255, 255, 255))
        if img.mode == "RGBA": bg.paste(img, ((cs-w)//2, (cs-h)//2), img)
        else: bg.paste(img, ((cs-w)//2, (cs-h)//2))
        return bg

    def apply_watermark(self, img, watermark_path):
        if not watermark_path or not os.path.exists(watermark_path): return img
        try:
            wm = PILImage.open(watermark_path).convert("RGBA"); iw, ih = img.size
            wm_max_w = max(iw // 4, 50); wm_max_h = max(ih // 4, 50); wm.thumbnail((wm_max_w, wm_max_h), PILImage.LANCZOS)
            wm_w, wm_h = wm.size; padding = max(iw // 50, 10); pos_x = iw - wm_w - padding; pos_y = ih - wm_h - padding
            alpha = wm.split()[3]; alpha = alpha.point(lambda p: int(p * 0.5)); wm.putalpha(alpha)
            if img.mode != "RGBA": img = img.convert("RGBA")
            img.paste(wm, (pos_x, pos_y), wm); wm.close(); return img
        except Exception as e: LOGGER.error(f"Watermark error: {e}"); return img

    def process(self, img, s, bg_type="white_bg", skip_center=False, threshold=255):
        if s.remove_meta: img = ImageOps.exif_transpose(img); img.info.pop("exif", None)
        if s.center_square and bg_type == "white_bg" and not skip_center:
            cropped = self._crop_to_content(img, threshold)
            if cropped is not None: img = cropped
            if img.size[0] > 0 and img.size[1] > 0: img = self._center_in_square(img, s.padding_pct)
        w, h = img.size
        if w == 0 or h == 0: return PILImage.new("RGBA", (1,1), (255,255,255,255)), "!ОШИБКА_"
        if s.replace_transparent and img.mode == "RGBA":
            bg = PILImage.new("RGB", img.size, (255, 255, 255)); bg.paste(img, mask=img.split()[3]); img = bg
        td = h if s.align == "height" else w
        if td == 0: td = 1
        scale = 1.0; pref = ""
        if s.min_px > 0 and td < s.min_px:
            pct = ((s.min_px - td) / td) * 100; scale = s.min_px / td
            if s.max_upscale_pct > 0 and pct > s.max_upscale_pct: pref = "!РАЗМЕР_"
        elif s.max_px > 0 and td > s.max_px: scale = s.max_px / td
        if scale != 1.0: img = img.resize((int(w*scale), int(h*scale)), PILImage.LANCZOS)
        if s.watermark_path: img = self.apply_watermark(img, s.watermark_path)
        return img, pref

    def process_with_analysis(self, img, s, force_type="", threshold=255):
        if force_type == "WHITE_BG":
            cropped = self._crop_to_content(img, threshold)
            if cropped is not None: img = cropped
            if img.size[0] > 0 and img.size[1] > 0 and s.center_square: img = self._center_in_square(img, s.padding_pct)
            p, _ = self.process(img, s, "white_bg", threshold=threshold); return p, "", "FORCED_WHITE"
        if force_type == "PDF_SQUARE":
            cropped = self._crop_to_content(img, threshold)
            if cropped is not None: img = cropped
            if img.size[0] > 0 and img.size[1] > 0 and s.center_square: img = self._center_in_square(img, s.padding_pct)
            p, _ = self.process(img, s, "white_bg", threshold=threshold); return p, "", "OK_PDF"
        if force_type == "SIZE":
            p, _ = self.process(img, s, "white_bg", threshold=threshold); return p, "", "FORCED_SIZE"
        bg_type = self.classify_white_background(img, threshold); p, sd = self.process(img, s, bg_type, threshold=threshold)
        if sd: return p, sd, "SIZE_FAIL"
        if bg_type == "questionable": return p, "!БЕЛЫЙ_", "QUESTION_WHITE"
        if bg_type == "white_bg": return p, "", "OK_WHITE"
        return p, "", "OK_INTERIOR"

# ============================================================
# CLOUD RESOLVER & EXCEL MANAGER
# ============================================================
class CloudResolver:
    def __init__(self, verify_ssl=True):
        self.session = requests.Session(); self.session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "*/*", "Origin": "https://cloud.mail.ru", "Referer": "https://cloud.mail.ru/"})
        adapter = HTTPAdapter(max_retries=Retry(total=3, backoff_factor=1, status_forcelist=[429,500,502,503,504], allowed_methods=["HEAD","GET","OPTIONS"]))
        self.session.mount("https://", adapter); self.session.mount("http://", adapter); self.verify = verify_ssl
    def resolve_yandex(self, public_key):
        api = f"https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key={quote(public_key, safe='')}"
        try:
            r = self.session.get(api, timeout=10, verify=self.verify)
            if r.status_code == 200 and "href" in r.json(): return [(r.json()["href"], None)]
        except: pass
        return [(public_key, None)]
    def resolve_mail(self, public_key):
        try: self.session.get(public_key, timeout=10, verify=self.verify)
        except: pass
        parsed = urlparse(public_key); path = parsed.path; weblink = path[len("/public"):] if path.startswith("/public/") else path
        try:
            r = self.session.get(f"https://cloud.mail.ru/api/v2/discovery?weblink={quote(weblink, safe='')}", timeout=10, verify=self.verify)
            if r.status_code == 200 and r.json().get("status") == 200:
                wg = r.json().get("body", {}).get("weblink_get", [])
                if wg and wg[0].get("url"): return [(wg[0]["url"], None)]
        except: pass
        cloclo_url = f"https://cloclo3.cloud.mail.ru/weblink/view{weblink}"
        try:
            r = self.session.head(cloclo_url, timeout=5, verify=self.verify, allow_redirects=True)
            if r.status_code == 200 and "image" in r.headers.get("Content-Type", "").lower(): return [(cloclo_url, None)]
        except: pass
        return [(public_key, None)]

class ExcelManager:
    def __init__(self, filepath): self.filepath = filepath; self.wb = None
    def load(self): self.wb = openpyxl.load_workbook(self.filepath); return self.wb
    def close(self):
        if self.wb: self.wb.close(); self.wb = None
    def count_urls(self, url_cols, article_col):
        total = 0
        for sheet in self.wb.worksheets:
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=max(article_col, max(url_cols)) + 1):
                if article_col >= len(row) or not row[article_col].value: continue
                for c in url_cols:
                    if c < len(row) and row[c].value: total += len(extract_urls(row[c].value))
        return total

# ============================================================
# PREVIEW WIDGETS
# ============================================================
class ThumbnailWidget(QFrame):
    selection_changed = Signal()
    def __init__(self, filepath, category, parent=None):
        super().__init__(parent); self.filepath = filepath; self.category = category
        self.setFixedSize(110, 140); self.setFrameStyle(QFrame.StyledPanel); self._update_style(False)
        layout = QVBoxLayout(self); layout.setContentsMargins(4,4,4,4); layout.setSpacing(2)
        self.checkbox = QCheckBox(); self.checkbox.setStyleSheet("QCheckBox::indicator { width: 20px; height: 20px; }")
        self.checkbox.stateChanged.connect(lambda: self._update_style(self.checkbox.isChecked())); self.checkbox.stateChanged.connect(self.selection_changed)
        cat_text = ""
        if category == "QUESTION_WHITE": cat_text = "❓ Белый"
        elif category == "SIZE_FAIL": cat_text = "📏 Размер"
        elif category == "ERROR": cat_text = "❌ Ошибка"
        elif category == "OK_WHITE": cat_text = "✅ Белый"
        top = QHBoxLayout(); top.addWidget(self.checkbox)
        if cat_text: top.addWidget(QLabel(cat_text)); top.addStretch()
        layout.addLayout(top)
        self.img_label = QLabel(); self.img_label.setFixedSize(90, 90); self.img_label.setAlignment(Qt.AlignCenter)
        self.img_label.setStyleSheet("background: #eee; border: 1px solid #ccc;"); self._load_thumbnail()
        layout.addWidget(self.img_label, alignment=Qt.AlignCenter)
        name = os.path.basename(filepath); dn = name[:15]+"…" if len(name)>15 else name
        self.name_label = QLabel(dn); self.name_label.setAlignment(Qt.AlignCenter); self.name_label.setStyleSheet("font-size: 9px;"); self.name_label.setWordWrap(True); self.name_label.setFixedHeight(25)
        layout.addWidget(self.name_label); self.img_label.mouseDoubleClickEvent = lambda e: QDesktopServices.openUrl(QUrl.fromLocalFile(filepath))
    def _load_thumbnail(self):
        # Оптимизация: используем нативный QPixmap вместо PIL для экономии памяти
        try:
            pm = QPixmap(self.filepath)
            if not pm.isNull():
                self.img_label.setPixmap(pm.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else: self.img_label.setText("Ошибка")
        except: self.img_label.setText("Ошибка")
    def _update_style(self, sel):
        if sel: self.setStyleSheet("QFrame { border: 3px solid #4CA3E0; border-radius: 4px; background: #e0f0fb; }")
        else: self.setStyleSheet("QFrame { border: 1px solid #ccc; border-radius: 4px; background: white; }")
    def is_selected(self): return self.checkbox.isChecked()
    def set_selected(self, v): self.checkbox.setChecked(v)

class ErrorBadgeWidget(QFrame):
    def __init__(self, url, status, message, parent=None):
        super().__init__(parent)
        self.setStyleSheet("QFrame { background-color: #FDEDEC; border: 1px solid #E74C3C; border-radius: 3px; padding: 2px; }")
        layout = QHBoxLayout(self); layout.setContentsMargins(6,3,6,3); layout.setSpacing(6)
        status_lbl = QLabel(f"<b>{escape(str(status))}</b>"); status_lbl.setStyleSheet("color: #C0392B; font-size: 11px;")
        safe_url = escape(str(url)); display = safe_url if len(safe_url) <= 50 else safe_url[:50] + "…"
        url_lbl = QLabel(f'<a href="{safe_url}">{display}</a>'); url_lbl.setOpenExternalLinks(True); url_lbl.setStyleSheet("color: #2980B9; font-size: 11px;")
        msg_text = escape(str(message)); 
        if len(msg_text) > 60: msg_text = msg_text[:60] + "…"
        msg_lbl = QLabel(msg_text); msg_lbl.setStyleSheet("color: #555; font-size: 10px;"); msg_lbl.setWordWrap(True)
        layout.addWidget(status_lbl); layout.addWidget(url_lbl); layout.addWidget(msg_lbl); layout.addStretch()
        self.setToolTip(f"{safe_url}\n{escape(str(message))}"); self.setMaximumHeight(50)

class ProductRowWidget(QFrame):
    def __init__(self, fn, parent=None):
        super().__init__(parent); self.folder_name = fn; self.thumbs = []; self.error_widgets = []
        self.setFrameStyle(QFrame.StyledPanel)
        self.setStyleSheet("QFrame { border: 1px solid #ddd; border-radius: 4px; background: #fafafa; padding: 4px; }")
        layout = QVBoxLayout(self); layout.setContentsMargins(5,5,5,5); layout.setSpacing(4)
        self.title_lbl = QLabel(f"📦 {fn}"); layout.addWidget(self.title_lbl)
        self.tl = QHBoxLayout(); self.tl.setSpacing(8); self.tl.setAlignment(Qt.AlignLeft); layout.addLayout(self.tl)
        self.el = QHBoxLayout(); self.el.setSpacing(6); self.el.setAlignment(Qt.AlignLeft); layout.addLayout(self.el)
    def add_thumbnail(self, t): self.thumbs.append(t); self.tl.addWidget(t)
    def add_error(self, e_widget): self.error_widgets.append(e_widget); self.el.addWidget(e_widget)
    def has_category(self, cats):
        if not cats: return True
        if any(t.category in cats for t in self.thumbs): return True
        if "DOWNLOAD_ERROR" in cats and self.error_widgets: return True
        return False
    def apply_filter(self, cats): self.setVisible(self.has_category(cats))

class ProductPreviewDialog(QDialog):
    files_selected = Signal(list)
    PAGE_SIZE = 30
    def __init__(self, folder_path, settings, parent=None, download_errors=None, threshold_level=0):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр"); self.resize(1100, 750)
        self.folder_path = folder_path; self.s = settings; self.product_rows = []
        self.all_data = []; self.filtered_data = []; self.current_page = 0; self.active_filter = None
        self.download_errors = download_errors or []; self.current_threshold = max(80, 255 - (threshold_level * 20))
        self.undo_mgr = UndoManager(max_steps=15)
        undo_dir = os.path.join(os.path.dirname(folder_path), ".valera_undo_" + os.path.basename(folder_path))
        self.undo_mgr.init_temp(undo_dir)
        self.shortcut_undo = QShortcut(QKeySequence("Ctrl+Z"), self); self.shortcut_undo.activated.connect(self.do_undo)
        layout = QVBoxLayout(self); layout.setContentsMargins(10,10,10,10)
        layout.addWidget(QLabel("📦 Галочки ➔ Действие снизу. Двойной клик ➔ Открыть. Красные блоки ➔ Ошибки загрузки."))
        fl = QHBoxLayout(); self.filter_buttons = {}
        filters = [("Все на белом",("OK_WHITE","QUESTION_WHITE")),("Размер",("SIZE_FAIL",)),("Проблемы",("QUESTION_WHITE","SIZE_FAIL","ERROR")),("Ошибки загрузки",("DOWNLOAD_ERROR",)),("Показать все",None)]
        for text, cats in filters:
            btn = QPushButton(text); btn.setCheckable(True); btn.setMinimumHeight(32); btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, c=cats, b=btn: self.apply_filter(c, b)); self.filter_buttons[cats] = btn; fl.addWidget(btn)
        fl.addStretch(); layout.addLayout(fl)
        pl = QHBoxLayout()
        self.btn_prev = QPushButton("◀ Назад"); self.btn_prev.clicked.connect(self._prev_page)
        self.lbl_page = QLabel("1 / 1"); self.lbl_page.setAlignment(Qt.AlignCenter); self.lbl_page.setMinimumWidth(80); self.lbl_page.setStyleSheet("font-weight: bold;")
        self.btn_next = QPushButton("Вперед ▶"); self.btn_next.clicked.connect(self._next_page)
        pl.addStretch(); pl.addWidget(self.btn_prev); pl.addWidget(self.lbl_page); pl.addWidget(self.btn_next); pl.addStretch(); layout.addLayout(pl)
        self.sa = QScrollArea(); self.sa.setWidgetResizable(True)
        self.lw = QWidget(); self.ll = QVBoxLayout(self.lw); self.ll.setSpacing(10); self.ll.setAlignment(Qt.AlignTop); self.sa.setWidget(self.lw)
        layout.addWidget(self.sa)
        self.cl = QLabel("Выбрано: 0"); self.cl.setStyleSheet("font-weight: bold; font-size: 12px;"); layout.addWidget(self.cl)
        bl = QHBoxLayout()
        b_sel = QPushButton("✅ Выбрать видимые"); b_sel.clicked.connect(self.select_all)
        b_desel = QPushButton("⬜ Снять"); b_desel.clicked.connect(self.deselect_all)
        b_undo = QPushButton("↩ Отменить (Ctrl+Z)"); b_undo.setMinimumHeight(40); b_undo.clicked.connect(self.do_undo)
        b_crop = QPushButton("✂ Обрезать поля"); b_crop.clicked.connect(self.action_crop)
        b_sq = QPushButton("⬜ Центрировать в квадрат"); b_sq.clicked.connect(self.action_square)
        b_reduce = QPushButton("🔍 Сократить отступ"); b_reduce.clicked.connect(self.action_reduce_padding)
        b_orig = QPushButton("↩ Исходник"); b_orig.clicked.connect(self.action_original)
        b_del = QPushButton("🗑 Удалить"); b_del.setStyleSheet("QPushButton{color:red;font-weight:bold}"); b_del.clicked.connect(self.action_delete)
        self.btn_proc = QPushButton(f"⚡ Принудительно преобразовать (порог: {self.current_threshold}) (0)")
        self.btn_proc.setMinimumHeight(40)
        self.btn_proc.setStyleSheet("QPushButton{font-weight:bold;font-size:14px;background-color:#4CA3E0;color:white;border-radius:5px}QPushButton:hover{background-color:#3C93D0}QPushButton:disabled{background-color:#ccc}")
        self.btn_proc.clicked.connect(self.accept_selection)
        b_cancel = QPushButton("Закрыть"); b_cancel.clicked.connect(self.reject)
        bl.addWidget(b_sel); bl.addWidget(b_desel); bl.addWidget(b_undo); bl.addSpacing(10)
        bl.addWidget(b_crop); bl.addWidget(b_sq); bl.addWidget(b_reduce); bl.addWidget(b_orig); bl.addWidget(b_del); bl.addStretch()
        bl.addWidget(self.btn_proc); bl.addWidget(b_cancel); layout.addLayout(bl)
        self._load_records()
        active_btn = self.filter_buttons.get(self.active_filter); self._update_filter_buttons(active_btn)

    def do_undo(self):
        ok, msg = self.undo_mgr.undo()
        if ok: self._load_records(); self._update_count(); QMessageBox.information(self, "Отмена", msg)
        else: QMessageBox.warning(self, "Отмена", msg)
    def _update_filter_buttons(self, active_btn):
        for btn in self.filter_buttons.values():
            if btn == active_btn: btn.setStyleSheet("QPushButton { background-color: #27AE60; color: white; font-weight: bold; border: 2px solid #1E8449; border-radius: 4px; padding: 4px 12px; } QPushButton:hover { background-color: #229954; }"); btn.setChecked(True)
            else: btn.setStyleSheet(""); btn.setChecked(False)
    def apply_filter(self, cats, btn=None):
        self.active_filter = cats
        if btn: self._update_filter_buttons(btn)
        if cats is None: self.filtered_data = self.all_data[:]
        else:
            self.filtered_data = []; cat_set = set(cats)
            for item in self.all_data:
                item_cats = item.get("categories", set())
                if item_cats & cat_set: self.filtered_data.append(item)
        self.current_page = 0; self._render_page(); self._update_count()
    def _render_page(self):
        while self.ll.count():
            item = self.ll.takeAt(0); w = item.widget()
            if w: w.deleteLater()
        self.product_rows.clear()
        total_pages = max(1, (len(self.filtered_data) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        self.current_page = min(self.current_page, total_pages - 1)
        start = self.current_page * self.PAGE_SIZE; end = min(start + self.PAGE_SIZE, len(self.filtered_data))
        self.lw.setUpdatesEnabled(False)
        for item in self.filtered_data[start:end]:
            rw = ProductRowWidget(item["name"], self)
            for fp, cat in item.get("files", []):
                t = ThumbnailWidget(fp, cat, self); t.selection_changed.connect(self._update_count); rw.add_thumbnail(t)
            for err in item.get("errors", []):
                ew = ErrorBadgeWidget(err["url"], err["status"], err["message"], self); rw.add_error(ew)
            self.product_rows.append(rw); self.ll.addWidget(rw)
        self.lw.setUpdatesEnabled(True); self._update_pagination_controls(total_pages)
    def _update_pagination_controls(self, total_pages):
        self.lbl_page.setText(f"{self.current_page + 1} / {total_pages}")
        self.btn_prev.setEnabled(self.current_page > 0); self.btn_next.setEnabled(self.current_page < total_pages - 1)
    def _prev_page(self):
        if self.current_page > 0: self.current_page -= 1; self._render_page(); self._update_count()
    def _next_page(self):
        total_pages = max(1, (len(self.filtered_data) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        if self.current_page < total_pages - 1: self.current_page += 1; self._render_page(); self._update_count()
    def _get_selected(self): return [t.filepath for row in self.product_rows for t in row.thumbs if t.is_selected() and t.isVisible()]
    
    def _process_edit_action(self, files, process_func):
        if not files: return
        proc = ImageProcessor(); ops = []
        for fp in files:
            backup = self.undo_mgr.backup_for_edit(fp)
            if backup: ops.append(("edit", fp, backup))
        if not ops: QMessageBox.warning(self, "Отмена", "Не удалось создать резервные копии"); return
        for fp in files:
            try:
                img = PILImage.open(fp).convert("RGBA")
                img_final = process_func(proc, img)
                if self.s.replace_transparent and img_final.mode == "RGBA":
                    bg = PILImage.new("RGB", img_final.size, (255,255,255)); bg.paste(img_final, mask=img_final.split()[3]); img_final = bg
                img_final.save(fp)
                # Принудительная очистка
                del img; del img_final
            except Exception as e: LOGGER.error(f"Editor error: {e}")
        self.undo_mgr.push_batch(ops)
        self.current_threshold = max(80, self.current_threshold - 20)
        self._load_records(); gc.collect()

    def action_crop(self):
        def process_func(proc, img):
            c = proc._crop_to_content(img, self.current_threshold)
            if c: img = c
            img_final, _ = proc.process(img, self.s, "white_bg", skip_center=True)
            return img_final
        self._process_edit_action(self._get_selected(), process_func)

    def action_square(self):
        def process_func(proc, img):
            c = proc._crop_to_content(img, self.current_threshold)
            if c: img = c
            sq = proc._center_in_square(img, self.s.padding_pct); img_final, _ = proc.process(sq, self.s, "white_bg", skip_center=True)
            return img_final
        self._process_edit_action(self._get_selected(), process_func)

    def action_reduce_padding(self):
        def process_func(proc, img):
            c = proc._crop_to_content(img, self.current_threshold)
            if c: img = c
            sq = proc._center_in_square(img, max(2, self.s.padding_pct // 3))
            img_final, _ = proc.process(sq, self.s, "white_bg", skip_center=True)
            return img_final
        self._process_edit_action(self._get_selected(), process_func)

    def action_original(self):
        files = self._get_selected()
        if not files: return
        raw_dir = self.folder_path.replace("обработано_", "скачано_")
        if not os.path.exists(raw_dir): QMessageBox.warning(self, "Ошибка", "Папка с исходниками не найдена!"); return
        ops = []
        for fp in files:
            backup = self.undo_mgr.backup_for_edit(fp)
            if backup: ops.append(("edit", fp, backup))
        if not ops: return
        for fp in files:
            rel = os.path.relpath(fp, self.folder_path); raw_fp = os.path.join(raw_dir, rel)
            raw_folder = os.path.dirname(raw_fp)
            if os.path.exists(raw_folder):
                raws = [f for f in os.listdir(raw_folder) if f.startswith("raw_")]
                if raws:
                    try: shutil.copy2(os.path.join(raw_folder, raws[0]), fp)
                    except Exception as e: LOGGER.error(f"Restore original error: {e}")
        self.undo_mgr.push_batch(ops); self._load_records()
    def action_delete(self):
        files = self._get_selected()
        if not files: return
        if QMessageBox.question(self, "Удаление", f"Удалить {len(files)} файлов? (можно отменить)") == QMessageBox.Yes:
            ops = []
            for fp in files:
                backup = self.undo_mgr.backup_for_delete(fp)
                if backup: ops.append(("delete", fp, backup))
            if ops: self.undo_mgr.push_batch(ops)
            self._load_records()
    def _load_records(self):
        self.all_data.clear()
        st = defaultdict(list)
        for root, _, fns in os.walk(self.folder_path):
            for fn in fns:
                if fn.lower().endswith((".png",".jpg",".jpeg",".webp")):
                    fp = os.path.join(root, fn); rel = os.path.relpath(fp, self.folder_path)
                    parts = rel.split(os.sep); folder_name = parts[1] if len(parts) >= 3 else (parts[0] if len(parts) == 2 else "Без группы")
                    cat = "OK_INTERIOR"
                    if fn.startswith("!БЕЛЫЙ_"): cat = "QUESTION_WHITE"
                    elif fn.startswith("!РАЗМЕР_"): cat = "SIZE_FAIL"
                    elif fn.startswith("!ОШИБКА_"): cat = "ERROR"
                    st[folder_name].append((fp, cat))
        err_by_folder = defaultdict(list)
        for err in self.download_errors:
            folder = err.get("folder", "Без группы"); err_by_folder[folder].append(err)
        all_folders = set(st.keys()) | set(err_by_folder.keys())
        for fn in sorted(all_folders):
            files = st.get(fn, []); errors = err_by_folder.get(fn, [])
            cats = set(cat for _, cat in files)
            if errors: cats.add("DOWNLOAD_ERROR")
            self.all_data.append({"type":"product","name":fn,"files":files,"errors":errors,"categories":cats})
        self.apply_filter(self.active_filter)
    def _update_count(self):
        total = len(self._get_selected())
        self.cl.setText(f"Выбрано: {total}")
        self.btn_proc.setEnabled(total > 0)
        self.btn_proc.setText(f"⚡ Преобразовать (порог:{self.current_threshold}) ({total})")
    def select_all(self):
        for r in self.product_rows:
            if r.isVisible():
                for t in r.thumbs: t.set_selected(True)
    def deselect_all(self):
        for r in self.product_rows:
            for t in r.thumbs: t.set_selected(False)
    def accept_selection(self):
        s = self._get_selected()
        if not s: QMessageBox.information(self, "Внимание", "Не выбрано файлов."); return
        self.files_selected.emit(s); self.accept()

# ============================================================
# LOG DIALOG & PROGRESS WIDGET
# ============================================================
class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent); self.setWindowTitle("Лог"); self.setMinimumSize(1000,700)
        lay = QVBoxLayout(self); self.table = QTableWidget(); self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Статус","Источник","Сообщение"])
        self.table.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents); self.table.horizontalHeader().setSectionResizeMode(1,QHeaderView.Interactive); self.table.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.setColumnWidth(1,500); lay.addWidget(self.table)
        b=QPushButton("Закрыть"); b.clicked.connect(self.close); lay.addWidget(b)
    def add_rows_batch(self, rows):
        self.table.setUpdatesEnabled(False)
        for s,so,m in rows:
            r=self.table.rowCount(); self.table.insertRow(r)
            for c,t in enumerate([s,so,m]): self.table.setItem(r,c,QTableWidgetItem(t))
        self.table.setUpdatesEnabled(True); self.table.scrollToBottom()

class ValeraProgressWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self.setMinimumHeight(70); self._v = 0; self._valera_pix = create_valera_pixmap(48)
    def setValue(self, v): self._v = max(0, min(100, v)); self.update()
    def paintEvent(self, e):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        track_y = 28; track_h = 18; margin = 55; track_rect = QRect(margin, track_y, self.width() - margin * 2, track_h)
        p.setBrush(QBrush(QColor(220, 220, 220))); p.setPen(QPen(QColor(200, 200, 200), 1)); p.drawRoundedRect(track_rect, 9, 9)
        if self._v > 0:
            fill_w = int(track_rect.width() * (self._v / 100.0)); fill_rect = QRect(track_rect.x(), track_rect.y(), max(fill_w, track_h), track_h)
            grad = QLinearGradient(fill_rect.topLeft(), fill_rect.topRight()); grad.setColorAt(0, QColor(76, 163, 224)); grad.setColorAt(1, QColor(130, 210, 255))
            p.setBrush(QBrush(grad)); p.setPen(Qt.NoPen); p.drawRoundedRect(fill_rect, 9, 9)
        vx = track_rect.x() + int(track_rect.width() * (self._v / 100.0)) - 24; vx = max(track_rect.x() - 24, min(vx, track_rect.right() - 24))
        if self._valera_pix: p.drawPixmap(int(vx), track_y - 28, self._valera_pix)
        p.setPen(QPen(QColor(80, 80, 80))); f = p.font(); f.setBold(True); f.setPointSize(10); p.setFont(f)
        p.drawText(QRect(0, track_y + track_h + 3, self.width(), 18), Qt.AlignCenter, f"Ход работ: {self._v}%"); p.end()

# ============================================================
# WORKER
# ============================================================
class Worker(QThread):
    progress = Signal(int); stats_updated = Signal(int,int,int,int,int,float)
    log_row = Signal(str,str,str); finished = Signal(dict); error = Signal(str); download_errors_ready = Signal(list)
    def __init__(self, settings):
        super().__init__()
        self.s=settings; self._p=False; self._c=False; self._m=QMutex(); self._w=QWaitCondition()
        self.stats={"ok":0,"defect_size":0,"defect_white":0,"fail":0,"bytes":0}
        self.processor=ImageProcessor(); self.resolver=CloudResolver(verify_ssl=settings.ssl_verify)
        self.download_errors = []; self.start_time = 0
    def pause(self): self._p=True
    def resume(self): self._p=False; self._w.wakeAll()
    def cancel(self): self._c=True; self.resume()
    def check_cancel(self): return self._c
    def check_pause(self):
        if self._p: self._m.lock(); self._w.wait(self._m); self._m.unlock()
    def emit_stats(self): 
        elapsed = time.time() - self.start_time
        self.stats_updated.emit(self.stats["ok"],self.stats["defect_size"],self.stats["defect_white"],self.stats["fail"],self.stats["bytes"], elapsed)
    def log(self, s, so, m=""): self.log_row.emit(s, so, m)
    def save_img(self, img, path):
        fmt = self.s.fmt
        try:
            if fmt=="jpg":
                if img.mode in ("RGBA","P"): rgb=PILImage.new("RGB",img.size,(255,255,255)); rgb.paste(img,mask=img.split()[3]) if img.mode=="RGBA" else rgb.paste(img); rgb.save(path,quality=92,optimize=True); rgb.close()
                else: img.save(path,quality=92,optimize=True)
            elif fmt=="webp": img.save(path,optimize=True,quality=90)
            else: img.save(path,optimize=True)
        finally:
            try: img.close()
            except: pass
    def process_pdf_bytes(self, content):
        try: import fitz
        except: raise Exception("pip install PyMuPDF")
        doc=fitz.open(stream=content,filetype="pdf"); imgs=[]
        for pn in range(len(doc)): pg=doc.load_page(pn); px=pg.get_pixmap(dpi=200); imgs.append((PILImage.frombytes("RGB",[px.width,px.height],px.samples).convert("RGBA"),pn+1))
        doc.close(); return imgs
    def is_pdf_multi_b(self, c):
        try: import fitz; doc=fitz.open(stream=c,filetype="pdf"); l=len(doc); doc.close(); return l>1
        except: return False
    def is_pdf_multi_f(self, fp):
        try: import fitz; doc=fitz.open(fp); l=len(doc); doc.close(); return l>1
        except: return False
    def run(self):
        try:
            self.start_time=time.time(); self.stats={"ok":0,"defect_size":0,"defect_white":0,"fail":0,"bytes":0}; self.download_errors.clear()
            if self.s.source_type=="Отбракованное": self.process_rejected()
            elif self.s.source_type=="Excel": self.process_excel()
            else: self.process_folder(self.s.source,self.s.out_dir)
            self.stats["time"]=time.time()-self.start_time; self.download_errors_ready.emit(self.download_errors); self.finished.emit(self.stats)
        except Exception as e: LOGGER.exception("Worker crash"); self.error.emit(traceback.format_exc())
    def process_rejected(self):
        files=self.s.selected_rejected_files
        if files: self.stats["processed_dir"] = os.path.dirname(files[0])
        if not files: return
        thr = self.s.white_threshold
        for i,fp in enumerate(files):
            if self.check_cancel(): break; self.check_pause(); self._proc_single(fp, thr); self.progress.emit(int((i+1)/len(files)*100))
        self.emit_stats(); self.progress.emit(100)
    def _proc_single(self, fp, threshold=255):
        bn=os.path.basename(fp); ft=""
        if bn.startswith("!РАЗМЕР_"): ft="SIZE"
        elif bn.startswith("!БЕЛЫЙ_"): ft="WHITE_BG"
        elif bn.startswith("!ОШИБКА_"): ft="RETRY"
        if not ft: return
        cn=bn
        for p in ["!РАЗМЕР_","!БЕЛЫЙ_","!ОШИБКА_"]:
            if cn.startswith(p): cn=cn[len(p):]; break
        fb=os.path.splitext(cn)[0]; td=os.path.dirname(fp); tp=os.path.join(td,f"{fb}_ИСПРАВЛЕНО.{self.s.fmt}")
        c=1
        while os.path.exists(tp): tp=os.path.join(td,f"{fb}_ИСПРАВЛЕНО_{c}.{self.s.fmt}"); c+=1
        try:
            with PILImage.open(fp) as img: img=img.convert("RGBA"); pi,_,_=self.processor.process_with_analysis(img,self.s,force_type=ft,threshold=threshold)
            self.save_img(pi,tp); self.stats["ok"]+=1
        except Exception as e: self.stats["fail"]+=1; self.log("[ОШИБКА]",bn,str(e))
        self.emit_stats()
    def process_folder(self, src, out):
        dd=os.path.join(out,"Готово"); self.stats["processed_dir"] = dd
        files=[os.path.join(r,f) for r,_,fs in os.walk(src) for f in fs if f.lower().endswith((".png",".jpg",".jpeg",".webp",".pdf"))]
        thr = self.s.white_threshold
        for i,fp in enumerate(files):
            if self.check_cancel(): break; self.check_pause()
            try:
                rp=os.path.relpath(fp,src); tfp=os.path.join(dd,rp); os.makedirs(os.path.dirname(tfp),exist_ok=True); fb=os.path.splitext(os.path.basename(fp))[0]
                if fp.lower().endswith(".pdf"): self._proc_pdf(fp,tfp,fb,thr)
                else: self._proc_img(fp,tfp,fb,thr)
            except: self.stats["fail"]+=1
            self.emit_stats(); self.progress.emit(int((i+1)/max(1,len(files))*100))
        gc.collect()
    def _proc_pdf(self, fp, tfp, fb, threshold=255):
        if self.is_pdf_multi_f(fp): shutil.copy2(fp,os.path.join(os.path.dirname(tfp),f"{fb}.pdf")); self.stats["ok"]+=1; return
        with open(fp,"rb") as f: c=f.read()
        for img,pn in self.process_pdf_bytes(c):
            if self.check_cancel(): img.close(); break
            pi,dp,cat=self.processor.process_with_analysis(img,self.s,force_type="PDF_SQUARE" if self.s.pdf_always_square else "",threshold=threshold)
            img.close(); self.save_img(pi,os.path.join(os.path.dirname(tfp),f"{dp}{fb}_стр{pn}.{self.s.fmt}")); self._log_def(fb+f"_стр{pn}",cat)
    def _proc_img(self, fp, tfp, fb, threshold=255):
        with PILImage.open(fp) as img: img=img.convert("RGBA"); pi,dp,cat=self.processor.process_with_analysis(img,self.s,threshold=threshold)
        self.save_img(pi,os.path.join(os.path.dirname(tfp),f"{dp}{fb}.{self.s.fmt}")); self._log_def(fb,cat); self.stats["bytes"]+=os.path.getsize(fp)
    def _log_def(self, n, cat):
        if cat=="SIZE_FAIL": self.stats["defect_size"]+=1; self.log("[ОТБРАКОВАНО]",n,"!РАЗМЕР_")
        elif cat=="QUESTION_WHITE": self.stats["defect_white"]+=1; self.log("[ПОД ВОПРОСОМ]",n,"!БЕЛЫЙ_")
        else: self.stats["ok"]+=1; self.log("[OK]",n,"Успешно")
        
    def process_excel(self):
        eb=os.path.splitext(os.path.basename(self.s.source))[0]
        dl_dir=os.path.join(self.s.out_dir,f"скачано_{eb}"); pr_dir=os.path.join(self.s.out_dir,f"обработано_{eb}")
        os.makedirs(dl_dir,exist_ok=True); os.makedirs(pr_dir,exist_ok=True)
        self.stats["processed_dir"] = pr_dir; xl=ExcelManager(self.s.source); wb=xl.load()
        fc=col_letter_to_index(self.s.url_from); tc=col_letter_to_index(self.s.url_to)
        ucols=list(range(fc,tc+1)); ac=col_letter_to_index(self.s.article_col); mc=max(ac,max(ucols))+1
        tu=xl.count_urls(ucols,ac)
        if tu==0: xl.close(); raise Exception("В Excel нет ссылок!")
        cm={}; proc=0

        def download_task(fetch_url, raw_folder, idx):
            os.makedirs(raw_folder, exist_ok=True)
            try:
                resp = self.resolver.session.get(fetch_url, timeout=15, headers={"Referer": urlparse(fetch_url).scheme + "://" + urlparse(fetch_url).netloc + "/"}, verify=self.s.ssl_verify)
                if resp.status_code != 200: raise Exception(f"Status {resp.status_code}")
                content = resp.content; ctype = resp.headers.get("Content-Type", "").lower()
                is_pdf = fetch_url.lower().endswith(".pdf") or "application/pdf" in ctype
                if is_pdf:
                    if self.is_pdf_multi_b(content):
                        pdf_path = os.path.join(raw_folder, f"raw_{idx}.pdf")
                        with open(pdf_path, "wb") as pf: pf.write(content)
                        return ("OK_PDF", os.path.relpath(pdf_path, dl_dir), None, len(content))
                    else:
                        paths = []
                        for img, pn in self.process_pdf_bytes(content):
                            fname = f"raw_{idx}_{pn}.{self.s.fmt}"; rp = os.path.join(raw_folder, fname)
                            self.save_img(img, rp); paths.append(os.path.relpath(rp, dl_dir))
                        return ("OK_PDF_IMG", paths, None, len(content))
                else:
                    # Оптимизация: просто сохраняем байты на диск без загрузки в PIL!
                    ext = ".jpg"
                    if "image/png" in ctype or fetch_url.lower().endswith(".png"): ext = ".png"
                    elif "image/webp" in ctype or fetch_url.lower().endswith(".webp"): ext = ".webp"
                    fname = f"raw_{idx}{ext}"; rp = os.path.join(raw_folder, fname)
                    with open(rp, "wb") as f: f.write(content)
                    return ("OK_IMG", os.path.relpath(rp, dl_dir), None, len(content))
            except Exception as e:
                return ("ERR", None, str(e), 0)

        # Уменьшено количество потоков с 4 до 2 для экономии ОЗУ
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = {}
            for sheet in wb.worksheets:
                ss = "".join(c if c.isalnum() else "_" for c in sheet.title)
                for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=mc):
                    if self.check_cancel(): break
                    if ac >= len(row) or not row[ac].value: continue
                    art = row[ac].value; er = row[0].row; sa = "".join(c if c.isalnum() else "_" for c in str(art))
                    fn = (f"{er:03d} - {sa}" if self.s.folder_sort == "В порядке Excel" else sa)
                    tb = os.path.join(dl_dir, ss); idx = 1
                    for col in ucols:
                        if col >= len(row) or not row[col].value: continue
                        cell = row[col]; urls = extract_urls(cell.value)
                        for val in urls:
                            iy = any(d in val for d in ["yadi.sk","disk.yandex","360.yandex"]); im = "cloud.mail.ru" in val
                            if iy: utf = self.resolver.resolve_yandex(val)
                            elif im: utf = self.resolver.resolve_mail(val)
                            else: utf = [(val, None)]
                            for fu, sf in utf:
                                rf = os.path.join(tb, fn)
                                if sf: rf = os.path.join(rf, sf)
                                future = executor.submit(download_task, fu, rf, idx); futures[future] = (cell, val, fu, fn); idx += 1
            for future in as_completed(futures):
                if self.check_cancel(): break
                cell, val, fu, folder_name = futures[future]
                try:
                    res_type, path, err, size = future.result(); self.stats["bytes"] += size
                    if res_type == "OK_IMG":
                        cm[path] = cell; self.log("[OK]", fu, f"Скачано {size/(1024*1024):.1f} МБ")
                    elif res_type == "OK_PDF":
                        cm[path] = cell; self.log("[OK]", fu, "PDF сохранён")
                    elif res_type == "OK_PDF_IMG":
                        for pp in path: cm[pp] = cell
                        self.log("[OK]", fu, f"PDF → {len(path)} фото")
                    else:
                        status_code = "404" if err and ("404" in err or "Status 404" in err) else "ERROR"
                        self.download_errors.append({"url": fu, "status": status_code, "message": err, "folder": folder_name})
                        cell.fill = PatternFill(start_color="FF6347", fill_type="solid"); self.log("[ОШИБКА]", fu, err); self.stats["fail"] += 1
                except Exception as e:
                    self.download_errors.append({"url": val, "status": "ERROR", "message": str(e), "folder": folder_name})
                    self.log("[ОШИБКА]", val, str(e)); self.stats["fail"] += 1
                proc += 1; self.emit_stats(); self.progress.emit(int(proc / tu * 50))
        gc.collect() # Очистка после скачивания

        # ФАЗА 2: Обработка скачанных файлов с диска
        self.log("[ИНФО]","Обработка","Начинаю обработку...")
        thr = self.s.white_threshold
        ftp=[os.path.join(r,f) for r,_,fs in os.walk(dl_dir) for f in fs if f.lower().endswith((".png",".jpg",".jpeg",".webp"))]
        for i,rp in enumerate(ftp):
            if self.check_cancel(): break; self.check_pause()
            rlp=os.path.relpath(rp,dl_dir); fd=os.path.join(pr_dir,os.path.dirname(rlp)); os.makedirs(fd,exist_ok=True)
            if rlp not in cm: continue
            cell=cm[rlp]
            if self.s.rename_mode=="article":
                fn=os.path.basename(os.path.dirname(rlp)); sa=fn.split(" - ",1)[-1] if " - " in fn else fn
                rn=os.path.splitext(os.path.basename(rlp))[0]; ip=rn.split("_")[-1]; fname=f"{sa}_{ip}.{self.s.fmt}"
            else: fname=os.path.basename(rlp)
            fp=os.path.join(fd,fname)
            if os.path.exists(fp): cell.fill=PatternFill(start_color="90EE90",fill_type="solid"); self.stats["ok"]+=1; self.emit_stats(); self.progress.emit(50+int((i+1)/len(ftp)*50)); continue
            try:
                with PILImage.open(rp) as img: img=img.convert("RGBA"); pi,dp,cat=self.processor.process_with_analysis(img,self.s,threshold=thr)
                self.save_img(pi,os.path.join(fd,f"{dp}{fname}"))
                if cat=="SIZE_FAIL": cell.fill=PatternFill(start_color="FFD700",fill_type="solid"); self.stats["defect_size"]+=1; self.log("[ОТБРАКОВАНО]",fname,"!РАЗМЕР_")
                elif cat=="QUESTION_WHITE": cell.fill=PatternFill(start_color="87CEEB",fill_type="solid"); self.stats["defect_white"]+=1; self.log("[ПОД ВОПРОСОМ]",fname,"!БЕЛЫЙ_")
                else: cell.fill=PatternFill(start_color="90EE90",fill_type="solid"); self.stats["ok"]+=1; self.log("[OK]",fname,"Успешно")
                del pi # Очистка памяти
            except Exception as e: cell.fill=PatternFill(start_color="FF6347",fill_type="solid"); self.stats["fail"]+=1; self.log("[ОШИБКА]",rp,str(e))
            self.emit_stats(); self.progress.emit(50+int((i+1)/len(ftp)*50))
        try:
            if self.s.report_copy: wb.save(os.path.join(os.path.dirname(self.s.source),f"Отчет_{os.path.basename(self.s.source)}"))
            else: wb.save(self.s.source)
            xl.close()
        except Exception as e: xl.close(); raise Exception(f"Ошибка Excel: {e}")
        if self.s.clean_raw and os.path.exists(dl_dir):
            try: shutil.rmtree(dl_dir)
            except: pass
        gc.collect() # Финальная очистка

# ============================================================
# AUTHOR DIALOG & MAIN WINDOW
# ============================================================
class AuthorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent); self.setWindowTitle("Автор"); self.setFixedSize(380, 220)
        self.setStyleSheet("QDialog { background: white; }")
        layout = QVBoxLayout(self); layout.setSpacing(12); layout.setContentsMargins(20, 20, 20, 15)
        vp = create_valera_pixmap(40)
        if not vp.isNull():
            lbl_img = QLabel(); lbl_img.setPixmap(vp); lbl_img.setAlignment(Qt.AlignCenter); layout.addWidget(lbl_img)
        t1 = QLabel('Гитхаб автора: <a href="https://github.com/allkirill" style="color:#4CA3E0;">https://github.com/allkirill</a>')
        t1.setOpenExternalLinks(True); t1.setStyleSheet("font-size: 12px;"); layout.addWidget(t1)
        t2 = QLabel('Еще несколько решений: <a href="https://vlookup-app.ru" style="color:#4CA3E0;">https://vlookup-app.ru</a>')
        t2.setOpenExternalLinks(True); t2.setStyleSheet("font-size: 12px;"); layout.addWidget(t2)
        t3 = QLabel("Фотоукладчик Валера всегда на связи 🛠")
        t3.setStyleSheet("font-weight: bold; font-size: 13px; color: #4CA3E0; margin-top: 6px;"); layout.addWidget(t3)
        layout.addStretch()
        btn = QPushButton("Закрыть"); btn.setFixedWidth(100); btn.clicked.connect(self.close); layout.addWidget(btn, alignment=Qt.AlignCenter)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Фотоукладчик Валера"); self.setMinimumSize(760,720); self.resize(860,780)
        self.setWindowIcon(QIcon(create_valera_pixmap(32)))
        self.source_path=None; self.output_dir=None; self.settings_io=QSettings("ValeraSoft","PhotoLoader")
        self.worker=None; self.log_dialog=None; self.setAcceptDrops(True); self._ap=False
        self._srf=[]; self.log_buffer=[]; self.log_timer=QTimer(); self.log_timer.timeout.connect(self.flush_log); self.log_timer.start(300)
        self.last_download_errors = []; self.last_processed_dir = None
        self.force_threshold_level = 0; self.watermark_path = ""
        self.build_ui(); self.load_settings()

    def build_ui(self):
        ml=QVBoxLayout(self); ml.setContentsMargins(10,10,10,10); ml.setSpacing(6)
        tg=QGridLayout(); tg.setSpacing(5)
        self.bi=QToolButton(); self.bi.setText("ℹ"); self.bi.setStyleSheet("border:none;color:blue;font-weight:bold;font-size:16px;"); self.bi.clicked.connect(self.show_info)
        self.stc=QComboBox(); self.stc.addItems(["Excel","Папка"]); self.stc.setFixedWidth(90); self.stc.activated.connect(self.on_stc)
        self.bs=QPushButton("Выбрать источник"); self.bs.setFixedWidth(140); self.ls=QLabel("Не выбран"); self.ls.setStyleSheet("color:gray;"); self.ls.setWordWrap(True)
        tg.addWidget(self.bi,0,0); tg.addWidget(self.stc,0,1); tg.addWidget(self.bs,0,2); tg.addWidget(self.ls,0,3)
        cl=generate_columns(); he=QHBoxLayout(); he.setSpacing(4)
        self.ca=QComboBox(); self.ca.addItems(cl); self.ca.setFixedWidth(50); self.ca.setEditable(True)
        self.cuf=QComboBox(); self.cuf.addItems(cl); self.cuf.setFixedWidth(50); self.cuf.setEditable(True)
        self.cut=QComboBox(); self.cut.addItems(cl); self.cut.setCurrentText("P"); self.cut.setFixedWidth(50); self.cut.setEditable(True)
        self.cr=QComboBox(); self.cr.addItems(["По артикулу","Оригинальное"]); self.cr.setFixedWidth(110)
        self.cfs=QComboBox(); self.cfs.addItems(["В порядке Excel","По алфавиту"]); self.cfs.setFixedWidth(120)
        he.addWidget(QLabel("Артикул:")); he.addWidget(self.ca); he.addSpacing(6); he.addWidget(QLabel("Ссылки от:")); he.addWidget(self.cuf); he.addWidget(QLabel("до:")); he.addWidget(self.cut); he.addSpacing(6); he.addWidget(QLabel("Имя:")); he.addWidget(self.cr); he.addSpacing(6); he.addWidget(QLabel("Папки:")); he.addWidget(self.cfs); he.addStretch()
        self.weo=QWidget(); self.weo.setLayout(he); tg.addWidget(self.weo,1,0,1,4)
        hr=QHBoxLayout(); self.chk_rc=QCheckBox("Отчёт в копию"); self.chk_cr=QCheckBox("Удалить исходники"); self.chk_ssl=QCheckBox("SSL"); self.chk_ssl.setChecked(True)
        b1=QToolButton(); b1.setText("ℹ"); b1.setStyleSheet("border:none;color:blue;"); b1.clicked.connect(lambda: QMessageBox.information(self,"Справка","Отчёт сохранится в копию, чтобы не сломать оригинал."))
        b2=QToolButton(); b2.setText("ℹ"); b2.setStyleSheet("border:none;color:blue;"); b2.clicked.connect(lambda: QMessageBox.information(self,"Справка","Удалит папку 'скачано' после обработки."))
        b3=QToolButton(); b3.setText("ℹ"); b3.setStyleSheet("border:none;color:blue;"); b3.clicked.connect(lambda: QMessageBox.information(self,"Справка","Отключите, если ошибки SSL в корпоративной сети."))
        hr.addWidget(self.chk_rc); hr.addWidget(b1); hr.addSpacing(10); hr.addWidget(self.chk_cr); hr.addWidget(b2); hr.addSpacing(10); hr.addWidget(self.chk_ssl); hr.addWidget(b3); hr.addStretch()
        self.wer=QWidget(); self.wer.setLayout(hr); tg.addWidget(self.wer,2,0,1,4)
        self.bd=QPushButton("Место назначения"); self.bd.setFixedWidth(140); self.ld=QLabel("Рабочий стол"); self.ld.setStyleSheet("color:gray;")
        tg.addWidget(QLabel(""),3,0); tg.addWidget(self.bd,3,2); tg.addWidget(self.ld,3,3)
        ml.addLayout(tg)
        pl=QHBoxLayout(); pl.addWidget(QLabel("💡 Пресет:")); self.pc=QComboBox(); self.pc.addItems(list(PRESETS.keys())); self.pc.activated.connect(self.on_ps); pl.addWidget(self.pc); pl.addStretch(); ml.addLayout(pl)
        gp=QGroupBox("Параметры"); gl=QHBoxLayout(gp); gl.setSpacing(5)
        self.fb=QComboBox(); self.fb.addItems(["jpg","png","webp"]); self.fb.setFixedWidth(60)
        self.al=QComboBox(); self.al.addItems(["По высоте","По ширине"]); self.al.setFixedWidth(85)
        self.mi=QLineEdit("0"); self.mi.setFixedWidth(45); self.ma=QLineEdit("4000"); self.ma.setFixedWidth(45); self.up=QLineEdit("50"); self.up.setFixedWidth(35)
        bu=QToolButton(); bu.setText("ℹ"); bu.setStyleSheet("border:none;color:blue;"); bu.clicked.connect(lambda: QMessageBox.information(self,"Справка","Если % увеличения больше лимита, фото пометится !РАЗМЕР_"))
        gl.addWidget(QLabel("Формат:")); gl.addWidget(self.fb); gl.addSpacing(4); gl.addWidget(self.al); gl.addSpacing(4); gl.addWidget(QLabel("Мин px:")); gl.addWidget(self.mi); gl.addSpacing(4); gl.addWidget(QLabel("Макс px:")); gl.addWidget(self.ma); gl.addSpacing(4); gl.addWidget(QLabel("Макс.увел %:")); gl.addWidget(self.up); gl.addWidget(bu)
        gl.addSpacing(12)
        self.btn_wm=QPushButton("Наложить водный знак"); self.btn_wm.setFixedWidth(160); self.btn_wm.setCursor(Qt.PointingHandCursor); self.btn_wm.clicked.connect(self.pick_watermark)
        self.btn_wm_clear=QPushButton("✕"); self.btn_wm_clear.setFixedWidth(22); self.btn_wm_clear.setVisible(False); self.btn_wm_clear.setStyleSheet("QPushButton{color:red;font-weight:bold;border:1px solid #ccc;border-radius:3px}QPushButton:hover{background:#ffcccc}"); self.btn_wm_clear.clicked.connect(self.clear_watermark)
        self.lbl_wm=QLabel(""); self.lbl_wm.setStyleSheet("color:gray;font-size:10px;")
        gl.addWidget(self.btn_wm); gl.addWidget(self.btn_wm_clear); gl.addWidget(self.lbl_wm)
        gl.addStretch(); ml.addWidget(gp)
        gw=QGroupBox("Белый фон и PDF"); wl=QHBoxLayout(gw); wl.setSpacing(5)
        self.chk_cs=QCheckBox("Центрировать в квадрат"); self.pi=QLineEdit("10"); self.pi.setFixedWidth(35)
        bpi=QToolButton(); bpi.setText("ℹ"); bpi.setStyleSheet("border:none;color:blue;"); bpi.clicked.connect(lambda: QMessageBox.information(self,"Справка","Процент отступа вокруг объекта при центрировании в квадрат.\nПри повторном нажатии 'Принудительно преобразовать' порог определения фона снижается (255→240→220…), что позволяет программе захватить всё более серые поля."))
        self.chk_rt=QCheckBox("Преобразовать прозрачный в белый"); self.chk_rm=QCheckBox("Очистить метаданные"); self.chk_pdf=QCheckBox("PDF в квадрат"); self.chk_pdf.setChecked(True)
        bpdf=QToolButton(); bpdf.setText("ℹ"); bpdf.setStyleSheet("border:none;color:blue;"); bpdf.clicked.connect(lambda: QMessageBox.information(self,"Справка","1-стр PDF → фото в квадрате. Много-стр PDF → просто скачается."))
        wl.addWidget(self.chk_cs); wl.addSpacing(4); wl.addWidget(QLabel("Отступ %:")); wl.addWidget(self.pi); wl.addWidget(bpi); wl.addSpacing(12); wl.addWidget(self.chk_rt); wl.addSpacing(12); wl.addWidget(self.chk_rm); wl.addSpacing(12); wl.addWidget(self.chk_pdf); wl.addWidget(bpdf); wl.addStretch(); ml.addWidget(gw)
        lh=QHBoxLayout(); lh.addWidget(QLabel("🖥 Лог:")); lh.addStretch(); self.bel=QToolButton(); self.bel.setText("⛶"); self.bel.clicked.connect(self.show_log_d); lh.addWidget(self.bel); ml.addLayout(lh)
        self.lt=QTableWidget(); self.lt.setColumnCount(3); self.lt.setHorizontalHeaderLabels(["Статус","Источник","Сообщение"])
        self.lt.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents); self.lt.horizontalHeader().setSectionResizeMode(1,QHeaderView.Interactive); self.lt.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch)
        self.lt.setEditTriggers(QAbstractItemView.NoEditTriggers); self.lt.setColumnWidth(1,380); self.lt.setMinimumHeight(110); self.lt.setMaximumHeight(180); ml.addWidget(self.lt)
        self.ls2=QLabel("✅ ОК: 0 | 📏 Размер: 0 | ❓ Белый: 0 | ❌ Ошибки: 0 | 💾 0.0 МБ | ⏱ 00:00")
        self.ls2.setStyleSheet("font-weight:bold;font-size:11px;"); ml.addWidget(self.ls2)
        self.pw=ValeraProgressWidget(); self.pw.setVisible(False); ml.addWidget(self.pw)
        
        bl=QHBoxLayout(); bl.setSpacing(6)
        self.btn_start=QPushButton("▶ ЗАПУСТИТЬ"); self.btn_start.setMinimumHeight(50); self.btn_start.setStyleSheet("QPushButton{font-weight:bold;font-size:16px;background-color:#4CA3E0;color:white;border-radius:5px}QPushButton:hover{background-color:#3C93D0}")
        self.btn_rej=QPushButton("⚡ Проверить отбракованные"); self.btn_rej.setMinimumHeight(50); self.btn_rej.setVisible(False); self.btn_rej.setStyleSheet("QPushButton{font-weight:bold;font-size:14px;background-color:#FF9800;color:white;border-radius:5px}QPushButton:hover{background-color:#F57C00}")
        self.btn_cancel=QPushButton("✖ ОТМЕНА"); self.btn_cancel.setMinimumHeight(50); self.btn_cancel.setVisible(False); self.btn_cancel.setStyleSheet("QPushButton{font-weight:bold;font-size:16px;background-color:#E74C3C;color:white;border-radius:5px}")
        bl.addWidget(self.btn_start,3); bl.addWidget(self.btn_rej,2); bl.addWidget(self.btn_cancel,1)
        ml.addLayout(bl)
        
        # Автор справа внизу
        bottom_row = QHBoxLayout()
        bottom_row.addStretch()
        self.lbl_author = QLabel('<a href="#" style="color: #999; text-decoration: none; font-size: 10px;">Автор</a>')
        self.lbl_author.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        self.lbl_author.linkActivated.connect(self.show_author)
        bottom_row.addWidget(self.lbl_author)
        ml.addLayout(bottom_row)
        
        self.bs.clicked.connect(self.pick_s); self.bd.clicked.connect(self.pick_d); self.btn_start.clicked.connect(self.toggle); self.btn_cancel.clicked.connect(self.cancel); self.btn_rej.clicked.connect(self.open_rejected)
        for w in [self.fb,self.al,self.mi,self.ma,self.up,self.chk_cs,self.pi,self.chk_rt,self.chk_rm,self.chk_pdf]: 
            if isinstance(w, (QComboBox, QCheckBox)): w.activated.connect(self.on_pm) if isinstance(w, QComboBox) else w.clicked.connect(self.on_pm)
            else: w.textEdited.connect(self.on_pm)

    def show_author(self): AuthorDialog(self).exec()
    def pick_watermark(self):
        f, _ = QFileDialog.getOpenFileName(self, "Выбрать водный знак", "", "Изображения (*.png *.jpg *.jpeg *.webp)")
        if f: self.watermark_path = f; self.lbl_wm.setText(os.path.basename(f)); self.lbl_wm.setStyleSheet("color:black;font-size:10px;"); self.btn_wm_clear.setVisible(True)
    def clear_watermark(self): self.watermark_path = ""; self.lbl_wm.setText(""); self.btn_wm_clear.setVisible(False)
    def on_pm(self):
        if self._ap: return
        if self.pc.currentText()!="Пользовательский": self.pc.blockSignals(True); self.pc.setCurrentText("Пользовательский"); self.pc.blockSignals(False)
    def on_ps(self): self._ap=True; self.apply_p(self.pc.currentText()); self._ap=False
    def apply_p(self, n):
        if n not in PRESETS: return
        p=PRESETS[n]; self.mi.setText(str(p.min_px)); self.ma.setText(str(p.max_px)); self.up.setText(str(p.max_upscale_pct)); self.al.setCurrentText("По высоте" if p.align=="height" else "По ширине"); self.fb.setCurrentText(p.fmt); self.chk_cs.setChecked(p.center_square); self.pi.setText(str(p.padding_pct)); self.chk_rm.setChecked(p.remove_meta); self.chk_rt.setChecked(p.replace_transparent)
    def on_stc(self):
        st=self.stc.currentText(); self.weo.setVisible(st=="Excel"); self.wer.setVisible(st=="Excel")
        self.source_path=None; self.ls.setText("Не выбран"); self.ls.setStyleSheet("color:gray;"); self.btn_start.setText("▶ ЗАПУСТИТЬ")
    def dragEnterEvent(self,e): 
        if e.mimeData().hasUrls(): e.acceptProposedAction()
    def dropEvent(self,e):
        p=e.mimeData().urls()[0].toLocalFile()
        if os.path.isfile(p) and p.lower().endswith((".xlsx",".xls")): self.stc.setCurrentText("Excel"); self.on_stc(); self.set_s(p)
        elif os.path.isdir(p): self.stc.setCurrentText("Папка"); self.on_stc(); self.set_s(p)
    def set_s(self, p): self.source_path=p; self.ls.setText(os.path.basename(p) if os.path.isfile(p) else p); self.ls.setStyleSheet("color:black;")
    def add_lr(self, s, so, m): self.log_buffer.append((s,so,m))
    def flush_log(self):
        if not self.log_buffer: return
        b=self.log_buffer[:]; self.log_buffer.clear(); self.lt.setUpdatesEnabled(False)
        for s,so,m in b:
            r=self.lt.rowCount(); self.lt.insertRow(r)
            for c,t in enumerate([s,so,m]): self.lt.setItem(r,c,QTableWidgetItem(t))
        self.lt.setUpdatesEnabled(True)
        if not (self.log_dialog and self.log_dialog.isVisible()): self.lt.scrollToBottom()
        if self.log_dialog and self.log_dialog.isVisible(): self.log_dialog.add_rows_batch(b)
    def show_log_d(self):
        if not self.log_dialog:
            self.log_dialog=LogDialog(self); rows=[]
            for r in range(self.lt.rowCount()): rows.append((self.lt.item(r,0).text(),self.lt.item(r,1).text(),self.lt.item(r,2).text()))
            self.log_dialog.add_rows_batch(rows)
        self.log_dialog.show(); self.log_dialog.raise_()
    def show_info(self): 
        QMessageBox.information(self,"Инструкция",
            "Общий алгоритм:\n"
            "1. Выберите источник (Excel с ссылками или Папка с фото).\n"
            "2. Укажите место назначения.\n"
            "3. Настройте параметры или выберите пресет.\n"
            "4. Нажмите ЗАПУСТИТЬ.\n"
            "5. Если есть отбракованные фото (серый фон, размер), появится оранжевая кнопка. Выберите проблемные фото и нажмите 'Принудительно преобразовать'.\n"
            "6. С каждым нажатием порог определения белого фона ослабляется (255→240→220...), что позволяет обрезать всё более серые поля.")
    def pick_s(self):
        if self.stc.currentText()=="Excel":
            f,_=QFileDialog.getOpenFileName(self,"Excel","","Excel (*.xlsx *.xls)")
            if f and os.path.exists(f): self.set_s(f)
        else:
            f=QFileDialog.getExistingDirectory(self,"Папка с фото")
            if f: self.set_s(f)
    def pick_d(self):
        f=QFileDialog.getExistingDirectory(self,"Назначение")
        if f: self.output_dir=f; self.ld.setText(f); self.ld.setStyleSheet("color:black;")
    def gather(self):
        try: return AppSettings(source=self.source_path or "",source_type=self.stc.currentText(),out_dir=self.output_dir or os.path.join(os.path.expanduser("~"),"Desktop"),article_col=self.ca.currentText().upper(),url_from=self.cuf.currentText().upper(),url_to=self.cut.currentText().upper(),rename_mode="article" if self.cr.currentText()=="По артикулу" else "original",folder_sort=self.cfs.currentText(),min_px=int(self.mi.text()),max_px=int(self.ma.text()),max_upscale_pct=int(self.up.text()),align="height" if self.al.currentText()=="По высоте" else "width",fmt=self.fb.currentText(),center_square=self.chk_cs.isChecked(),padding_pct=int(self.pi.text()),replace_transparent=self.chk_rt.isChecked(),remove_meta=self.chk_rm.isChecked(),report_copy=self.chk_rc.isChecked(),clean_raw=self.chk_cr.isChecked(),ssl_verify=self.chk_ssl.isChecked(),preset_name=self.pc.currentText(),pdf_always_square=self.chk_pdf.isChecked(),selected_rejected_files=self._srf,white_threshold=max(80, 255 - (self.force_threshold_level * 20)),watermark_path=self.watermark_path)
        except: raise ValueError("Ошибка в числах!")
    def toggle(self):
        if self.worker and self.worker.isRunning():
            if self.worker._p: self.worker.resume(); self.btn_start.setText("⏸ ОСТАНОВИТЬ")
            else: self.worker.pause(); self.btn_start.setText("▶ ПРОДОЛЖИТЬ")
            return
        if not self.source_path: QMessageBox.warning(self,"Ошибка","Выберите источник"); return
        if not self.output_dir:
            if QMessageBox.question(self,"Внимание","Сохранить на Рабочий стол?")==QMessageBox.No: return
            self.output_dir=os.path.join(os.path.expanduser("~"),"Desktop")
        if self.stc.currentText()=="Excel" and is_excel_locked(self.source_path): QMessageBox.warning(self,"Внимание","Закройте Excel!"); return
        try: s=self.gather()
        except ValueError as e: QMessageBox.warning(self,"Ошибка",str(e)); return
        self.force_threshold_level = 0
        self.start_w(s)
    def start_w(self, s):
        self.lt.setRowCount(0); self.pw.setVisible(True); self.pw.setValue(0); self.btn_cancel.setVisible(True); self.btn_rej.setVisible(False); self.btn_start.setText("⏸ ОСТАНОВИТЬ")
        self.worker=Worker(s)
        self.worker.progress.connect(self.upd_p); self.worker.stats_updated.connect(self.upd_s); self.worker.log_row.connect(self.add_lr); self.worker.finished.connect(self.done); self.worker.error.connect(self.fail); self.worker.download_errors_ready.connect(self.on_download_errors)
        self.worker.start()
    def cancel(self):
        if self.worker and self.worker.isRunning(): self.worker.cancel(); self.btn_cancel.setEnabled(False)
    def upd_p(self, v): self.pw.setValue(v)
    def upd_s(self, ok, ds, dw, f, b, t): 
        mins = int(t) // 60; secs = int(t) % 60
        self.ls2.setText(f"✅ ОК: {ok} | 📏 Размер: {ds} | ❓ Белый: {dw} | ❌ Ошибки: {f} | 💾 {b/(1024*1024):.1f} МБ | ⏱ {mins:02d}:{secs:02d}")
    def on_download_errors(self, errors): self.last_download_errors = errors
    def done(self, stats):
        self.pw.setValue(100); self.btn_start.setText("▶ ЗАПУСТИТЬ"); self.btn_cancel.setVisible(False); self.btn_cancel.setEnabled(True); self.save_s()
        self.last_processed_dir = stats.get("processed_dir")
        self.add_lr("[СИСТЕМА]","Готово",f"Время {int(stats.get('time',0))//60:02d}:{int(stats.get('time',0))%60:02d}"); self.flush_log()
        has_issues = (stats.get("defect_white",0) > 0 or stats.get("defect_size",0) > 0 or self.last_download_errors)
        if has_issues: self.btn_rej.setVisible(True)
    def fail(self, msg):
        self.pw.setVisible(False); self.btn_start.setText("▶ ЗАПУСТИТЬ"); self.btn_cancel.setVisible(False); self.add_lr("[ОШИБКА]","Критическая",msg[:300]); self.flush_log(); QMessageBox.critical(self,"Ошибка",msg)
    def open_rejected(self):
        target_dir = self.last_processed_dir
        if not target_dir or not os.path.exists(target_dir):
            if not self.output_dir or not os.path.exists(self.output_dir): return
            dirs = [os.path.join(self.output_dir, d) for d in os.listdir(self.output_dir) if d.startswith("обработано_")]
            if dirs: target_dir = max(dirs, key=os.path.getmtime)
            else:
                goto = os.path.join(self.output_dir, "Готово")
                if os.path.exists(goto): target_dir = goto
                else: return
        try: s=self.gather()
        except: return
        dlg=ProductPreviewDialog(target_dir, s, self, download_errors=self.last_download_errors, threshold_level=self.force_threshold_level)
        dlg.files_selected.connect(self.on_rej_sel); dlg.exec()
    def on_rej_sel(self, files):
        self._srf=files
        if not files: return
        try: 
            self.force_threshold_level += 1
            s=self.gather(); s.source_type="Отбракованное"
        except: return
        self.start_w(s)
    def save_s(self):
        s=self.settings_io; s.setValue("source",self.source_path or ""); s.setValue("out_dir",self.output_dir or ""); s.setValue("st",self.stc.currentText()); s.setValue("art",self.ca.currentText()); s.setValue("uf",self.cuf.currentText()); s.setValue("ut",self.cut.currentText()); s.setValue("rm",self.cr.currentText()); s.setValue("fs",self.cfs.currentText()); s.setValue("mi",self.mi.text()); s.setValue("ma",self.ma.text()); s.setValue("up",self.up.text()); s.setValue("fmt",self.fb.currentText()); s.setValue("al",self.al.currentText()); s.setValue("meta",self.chk_rm.isChecked()); s.setValue("cs",self.chk_cs.isChecked()); s.setValue("pp",self.pi.text()); s.setValue("rt",self.chk_rt.isChecked()); s.setValue("rc",self.chk_rc.isChecked()); s.setValue("clr",self.chk_cr.isChecked()); s.setValue("ssl",self.chk_ssl.isChecked()); s.setValue("pr",self.pc.currentText()); s.setValue("pdf",self.chk_pdf.isChecked())
    def load_settings(self):
        s=self.settings_io; self.source_path=s.value("source",""); self.output_dir=s.value("out_dir","")
        self.stc.setCurrentText(s.value("st","Excel")); self.on_stc()
        if self.source_path and os.path.exists(self.source_path): self.ls.setText(os.path.basename(self.source_path)); self.ls.setStyleSheet("color:black;")
        if self.output_dir and os.path.exists(self.output_dir): self.ld.setText(self.output_dir); self.ld.setStyleSheet("color:black;")
        self.ca.setCurrentText(s.value("art","A")); self.cuf.setCurrentText(s.value("uf","B")); self.cut.setCurrentText(s.value("ut","P")); self.cr.setCurrentText(s.value("rm","По артикулу")); self.cfs.setCurrentText(s.value("fs","В порядке Excel")); self.mi.setText(s.value("mi","0")); self.ma.setText(s.value("ma","4000")); self.up.setText(s.value("up","50")); self.fb.setCurrentText(s.value("fmt","jpg")); self.al.setCurrentText(s.value("al","По высоте")); self.chk_rm.setChecked(s.value("meta",False,type=bool)); self.chk_cs.setChecked(s.value("cs",True,type=bool)); self.pi.setText(s.value("pp","10")); self.chk_rt.setChecked(s.value("rt",True,type=bool)); self.chk_rc.setChecked(s.value("rc",False,type=bool)); self.chk_cr.setChecked(s.value("clr",False,type=bool)); self.chk_ssl.setChecked(s.value("ssl",True,type=bool)); self.chk_pdf.setChecked(s.value("pdf",True,type=bool))
        pr=s.value("pr","santehnica.ru")
        if pr in PRESETS: self._ap=True; self.pc.setCurrentText(pr); self.apply_p(pr); self._ap=False

if __name__ == "__main__":
    sys.excepthook = lambda t,v,b: (logging.getLogger("Valera").critical("".join(traceback.format_exception(t,v,b))), QMessageBox.critical(None,"Ошибка","Критическая ошибка!"))
    app=QApplication(sys.argv)
    setup_logging()
    apply_light_theme(app)
    w=MainWindow()
    w.show()
    sys.exit(app.exec())