import sys
import os
import json
import ctypes
from ctypes import wintypes
import win32gui, win32con, win32process
import psutil
from PyQt5 import QtCore, QtGui, QtWidgets

# ===== DPI対応 =====
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    pass

# ===== DWM API =====
dwmapi = ctypes.windll.dwmapi

class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)]

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_int), ("cy", ctypes.c_int)]

class DWM_THUMBNAIL_PROPERTIES(ctypes.Structure):
    _fields_ = [
        ("dwFlags", ctypes.c_uint),
        ("rcDestination", RECT),
        ("rcSource", RECT),
        ("opacity", ctypes.c_ubyte),
        ("fVisible", wintypes.BOOL),
        ("fSourceClientAreaOnly", wintypes.BOOL),
    ]

DwmRegisterThumbnail         = dwmapi.DwmRegisterThumbnail
DwmUnregisterThumbnail       = dwmapi.DwmUnregisterThumbnail
DwmQueryThumbnailSourceSize  = dwmapi.DwmQueryThumbnailSourceSize
DwmUpdateThumbnailProperties = dwmapi.DwmUpdateThumbnailProperties

DWM_TNP_RECTDESTINATION      = 0x00000001
DWM_TNP_RECTSOURCE           = 0x00000002
DWM_TNP_OPACITY              = 0x00000004
DWM_TNP_VISIBLE              = 0x00000008
DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010

# ===== ディレクトリ構成 =====
BASE_DIR = "overlay_settings"
WINDOW_DIR = os.path.join(BASE_DIR, "window_config")
EXE_DIR = os.path.join(BASE_DIR, "exe_config")
os.makedirs(WINDOW_DIR, exist_ok=True)
os.makedirs(EXE_DIR, exist_ok=True)

# ===== 設定ロード／保存 =====
def sanitize_filename(name: str):
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name.strip() or "noname"

def load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_json(data, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_config(title: str, exe: str):
    title_path = os.path.join(WINDOW_DIR, f"{sanitize_filename(title)}.json")
    exe_path = os.path.join(EXE_DIR, f"{sanitize_filename(exe)}.json")
    if os.path.exists(title_path):
        print(f"📄 ウィンドウ設定ロード: {title_path}")
        return load_json(title_path)
    elif os.path.exists(exe_path):
        print(f"📄 EXE設定ロード: {exe_path}")
        return load_json(exe_path)
    return {}

def save_config(data: dict, title: str, exe: str):
    title_path = os.path.join(WINDOW_DIR, f"{sanitize_filename(title)}.json")
    exe_path = os.path.join(EXE_DIR, f"{sanitize_filename(exe)}.json")
    save_json(data, exe_path)
    save_json(data, title_path)
    print(f"💾 設定保存: {title_path} / {exe_path}")

# ===== ウィンドウ列挙 =====
def list_visible_windows():
    """有効なアプリケーションウィンドウのみ列挙"""
    windows = []

    def enum_cb(hwnd, _):
        # 非表示・最小化ウィンドウを除外
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
            return

        # タイトル取得（空白は除外）
        title = win32gui.GetWindowText(hwnd)
        if not title or not title.strip():
            return

        # サイズ取得（極小は除外）
        try:
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            if width < 80 or height < 80:
                return
        except Exception:
            return

        # 子ウィンドウや内部UIを除外
        GA_ROOT = 2
        root = ctypes.windll.user32.GetAncestor(hwnd, GA_ROOT)
        if root != hwnd:
            return

        # 実行ファイル名取得
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            exe = psutil.Process(pid).name()
        except Exception:
            exe = "Unknown"

        windows.append((hwnd, exe, title))

    win32gui.EnumWindows(enum_cb, None)

    # 重複や無効タイトル除外
    unique = []
    seen = set()
    for hwnd, exe, title in windows:
        key = (exe.lower(), title)
        if key not in seen:
            seen.add(key)
            unique.append((hwnd, exe, title))
    return unique


# ===== Overlay =====
class Overlay(QtWidgets.QWidget):
    def __init__(self, target_hwnd, exe_name, title):
        super().__init__()
        self.target_hwnd = target_hwnd
        self.exe_name = exe_name
        self.title = title
        self.hthumb = wintypes.HANDLE(0)
        self.src_size = SIZE(0, 0)

        # 状態変数
        self.dragging = False
        self.mode = None
        self.start_pos = QtCore.QPoint(0, 0)
        self.start_size = QtCore.QSize(0, 0)
        self.start_win_pos = QtCore.QPoint(0, 0)
        self.start_crop = QtCore.QRect(0, 0, 0, 0)
        self.start_opacity = 0.85
        self.crop = QtCore.QRect(0, 0, 640, 360)

        # UI
        self.top_margin = 30
        self.ctrl_rect = QtCore.QRect(8, 8, 28, 28)
        self.ctrl_hover = False

        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint |
                            QtCore.Qt.FramelessWindowHint |
                            QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # 設定をロード
        cfg = load_config(title, exe_name)
        if cfg:
            x, y = cfg.get("pos", [100, 100])
            w, h = cfg.get("size", [640, 360])
            crop = cfg.get("crop", [0, 0, w, h])
            opacity = cfg.get("opacity", 0.85)
            self.setGeometry(x, y, w, h)
            self.crop = QtCore.QRect(*crop)
            self.setWindowOpacity(opacity)
            self.update_thumbnail_props()  # ←★ これを追加！
        else:
            self.setGeometry(100, 100, 640, 380)
            self.setWindowOpacity(0.85)
            self.update_thumbnail_props()


        self.register_thumbnail()
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.refresh)
        self.timer.start(100)

    def register_thumbnail(self):
        dest_hwnd = int(self.winId())
        hr = DwmRegisterThumbnail(dest_hwnd, self.target_hwnd, ctypes.byref(self.hthumb))
        if hr != 0 or not self.hthumb.value:
            raise RuntimeError("DwmRegisterThumbnail failed")

        DwmQueryThumbnailSourceSize(self.hthumb, ctypes.byref(self.src_size))

        # 👇ここを変更！
        # 保存されたcropがまだデフォルト(0,0,640,360)なら、初期化してOK
        if self.crop == QtCore.QRect(0, 0, 640, 360) or self.crop.isNull():
            self.crop = QtCore.QRect(0, 0, self.src_size.cx, self.src_size.cy)

        self.update_thumbnail_props()


    def reselect_window(self):
        """R+クリック: 新しいウィンドウを再選択（保存＋再読み込み対応）"""
        # --- 現在の設定を保存しておく ---
        current_data = {
            "crop": [self.crop.left(), self.crop.top(), self.crop.right(), self.crop.bottom()],
            "pos": [self.x(), self.y()],
            "size": [self.width(), self.height()],
            "opacity": self.windowOpacity(),
        }
        save_config(current_data, self.title, self.exe_name)
        print(f"💾 現在の設定を保存: [{self.exe_name}] {self.title}")

        wins = list_visible_windows()
        if not wins:
            print("ウィンドウが見つかりません")
            return

        # 非モーダル選択ダイアログ
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("ウィンドウ再選択")
        dialog.setWindowModality(QtCore.Qt.NonModal)
        dialog.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        dialog.setGeometry(self.x() + 50, self.y() + 50, 400, 300)

        layout = QtWidgets.QVBoxLayout(dialog)
        listbox = QtWidgets.QListWidget()
        for hwnd, exe, title in wins:
            listbox.addItem(f"[{exe}] {title}")
        layout.addWidget(listbox)

        btn_ok = QtWidgets.QPushButton("OK")
        btn_cancel = QtWidgets.QPushButton("キャンセル")
        btns = QtWidgets.QHBoxLayout()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)

        def on_ok():
            sel = listbox.currentRow()
            if sel >= 0:
                hwnd, exe, title = wins[sel]
                if self.hthumb.value:
                    DwmUnregisterThumbnail(self.hthumb)
                self.target_hwnd = hwnd
                self.exe_name = exe
                self.title = title
                self.register_thumbnail()
                print(f"🔁 再選択: [{exe}] {title}")

                # --- 新しいウィンドウの設定をロード ---
                cfg = load_config(title, exe)
                if cfg:
                    x, y = cfg.get("pos", [self.x(), self.y()])
                    w, h = cfg.get("size", [self.width(), self.height()])
                    crop = cfg.get("crop", [0, 0, w, h])
                    opacity = cfg.get("opacity", self.windowOpacity())
                    self.setGeometry(x, y, w, h)
                    self.crop = QtCore.QRect(*crop)
                    self.setWindowOpacity(opacity)
                    print(f"📄 設定を読み込み: [{exe}] {title}")
                else:
                    print(f"⚠ 新しい設定が見つかりません。デフォルトで開始。")
            dialog.close()

        def on_cancel():
            dialog.close()

        btn_ok.clicked.connect(on_ok)
        btn_cancel.clicked.connect(on_cancel)

        dialog.show()

        
    def show_shortcuts(self):
        print("""
===== 🎮 Overlay 操作一覧 =====
🟥 左クリック＋ドラッグ        : 移動
🟥 Ctrl＋ドラッグ             : サイズ変更
🟥 Alt＋ドラッグ              : トリミング（右ドラッグで左を削る）
🟥 X＋ドラッグ                : 透過率変更（上で濃く、下で薄く）
🟥 Z＋クリック                : トリミング初期化
🟥 R＋クリック                : ウィンドウ再選択
🟥 H＋クリック                : このヘルプを表示
🟥 C＋クリック                : 終了
==============================
""")

    def update_thumbnail_props(self):
        w, h = self.width(), self.height()
        dest = RECT(0, self.top_margin, w, h)
        src = RECT(self.crop.left(), self.crop.top(),
                   self.crop.right(), self.crop.bottom())
        props = DWM_THUMBNAIL_PROPERTIES()
        props.dwFlags = (DWM_TNP_RECTDESTINATION | DWM_TNP_RECTSOURCE |
                         DWM_TNP_VISIBLE | DWM_TNP_OPACITY)
        props.rcDestination = dest
        props.rcSource = src
        props.opacity = 255
        props.fVisible = True
        props.fSourceClientAreaOnly = False
        DwmUpdateThumbnailProperties(self.hthumb, ctypes.byref(props))

    def refresh(self):
        """定期的にサムネイルを更新。止まった場合は自動で復旧"""
        if not win32gui.IsWindow(self.target_hwnd):
            return

        # ウィンドウが最小化・非表示になった場合
        if win32gui.IsIconic(self.target_hwnd) or not win32gui.IsWindowVisible(self.target_hwnd):
            if hasattr(self, "_minimized") and not self._minimized:
                print("🟡 ウィンドウが非表示/最小化されました。一時停止中...")
            self._minimized = True
            return
        else:
            if hasattr(self, "_minimized") and self._minimized:
                print("🟢 ウィンドウが再表示されました。サムネイル再登録中...")
                try:
                    if self.hthumb.value:
                        DwmUnregisterThumbnail(self.hthumb)
                except Exception:
                    pass
                self.register_thumbnail()
                self._minimized = False

        # DWMサムネイルの有効性を確認（途切れたら再登録）
        try:
            test_props = DWM_THUMBNAIL_PROPERTIES()
            DwmUpdateThumbnailProperties(self.hthumb, ctypes.byref(test_props))
        except Exception:
            print("🔄 サムネイルが無効になったため再登録します")
            try:
                DwmUnregisterThumbnail(self.hthumb)
            except Exception:
                pass
            self.register_thumbnail()

        # 通常更新
        self.update_thumbnail_props()

    def nativeEvent(self, eventType, message):
        if eventType == "windows_generic_MSG":
            msg = ctypes.wintypes.MSG.from_address(message.__int__())
            if msg.message == win32con.WM_NCHITTEST:
                pos = QtGui.QCursor.pos()
                p = self.mapFromGlobal(pos)
                if self.ctrl_rect.contains(p):
                    return True, win32con.HTCLIENT
                else:
                    return True, win32con.HTTRANSPARENT
        return False, 0

    # ==== マウス ====
    def mousePressEvent(self, e):
        if e.button() != QtCore.Qt.LeftButton or not self.ctrl_rect.contains(e.pos()):
            return

        get = lambda k: bool(ctypes.windll.user32.GetAsyncKeyState(ord(k)) & 0x8000)
        z, c, r, h = map(get, "ZCRH")

        if h:
            self.show_shortcuts()
            return
        if r:
            # ✅ 非モーダルで再選択（終了しない）
            self.reselect_window()
            return
        if z:
            self.crop = QtCore.QRect(0, 0, self.src_size.cx, self.src_size.cy)
            self.update_thumbnail_props()
            print("✅ トリミング初期化")
            return
        if c:
            print("🛑 終了")
            self.close()
            QtCore.QCoreApplication.quit()
            return
            
        # 通常操作開始
        self.dragging = True
        mods = QtWidgets.QApplication.queryKeyboardModifiers()
        x_pressed = get("X")
        if mods == QtCore.Qt.ControlModifier:
            self.mode = "resize"
        elif mods == QtCore.Qt.AltModifier:
            self.mode = "trim"
        elif x_pressed:
            self.mode = "opacity"
        else:
            self.mode = "move"

        self.start_pos = e.globalPos()
        self.start_size = self.size()
        self.start_win_pos = self.pos()
        self.start_crop = QtCore.QRect(self.crop)
        self.start_opacity = self.windowOpacity()

    def mouseMoveEvent(self, e):
        self.ctrl_hover = self.ctrl_rect.contains(e.pos())
        if not self.dragging:
            self.update()
            return
        delta = e.globalPos() - self.start_pos
        dx, dy = delta.x(), delta.y()
        if self.mode == "move":
            self.move(self.start_win_pos + delta)
        elif self.mode == "resize":
            self.resize(max(100, self.start_size.width() + dx),
                        max(80, self.start_size.height() + dy))
            self.update_thumbnail_props()
        elif self.mode == "trim":
            self.adjust_crop(dx, dy)
        elif self.mode == "opacity":
            new_opacity = self.start_opacity - (dy * 0.003)
            new_opacity = max(0.2, min(1.0, new_opacity))
            self.setWindowOpacity(new_opacity)
        self.update()

    def adjust_crop(self, dx, dy):
        left, top, right, bottom = (
            self.start_crop.left(),
            self.start_crop.top(),
            self.start_crop.right(),
            self.start_crop.bottom(),
        )
        if dx > 0: left += dx
        elif dx < 0: right += dx
        if dy > 0: top += dy
        elif dy < 0: bottom += dy
        min_w, min_h = 50, 50
        if right < left + min_w: right = left + min_w
        if bottom < top + min_h: bottom = top + min_h
        left   = max(0, min(left,   self.src_size.cx-1))
        right  = max(1, min(right,  self.src_size.cx))
        top    = max(0, min(top,    self.src_size.cy-1))
        bottom = max(1, min(bottom, self.src_size.cy))
        self.crop = QtCore.QRect(QtCore.QPoint(left, top),
                                 QtCore.QPoint(right, bottom))
        self.update_thumbnail_props()

    def mouseReleaseEvent(self, e):
        self.dragging = False
        self.mode = None

    def paintEvent(self, e):
        painter = QtGui.QPainter(self)
        color = QtGui.QColor(255, 100, 100, 255 if self.ctrl_hover else 210)
        painter.fillRect(self.ctrl_rect, color)
        painter.setPen(QtGui.QPen(QtGui.QColor(255,255,255,200)))
        font = painter.font()
        font.setPointSize(8)
        painter.setFont(font)
        painter.drawText(
            self.ctrl_rect.adjusted(32, 6, 0, 0),
            QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter,
            f"[{self.exe_name}] {self.title[:40]} | Alt:トリム / Ctrl:サイズ / X:透過 / R:再選択 / H:ヘルプ / C:終了"
        )

    def closeEvent(self, e):
        if self.hthumb.value:
            DwmUnregisterThumbnail(self.hthumb)
        data = {
            "crop": [self.crop.left(), self.crop.top(), self.crop.right(), self.crop.bottom()],
            "pos": [self.x(), self.y()],
            "size": [self.width(), self.height()],
            "opacity": self.windowOpacity(),
        }
        save_config(data, self.title, self.exe_name)
        super().closeEvent(e)


# ===== 実行 =====
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    wins = list_visible_windows()
    if not wins:
        print("ウィンドウが見つかりません")
        sys.exit(1)
    items = [f"[{exe}] {title}" for hwnd, exe, title in wins]
    item, ok = QtWidgets.QInputDialog.getItem(None, "ウィンドウ選択",
                                              "プレビューするウィンドウを選んでください:",
                                              items, 0, False)
    if not ok:
        sys.exit(0)
    hwnd, exe, title = wins[items.index(item)]
    overlay = Overlay(hwnd, exe, title)
    overlay.show()
    sys.exit(app.exec_())
