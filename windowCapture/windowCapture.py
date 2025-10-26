import sys, ctypes, asyncio, numpy as np
import win32gui, win32process, psutil, win32con
from PyQt5 import QtCore, QtGui, QtWidgets
from ctypes import wintypes
import winrt.windows.graphics.capture as wgc
import winrt.windows.graphics.directx.direct3d11.interop as d3d11_interop
import winrt.windows.graphics.capture.interop as capture_interop
import winrt.windows.graphics.imaging as imaging
import winrt.windows.storage.streams as streams


# =======================================================
# D3D / WinRT capture
# =======================================================
def create_d3d_device_idirect3d():
    from ctypes import c_void_p, c_uint, POINTER, byref
    d3d11 = ctypes.windll.d3d11
    D3D_DRIVER_TYPE_HARDWARE = 1
    D3D11_CREATE_DEVICE_BGRA_SUPPORT = 0x20
    D3D11_SDK_VERSION = 7
    D3D11CreateDevice = d3d11.D3D11CreateDevice
    D3D11CreateDevice.argtypes = [
        c_void_p, c_uint, c_void_p, c_uint, c_void_p, c_uint, c_uint,
        POINTER(c_void_p), POINTER(c_uint), POINTER(c_void_p)
    ]
    pDev, pCtx = c_void_p(), c_void_p()
    feat = c_uint()
    hr = D3D11CreateDevice(None, D3D_DRIVER_TYPE_HARDWARE, None,
                           D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                           None, 0, D3D11_SDK_VERSION,
                           byref(pDev), byref(feat), byref(pCtx))
    if hr != 0:
        raise OSError(f"D3D11CreateDevice failed (HRESULT=0x{hr:08X})")
    return d3d11_interop.create_direct3d11_device_from_dxgi_device(pDev.value)


def softwarebitmap_to_numpy(sb: imaging.SoftwareBitmap):
    try:
        if sb.bitmap_pixel_format != imaging.BitmapPixelFormat.BGRA8:
            sb = imaging.SoftwareBitmap.convert(sb, imaging.BitmapPixelFormat.BGRA8)
        h, w = sb.pixel_height, sb.pixel_width
        buf = streams.Buffer(w * h * 4)
        sb.copy_to_buffer(buf)
        sb.close()
        reader = streams.DataReader.from_buffer(buf)
        ibuf = reader.read_buffer(int(buf.length))
        reader2 = streams.DataReader.from_buffer(ibuf)
        data_bytes = bytearray(int(ibuf.length))
        reader2.read_bytes(data_bytes)
        arr = np.frombuffer(data_bytes, dtype=np.uint8).reshape((h, w, 4))
        return arr[:, :, :3][:, :, ::-1].copy()
    except Exception as e:
        print("bitmap convert error:", e)
        return np.zeros((100, 100, 3), np.uint8)


class WinRTCapture(QtCore.QThread):
    new_frame = QtCore.pyqtSignal(np.ndarray)
    def __init__(self, hwnd): super().__init__(); self.hwnd = hwnd; self.running = True
    def run(self):
        asyncio.set_event_loop(asyncio.new_event_loop())
        device = create_d3d_device_idirect3d()
        item = capture_interop.create_for_window(self.hwnd)
        size = item.size
        pool = wgc.Direct3D11CaptureFramePool.create(device, 87, 2, size)
        session = pool.create_capture_session(item)
        try:
            session.is_cursor_capture_enabled = False
            session.is_border_required = False
        except: pass
        session.start_capture()
        while self.running:
            frame = pool.try_get_next_frame()
            if frame:
                try:
                    op = imaging.SoftwareBitmap.create_copy_from_surface_async(frame.surface)
                    sb = asyncio.get_event_loop().run_until_complete(op)
                    img = softwarebitmap_to_numpy(sb)
                    self.new_frame.emit(img)
                except Exception as e:
                    print("frame error:", e)
                finally:
                    frame.close()
            self.msleep(16)
        session.close(); pool.close()
    def stop(self): self.running = False


# =======================================================
# Overlay window (transparent capture display)
# =======================================================
class Overlay(QtWidgets.QWidget):
    def __init__(self, hwnd, exe, title):
        super().__init__()
        self.hwnd, self.exe, self.title = hwnd, exe, title
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint |
                            QtCore.Qt.Tool |
                            QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setGeometry(100, 100, 800, 480)

        # キャプチャ開始
        self.frame_pix = QtGui.QPixmap()
        self.cap = WinRTCapture(hwnd)
        self.cap.new_frame.connect(self.on_frame)
        self.cap.start()

        # 全体クリック透過ON
        self.set_click_through(True)

        # 操作用ボタンウィンドウを別ウィンドウとして生成
        self.ctrl_window = ControlWindow(self)
        self.ctrl_window.show()
        self.ctrl_window.raise_()
        self.ctrl_window.activateWindow()
        print("[UI] ControlWindow created and raised to front")

    def set_click_through(self, enable: bool):
        hwnd = int(self.winId())
        GWL_EXSTYLE = -20
        WS_EX_TRANSPARENT = 0x20
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if enable:
            style |= WS_EX_TRANSPARENT
        else:
            style &= ~WS_EX_TRANSPARENT
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)

    def on_frame(self, arr):
        h, w, _ = arr.shape
        img = QtGui.QImage(arr.tobytes(), w, h, 3*w, QtGui.QImage.Format_RGB888)
        self.frame_pix = QtGui.QPixmap.fromImage(img)
        self.update()

    def paintEvent(self, e):
        p = QtGui.QPainter(self)
        if not self.frame_pix.isNull():
            p.drawPixmap(self.rect(), self.frame_pix)
        p.setPen(QtGui.QPen(QtGui.QColor("white")))
        f = p.font(); f.setPointSize(8); p.setFont(f)
        p.drawText(45, 25, f"[{self.exe}] {self.title[:40]}")

    def closeEvent(self, e):
        if self.cap and self.cap.isRunning():
            self.cap.stop(); self.cap.wait()
        self.ctrl_window.close()
        e.accept()


# =======================================================
# Control Window (red square, independent, clickable)
# =======================================================
class ControlWindow(QtWidgets.QWidget):
    dragStarted = QtCore.pyqtSignal(QtCore.QPoint)
    dragUpdated = QtCore.pyqtSignal(QtCore.QPoint)
    dragFinished = QtCore.pyqtSignal(QtCore.QPoint)

    def __init__(self, overlay):
        super().__init__()
        self.overlay = overlay
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint |
                            QtCore.Qt.Tool |
                            QtCore.Qt.WindowStaysOnTopHint |
                            QtCore.Qt.X11BypassWindowManagerHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAttribute(QtCore.Qt.WA_ShowWithoutActivating)

        # 表示矩形を操作ウィンドウの近くに配置
        target_pos = overlay.frameGeometry().topLeft() + QtCore.QPoint(40, 40)
        self.setGeometry(QtCore.QRect(target_pos, QtCore.QSize(36, 36)))
        self.setFixedSize(36, 36)

        # 実際に描画される赤い四角は子ウィジェットとして作成
        self.square = QtWidgets.QFrame(self)
        self.square.setObjectName("controlSquare")
        self.square.setGeometry(self.rect())
        self.square.setStyleSheet(
            "#controlSquare {"
            "background-color: rgba(255,60,60,220);"
            "border-radius: 5px;"
            "}"
        )

        # 最前面化を維持するため、一定間隔で SetWindowPos を呼び出す
        self._raise_timer = QtCore.QTimer(self)
        self._raise_timer.setInterval(750)
        self._raise_timer.timeout.connect(self.raise_to_top)
        self._raise_timer.start()
        QtCore.QTimer.singleShot(0, self.raise_to_top)

        self._dragging = False
        self._press_global = QtCore.QPoint()

    def raise_to_top(self):
        if not self.isVisible():
            self.show()
        self.raise_()
        hwnd = int(self.winId())
        ctypes.windll.user32.SetWindowPos(
            hwnd,
            win32con.HWND_TOPMOST,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
        )
        # 透明ウィンドウは描画更新を促さないと表示されない場合があるため
        self.square.update()

    def mousePressEvent(self, e):
        if e.button() == QtCore.Qt.LeftButton:
            self._dragging = True
            self._press_global = e.globalPos()
            print("🟥 操作用ドラッグ開始")
            self.dragStarted.emit(self._press_global)
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self._dragging and (e.buttons() & QtCore.Qt.LeftButton):
            delta = e.globalPos() - self._press_global
            self.dragUpdated.emit(delta)
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self._dragging and e.button() == QtCore.Qt.LeftButton:
            self._dragging = False
            delta = e.globalPos() - self._press_global
            print("🟥 操作用ドラッグ終了")
            self.dragFinished.emit(delta)
        super().mouseReleaseEvent(e)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "square"):
            self.square.setGeometry(self.rect())

    def closeEvent(self, event):
        if hasattr(self, "_raise_timer") and self._raise_timer.isActive():
            self._raise_timer.stop()
        super().closeEvent(event)

# =======================================================
# Window list & entry point
# =======================================================
def list_visible_windows():
    wins = []
    def enum_cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd): return
        title = win32gui.GetWindowText(hwnd)
        if not title.strip(): return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            exe = psutil.Process(pid).name()
        except: exe = "Unknown"
        wins.append((hwnd, exe, title))
    win32gui.EnumWindows(enum_cb, None)
    return wins


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    wins = list_visible_windows()
    if not wins:
        print("No windows found."); sys.exit(1)
    items = [f"[{exe}] {title}" for hwnd, exe, title in wins]
    item, ok = QtWidgets.QInputDialog.getItem(None, "Select window", "Capture target:", items, 0, False)
    if not ok: sys.exit(0)
    hwnd, exe, title = wins[items.index(item)]
    print(f"🎬 Target: {exe} - {title}")
    overlay = Overlay(hwnd, exe, title)
    overlay.show()
    sys.exit(app.exec_())
