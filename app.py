# -*- coding: utf-8 -*-
"""视觉拟人鼠标：截取任意窗口 -> OCR 识别文字位置 -> 按类人轨迹移动鼠标并单击。

依赖：pywin32、Pillow、Windows 自带 OCR（ocr-lines.ps1）。
用法：python app.py
"""

import ctypes
import json
import math
import os
import random
import statistics
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import urllib.request
from tkinter import ttk, messagebox, simpledialog

from PIL import Image, ImageTk

try:
    import numpy as np
except ImportError:
    np = None

# —— 版本 & 更新源 ——
__version__ = "1.0.0"
GITHUB_OWNER = "Jaymin-chang"
GITHUB_REPO = "VisionMouse"
GITHUB_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"

APP_DIR = os.path.dirname(os.path.abspath(__file__))
OCR_PS1 = os.path.join(APP_DIR, "ocr-lines.ps1")
CREATE_NO_WINDOW = 0x08000000
NEAR_BLACK = 5          # 灰度值 <= 5 视为纯黑，用来判断 PrintWindow 是否拿到画面


def dpi_aware():
    """坐标和截图都走物理像素，175% 缩放下也不会错位。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


dpi_aware()      # 导入即生效，任何入口拿到的都是物理像素坐标


# ── 窗口与截图 ────────────────────────────────────────────────────────────
def list_windows():
    """列出可截图的顶层窗口：[(hwnd, 标题)]。"""
    import win32con
    import win32gui
    found = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h) or win32gui.IsIconic(h):
                return True
            title = win32gui.GetWindowText(h).strip()
            if not title:
                # 自绘标题栏的窗口（如雷电「位置设置」WKE 对话框）GetWindowText 为空，
                # 用窗口类名做标签，不再跳过
                cls = win32gui.GetClassName(h).strip()
                if not cls:
                    return True
                title = "[%s]" % cls
            r = win32gui.GetWindowRect(h)
            if r[2] - r[0] < 120 or r[3] - r[1] < 120:
                return True
            found.append((h, title))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    # 再补收各顶层窗口内「有标题、可见、够大」的子窗口：
    # 雷电「位置设置」这类对话框是主窗体的子窗口，EnumWindows 枚举不到
    seen = {h for h, _ in found}
    for hwnd, _ in list(found):
        def child_cb(h, _):
            try:
                if h in seen or not win32gui.IsWindowVisible(h):
                    return True
                title = win32gui.GetWindowText(h).strip()
                if not title:
                    return True
                r = win32gui.GetWindowRect(h)
                if r[2] - r[0] < 120 or r[3] - r[1] < 120:
                    return True
                seen.add(h)
                found.append((h, title))
            except Exception:
                pass
            return True
        try:
            win32gui.EnumChildWindows(hwnd, child_cb, None)
        except Exception:
            pass
    return found


def _print_window(hwnd, w, h):
    """PrintWindow(PW_RENDERFULLCONTENT) 截整窗，窗口滑出屏幕/被遮挡也能截全。"""
    import win32gui
    import win32ui
    hdc = save = bmp = None
    try:
        hdc = win32gui.GetWindowDC(hwnd)
        src = win32ui.CreateDCFromHandle(hdc)
        save = src.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(src, w, h)
        save.SelectObject(bmp)
        if not ctypes.windll.user32.PrintWindow(hwnd, save.GetSafeHdc(), 2):
            return None
        info = bmp.GetInfo()
        img = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                               bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
        gray = img.convert("L").histogram()
        if sum(gray[:NEAR_BLACK]) > w * h * 0.9:      # 全黑说明没渲染出来
            return None
        return img.copy()
    except Exception:
        return None
    finally:
        try:
            if bmp is not None:
                win32gui.DeleteObject(bmp.GetHandle())
            if save is not None:
                save.DeleteDC()
        except Exception:
            pass
        try:
            if hdc:
                win32gui.ReleaseDC(hwnd, hdc)
        except Exception:
            pass


def capture_window(hwnd):
    """返回 (截图, 窗口矩形)。PrintWindow 失败时退回屏幕可见区域抓取。"""
    import win32gui
    r = list(win32gui.GetWindowRect(hwnd))
    w, h = r[2] - r[0], r[3] - r[1]
    if w < 120 or h < 120:
        return None, None
    img = _print_window(hwnd, w, h)
    if img is not None:
        return img, r
    from PIL import ImageGrab
    try:
        img = ImageGrab.grab(bbox=(r[0], r[1], r[2], r[3]))
    except Exception:
        return None, None
    return (img, r) if img is not None and img.size[0] > 10 else (None, None)


# ── 图像配准：相位相关求平移，比 OCR 文字坐标稳得多 ──────────────────────
def fft_shift(ref, cur):
    """ref/cur 为同尺寸灰度 float 数组。返回 (dx, dy, 置信度)：
    cur 的内容相当于 ref 平移了 (dx,dy)。置信度越高匹配越可靠。"""
    if np is None or ref.shape != cur.shape:
        return None
    f = np.fft.fft2(ref) * np.conj(np.fft.fft2(cur))
    m = np.abs(f)
    f /= np.where(m < 1e-9, 1e-9, m)         # 相位相关，抗亮度/对比度变化
    corr = np.abs(np.fft.ifft2(f))
    py, px = np.unravel_index(int(np.argmax(corr)), corr.shape)
    h, w = ref.shape
    dx = px - w if px > w // 2 else px
    dy = py - h if py > h // 2 else py
    conf = float(corr.max() / (corr.mean() + 1e-9))
    return float(-dx), float(-dy), conf      # 取负：cur 内容 = ref 平移 (dx,dy)


# ── OCR ─────────────────────────────────────────────────────────────────
_OCR_LOCK = threading.Lock()   # OCR 串行执行：实时同步和自动确认同时跑会抢 CPU、卡界面


def ocr_lines(img):
    """返回 [{'text','clean','x','y','w','h'}]，坐标为截图内的像素。"""
    with _OCR_LOCK:
        return _ocr_lines_impl(img)


def _ocr_lines_impl(img):
    scale = 1
    if max(img.size) < 2600:
        scale = 2                          # 小图放大 2 倍再 OCR，小字也能认出来
        img = img.resize((img.size[0] * scale, img.size[1] * scale), Image.LANCZOS)
    tag = "vision-mouse-%d-%d" % (os.getpid(), threading.get_ident())   # 每线程独立临时文件
    tmp = os.path.join(tempfile.gettempdir(), tag + ".png")
    out = os.path.join(tempfile.gettempdir(), tag + ".json")
    img.save(tmp)
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        OCR_PS1, "-ImagePath", tmp, "-OutFile", out],
                       capture_output=True, timeout=90, creationflags=CREATE_NO_WINDOW)
        try:
            with open(out, encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:
            return []
    finally:
        for p in (tmp, out):
            try:
                os.remove(p)
            except OSError:
                pass
    if isinstance(data, dict):          # 只认出一行时 ConvertTo-Json 会退化成对象
        data = [data]
    if not isinstance(data, list):
        return []
    rows = []
    for line in data:
        words = line.get("words") or []
        if not words:
            continue
        x = min(w["x"] for w in words) / scale
        y = min(w["y"] for w in words) / scale
        right = max(w["x"] + w["w"] for w in words) / scale
        h = max(w["h"] for w in words) / scale
        text = line.get("text") or ""
        rows.append({"text": text, "clean": "".join(text.split()),
                     "x": x, "y": y, "w": right - x, "h": h})
    return rows


# ── 鼠标动作 ────────────────────────────────────────────────────────────
def esc_pressed():
    import win32api
    return bool(win32api.GetAsyncKeyState(0x1B) & 0x8000)


def human_move(target, duration=None):
    """按三次贝塞尔 + 缓入缓出把鼠标移到 target，落点带轻微过冲回落。ESC 可中止。"""
    import win32api
    x0, y0 = win32api.GetCursorPos()
    x1, y1 = int(target[0]), int(target[1])
    dx, dy = x1 - x0, y1 - y0
    dist = math.hypot(dx, dy)
    if dist < 1:
        win32api.SetCursorPos((x1, y1))
        return True
    if duration is None:                # 远距离慢一点，近距离也留出人的手感
        duration = min(1.15, max(0.33, dist / 1300.0 + random.uniform(0.18, 0.40)))
    nx, ny = -dy, dx                    # 法向量，控制点沿垂直方向随机偏移走弧线
    n = math.hypot(nx, ny) or 1.0
    bend = min(dist * 0.30, 90.0)
    c1 = (x0 + dx * 0.30 + nx / n * random.uniform(-bend, bend) * 0.6,
          y0 + dy * 0.30 + ny / n * random.uniform(-bend, bend) * 0.6)
    c2 = (x0 + dx * 0.72 + nx / n * random.uniform(-bend, bend),
          y0 + dy * 0.72 + ny / n * random.uniform(-bend, bend))
    steps = max(14, int(duration * 120))
    for i in range(steps + 1):
        if esc_pressed():
            return False
        t = i / float(steps)
        e = 0.5 - 0.5 * math.cos(math.pi * t)     # 缓入缓出
        u = 1.0 - e
        x = u * u * u * x0 + 3 * u * u * e * c1[0] + 3 * u * e * e * c2[0] + e * e * e * x1
        y = u * u * u * y0 + 3 * u * u * e * c1[1] + 3 * u * e * e * c2[1] + e * e * e * y1
        win32api.SetCursorPos((int(round(x)), int(round(y))))
        time.sleep(duration / steps)
    return True


def settle(target):
    """落点先过冲几像素再回正，避免像机器一样精准落在同一点。"""
    import win32api
    x1, y1 = int(target[0]), int(target[1])
    win32api.SetCursorPos((x1 + random.randint(-3, 3), y1 + random.randint(-3, 3)))
    time.sleep(0.06)
    win32api.SetCursorPos((x1, y1))


def left_click(hold=None):
    import win32api
    import win32con
    hold = hold or random.uniform(0.07, 0.14)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(hold)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def bring_front(hwnd):
    import win32gui
    import win32con
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.25)


# ── 界面 ────────────────────────────────────────────────────────────────
class VisionMouseApp:
    def __init__(self, root):
        self.root = root
        self.img = None
        self.rect = None
        self.rows = []
        self.shown = []
        self.target = None          # 相对窗口的 (rx, ry)
        self.photo = None
        self.scale = 1.0
        self.offset = (0, 0)
        self.win_map = {}
        self.busy = False
        self.route = []                     # 记录连线的相对坐标序列 [(rx, ry), ...]
        self.avoid_rects = []               # 手动框选的禁区（记录帧归一化坐标）
        self.skip_idx = set()               # 点选跳过的标点序号
        self.mark_mode = None               # None / "rect" 框选 / "point" 点选
        self.rect_from = None
        self.recording = False
        self.looping = False
        self.stop_loop = threading.Event()
        self.live = False                   # 实时同步开关
        self.live_stop = threading.Event()
        self._live_sig = None               # 上一轮 OCR 结果签名，变了才刷新列表
        self.anchor_labels = None           # 地图锚定：记录时的地标快照 [(clean, cx, cy)]
        self.anchor_img = None              # 地图锚定：滚动基准画面（跑到远处才换）
        self.anchor0_img = None             # 地图锚定：记录时的原画面，绝对基准，不随跑动刷新
        self.anchor_base = (0.0, 0.0)       # 滚动基准画面本身对应的总偏移
        self.map_shift = (0.0, 0.0)         # 地图平移量（图像像素），渲染和点击都加它
        self.calibrating = False            # 两点校准模式
        self.calib = []

        root.title("视觉拟人鼠标 · 截图识别 + 类人点击")
        root.geometry("1600x1000")
        root.minsize(1100, 700)

        top = ttk.Frame(root, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="目标窗口").pack(side="left")
        self.combo = ttk.Combobox(top, state="readonly", width=52)
        self.combo.pack(side="left", padx=6)
        ttk.Button(top, text="刷新", command=self.refresh_windows).pack(side="left")
        ttk.Button(top, text="截图识别", command=self.start_capture).pack(side="left", padx=(10, 0))
        self.live_btn = ttk.Button(top, text="▶ 实时同步", command=self.toggle_live)
        self.live_btn.pack(side="left", padx=(6, 0))
        self.status = tk.StringVar(value="选窗口后点「截图识别」")
        ttk.Label(top, textvariable=self.status, foreground="#555").pack(side="left", padx=12)
        ttk.Button(top, text=f"关于 v{__version__}", command=self._about).pack(side="right")

        body = ttk.Frame(root, padding=(10, 0))
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        self.canvas = tk.Canvas(left, bg="#1b1b1f", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Configure>", lambda e: self.render())

        # 右侧设置栏放进可滚动容器：窗口矮的时候也能滑到下面的跑圈设置
        outer = ttk.Frame(body, width=366)
        outer.pack(side="right", fill="y", padx=(10, 0))
        outer.pack_propagate(False)
        rcv = tk.Canvas(outer, width=346, highlightthickness=0,
                        bg=ttk.Style().lookup("TFrame", "background") or "#f0f0f0")
        vsb = ttk.Scrollbar(outer, orient="vertical", command=rcv.yview)
        rcv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        rcv.pack(side="left", fill="both", expand=True)
        right = ttk.Frame(rcv)
        rwin = rcv.create_window((0, 0), window=right, anchor="nw")
        right.bind("<Configure>",
                   lambda e: rcv.configure(scrollregion=rcv.bbox("all")))
        rcv.bind("<Configure>", lambda e: rcv.itemconfigure(rwin, width=e.width))

        def _rpanel_wheel(e):               # 只有鼠标在设置栏范围内时滚轮才滚设置栏
            px, py = rcv.winfo_pointerxy()
            x0, y0 = rcv.winfo_rootx(), rcv.winfo_rooty()
            if x0 <= px <= x0 + rcv.winfo_width() and y0 <= py <= y0 + rcv.winfo_height():
                rcv.yview_scroll(-1 if e.delta > 0 else 1, "units")
                return "break"

        rcv.bind_all("<MouseWheel>", _rpanel_wheel)

        search = ttk.Frame(right)
        search.pack(fill="x")
        ttk.Label(search, text="搜索文字").pack(side="left")
        self.search = tk.StringVar()
        entry = ttk.Entry(search, textvariable=self.search)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        self.search.trace_add("write", lambda *_: self.fill_list())

        cols = ("text", "pos", "size")
        self.tree = ttk.Treeview(right, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("text", text="识别到的文字")
        self.tree.heading("pos", text="坐标")
        self.tree.heading("size", text="字高")
        self.tree.column("text", width=186)
        self.tree.column("pos", width=80, anchor="center")
        self.tree.column("size", width=42, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=6)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        self.target_var = tk.StringVar(value="目标：未选择")
        ttk.Label(right, textvariable=self.target_var, foreground="#0a7").pack(anchor="w")

        self.click_btn = ttk.Button(right, text="拟人单击（ESC 中止）", command=self.do_click)
        self.click_btn.pack(fill="x", pady=6)

        route_box = ttk.LabelFrame(right, text="记录连线 · 反复执行", padding=6)
        route_box.pack(fill="x", pady=(8, 0))

        row1 = ttk.Frame(route_box)
        row1.pack(fill="x")
        self.rec_btn = ttk.Button(row1, text="● 开始记录", width=9, command=self.toggle_record)
        self.rec_btn.pack(side="left")
        ttk.Button(row1, text="撤销点", width=6, command=self.undo_point).pack(side="left", padx=4)
        ttk.Button(row1, text="清空", width=5, command=self.clear_route).pack(side="left")

        self.route_var = tk.StringVar(value="已记录 0 个点")
        ttk.Label(route_box, textvariable=self.route_var, foreground="#06c").pack(anchor="w", pady=(4, 0))

        self.interval = tk.StringVar(value="3.0")
        self.round_interval = tk.StringVar(value="3.0")
        for text, var, top in (("点击间隔", self.interval, 4), ("每轮间隔", self.round_interval, 0)):
            row2 = ttk.Frame(route_box)
            row2.pack(fill="x", pady=(top, 2))
            ttk.Label(row2, text=text).pack(side="left")
            ttk.Spinbox(row2, from_=0 if text == "每轮间隔" else 0.1, to=600,
                        increment=0.5 if text == "每轮间隔" else 0.1, width=5,
                        textvariable=var).pack(side="left", padx=(2, 4))
            ttk.Label(row2, text="秒").pack(side="left")

        self.auto_confirm = tk.BooleanVar(value=True)
        ttk.Checkbutton(route_box, text="点击后自动找「确定」确认",
                        variable=self.auto_confirm).pack(anchor="w", pady=(2, 0))
        self.avoid_red = tk.BooleanVar(value=False)      # 默认关闭：勾上才生效
        ttk.Checkbutton(route_box, text="自动识别红色定位标（跳过）",
                        variable=self.avoid_red).pack(anchor="w")
        row3 = ttk.Frame(route_box)
        row3.pack(fill="x", pady=(2, 0))
        self.rect_btn = ttk.Button(row3, text="框选范围", width=9, command=self._toggle_rect_mode)
        self.rect_btn.pack(side="left")
        self.point_btn = ttk.Button(row3, text="点选跳过", width=9, command=self._toggle_point_mode)
        self.point_btn.pack(side="left", padx=3)
        ttk.Button(row3, text="清除", width=6, command=self._clear_avoid).pack(side="left")
        self.avoid_info = tk.StringVar(value="范围 0 · 跳过 0 点（手动生效）")
        ttk.Label(route_box, textvariable=self.avoid_info, foreground="#888").pack(anchor="w")

        self.loop_btn = ttk.Button(route_box, text="▶ 开始反复", command=self.toggle_loop)
        self.loop_btn.pack(fill="x", pady=(2, 0))

        track_box = ttk.LabelFrame(right, text="跑圈模板 · 地图锚定", padding=6)
        track_box.pack(fill="x", pady=(8, 0))

        tr1 = ttk.Frame(track_box)
        tr1.pack(fill="x")
        ttk.Label(tr1, text="像素/米").pack(side="left")
        self.ppm_var = tk.StringVar(value="2.0")
        ttk.Entry(tr1, width=7, textvariable=self.ppm_var).pack(side="left", padx=(2, 8))
        ttk.Button(tr1, text="两点校准", command=self.start_calib).pack(side="left")

        fields = (("直道", "len_var", "100", 5), ("半径", "wid_var", "45", 5),
                  ("角度", "ang_var", "300", 4), ("点数", "n_var", "24", 4))
        for start in (0, 2):                    # 分两行摆，窄一点的设置栏也不会被截断
            tr2 = ttk.Frame(track_box)
            tr2.pack(fill="x", pady=(4 if start == 0 else 0, 0))
            for label, attr, default, width in fields[start:start + 2]:
                ttk.Label(tr2, text=label).pack(side="left", padx=(0, 2))
                var = tk.StringVar(value=default)
                setattr(self, attr, var)
                ttk.Entry(tr2, width=width, textvariable=var).pack(side="left", padx=(0, 12))

        tr3 = ttk.Frame(track_box)
        tr3.pack(fill="x")
        ttk.Button(tr3, text="400米操场", command=self.preset_400).pack(side="left")
        ttk.Button(tr3, text="生成连线", command=self.gen_track).pack(side="left", padx=4)

        tr4 = ttk.Frame(track_box)
        tr4.pack(fill="x", pady=(4, 0))
        self.pace_on = tk.BooleanVar(value=True)
        ttk.Checkbutton(tr4, text="按圈速", variable=self.pace_on).pack(side="left")
        self.pace_min = tk.StringVar(value="4")
        ttk.Entry(tr4, width=3, textvariable=self.pace_min).pack(side="left", padx=(2, 0))
        ttk.Label(tr4, text="分").pack(side="left")
        self.pace_sec = tk.StringVar(value="30")
        ttk.Entry(tr4, width=4, textvariable=self.pace_sec).pack(side="left", padx=(2, 0))
        ttk.Label(tr4, text="秒/圈").pack(side="left")

        self.anchor_var = tk.StringVar(value="锚点：未启用")
        ttk.Label(track_box, textvariable=self.anchor_var, foreground="#888").pack(anchor="w", pady=(4, 0))

        tip = ttk.Label(right, text="提示：也可直接在左侧画面上点一下选目标",
                        foreground="#777", wraplength=320)
        tip.pack(anchor="w")

        self.log_box = tk.Text(right, height=8, wrap="word", state="disabled",
                               bg="#f7f7f7", relief="flat")
        self.log_box.pack(fill="x", pady=(10, 0))

        self.refresh_windows()

        # 启动后后台静默检查一次更新（不弹窗，有新版本才通知）
        threading.Thread(target=self._check_update, kwargs={"silent": True}, daemon=True).start()

    # ── 关于 & 更新 ──────────────────────────────────────────
    def _open_url(self, url):
        """Windows 上用默认浏览器打开链接。"""
        try:
            os.startfile(url)
        except Exception:
            pass

    @staticmethod
    def _version_tuple(v):
        """把 '1.0.0' 拆成 (1, 0, 0) 用于比较，忽略 v 前缀和非数字尾巴。"""
        v = v.lstrip("vV").split("-")[0]
        try:
            return tuple(int(x) for x in v.split("."))
        except ValueError:
            return (0, 0, 0)

    def _check_update(self, silent=False):
        """后台线程查 GitHub Releases，有新版本弹提示；silent=True 时无新版本不弹。"""
        try:
            req = urllib.request.Request(
                GITHUB_API,
                headers={"Accept": "application/vnd.github+json",
                         "User-Agent": f"VisionMouse/{__version__}"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode())
            tag = data.get("tag_name", "").lstrip("vV")
            latest = self._version_tuple(tag)
            current = self._version_tuple(__version__)
            release_page = data.get("html_url", GITHUB_RELEASES_PAGE)
            body = data.get("body", "")[:400]
        except Exception as e:
            if not silent:
                self.root.after(0, lambda: messagebox.showwarning(
                    "检查更新", f"检查失败：{e}\n\n请确认网络连接后重试。"))
            return

        if latest > current:
            self.root.after(0, lambda: self._update_dialog(tag, body, release_page))
        elif not silent:
            self.root.after(0, lambda: messagebox.showinfo(
                "检查更新", f"当前已是最新版本 v{__version__}"))

    def _update_dialog(self, latest_tag, body, url):
        """发现新版本时的提示弹窗。"""
        win = tk.Toplevel(self.root)
        win.title("发现新版本")
        win.geometry("460x300")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text=f"有新版本可用：v{latest_tag}",
                  font=("", 12, "bold")).pack(pady=(16, 4))
        ttk.Label(win, text=f"当前版本：v{__version__}", foreground="#666").pack()

        frame = ttk.Frame(win, padding=(20, 10))
        frame.pack(fill="both", expand=True)
        tb = tk.Text(frame, height=7, wrap="word", state="disabled",
                     bg="#f7f7f7", relief="flat")
        tb.pack(fill="both", expand=True)
        tb.config(state="normal")
        tb.insert("1.0", body or "无更新说明")
        tb.config(state="disabled")

        btns = ttk.Frame(win, padding=(10, 0))
        btns.pack(fill="x", pady=(0, 12))
        ttk.Button(btns, text="前往下载",
                   command=lambda: (self._open_url(url), win.destroy())).pack(side="right", padx=6)
        ttk.Button(btns, text="跳过此版本",
                   command=win.destroy).pack(side="right")

    def _about(self):
        """关于窗口：版本号 + 检查更新按钮。"""
        win = tk.Toplevel(self.root)
        win.title("关于")
        win.geometry("420x260")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="视觉拟人鼠标",
                  font=("", 14, "bold")).pack(pady=(20, 2))
        ttk.Label(win, text=f"版本 v{__version__}", foreground="#666").pack()
        ttk.Label(win, text="截图识别 · OCR 定位 · 类人点击",
                  foreground="#888").pack(pady=(0, 16))

        info = ttk.LabelFrame(win, text="项目地址", padding=8)
        info.pack(fill="x", padx=20)
        link = ttk.Label(info, text=GITHUB_RELEASES_PAGE,
                         foreground="#0066cc", cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda e: self._open_url(GITHUB_RELEASES_PAGE))

        btns = ttk.Frame(win, padding=(10, 16))
        btns.pack(fill="x")
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="检查更新",
                   command=lambda: threading.Thread(
                       target=self._check_update, kwargs={"silent": False},
                       daemon=True).start()).pack(side="right")

    # ── 窗口列表 ──────────────────────────────────────────────
    def refresh_windows(self):
        wins = list_windows()
        self.win_map = {}
        labels = []
        for hwnd, title in wins:
            label = "%s  ·  #%d" % (title[:46], hwnd)
            self.win_map[label] = hwnd
            labels.append(label)
        self.combo["values"] = labels
        if labels:
            prefer = next((x for x in labels if "抖音" in x or "douyin" in x.lower()), labels[0])
            self.combo.set(prefer)
        self.log("窗口列表已刷新，共 %d 个" % len(labels))

    def log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", time.strftime("[%H:%M:%S] ") + msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ── 截图 + OCR ────────────────────────────────────────────
    def start_capture(self):
        if self.busy:
            return
        hwnd = self.win_map.get(self.combo.get())
        if not hwnd:
            messagebox.showwarning("提示", "先选一个目标窗口")
            return
        self.busy = True
        self.status.set("正在截图并识别…")
        threading.Thread(target=self._capture_worker, args=(hwnd,), daemon=True).start()

    def _capture_worker(self, hwnd):
        t0 = time.time()
        img, rect = capture_window(hwnd)
        if img is None:
            self.root.after(0, self._capture_done, None, None, [], t0, "截图失败：窗口可能已最小化或被遮挡")
            return
        rows = ocr_lines(img)
        self.root.after(0, self._capture_done, img, rect, rows, t0, None)

    def _capture_done(self, img, rect, rows, t0, err):
        self.busy = False
        if err:
            self.status.set(err)
            self.log(err)
            return
        self.img, self.rect, self.rows = img, rect, rows
        self.target = None
        self.target_var.set("目标：未选择")
        self.status.set("识别到 %d 行文字，用时 %.1f 秒" % (len(rows), time.time() - t0))
        self.log("窗口 #%d 截到 %dx%d，识别 %d 行，用时 %.1f 秒" % (
            self.win_map.get(self.combo.get(), 0), img.size[0], img.size[1],
            len(rows), time.time() - t0))
        self.fill_list()
        self._update_anchor_shift(img)          # 重新量偏移：标点/禁区钉在地图上，不留在旧屏幕位置
        self.render()

    # ── 实时同步 ──────────────────────────────────────────────
    def toggle_live(self):
        if self.live:
            self.live_stop.set()
            self.status.set("正在停止同步…")
            return
        hwnd = self.win_map.get(self.combo.get())
        if not hwnd:
            messagebox.showwarning("提示", "先选一个目标窗口")
            return
        self.live = True
        self.live_stop.clear()
        self._live_sig = None
        self.live_btn.configure(text="■ 停止同步")
        self.log("开启实时同步：%s" % self.combo.get())
        threading.Thread(target=self._live_worker, args=(hwnd,), daemon=True).start()

    def _live_worker(self, hwnd):
        fails = 0
        while not self.live_stop.is_set():
            img, rect = capture_window(hwnd)
            if img is None:
                fails += 1
                if fails >= 5:
                    self.root.after(0, self.log, "连续截图失败，自动停止同步（窗口可能已关闭）")
                    break
            else:
                fails = 0
                # 跑圈时跳过 OCR：识别很吃 CPU，和跑圈线程抢会卡界面、也会卡住自动确认
                rows = self.rows if self.looping else ocr_lines(img)
                self.root.after(0, self._live_update, img, rect, rows)
            self.live_stop.wait(0.4)
        self.root.after(0, self._live_done)

    def _live_update(self, img, rect, rows):
        if not self.live:
            return
        self.img, self.rect = img, rect
        sig = [r["clean"] for r in rows]
        if sig != self._live_sig:               # 识别结果变了才刷列表，避免选中项闪烁
            self._live_sig = sig
            sel_text = None
            sel = self.tree.selection()
            if sel and self.shown:
                sel_text = self.shown[int(sel[0])]["clean"]
            self.rows = rows
            self.fill_list()
            if sel_text:                        # 尽量保持原选中项
                for i, r in enumerate(self.shown):
                    if r["clean"] == sel_text:
                        self.tree.selection_set(str(i))
                        break
        if not self.looping:                    # 跑圈时偏移由跑圈线程自己算，别两处抢着改
            self._update_anchor_shift(img)      # 每轮都算地图平移量（优先图像配准）
        self.status.set("实时同步中 · %d 行文字 · %s" % (len(rows), time.strftime("%H:%M:%S")))
        self.render()                           # 目标十字和记录连线照常叠加

    def _live_done(self):
        self.live = False
        self.live_stop.clear()
        self.live_btn.configure(text="▶ 实时同步")
        self.status.set("实时同步已停止")
        self.log("实时同步已停止")

    # ── 列表与画面 ────────────────────────────────────────────
    def fill_list(self):
        q = "".join(self.search.get().split()).lower()
        self.shown = [r for r in self.rows if q in r["clean"].lower()] if q else list(self.rows)
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.shown):
            self.tree.insert("", "end", iid=str(i),
                             values=(r["clean"][:60], "%d,%d" % (r["x"], r["y"]), r["h"]))
        if not self.shown and self.rows:
            self.log("没有匹配「%s」的文字" % self.search.get().strip())

    def on_row_select(self, _=None):
        sel = self.tree.selection()
        if not sel or not self.img:
            return
        r = self.shown[int(sel[0])]
        self.set_target_abs(r["x"] + r["w"] / 2.0, r["y"] + r["h"] / 2.0)
        self.render()

    def _toggle_rect_mode(self):
        self.mark_mode = None if self.mark_mode == "rect" else "rect"
        self.rect_from = None
        self._mark_btn_state()
        if self.mark_mode == "rect":
            self.status.set("框选范围：在画面上拖一个矩形，只有框内的标点才避让红色定位标")
        else:
            self.status.set("已退出框选")

    def _toggle_point_mode(self):
        self.mark_mode = None if self.mark_mode == "point" else "point"
        self.rect_from = None
        self._mark_btn_state()
        if self.mark_mode == "point":
            self.status.set("点选跳过：点画面上的轨迹圆点即切换跳过/恢复，再点按钮退出")
        else:
            self.status.set("已退出点选")

    def _mark_btn_state(self):
        self.rect_btn.configure(text="框选中…" if self.mark_mode == "rect" else "框选范围")
        self.point_btn.configure(text="点选中…" if self.mark_mode == "point" else "点选跳过")

    def _clear_avoid(self):
        self.avoid_rects = []
        self.skip_idx = set()
        self._update_avoid_info()
        self.render()
        self.log("已清除全部作用范围与跳过点")

    def _update_avoid_info(self):
        self.avoid_info.set("范围 %d · 跳过 %d 点（手动生效）" % (
            len(self.avoid_rects), len(self.skip_idx)))

    def on_canvas_drag(self, event):
        """框选时画橡皮筋。"""
        if self.mark_mode != "rect" or not self.rect_from:
            return
        self.canvas.delete("band")
        self.canvas.create_rectangle(self.rect_from[0], self.rect_from[1],
                                     event.x, event.y, outline="#ff5252",
                                     width=2, dash=(4, 3), tags="band")

    def on_canvas_release(self, event):
        if self.mark_mode != "rect" or not self.rect_from or not self.img:
            return
        start = self.rect_from
        self.rect_from = None
        self.canvas.delete("band")
        ix, iy = (event.x - self.offset[0]) / self.scale, (event.y - self.offset[1]) / self.scale
        x0, x1 = sorted((start[0], ix))
        y0, y1 = sorted((start[1], iy))
        if x1 - x0 < 6 or y1 - y0 < 6:          # 太小当成误点
            return
        iw, ih = self.img.size
        sx, sy = self.map_shift              # 框是在「当前画面」上框的，存成记录帧坐标
        self.avoid_rects.append(((x0 - sx) / iw, (y0 - sy) / ih,
                                 (x1 - sx) / iw, (y1 - sy) / ih))
        self._update_avoid_info()
        self.render()
        self.log("已加作用范围：窗口内 (%d,%d)-(%d,%d)" % (x0, y0, x1, y1))

    def on_canvas_click(self, event):
        if not self.img:
            return
        ix = (event.x - self.offset[0]) / self.scale
        iy = (event.y - self.offset[1]) / self.scale
        if not (0 <= ix < self.img.size[0] and 0 <= iy < self.img.size[1]):
            return
        if self.mark_mode == "rect":                # 框选：按下记起点，拖拽画框，抬起入库
            self.rect_from = (event.x, event.y)
            return
        if self.mark_mode == "point":               # 点选跳过：点最近的轨迹点切换
            if not self.route:
                return
            iw, ih = self.img.size
            sx, sy = self.map_shift
            best_i, best_d = -1, 1e9
            for i, p in enumerate(self.route):
                d = (p[0] * iw + sx - ix) ** 2 + (p[1] * ih + sy - iy) ** 2
                if d < best_d:
                    best_i, best_d = i, d
            if best_i >= 0 and best_d <= 16 * 16:
                self.skip_idx.symmetric_difference_update({best_i})
                self._update_avoid_info()
                self.render()
                self.log("第 %d 点：%s" % (best_i + 1,
                                           "跳过" if best_i in self.skip_idx else "恢复参与跑圈"))
            return
        if self.calibrating:                # 两点校准模式：收集两个点，问实际距离
            self.calib.append((ix, iy))
            self.log("校准点 %d/2：(%d, %d)" % (len(self.calib), int(ix), int(iy)))
            if len(self.calib) == 1:
                self.status.set("校准：再点第二个点")
            else:
                self.calibrating = False
                d = math.hypot(self.calib[1][0] - self.calib[0][0],
                               self.calib[1][1] - self.calib[0][1])
                meters = simpledialog.askfloat("两点校准", "这两点间的实际距离（米）：",
                                               minvalue=0.1, parent=self.root)
                if meters:
                    self.ppm_var.set("%.3f" % (d / meters))
                    self.log("校准完成：%.1f 像素 / %g 米 = %.3f 像素每米" % (d, meters, d / meters))
                self.status.set("校准结束")
            return
        best, best_d = None, 28 * 28
        for r in self.rows:
            cx, cy = r["x"] + r["w"] / 2.0, r["y"] + r["h"] / 2.0
            d = (cx - ix) ** 2 + (cy - iy) ** 2
            if d < best_d:
                best, best_d = r, d
        if best:
            self.set_target_abs(best["x"] + best["w"] / 2.0, best["y"] + best["h"] / 2.0)
        else:
            self.set_target_abs(ix, iy)
        self.render()
        idx = self.shown.index(best) if best in self.shown else -1
        if idx >= 0:
            self.tree.selection_set(str(idx))
            self.tree.see(str(idx))

    def set_target_abs(self, x, y):
        w, h = self.img.size
        self.target = (x / float(w), y / float(h))
        if self.recording and not self.looping:
            self.route.append(self.target)
            self.route_var.set("已记录 %d 个点" % len(self.route))
            self.target_var.set("目标：%d,%d → 记为第 %d 点" % (int(x), int(y), len(self.route)))
            self.log("记录第 %d 点：窗口内 (%d, %d)" % (len(self.route), int(x), int(y)))
        else:
            self.target_var.set("目标：%d,%d（窗口内）" % (int(x), int(y)))
            self.log("选中目标：窗口内 (%d, %d)" % (int(x), int(y)))

    def render(self):
        self.canvas.delete("all")
        if not self.img:
            return
        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)
        iw, ih = self.img.size
        self.scale = min(cw / float(iw), ch / float(ih))
        dw, dh = int(iw * self.scale), int(ih * self.scale)
        self.offset = ((cw - dw) // 2, (ch - dh) // 2)
        self.photo = ImageTk.PhotoImage(self.img.resize((dw, dh), Image.LANCZOS))
        self.canvas.create_image(self.offset[0], self.offset[1], anchor="nw", image=self.photo)
        for r in self.shown:
            x0 = self.offset[0] + r["x"] * self.scale
            y0 = self.offset[1] + r["y"] * self.scale
            x1 = x0 + r["w"] * self.scale
            y1 = y0 + r["h"] * self.scale
            self.canvas.create_rectangle(x0, y0, x1, y1, outline="#3ecf8e", width=1)
        sx, sy = self.map_shift                        # 画的东西都随地图平移
        for rx0, ry0, rx1, ry1 in self.avoid_rects:     # 手动禁区
            self.canvas.create_rectangle(
                self.offset[0] + (rx0 * iw + sx) * self.scale,
                self.offset[1] + (ry0 * ih + sy) * self.scale,
                self.offset[0] + (rx1 * iw + sx) * self.scale,
                self.offset[1] + (ry1 * ih + sy) * self.scale,
                outline="#ff5252", width=2, dash=(4, 3))
        if self.route:              # 记录连线：带方向的折线 + 序号圆点
            pts = [(self.offset[0] + (p[0] * iw + sx) * self.scale,
                    self.offset[1] + (p[1] * ih + sy) * self.scale) for p in self.route]
            for i in range(len(pts) - 1):
                self.canvas.create_line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1],
                                        fill="#ffb300", width=2, arrow="last")
            for i, (px, py) in enumerate(pts):
                off = i in self.skip_idx                  # 点选跳过的点画成灰的
                self.canvas.create_oval(px - 9, py - 9, px + 9, py + 9,
                                        fill="#6b6b72" if off else "#ffb300",
                                        outline="#3a3a40" if off else "#7a5200")
                self.canvas.create_text(px, py, text=str(i + 1),
                                        fill="#d0d0d0" if off else "#1b1b1f",
                                        font=("Segoe UI", 8, "bold"))
        if self.target:
            tx = self.offset[0] + self.target[0] * iw * self.scale
            ty = self.offset[1] + self.target[1] * ih * self.scale
            self.canvas.create_line(tx - 9, ty, tx + 9, ty, fill="#ff5252", width=2)
            self.canvas.create_line(tx, ty - 9, tx, ty + 9, fill="#ff5252", width=2)
            self.canvas.create_oval(tx - 12, ty - 12, tx + 12, ty + 12, outline="#ff5252", width=2)

    # ── 地图锚定与跑圈模板 ──────────────────────────────────
    def _snapshot_anchors(self):
        """从当前 OCR 结果挑稳定的地名做锚点快照。"""
        cands = [(r["clean"], r["x"] + r["w"] / 2.0, r["y"] + r["h"] / 2.0)
                 for r in self.rows if len(r["clean"]) >= 3]
        cands.sort(key=lambda c: -len(c[0]))
        return cands[:12] or None

    def _compute_shift(self, rows):
        """用锚点地标在 rows 中找同名文字，返回 (dx, dy, 匹配数)；不足 2 个返回 None。"""
        if not self.anchor_labels:
            return None
        cur = {}
        for r in rows:
            cur.setdefault(r["clean"], (r["x"] + r["w"] / 2.0, r["y"] + r["h"] / 2.0))
        dxs, dys = [], []
        for clean, ax, ay in self.anchor_labels:
            if clean in cur:
                cx, cy = cur[clean]
                dxs.append(cx - ax)
                dys.append(cy - ay)
        if len(dxs) >= 2:
            return (statistics.median(dxs), statistics.median(dys), len(dxs))
        return None

    def _gray_center(self, img):
        """裁中心 2/3 并转灰度、缩到 400 宽，作为配准基准图。"""
        w, h = img.size
        crop = img.crop((w // 6, h // 6, w * 5 // 6, h * 5 // 6)).convert("L")
        if crop.size[0] > 400:
            s = 400.0 / crop.size[0]
            crop = crop.resize((400, max(1, int(crop.size[1] * s))), Image.LANCZOS)
        return np.asarray(crop, dtype=np.float32)

    def _snapshot_anchor_img(self):
        """对当前画面做灰度快照，作为 FFT 配准基准（绝对基准 + 滚动基准都重置）。"""
        if np is None or self.img is None:
            self.anchor_img = None
            self.anchor0_img = None
            return
        self.anchor0_img = self._gray_center(self.img)
        self.anchor_img = self.anchor0_img
        self.anchor_base = (0.0, 0.0)

    def _reg_shift(self, ref, img, min_conf=60.0, max_resid=12.0):
        """FFT 相位相关算 img 相对基准图 ref 的平移，像素级精度。不可靠返回 None。
        主判据是峰值置信度：实测真配准 114~11万，假配准（白屏/缩放态/低纹理）只有 15~28。
        残差只做宽松兜底，防止平移绕了半屏那种混叠。"""
        if np is None or ref is None or img is None:
            return None
        w, h = img.size
        crop = img.crop((w // 6, h // 6, w * 5 // 6, h * 5 // 6)).convert("L")
        cw, ch = ref.shape[1], ref.shape[0]
        if crop.size != (cw, ch):
            crop = crop.resize((cw, ch), Image.LANCZOS)
        res = fft_shift(ref, np.asarray(crop, dtype=np.float32))
        if res is None or res[2] < min_conf:
            return None
        if abs(res[0]) > cw * 0.44 or abs(res[1]) > ch * 0.44:
            return None                                 # 接近半屏周期 = 可能是混叠出来的假峰
        sx = (w * 2.0 / 3.0) / cw                       # 配准图缩过，换回原图像素
        sy = (h * 2.0 / 3.0) / ch
        dx, dy = res[0] * sx, res[1] * sy
        if max_resid is not None and self._residual(img, dx, dy, ref) > max_resid:
            return None                                 # 补偿后仍对不上 = 尺度/视野真的变了
        return (dx, dy, res[2])

    def _compute_shift_img(self, img, min_conf=60.0, max_resid=12.0):
        """相对滚动基准的配准（兼容旧调用）。"""
        return self._reg_shift(self.anchor_img, img, min_conf, max_resid)

    def _locate(self, img, reanchor=80.0):
        """给出当前画面相对「记录截图」的总偏移，返回 (dx, dy, 置信度) 或 None。
        先用滚动基准量小增量再累加（增量永远很小，撞不上 FFT 的半屏周期混叠）；
        滚动基准跟丢时才用记录原图做绝对对位兜底。偏移再大也量不出假峰。"""
        if np is None or img is None:
            return None
        # 先跟滚动基准：每步只量一小段增量，离半屏周期很远，绝不会量出混叠的假峰
        if self.anchor_img is not None:
            res = self._reg_shift(self.anchor_img, img)
            if res is not None:
                bx, by = self.anchor_base
                dx, dy = bx + res[0], by + res[1]
                if abs(res[0]) >= reanchor or abs(res[1]) >= reanchor:   # 走远了换基准
                    self.anchor_img = self._gray_center(img)
                    self.anchor_base = (dx, dy)
                return (dx, dy, res[2])
        # 滚动基准跟丢了（画面大变），才用记录原图做绝对对位兜底
        if self.anchor0_img is not None:
            res = self._reg_shift(self.anchor0_img, img)
            if res is not None:
                self.anchor_img = self._gray_center(img)
                self.anchor_base = (res[0], res[1])
                return res
        return None

    def _residual(self, img, dx, dy, ref=None):
        """把当前图按 (dx,dy) 平移回去后与基准图的平均灰度差。真对齐接近 0，
        假对齐（如平移绕了半屏）会很大，用来防止 FFT 混叠造成「假恢复」。"""
        ref = self.anchor_img if ref is None else ref
        if np is None or ref is None or img is None:
            return 255.0
        w, h = img.size
        crop = img.crop((w // 6, h // 6, w * 5 // 6, h * 5 // 6)).convert("L")
        ch, cw = ref.shape
        if crop.size != (cw, ch):
            crop = crop.resize((cw, ch), Image.LANCZOS)
        cur = np.asarray(crop, dtype=np.float32)
        isx = int(round(dx * cw / (w * 2.0 / 3.0)))     # 原图像素 → 配准图像素
        isy = int(round(dy * ch / (h * 2.0 / 3.0)))
        cur = np.roll(cur, (-isy, -isx), axis=(0, 1))
        x0, x1 = max(0, isx), cw - max(0, -isx)         # 只比真正重叠的部分，
        y0, y1 = max(0, isy), ch - max(0, -isy)         # 排除循环移位绕回来的边缘
        if x1 - x0 < cw // 4 or y1 - y0 < ch // 4:
            return 255.0                                # 重叠太少，判不可信
        d = np.abs(cur[y0:y1, x0:x1] - ref[y0:y1, x0:x1])
        return float(d.mean())

    def _find_red_marks(self, img):
        """找地图上的红色定位标（百度的红色图钉/搜索标）。返回 [(x,y)] 图像坐标。
        判据是 45×45 窗口里的红色像素数：图钉是实心红块（实测 240+），红字/红线凑不够。"""
        if np is None or img is None:
            return []
        w, h = img.size
        box = (int(w * 0.04), int(h * 0.14), int(w * 0.96), int(h * 0.96))
        a = np.asarray(img.convert("RGB").crop(box), dtype=np.int16)
        r, g, b = a[:, :, 0], a[:, :, 1], a[:, :, 2]
        m = ((r >= 160) & (g <= 85) & (b <= 85) & (r - g >= 80) & (r - b >= 80)).astype(np.int32)
        ch, cw = m.shape
        k = 45
        if ch <= k or cw <= k:
            return []
        ii = np.zeros((ch + 1, cw + 1), np.int32)
        ii[1:, 1:] = np.cumsum(np.cumsum(m, 0), 1)
        s = ii[k:, k:] - ii[:-k, k:] - ii[k:, :-k] + ii[:-k, :-k]
        out = []
        for _ in range(8):                               # 一屏最多认 8 个标
            py, px = np.unravel_index(int(s.argmax()), s.shape)
            if s[py, px] < 60:                           # 红点太少 = 不是图钉
                break
            ys, xs = np.nonzero(m[py:py + k, px:px + k])
            out.append((box[0] + px + float(xs.mean()), box[1] + py + float(ys.mean())))
            s[max(0, py - k):py + k, max(0, px - k):px + k] = 0
        return out

    def _avoided(self, px, py, w, h, idx):
        """这一点的「不跳过」判定：被点选成跳过的点直接跳过。返回原因，正常返回 None。"""
        if idx in self.skip_idx:
            return "被点选跳过"
        return None

    def _in_zone(self, px, py, w, h):
        """这一点是不是落在框选的作用区里。框选范围存的是「记录帧」坐标，
        这里先把当前坐标减掉地图偏移再比。框内才做红标避让，框外照常点。"""
        if not self.avoid_rects:
            return False
        sx, sy = self.map_shift
        nx, ny = (px - sx) / float(w), (py - sy) / float(h)
        return any(x0 <= nx <= x1 and y0 <= ny <= y1 for x0, y0, x1, y1 in self.avoid_rects)

    def _red_covers(self, px, py, marks, radius=30.0):
        """这个点是否落在红色定位标的作用区里（图钉本身再留一点余量）。
        落在里面的点直接不点，避免把它当成地图空白处点出新钉/弹气泡。"""
        return any(math.hypot(px - mx, py - my) < radius for mx, my in marks or [])

    def _wait_still(self, hwnd, timeout=1.2, gap=0.13, tol=2.0):
        """等地图真的停住，再拿它量偏移、点下去。雷电确认钉后视野是平滑挪过去的，
        边挪边点会把钉落在半路，一步偏一点、整圈就跑偏。返回 (停住的那一帧, 等了多少秒)。"""
        t0 = time.time()
        last, _ = capture_window(hwnd)
        if last is None or np is None:
            return last, 0.0
        while time.time() - t0 < timeout:
            time.sleep(gap)
            cur, _ = capture_window(hwnd)
            if cur is None:
                return last, time.time() - t0
            if cur.size != last.size:
                last = cur
                continue
            d = float(np.abs(np.asarray(cur.convert("L"), dtype=np.float32) -
                             np.asarray(last.convert("L"), dtype=np.float32)).mean())
            if d < tol:                     # 整幅画面基本没动 = 视野停了
                return cur, time.time() - t0
            last = cur
        return last, time.time() - t0       # 超时就用最新这一帧，别把循环卡死

    def _trace(self, msg):
        """每步的实测偏移、落钉情况写进临时目录的 trace 日志，轨迹跑偏了能查账。"""
        try:
            with open(os.environ.get("TEMP", ".") + "/visionmouse_trace.log", "a",
                      encoding="utf-8") as f:
                f.write("%s %s\n" % (time.strftime("%H:%M:%S"), msg))
        except OSError:
            pass

    def _apply_shift(self, nx, ny, matches):
        """带 3px 死区应用新偏移，防止 OCR 坐标抖动让点来回跳。"""
        ox, oy = self.map_shift
        if abs(nx - ox) < 3 and abs(ny - oy) < 3:
            return
        self.map_shift = (nx, ny)
        self.root.after(0, self.anchor_var.set,
                        "锚点：匹配 %d 个地标，偏移 %+d,%+d" % (matches, nx, ny))

    def _update_anchor_shift(self, img=None):
        if not self.route:
            self.map_shift = (0.0, 0.0)
            return
        res = self._locate(img) if img is not None else None              # 优先图像配准
        if res is None and self.anchor_img is None:
            res = self._compute_shift(self.rows)      # 只有完全没有图像锚定时才退回文字地标
        if res:
            self._apply_shift(res[0], res[1], res[2])
        elif self.anchor_labels is not None or self.anchor_img is not None:
            self.anchor_var.set("锚点：匹配不足，保持偏移 %+d,%+d" % self.map_shift)

    def start_calib(self):
        if not self.img:
            messagebox.showinfo("提示", "先截图或开实时同步")
            return
        self.calibrating = True
        self.calib = []
        self.status.set("校准：在地图上点两个已知距离的点（如操场两端）")
        self.log("开始两点校准")

    def preset_400(self):
        self.len_var.set("84")      # 标准 400 米操场：直道 84.4m + 弯道半径 36.5m
        self.wid_var.set("36.5")
        self.log("已填入 400 米操场尺寸：直道 84m，弯道半径 36.5m")

    def gen_track(self):
        if not self.img or not self.target:
            messagebox.showinfo("提示", "先截图/同步，并在地图上点选操场中心点")
            return
        try:
            straight = float(self.len_var.get())        # 直道长度（米）
            radius = float(self.wid_var.get())          # 弯道半径（米）
            ang = math.radians(float(self.ang_var.get()))
            ppm = float(self.ppm_var.get())
            if straight <= 0 or radius <= 0 or ppm <= 0:
                raise ValueError
            if self.pace_on.get():                      # 按圈速：点数×点击间隔 + 每轮间隔 = 一圈
                pace = int(self.pace_min.get() or 0) * 60 + float(self.pace_sec.get())
                per = max(0.5, float(self.interval.get()))
                round_iv = max(0.0, float(self.round_interval.get()))
                if pace - round_iv <= 0:
                    raise ValueError
                n = max(6, int(round((pace - round_iv) / per)))
            else:
                n = max(6, int(float(self.n_var.get())))
        except (ValueError, tk.TclError):
            messagebox.showwarning("提示", "参数需为正数，且圈速须大于每轮间隔")
            return
        iw, ih = self.img.size
        cx, cy = self.target[0] * iw, self.target[1] * ih
        ca, sa = math.cos(ang), math.sin(ang)
        half = straight / 2.0
        arc = math.pi * radius
        perim = 2 * straight + 2 * arc                  # 周长，点沿周长均布
        for i in range(n):
            d = perim * i / n
            if d < straight:                            # 上直道：向左
                mx, my = half - d, -radius
            elif d < straight + arc:                    # 左端半圆：向下绕
                t = (d - straight) / radius
                mx, my = -half - radius * math.sin(t), -radius * math.cos(t)
            elif d < 2 * straight + arc:                # 下直道：向右
                mx, my = -half + (d - straight - arc), radius
            else:                                       # 右端半圆：向上绕
                t = (d - 2 * straight - arc) / radius
                mx, my = half + radius * math.sin(t), radius * math.cos(t)
            ex, ey = mx * ppm, my * ppm                 # 米 → 像素，再旋转
            x = cx + ex * ca - ey * sa
            y = cy + ex * sa + ey * ca
            self.route.append((x / float(iw), y / float(ih)))
        if not self.anchor_labels:
            self.anchor_labels = self._snapshot_anchors()
        if self.anchor_img is None:
            self._snapshot_anchor_img()
        self.route_var.set("已记录 %d 个点" % len(self.route))
        extra = ""
        if self.pace_on.get():
            extra = "（(%d分%g秒−轮间隔%g秒) ÷ 每点%.1f秒，一圈 = %d点×间隔+轮间隔）" % (
                int(self.pace_min.get() or 0), float(self.pace_sec.get()), round_iv, per, n)
        self.log("生成跑圈 %d 点%s：直道 %gm 半径 %gm 角度 %g° 中心 (%d,%d)" % (
            n, extra, straight, radius, math.degrees(ang), cx, cy))
        self.render()

    # ── 记录连线与反复执行 ──────────────────────────────────
    def toggle_record(self):
        if self.looping:
            messagebox.showinfo("提示", "反复执行中，先停止再记录")
            return
        self.recording = not self.recording
        if self.recording:
            if not self.route:
                self.anchor_labels = self._snapshot_anchors()
                self._snapshot_anchor_img()
            self.rec_btn.configure(text="■ 停止记录")
            self.status.set("记录中：点列表或画面即追加到连线")
            self.log("开始记录连线")
        else:
            self.rec_btn.configure(text="● 开始记录")
            self.status.set("记录结束，共 %d 个点" % len(self.route))
            self.log("结束记录，共 %d 个点" % len(self.route))

    def undo_point(self):
        if self.route:
            self.route.pop()
            self.route_var.set("已记录 %d 个点" % len(self.route))
            self.log("撤销一个点，剩 %d 个" % len(self.route))
            self.render()

    def clear_route(self):
        if self.route:
            self.route = []
            self.skip_idx = set()
            self.anchor_labels = None
            self.anchor_img = None
            self.anchor0_img = None
            self.anchor_base = (0.0, 0.0)
            self.map_shift = (0.0, 0.0)
            self.anchor_var.set("锚点：未启用")
            self.route_var.set("已记录 0 个点")
            self.log("已清空记录连线")
            self.render()

    def toggle_loop(self):
        if self.looping:
            self.stop_loop.set()
            self.status.set("正在停止…")
            return
        if self.busy or not self.route:
            messagebox.showinfo("提示", "先记录至少 1 个点")
            return
        hwnd = self.win_map.get(self.combo.get())
        if not hwnd:
            return
        try:
            interval = max(0.05, float(self.interval.get()))
            round_iv = max(0.0, float(self.round_interval.get()))
            pace = 0.0
            if self.pace_on.get():
                pace = int(self.pace_min.get() or 0) * 60 + float(self.pace_sec.get())
                if pace <= 0:
                    raise ValueError
        except (ValueError, tk.TclError):
            messagebox.showwarning("提示", "间隔/圈速请输入数字")
            return
        self.looping = True
        self.busy = True
        self.stop_loop.clear()
        self._low_conf = 0                  # 连续锚定失败计数，≥3 熔断
        self.loop_btn.configure(text="■ 停止反复")
        self.rec_btn.configure(state="disabled")
        self.click_btn.configure(state="disabled")
        if pace:
            self.log("开始反复执行 %d 个点：%d点×间隔%.1f秒 + 轮间隔%.1f秒 ≈ 圈速 %d分%g秒" % (
                len(self.route), len(self.route), interval, round_iv,
                int(pace // 60), pace % 60))
        else:
            self.log("开始反复执行 %d 个点，点击间隔 %.1f 秒，每轮间隔 %.1f 秒" % (
                len(self.route), interval, round_iv))
        if self.avoid_red.get():
            self.log("已开启「不触碰红色定位标」：压在标上的点会直接忽略")
        # 勾选状态在主线程读出来传进去：tkinter 变量不能在子线程读，偶发卡死的元凶
        threading.Thread(target=self._loop_worker,
                         args=(hwnd, interval, round_iv, pace,
                               self.auto_confirm.get(), self.avoid_red.get()),
                         daemon=True).start()

    def _confirm_popup(self, hwnd, clicked=None):
        """点轨迹点后弹出的「确定」气泡：重新截图 OCR 找到并拟人点击。找不到就跳过。
        气泡贴着钉出现，所以气泡与点钉位置的距离能反过来说明钉到底落在哪，写进 trace 好查账。"""
        import win32gui
        time.sleep(0.5)                     # 等气泡弹出
        img, rect = capture_window(hwnd)
        if img is None:
            return
        rows = ocr_lines(img)
        hit = next((r for r in rows if "确定" in r["clean"]), None)
        if not hit:
            self.root.after(0, self.log, "未识别到「确定」气泡，跳过确认")
            return
        r = win32gui.GetWindowRect(hwnd)
        cx = r[0] + hit["x"] + hit["w"] // 2
        cy = r[1] + hit["y"] + hit["h"] // 2
        if clicked:
            self._trace("  确定气泡 %+d,%+d（相对落钉处）" % (
                cx - clicked[0], cy - clicked[1]))
        if not human_move((cx, cy)):
            self.stop_loop.set()
            return
        settle((cx, cy))
        left_click()
        self.root.after(0, self.log, "已自动点击「确定」(%d,%d)" % (cx, cy))
        time.sleep(0.3)

    def _loop_worker(self, hwnd, interval, round_iv, pace=0.0,
                     auto_confirm=True, avoid_red=True):
        import win32gui
        rounds = 0
        total = 0                       # 已点总次数，圈速模式用它排绝对时刻表
        t0 = time.time()
        stopped = "手动停止"
        try:
            while not self.stop_loop.is_set():
                rounds += 1
                bring_front(hwnd)
                n = len(self.route)
                slot = interval if pace > 0 else 0.0   # 圈速模式：每点节奏=点击间隔，一圈=n×间隔+轮间隔
                for idx, target in enumerate(list(self.route)):
                    if self.stop_loop.is_set():
                        break
                    if esc_pressed():
                        stopped = "已按 ESC 中止"
                        self.stop_loop.set()
                        break
                    if slot:                    # 圈速模式：等排到该点的时刻，开销自动扣除
                        wait = t0 + total * slot - time.time()
                        if wait > 0 and self.stop_loop.wait(wait):
                            break
                    r = win32gui.GetWindowRect(hwnd)
                    w, h = r[2] - r[0], r[3] - r[1]
                    skip = False
                    img_now, waited = self._wait_still(hwnd)                # 等视野停住再量、再点
                    want_red = avoid_red or bool(self.avoid_rects) or bool(self.skip_idx)
                    marks = self._find_red_marks(img_now) if (want_red and img_now is not None) else None
                    if self.anchor_labels or self.anchor_img is not None:   # 点击前现场校正地图偏移
                        res = self._locate(img_now) if img_now is not None else None
                        if res is None:                                 # 瓦片刚刷新时可能差一点，重截一次
                            time.sleep(0.2)
                            img_now, waited = self._wait_still(hwnd)
                            if img_now is not None:
                                res = self._locate(img_now)
                                marks = self._find_red_marks(img_now)
                        if res is None and self.anchor_img is None and self.anchor_labels:
                            # 仅在没有图像锚定（无 numpy）时才退回文字地标：
                            # 缩放后地标位置全变，OCR 兜底会算出错误偏移、把点带偏
                            if img_now is not None:
                                res = self._compute_shift(ocr_lines(img_now))
                        if res:
                            self._apply_shift(res[0], res[1], res[2])
                            self._low_conf = 0
                            self._trace("点%d 记录(%d,%d) 实测偏移%+.0f,%+.0f conf=%.0f 应用%+.0f,%+.0f 静帧%.2fs" % (
                                idx + 1, int(target[0] * w), int(target[1] * h),
                                res[0], res[1], res[2], self.map_shift[0], self.map_shift[1], waited))
                        else:
                            # 不再去点缩放按钮：视野跟不上就停下来等重截图，免得把地图越点越乱
                            self._low_conf += 1
                            skip = True                                     # 匹配失败绝不用旧偏移盲点
                            if self._low_conf >= 3:                         # 连续失败 = 画面大变
                                stopped = "地图视野变化且无法配准，已停止：请重新截图并生成连线"
                                self.root.after(0, self.log, stopped)
                                self.root.after(0, messagebox.showwarning, "已停止", stopped)
                                self.stop_loop.set()
                                break
                    if skip:
                        continue
                    x = r[0] + int(target[0] * w) + int(self.map_shift[0])
                    y = r[1] + int(target[1] * h) + int(self.map_shift[1])
                    # 红标避让范围：全局开关打开时全图生效；没开时只有框选/点选的范围内才生效
                    zone = self._in_zone(x - r[0], y - r[1], w, h) or (idx in self.skip_idx)
                    if (avoid_red or zone) and self._red_covers(x - r[0], y - r[1], marks):
                        self.root.after(0, self.log,
                                        "第 %d 点落在红色定位标作用区，忽略该点" % (idx + 1))
                        continue
                    ix, iy = x - r[0], y - r[1]
                    why = self._avoided(ix, iy, w, h, idx)                  # 手动禁区/跳过点：独立生效
                    if why:
                        self.root.after(0, self.log, "第 %d 点%s，忽略该点" % (idx + 1, why))
                        continue
                    if ix < 60 or iy < 50 or (ix > w - 160 and iy < 170):   # 左上缩放控件/右上功能按钮区
                        self.root.after(0, self.log,
                                        "第 %d 点落在地图控件区，跳过（防误触缩放）" % (idx + 1))
                        continue
                    self.root.after(0, self.status.set,
                                    "第 %d 轮 · 第 %d/%d 点" % (rounds, idx + 1, n))
                    if not human_move((x, y)):
                        stopped = "已按 ESC 中止"
                        self.stop_loop.set()
                        break
                    settle((x, y))
                    if self.anchor0_img is not None or self.anchor_img is not None:
                        # 光标走过去了，再确认一次视野有没有在这段时间里挪动，挪了就按新偏移补正
                        img2, _ = self._wait_still(hwnd, timeout=0.9)
                        fix = self._locate(img2) if img2 is not None else None
                        if fix:
                            self._apply_shift(fix[0], fix[1], fix[2])
                            nx = r[0] + int(target[0] * w) + int(fix[0])
                            ny = r[1] + int(target[1] * h) + int(fix[1])
                            if abs(nx - x) >= 3 or abs(ny - y) >= 3:
                                self._trace("  光标到位后视野又挪了 %+d,%+d，按新偏移补正" % (
                                    nx - x, ny - y))
                                if not human_move((nx, ny)):
                                    stopped = "已按 ESC 中止"
                                    self.stop_loop.set()
                                    break
                                settle((nx, ny))
                                x, y = nx, ny
                    left_click()
                    self._trace("  落钉于 (%d,%d)" % (x, y))
                    total += 1
                    if auto_confirm:
                        self._confirm_popup(hwnd, (x, y))
                    if not slot and self.stop_loop.wait(interval):
                        break
                if self.stop_loop.is_set():
                    break
                self.root.after(0, self.log, "第 %d 轮完成" % rounds)
                if self.stop_loop.wait(round_iv):
                    break
                if slot and round_iv > 0:       # 轮间有休息则下一轮重新对表
                    t0, total = time.time(), 0
        except Exception as exc:
            stopped = "反复执行出错：%s" % exc
            self.root.after(0, self.log, stopped)
        finally:
            self.root.after(0, self._loop_done, rounds, stopped)

    def _loop_done(self, rounds, stopped):
        self.looping = False
        self.busy = False
        self.stop_loop.clear()
        self.loop_btn.configure(text="▶ 开始反复")
        self.rec_btn.configure(state="normal")
        self.click_btn.configure(state="normal")
        msg = "反复执行结束（%s），共完成 %d 轮" % (stopped, rounds)
        self.status.set(msg)
        self.log(msg)

    # ── 点击 ─────────────────────────────────────────────────
    def do_click(self):
        if self.busy or not self.target:
            messagebox.showinfo("提示", "先截图并选一个目标")
            return
        hwnd = self.win_map.get(self.combo.get())
        if not hwnd:
            return
        self.busy = True
        self.click_btn.configure(state="disabled")
        threading.Thread(target=self._click_worker, args=(hwnd, self.target), daemon=True).start()

    def _click_worker(self, hwnd, target):
        try:
            import win32gui
            r = win32gui.GetWindowRect(hwnd)
            w, h = r[2] - r[0], r[3] - r[1]
            x = r[0] + int(target[0] * w) + int(self.map_shift[0])
            y = r[1] + int(target[1] * h) + int(self.map_shift[1])
            bring_front(hwnd)
            if win32gui.GetForegroundWindow() != hwnd:
                self.root.after(0, self._click_done, "目标窗口没能到前台，已取消点击", 0)
                return
            t0 = time.time()
            if not human_move((x, y)):
                self.root.after(0, self._click_done, "已按 ESC 中止", 0)
                return
            settle((x, y))
            left_click()
            self.root.after(0, self._click_done,
                            "已在 %d,%d 单击（移动 %.2f 秒）" % (x, y, time.time() - t0), time.time() - t0)
        except Exception as exc:
            self.root.after(0, self._click_done, "点击失败：%s" % exc, 0)

    def _click_done(self, msg, _cost):
        self.busy = False
        self.click_btn.configure(state="normal")
        self.status.set(msg)
        self.log(msg)


if __name__ == "__main__":
    dpi_aware()
    root = tk.Tk()
    VisionMouseApp(root)
    root.mainloop()
