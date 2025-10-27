#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import json
import threading
import subprocess
import queue
import csv
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import math
import time
import re
# ---------- Optional deps ----------
try:
    import pandas as pd
except Exception:
    pd = None

try:
    import requests
except Exception:
    requests = None

# Voice (Windows SAPI via pyttsx3). If not available, app still runs silently.
try:
    import pyttsx3
    _HAS_TTS = True
except Exception:
    pyttsx3 = None
    _HAS_TTS = False

# Better image scaling (optional)
try:
    from PIL import Image, ImageTk, ImageEnhance
    _HAS_PIL = True
except Exception:
    Image = ImageTk = None
    _HAS_PIL = False

APP_TITLE = "Amsons Product Management"
APP_DIRNAME = "AmsonsPM"
CONFIG_BASENAME = "config.json"
GUIDELINES_BASENAME = "guidelines.pptx"

# ============ THEME & SPACING ============

class UI:
    # brand
    BRAND = "#B7955C"       # Amsons gold
    BRAND_DARK = "#9F7E45"
    INK = "#0e1525"
    BG = "#EDEFF3"
    WHITE = "#FFFFFF"
    MUTED = "#7a7a7a"
    SUCCESS = "#1a7f37"
    ERROR = "#d32f2f"

    # spacing scale (in px)
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 22
    XXL = 28

    # radii (simulated through padding/frames)
    R = 12
    R2 = 16


def set_ttk_theme(root: tk.Tk):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    # base fonts
    style.configure(".", font=("Segoe UI", 14))

    # headers
    style.configure("H2.TLabel", font=("Segoe UI Semibold", 15), foreground=UI.INK)
    style.configure("Muted.TLabel", foreground=UI.MUTED)

    # buttons
    common_pad = {"padding": (UI.MD, UI.SM)}
    style.configure("TButton", **common_pad)
    style.map("TButton",
              foreground=[("disabled", "#aaa"),
                          ("pressed", UI.WHITE),
                          ("active", UI.INK)],
              background=[("active", "#f2f2f2")])

    style.configure("Accent.TButton",
                    background=UI.BRAND, foreground=UI.WHITE,
                    **common_pad)
    style.map("Accent.TButton",
              background=[("active", UI.BRAND_DARK), ("pressed", UI.BRAND_DARK)])

    style.configure("Secondary.TButton",
                    background="#B7955C", foreground=UI.WHITE,
                    **common_pad)
    style.map("Secondary.TButton",
              background=[("active", "#B7955C"), ("pressed", "#B7955C")])

    # nav buttons
    style.configure("Nav.TButton",
                    background=UI.WHITE, foreground=UI.INK,
                    padding=(UI.MD, UI.SM), anchor="w")
    style.map("Nav.TButton",
              background=[("active", "#f7f7f8")])

    # labelframes
    style.configure("TLabelframe", background=UI.WHITE)
    style.configure("TLabelframe.Label", background=UI.WHITE, foreground="#374151")

    # entries
    style.configure("TEntry", padding=(8, 6))
    style.map("TEntry",
              fieldbackground=[("!disabled", "#ffffff")])

def _looks_like_placeholder_body(s: str) -> bool:
    """
    Returns True if Body (HTML) looks like a placeholder:
    - contains 'lorem ipsum', 'placeholder', 'tbd', etc.
    - only punctuation/dashes/bullets
    - real text length (after stripping HTML/entities) < 20 chars
    """
    t = str(s or "").strip().lower()
    # strip HTML tags
    t = re.sub(r"<[^>]*>", "", t)
    # normalize common entities and whitespace
    t = t.replace("&nbsp;", " ").replace("&#160;", " ")
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return False  # blank is handled by Error 107
    # too short after cleanup
    if len(t) < 20:
        return True
    # common placeholders
    bad_tokens = ("lorem ipsum", "placeholder", "coming soon", "tbd", "to be decided", "to be defined")
    if any(tok in t for tok in bad_tokens):
        return True
    # just punctuation/dashes/bullets
    if re.fullmatch(r"[-–—•\s]+", t):
        return True
    return False

# ============ Utilities ============

def resource_path(p: str) -> str:
    try:
        base = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, p)

def appdata_dir() -> Path:
    base = os.getenv("APPDATA") or os.path.expanduser("~")
    d = Path(base) / APP_DIRNAME
    d.mkdir(parents=True, exist_ok=True)
    return d

def config_path() -> Path:
    return appdata_dir() / CONFIG_BASENAME

def guidelines_storage_path() -> Path:
    return appdata_dir() / GUIDELINES_BASENAME

def load_config() -> dict:
    p = config_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    cfg = {
        "users": [
            {"username": "admin", "password": "amsons123"}  # change in Settings later if you want
        ],
        "remember_last_user": False,
        "last_user": ""
    }
    try:
        p.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass
    return cfg

def save_config(cfg: dict):
    try:
        config_path().write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass

def blend_hex(c1: str, c2: str, t: float) -> str:
    c1 = c1.lstrip("#"); c2 = c2.lstrip("#")
    r1,g1,b1 = int(c1[0:2],16), int(c1[2:4],16), int(c1[4:6],16)
    r2,g2,b2 = int(c2[0:2],16), int(c2[2:4],16), int(c2[4:6],16)
    r = int(r1 + (r2-r1)*t); g = int(g1 + (g2-g1)*t); b = int(b1 + (b2-b1)*t)
    return f"#{r:02x}{g:02x}{b:02x}"

def slugify_like(s: str) -> str:
    import re
    s = str(s or "").strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    s = re.sub(r"-+", "-", s).strip("-")
    return s[:255]

def is_valid_handle(s: str) -> bool:
    """
    Valid Shopify-style handle: lowercase letters/numbers and hyphens only,
    no leading/trailing hyphen, no spaces, no uppercase, no symbols.
    Empty string is allowed (it's optional and can be auto-generated later).
    """
    s = str(s or "").strip()
    if not s:
        return True  # optional field
    import re
    if len(s) > 255:
        return False
    # allow segments of [a-z0-9] separated by single hyphens
    return bool(re.match(r"^[a-z0-9]+(?:-[a-z0-9]+)*$", s))

def is_url(s: str) -> bool:
    import re
    return bool(re.match(r"^https?://", str(s or "").strip(), re.I))

def looks_like_image_url(s: str) -> bool:
    import re
    return bool(re.search(r"\.(jpg|jpeg|png|gif|webp|tiff?)($|\?)", str(s or "").lower()))

def check_image_url(url: str, timeout=8):
    if not url or not is_url(url):
        return False, "Not a URL"
    if requests is None:
        return (looks_like_image_url(url), "requests not installed; extension check")
    try:
        resp = requests.head(url, allow_redirects=True, timeout=timeout)
        if resp.status_code == 405:
            resp = requests.get(url, stream=True, timeout=timeout)
        if resp.status_code != 200:
            return False, f"HTTP {resp.status_code}"
        ctype = (resp.headers.get("Content-Type") or "").lower()
        if "image" not in ctype:
            return False, f"Content-Type '{ctype}' not image"
        return True, "OK"
    except Exception as e:
        return False, f"Error: {e}"

def is_valid_positive_price_token(x: str) -> bool:
    """
    Accepts strings like 10, 10.0, 19.99 (no commas/currency).
    Must be strictly > 0.
    """
    s = str(x or "").strip()
    if not s:
        return False
    import re
    if not re.match(r"^\d+(\.\d+)?$", s):  # reject commas, currency, letters, etc.
        return False
    try:
        return float(s) > 0.0
    except Exception:
        return False

def draw_vertical_gradient(canvas: tk.Canvas, color_top="#000000", color_bottom="#000000"):
    canvas.delete("grad")
    w = canvas.winfo_width()
    h = canvas.winfo_height()
    if h <= 0 or w <= 0:
        return
    r1, g1, b1 = canvas.winfo_rgb(color_top)
    r2, g2, b2 = canvas.winfo_rgb(color_bottom)
    r_ratio = (r2 - r1) / max(h, 1)
    g_ratio = (g2 - g1) / max(h, 1)
    b_ratio = (b2 - b1) / max(h, 1)
    for i in range(h):
        nr = int(r1 + (r_ratio * i))
        ng = int(g1 + (g_ratio * i))
        nb = int(b1 + (b_ratio * i))
        color = f"#{nr>>8:02x}{ng>>8:02x}{nb>>8:02x}"
        canvas.create_line(0, i, w, i, fill=color, tags=("grad",))

# ============ SKU helpers for validation ============
STRICT_SKU_RE = None
def _sku_regex():
    global STRICT_SKU_RE
    import re
    if STRICT_SKU_RE is None:
        STRICT_SKU_RE = re.compile(r"^(?P<base>\d{6})(?:-(?P<idx>\d{2}))?$")
    return STRICT_SKU_RE

def extract_base_6(s: str):
    m = _sku_regex().match(str(s or "").strip())
    if not m:
        return None
    return int(m.group("base"))

def load_prev_highest_base(prev_path: Path) -> int:
    """Used only to validate presence of a highest SKU (Error 103)."""
    if pd is None or not prev_path or not prev_path.exists():
        return 0
    try:
        if prev_path.suffix.lower() in {".xlsx", ".xls"}:
            pdf = pd.read_excel(prev_path, dtype=str)
        else:
            pdf = pd.read_csv(prev_path, dtype=str)
    except Exception:
        return 0
    if pdf is None or pdf.empty:
        return 0
    pdf = pdf.fillna("")
    if "Variant SKU" not in pdf.columns:
        return 0
    col = pdf["Variant SKU"].astype(str).str.strip().tolist()
    top_cell = next((v for v in col if v and v.lower() != "nan"), None)
    if top_cell:
        b = extract_base_6(top_cell)
        if b is not None:
            return b
    bases = [extract_base_6(v) for v in col if v]
    bases = [b for b in bases if b is not None]
    return max(bases) if bases else 0

# ===== “How to fix” tips =====
def build_fix_tips(active_codes):
    tips = {
        "101": (
            "How to fix Error 101 (Broken Image Link):\n"
            "- Use public https URLs that open in a browser and end with .jpg/.jpeg/.png/.gif/.webp/.tiff\n"
            "- Avoid Google Drive viewer links or local file paths\n"
            "- Upload to Shopify Files and use that URL if needed\n"
            "- If you saw 'requests not installed', run: pip install requests"
        ),
        "102": (
            "How to fix Error 102 (Duplicate Titles):\n"
            "- Make Title unique OR provide a unique Handle (optional)\n"
            "- For updates, put the existing product Handle so import updates instead of duplicating\n"
            "- Remove accidental duplicate rows\n"
            "- Clean duplicates in the previous export / Shopify if they already exist there"
        ),
        "103": (
            "How to fix Error 103 (Unable to find Highest SKU):\n"
            "- Pick a recent Shopify export where 'Variant SKU' exists\n"
            "- Ensure first data row under 'Variant SKU' looks like 110357 or 110357-01\n"
            "- Header must be exactly 'Variant SKU'\n"
            "- If first run without history, leave previous export empty"
        ),
        "104": (
            "How to fix Error 104 (Blank/Empty Import or Previous Export):\n"
            "- Check file paths and sheet name\n"
            "- Make sure the sheet/file has rows and is not protected\n"
            "- Re-export products from Shopify if the export is empty"
        ),
        "105": (
            "How to fix Error 105 (Missing mandatory Shopify fields):\n"
            "- Fill Title*, Vendor*, and Variant Price*\n"
            "- For variants, set Option1 Name (e.g., Size) and Option1 Values (e.g., 52|54|56|58)"
        ),
        "106": (
            "How to fix Error 106 (Missing SEO Title/Description on rows):\n"
            "- For every product row, fill both 'SEO Title' and 'SEO Description'\n"
            "- Recommended: Title ≤ 60 chars, Description ≤ 320 chars\n"
            "- If the columns aren't present, add them to your template so you can fill them"
        ),
        "107": (
            "How to fix Error 107 (Missing Body (HTML) on rows):\n"
            "- For every product with a Title*, enter a product description in 'Body (HTML)'\n"
            "- Plain text is fine, or include simple HTML if you want formatting\n"
            "- Even a short sentence is okay as a placeholder to clear this error"
        ),
        "108":(
            "How to fix Error 108 (Invalid Price):\n"
            "- 'Variant Price*' must be a positive number (e.g., 19.99)\n"
            "- Do NOT include currency symbols (₹, £, $, etc.) or commas\n"
            "- No zeros or negatives; enter a value > 0"
        ),
        "109":(
            "How to fix Error 109 (Bad Handle Format):\n"
            "- Use lowercase letters and numbers only, separated by hyphens\n"
            "- No spaces, uppercase, or symbols; e.g., 'my-product-name'\n"
            "- Remove accents (use plain a-z) and keep length ≤ 255"
        ),
        "110":(
            "How to fix Error 110 (Variant Options Mismatch):\n"
            "- If you enter values (e.g., S|M|L), you must provide an Option1 Name (e.g., Size)\n"
            "- If there are no variants, leave BOTH Option1 Name and Option1 Values blank\n"
            "- Do not mix one without the other"
        ),
        "111":(
            "How to fix Error 111 (SEO Length Limits):\n"
            "- Keep 'SEO Title' ≤ ~60 characters\n"
            "- Keep 'SEO Description' ≤ ~320 characters\n"
            "- Trim extra words so search results don’t truncate"
        ),
        "112":(
            "How to fix Error 112 (Very Short/Placeholder Description):\n"
            "- Add a meaningful product description in 'Body (HTML)'.\n"
            "- Avoid placeholders like 'lorem ipsum' or just dashes.\n"
            "- Aim for at least 20+ characters; 1–2 clear sentences is fine."
        )
    }
    order = ["101","102","103","104","105","106","107","108","109","110","111","112"]
    blocks = [tips[c] for c in order if c in active_codes]
    if blocks:
        return "\n\n" + "\n\n".join(blocks)
    return ""

def sanitize_filename_part(s: str) -> str:
    """Remove Windows-illegal filename chars and trim length."""
    import re
    s = str(s or "").strip()
    s = re.sub(r'[<>:"/\\|?*]+', "-", s)
    s = re.sub(r"\s+", " ", s)
    return s[:80] if s else s

# ============ Small helpers (UI) ============

class Tooltip:
    def __init__(self, widget, text, delay=600):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip = None
        self.after_id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)

    def _schedule(self, _=None):
        self.after_id = self.widget.after(self.delay, self._show)

    def _show(self):
        if self.tip:
            return
        x, y, cx, cy = self.widget.bbox("insert") if self.widget.winfo_ismapped() else (0,0,0,0)
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20
        self.tip = tk.Toplevel(self.widget)
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        lbl = tk.Label(self.tip, text=self.text, bg="#111827", fg="#fff",
                       padx=8, pady=6, font=("Segoe UI", 9))
        lbl.pack()
        self.tip.geometry(f"+{x}+{y}")

    def _hide(self, _=None):
        if self.after_id:
            try: self.widget.after_cancel(self.after_id)
            except Exception: pass
        if self.tip:
            self.tip.destroy()
            self.tip = None

# ============ App Frames ============

class LoginFrame(ttk.Frame):
    """
    Two-panel Sign-In:
    - LEFT: background image (left_bg.jpg/png) if present, else gradient; centered logo (200x100) + title
    - RIGHT: sign-in form with Sign In and Sign Up
    """
    def __init__(self, master, on_login):
        super().__init__(master)
        self.master = master
        self.on_login = on_login
        self.cfg = load_config()

        self.username = tk.StringVar(value=self.cfg.get("last_user") if self.cfg.get("remember_last_user") else "")
        self.password = tk.StringVar(value="")
        self.remember = tk.BooleanVar(value=self.cfg.get("remember_last_user", False))
        self.show_pwd = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self):
        self.pack(fill="both", expand=True)
        # background
        self.bg = tk.Canvas(self, highlightthickness=0, bd=0, bg=UI.INK)
        self.bg.pack(fill="both", expand=True)

        # Centered elevated card container
        card = tk.Frame(self.bg, bg=UI.WHITE, highlightthickness=1, highlightbackground="#e5e7eb")
        self.bg.create_window(self.bg.winfo_reqwidth() // 2, self.bg.winfo_reqheight() // 2,
                              window=card, anchor="center", tags=("card",))

        def _center_card(_=None):
            w, h = self.bg.winfo_width(), self.bg.winfo_height()
            cw = int(min(1080, max(860, w * 0.86)))
            ch = int(min(560, max(440, h * 0.68)))
            self.bg.coords("card", w // 2, h // 2)
            card.config(width=cw, height=ch)

        self.bg.bind("<Configure>", _center_card)

        # Split card into two columns
        card.grid_propagate(False)
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)
        card.grid_rowconfigure(0, weight=1)

        # LEFT panel with image + subtle animated vignette
        left = tk.Canvas(card, highlightthickness=0, bd=0, bg="#000000")
        left.grid(row=0, column=0, sticky="nsew")

        self._left_bg_img = None
        self._left_logo_img = None
        self._logo_float_phase = 0.0

        def _draw_left(_evt=None):
            left.delete("all")
            w, h = max(1, left.winfo_width()), max(1, left.winfo_height())

            # background image if present
            bg_path = None
            for name in ("left_bg.jpg", "left_bg.png", "background_left.jpg", "background_left.png"):
                p = resource_path(name)
                if os.path.exists(p):
                    bg_path = p
                    break

            if _HAS_PIL and bg_path:
                try:
                    bg0 = Image.open(bg_path).convert("RGB")
                    bw, bh = bg0.size
                    scale = max(w / float(bw), h / float(bh))
                    bg0 = bg0.resize((int(bw*scale), int(bh*scale)), Image.LANCZOS)
                    x0 = (bg0.size[0] - w) // 2
                    y0 = (bg0.size[1] - h) // 2
                    bg0 = bg0.crop((x0, y0, x0 + w, y0 + h))
                    # subtle darken for better contrast
                    enhancer = ImageEnhance.Brightness(bg0)
                    bg0 = enhancer.enhance(0.7)
                    self._left_bg_img = ImageTk.PhotoImage(bg0)
                    left.create_image(0, 0, image=self._left_bg_img, anchor="nw")
                except Exception:
                    draw_vertical_gradient(left, "#000000", "#000000")
            else:
                draw_vertical_gradient(left, "#000000", "#000000")

            # Logo 200x100 centered
            TARGET_W, TARGET_H = 200, 100
            cx = w // 2
            # float the logo slightly up/down
            dy = int(6 * math.sin(self._logo_float_phase))
            cy = int(h * 0.42) + dy
            logo_path = resource_path("amsons.png")
            img_obj = None
            if os.path.exists(logo_path):
                try:
                    if _HAS_PIL:
                        img0 = Image.open(logo_path).convert("RGBA").resize((TARGET_W, TARGET_H), Image.LANCZOS)
                        img_obj = ImageTk.PhotoImage(img0)
                    else:
                        img_obj = tk.PhotoImage(file=logo_path)
                        if img_obj.height() > TARGET_H:
                            f = max(1, img_obj.height() // TARGET_H)
                            img_obj = img_obj.subsample(f, f)
                except Exception:
                    img_obj = None

            self._left_logo_img = img_obj
            if self._left_logo_img:
                left.create_image(cx, cy, image=self._left_logo_img, anchor="center")

            # Title
            title_y = cy + (TARGET_H // 2 + 20 if self._left_logo_img else 0)
            left.create_text(
                cx, title_y,
                text="Amsons Shopify Bulk Product Import Generator",
                fill="white",
                font=("Segoe UI Semibold", 18)
            )

        def _tick():
            # animate logo float
            self._logo_float_phase += 0.06
            _draw_left()
            self.after(50, _tick)

        left.bind("<Configure>", _draw_left)
        self.after(200, _tick)

        # RIGHT: white panel with uniform padding
        right = tk.Frame(card, bg=UI.WHITE)
        right.grid(row=0, column=1, sticky="nsew")
        for i in range(12):
            right.grid_rowconfigure(i, weight=0)
        right.grid_rowconfigure(11, weight=1)
        right.grid_columnconfigure(0, weight=1)
        right.grid_columnconfigure(1, weight=1)

        padx = (UI.XL, UI.XL)
        top_pad = (UI.XXL, UI.SM)

        tk.Label(right, text="Sign In", bg=UI.WHITE, fg=UI.BRAND, font=("Segoe UI Black", 22)) \
            .grid(row=0, column=0, columnspan=2, pady=top_pad)

        ttk.Label(right, text="Email / Username:", background=UI.WHITE)\
            .grid(row=1, column=0, columnspan=2, sticky="w", padx=padx, pady=(0, UI.XS))
        e_user = ttk.Entry(right, textvariable=self.username, width=38)
        e_user.grid(row=2, column=0, columnspan=2, sticky="we", padx=padx, pady=(0, UI.SM))

        ttk.Label(right, text="Password:", background=UI.WHITE)\
            .grid(row=3, column=0, columnspan=2, sticky="w", padx=padx, pady=(0, UI.XS))
        self.pw_entry = ttk.Entry(right, textvariable=self.password, show="•", width=38)
        self.pw_entry.grid(row=4, column=0, columnspan=2, sticky="we", padx=padx, pady=(0, UI.SM))

        chkrow = tk.Frame(right, bg=UI.WHITE)
        chkrow.grid(row=5, column=0, columnspan=2, sticky="we", padx=padx, pady=(0, UI.SM))
        ttk.Checkbutton(chkrow, text="Show password", variable=self.show_pwd, command=self._toggle_pwd).pack(side="left")
        ttk.Checkbutton(chkrow, text="Remember me", variable=self.remember).pack(side="right")

        # buttons row – same sizes
        btnrow = tk.Frame(right, bg=UI.WHITE)
        btnrow.grid(row=6, column=0, columnspan=2, sticky="we", padx=padx, pady=(UI.MD, UI.SM))
        btnrow.grid_columnconfigure(0, weight=1)
        btnrow.grid_columnconfigure(1, weight=1)

        b1 = ttk.Button(btnrow, text="Sign In", style="Accent.TButton", command=self._login)
        b2 = ttk.Button(btnrow, text="Sign Up", style="Secondary.TButton", command=self._open_signup)
        b1.grid(row=0, column=0, sticky="we", padx=(0, UI.SM))
        b2.grid(row=0, column=1, sticky="we", padx=(UI.SM, 0))
        Tooltip(b1, "Login with your credentials")
        Tooltip(b2, "Create a new account")

        tk.Label(right, text="Default: admin / amsons123 — change later in Settings.",
                 bg=UI.WHITE, fg="#778", font=("Segoe UI", 9)).grid(row=7, column=0, columnspan=2, pady=(UI.SM, 0))

        # keyboard shortcuts
        self.pw_entry.bind("<Return>", lambda e: self._login())
        self.bind_all("<Escape>", lambda e: self.master.focus_set())

    def _toggle_pwd(self):
        self.pw_entry.config(show="" if self.show_pwd.get() else "•")

    # --- Sign Up logic ---
    def _open_signup(self):
        dlg = tk.Toplevel(self)
        dlg.title("Create Account")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.configure(bg=UI.WHITE)

        frm = ttk.Frame(dlg, padding=UI.LG)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Create a new account", style="H2.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, UI.SM))

        ttk.Label(frm, text="Username:").grid(row=1, column=0, sticky="e", padx=(0, UI.SM), pady=UI.SM)
        uvar = tk.StringVar()
        ttk.Entry(frm, textvariable=uvar, width=28).grid(row=1, column=1, sticky="w", pady=UI.SM)

        ttk.Label(frm, text="Password:").grid(row=2, column=0, sticky="e", padx=(0, UI.SM), pady=UI.SM)
        pvar = tk.StringVar()
        ep = ttk.Entry(frm, textvariable=pvar, show="•", width=28)
        ep.grid(row=2, column=1, sticky="w", pady=UI.SM)

        ttk.Label(frm, text="Confirm Password:").grid(row=3, column=0, sticky="e", padx=(0, UI.SM), pady=UI.SM)
        cvar = tk.StringVar()
        ec = ttk.Entry(frm, textvariable=cvar, show="•", width=28)
        ec.grid(row=3, column=1, sticky="w", pady=UI.SM)

        show = tk.BooleanVar(value=False)
        def _toggle():
            show_now = "" if show.get() else "•"
            ep.config(show=show_now)
            ec.config(show=show_now)
        ttk.Checkbutton(frm, text="Show passwords", variable=show, command=_toggle).grid(row=4, column=1, sticky="w")

        def _create():
            user = uvar.get().strip()
            pw   = pvar.get().strip()
            pw2  = cvar.get().strip()

            if not user or not pw or not pw2:
                messagebox.showerror(APP_TITLE, "Please fill all fields.")
                return
            if len(pw) < 6:
                messagebox.showerror(APP_TITLE, "Password must be at least 6 characters.")
                return
            if pw != pw2:
                messagebox.showerror(APP_TITLE, "Passwords do not match.")
                return

            cfg = load_config()
            users = cfg.get("users", [])
            if any(rec.get("username") == user for rec in users):
                messagebox.showerror(APP_TITLE, "This username already exists. Choose another.")
                return

            users.append({"username": user, "password": pw})
            cfg["users"] = users
            save_config(cfg)

            # Pre-fill login form so user can sign in immediately
            self.username.set(user)
            self.password.set(pw)
            messagebox.showinfo(APP_TITLE, "Account created. You can now sign in.")
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(UI.MD,0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=(0,UI.SM))
        okb = ttk.Button(btns, text="Create Account", style="Accent.TButton", command=_create)
        okb.pack(side="left")
        dlg.bind("<Return>", lambda e: _create())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _login(self):
        u = self.username.get().strip()
        p = self.password.get().strip()

        ok = False
        for rec in load_config().get("users", []):
            if rec.get("username") == u and rec.get("password") == p:
                ok = True
                break

        if not ok:
            messagebox.showerror(APP_TITLE, "Invalid username or password.")
            return

        # Remember choice
        cfg = load_config()
        cfg["remember_last_user"] = bool(self.remember.get())
        cfg["last_user"] = u if self.remember.get() else ""
        save_config(cfg)

        # Voice greeting (non-blocking)
        if _HAS_TTS:
            threading.Thread(target=self._speak, args=("Welcome to Amsons Shopify Bulk Product Import Generator",), daemon=True).start()

        # switch to dashboard
        self.on_login(u)

    def _speak(self, text):
        try:
            eng = pyttsx3.init()
            eng.say(text)
            eng.runAndWait()
        except Exception:
            pass


class DashboardFrame(ttk.Frame):
    def __init__(self, master, username: str):
        super().__init__(master)
        self.master = master
        self.username = username

        # Hidden script path (auto-use file near this app)
        self.script_path = tk.StringVar(value=str(Path(__file__).with_name("final-script.py")))
        self.input_path = tk.StringVar(value="")
        self.prev_path = tk.StringVar(value="")
        self.out_dir = tk.StringVar(value=str(Path.cwd() / "out"))
        self.sheet_name = tk.StringVar(value="Products")
        self.respect_existing = tk.BooleanVar(value=True)
        self.status_choice = tk.StringVar(value="active")
        self.proceed_despite_errors = tk.BooleanVar(value=False)
        self.last_validation = {"ran": False, "has_errors": False, "summary": "", "codes": set(), "broken_titles": set()}
        self._run_custom_label = ""

        self.proc = None
        self.q = queue.Queue()
        self.phase = "idle"
        self._last_exit_code = None
        self._current_outdir = None

        self._header_anim_t = 0.0
        self._build_ui()

    def _build_ui(self):
        self.pack(fill="both", expand=True)

        # HERO header
        self.header = tk.Canvas(self, height=150, highlightthickness=0, bd=0, bg="#000")
        self.header.pack(fill="x")
        self.header.bind("<Configure>", self._redraw_header)
        self.logo_img = None
        self.after(40, self._animate_header)  # shimmer animation

        # Main layout with side nav + content card
        container = tk.Frame(self, bg=UI.BG)
        container.pack(fill="both", expand=True)

        # Side nav
        nav = tk.Frame(container, width=220, bg=UI.WHITE, bd=0, highlightthickness=1, highlightbackground="#e5e7eb")
        nav.pack(side="left", fill="y", padx=(UI.MD,UI.SM), pady=UI.MD)

        tk.Label(nav, text=f"Welcome, {self.username}", bg=UI.WHITE, fg=UI.BRAND,
                 font=("Segoe UI Semibold", 12)).pack(anchor="w", padx=UI.MD, pady=(UI.LG, UI.SM))
        ttk.Separator(nav, orient="horizontal").pack(fill="x", padx=UI.MD, pady=(0,UI.MD))

        def _nav_btn(txt, cmd):
            b = ttk.Button(nav, text=txt, style="Nav.TButton", command=cmd)
            b.pack(fill="x", padx=UI.MD, pady=UI.XS, ipady=6)
            return b

        _nav_btn("Dashboard", lambda: None)
        _nav_btn("New Template", self._new_template)
        _nav_btn("Guide lines", self._guidelines)  # <-- PPTX store/retrieve
        _nav_btn("Settings (Change Password)", self._open_change_password)
        _nav_btn("Open Output Folder", self._open_outdir)
        _nav_btn("Logout", self._logout)

        # Content card
        card = tk.Frame(container, bg=UI.WHITE, bd=0, highlightthickness=1, highlightbackground="#e5e7eb")
        card.pack(side="left", fill="both", expand=True, padx=(UI.SM,UI.MD), pady=UI.MD)
        card.grid_columnconfigure(1, weight=1)

        row = 0
        ttk.Label(card, text="Import Builder", style="H2.TLabel", background=UI.WHITE)\
            .grid(row=row, column=0, columnspan=3, sticky="w", padx=(UI.LG,UI.LG), pady=(UI.LG,UI.SM))

        # Input
        row += 1
        ttk.Label(card, text="Input Excel:", background=UI.WHITE).grid(row=row, column=0, sticky="e", padx=UI.LG, pady=UI.SM)
        e1 = ttk.Entry(card, textvariable=self.input_path)
        e1.grid(row=row, column=1, sticky="we", pady=UI.SM)
        b1 = ttk.Button(card, text="Browse", style="TButton", command=self._pick_input)
        b1.grid(row=row, column=2, padx=UI.LG, pady=UI.SM, sticky="we")
        Tooltip(b1, "Select the source .xlsx/.xls")

        # Prev
        row += 1
        ttk.Label(card, text="Previous Export (CSV/XLSX):", background=UI.WHITE).grid(row=row, column=0, sticky="e", padx=UI.LG, pady=UI.SM)
        e2 = ttk.Entry(card, textvariable=self.prev_path)
        e2.grid(row=row, column=1, sticky="we", pady=UI.SM)
        b2 = ttk.Button(card, text="Browse", command=self._pick_prev)
        b2.grid(row=row, column=2, padx=UI.LG, pady=UI.SM, sticky="we")
        Tooltip(b2, "Select a previous Shopify export to check duplicates / SKUs")

        # Out dir
        row += 1
        ttk.Label(card, text="Output Folder:", background=UI.WHITE).grid(row=row, column=0, sticky="e", padx=UI.LG, pady=UI.SM)
        e3 = ttk.Entry(card, textvariable=self.out_dir)
        e3.grid(row=row, column=1, sticky="we", pady=UI.SM)
        b3 = ttk.Button(card, text="Browse", command=self._pick_outdir)
        b3.grid(row=row, column=2, padx=UI.LG, pady=UI.SM, sticky="we")
        Tooltip(b3, "Choose where the files will be written")

        # Status + proceed
        row += 1
        ctl = tk.Frame(card, bg=UI.WHITE)
        ctl.grid(row=row, column=0, columnspan=3, sticky="we", padx=UI.LG, pady=(UI.MD,UI.SM))

        box = ttk.LabelFrame(ctl, text="Import Status", padding=UI.SM)
        box.pack(side="left", padx=(0, UI.LG))
        ttk.Radiobutton(box, text="Active", value="active", variable=self.status_choice).pack(side="left", padx=UI.SM, ipady=2)
        ttk.Radiobutton(box, text="Draft",  value="draft",  variable=self.status_choice).pack(side="left", padx=UI.SM, ipady=2)

        ttk.Checkbutton(ctl, text="Allow to proceed without images", variable=self.proceed_despite_errors).pack(side="left", padx=UI.LG)

        # Buttons — equal widths & spacing
        row += 1
        btns = tk.Frame(card, bg=UI.WHITE)
        btns.grid(row=row, column=0, columnspan=3, sticky="we", padx=UI.LG, pady=(UI.LG,UI.MD))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)
        btns.grid_columnconfigure(2, weight=1)

        self.btn_validate = ttk.Button(btns, text="Validate", style="Secondary.TButton", command=self._validate_only)
        self.btn_run = ttk.Button(btns, text="Run", style="Accent.TButton", command=self._run_only)
        self.btn_open_out = ttk.Button(btns, text="Open Output Folder", style="TButton", command=self._open_outdir, state="disabled")

        self.btn_validate.grid(row=0, column=0, sticky="we", padx=(0, UI.SM))
        self.btn_run.grid(row=0, column=1, sticky="we", padx=(UI.SM, UI.SM))
        self.btn_open_out.grid(row=0, column=2, sticky="we", padx=(UI.SM, 0))

        Tooltip(self.btn_validate, "Check your sheet for issues (fast)")
        Tooltip(self.btn_run, "Build the import files")
        Tooltip(self.btn_open_out, "Open the generated files folder")

        # Progress + log
        self.prog = ttk.Progressbar(card, mode="indeterminate")
        self.prog.grid(row=row+1, column=0, columnspan=3, sticky="we", padx=UI.LG)

        row += 2
        ttk.Label(card, text="Live Log:", background=UI.WHITE).grid(row=row, column=0, sticky="w", padx=UI.LG, pady=(UI.MD, UI.XS))
        row += 1
        self.txt = tk.Text(card, height=16, wrap="word", font=("Consolas", 12), bd=0, relief="flat", bg="#0b1020", fg="#e5e7eb", insertbackground="#e5e7eb")
        self.txt.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=UI.LG, pady=(0,UI.LG))
        card.grid_rowconfigure(row, weight=1)

        # Status bar
        self.status_bar = tk.Label(self, text="Ready", anchor="w", bg="#F8F8F9", fg="#111827", padx=UI.MD)
        self.status_bar.pack(fill="x")

    # ----- header -----
    def _redraw_header(self, _evt=None):
        c = self.header
        c.delete("all")
        w,h = c.winfo_width(), c.winfo_height()

        # animated shimmer across black
        t = (math.sin(self._header_anim_t) + 1) / 2  # 0..1
        top = blend_hex("#000000", "#111111", t*0.6)
        bot = blend_hex("#000000", "#161616", t*0.6)
        for i in range(max(1,h)):
            u = i/float(h-1) if h>1 else 0
            c.create_line(0,i,w,i,fill=blend_hex(top,bot,u))

        # center logo
        logo_path = resource_path("amsons.png")
        img = None
        if os.path.exists(logo_path):
            try:
                if _HAS_PIL:
                    img0 = Image.open(logo_path).convert("RGBA")
                    max_h = 80
                    W,H = img0.size
                    if H > max_h:
                        scale = max_h/float(H)
                        img0 = img0.resize((int(W*scale), int(H*scale)), Image.LANCZOS)
                    img = ImageTk.PhotoImage(img0)
                else:
                    img = tk.PhotoImage(file=logo_path)
                    if img.height()>80:
                        factor = max(2, img.height()//80)
                        img = img.subsample(factor, factor)
            except Exception:
                img = None
        self.logo_img = img
        cy = int(h*0.52)
        if self.logo_img:
            c.create_image(w//2, cy+8, image=self.logo_img, anchor="s")
            y = cy + 8
        else:
            y = cy
        c.create_text(w//2, y, text="Amsons Shopify Bulk Product Import Generator (V-1.1)", fill=UI.WHITE,
                      font=("Segoe UI Semibold", 18), anchor="n")
        c.create_text(w//2, y+35, text="Import builder • validator • image checks",
                      fill="#E9FFFB", font=("Segoe UI", 10), anchor="n")

    def _animate_header(self):
        # gentle shimmer
        self._header_anim_t += 0.04
        self._redraw_header()
        self.after(50, self._animate_header)

    # ----- pickers -----
    def _pick_script(self):
        p = filedialog.askopenfilename(title="Select final-script.py", filetypes=[("Python","*.py"),("All","*.*")])
        if p: self.script_path.set(p)
    def _pick_input(self):
        p = filedialog.askopenfilename(title="Select Input Excel", filetypes=[("Excel","*.xlsx;*.xls"),("All","*.*")])
        if p: self.input_path.set(p)
    def _pick_prev(self):
        p = filedialog.askopenfilename(title="Select Previous Export (CSV/XLSX)", filetypes=[("CSV","*.csv"),("Excel","*.xlsx;*.xls"),("All","*.*")])
        if p: self.prev_path.set(p)
    def _pick_outdir(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d: self.out_dir.set(d)

    # ----- NEW TEMPLATE -----
    def _new_template(self):
        default_name = "Amsons_Products_Template.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save New Template",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook","*.xlsx"), ("CSV (fallback)","*.csv")]
        )
        if not path:
            return

        cols = self._make_template_columns()
        rows = self._make_template_example_rows(cols)
        try:
            if Path(path).suffix.lower() == ".csv" or pd is None:
                self._save_template_csv(path, cols, rows)
                messagebox.showinfo(APP_TITLE, f"Template (CSV) saved with examples:\n{path}")
                self._log(f"New template saved (CSV): {path}")
            else:
                self._save_template_excel(path, cols, rows, self.sheet_name.get().strip() or "Products")
                messagebox.showinfo(APP_TITLE, f"Template (XLSX) saved with examples:\n{path}")
                self._log(f"New template saved (XLSX): {path}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to create template:\n\n{e}")

    def _make_template_columns(self):
        base_cols = [
            "Title*", "Vendor*", "Body (HTML)",
            "Handle (optional)",
            "SEO Title", "SEO Description",
            "Variant Price*",
            "Option1 Name", "Option1 Values",
            "Tags"
        ]
        img_cols = [f"Image URL {i}" for i in range(1, 9)]
        return base_cols + img_cols

    def _make_template_example_rows(self, columns):
        simple = {
            "Title*": "Demo Ceramic Mug",
            "Vendor*": "Amsons",
            "Body (HTML)": "A durable ceramic mug perfect for everyday use.",
            "Handle (optional)": "demo-ceramic-mug",
            "SEO Title": "Ceramic Mug - Demo",
            "SEO Description": "A simple ceramic mug for hot drinks.",
            "Variant Price*": "9.99",
            "Option1 Name": "",
            "Option1 Values": "",
            "Tags": "demo,mug,kitchen",
            "Image URL 1": "https://dummyimage.com/800x800/222/fff.jpg&text=Mug",
        }

        variant = {
            "Title*": "Demo Cotton T-Shirt",
            "Vendor*": "Amsons",
            "Body (HTML)": "Soft cotton tee available in multiple sizes.",
            "Handle (optional)": "demo-cotton-tshirt",
            "SEO Title": "Cotton T-Shirt - Demo",
            "SEO Description": "A comfy cotton t-shirt in S, M, L.",
            "Variant Price*": "19.99",
            "Option1 Name": "Size",
            "Option1 Values": "S|M|L",
            "Tags": "demo,apparel,tshirt",
            "Image URL 1": "https://dummyimage.com/800x800/333/fff.jpg&text=T-Shirt",
        }

        def normalize(row):
            return {c: row.get(c, "") for c in columns}
        return [normalize(simple), normalize(variant)]

    def _save_template_csv(self, path: str, columns, rows):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            for r in rows:
                writer.writerow(r)

    def _save_template_excel(self, path: str, columns, rows, sheet_name: str):
        if pd is None:
            raise RuntimeError("pandas is required to write XLSX. Choose CSV or install: pip install pandas openpyxl")
        import importlib
        engine = None
        try:
            importlib.import_module("openpyxl")
            engine = "openpyxl"
        except Exception:
            engine = None
        df = pd.DataFrame(rows, columns=columns)
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path, engine=engine) as xw:
            df.to_excel(xw, index=False, sheet_name=sheet_name)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path)
            ws = wb[sheet_name]
            ws.freeze_panes = "A2"
            width_map = {
                "A": 28, "B": 18, "C": 40, "D": 24, "E": 28,
                "F": 40, "G": 16, "H": 16, "I": 24, "J": 18
            }
            for col, w in width_map.items():
                if col in ws.column_dimensions:
                    ws.column_dimensions[col].width = w
            wb.save(path)
        except Exception:
            pass

    # ----- GUIDELINES (PowerPoint) -----
    def _guidelines(self):
        storage = guidelines_storage_path()
        if not storage.exists():
            src = filedialog.askopenfilename(
                title="Upload your Guidelines PowerPoint (.pptx) to store",
                filetypes=[("PowerPoint Presentation", "*.pptx")]
            )
            if not src:
                messagebox.showinfo(APP_TITLE, "No file selected.")
                return
            try:
                shutil.copyfile(src, storage)
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Failed to store the Guidelines file:\n\n{e}")
                return
            messagebox.showinfo(APP_TITLE, "Guidelines uploaded and stored. You can now download it anytime.")

        dest = filedialog.asksaveasfilename(
            title="Save Guide lines",
            defaultextension=".pptx",
            initialfile="Amsons_Guidelines.pptx",
            filetypes=[("PowerPoint Presentation", "*.pptx")]
        )
        if not dest:
            return
        try:
            shutil.copyfile(storage, dest)
            self._log(f"Guidelines saved to: {dest}")
            messagebox.showinfo(APP_TITLE, f"Guide lines saved:\n{dest}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to save Guide lines:\n\n{e}")

    # ----- error dialog (red X with scrollable details) -----
    def _show_error_dialog(self, text: str):
        dlg = tk.Toplevel(self)
        dlg.title(APP_TITLE)
        dlg.grab_set()
        dlg.geometry("820x520")
        dlg.minsize(680, 420)
        dlg.configure(bg=UI.WHITE)

        top = tk.Frame(dlg, bg=UI.WHITE)
        top.pack(fill="both", expand=True, padx=UI.LG, pady=UI.LG)
        top.grid_columnconfigure(1, weight=1)
        top.grid_rowconfigure(0, weight=1)

        # Red circle with white X
        cnv = tk.Canvas(top, width=80, height=80, highlightthickness=0, bg=UI.WHITE, bd=0)
        cnv.grid(row=0, column=0, sticky="n", padx=(0, UI.MD))
        cnv.create_oval(5,5,75,75, fill=UI.ERROR, outline="")
        cnv.create_line(22,22,58,58, width=8, fill="white", capstyle="round")
        cnv.create_line(58,22,22,58, width=8, fill="white", capstyle="round")

        # Scrollable text
        frm = tk.Frame(top, bg=UI.WHITE)
        frm.grid(row=0, column=1, sticky="nsew")
        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(0, weight=1)
        txt = tk.Text(frm, wrap="word", bg="#0b1020", fg="#e5e7eb", insertbackground="#e5e7eb")
        txt.grid(row=0, column=0, sticky="nsew")
        yb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        yb.grid(row=0, column=1, sticky="ns")
        txt.configure(yscrollcommand=yb.set)
        txt.insert("1.0", text)
        txt.configure(state="disabled")

        # Buttons
        ttk.Button(dlg, text="OK", style="Accent.TButton", command=dlg.destroy).pack(pady=(UI.SM,UI.SM))
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    # ----- validation -----
    def _validate_only(self):
        if pd is None:
            messagebox.showerror(APP_TITLE, "pandas is required. Install: pip install pandas openpyxl")
            return
        script = Path(self.script_path.get())
        if not script.exists():
            alt = filedialog.askopenfilename(title="Locate final-script.py", filetypes=[("Python","*.py"),("All","*.*")])
            if not alt:
                messagebox.showerror(APP_TITLE, "Cannot continue without final-script.py")
                return
            self.script_path.set(alt)

        if not self.input_path.get().strip():
            messagebox.showerror(APP_TITLE, "Choose an Input Excel file.")
            return

        self.btn_validate.config(state="disabled")
        self.btn_run.config(state="disabled")
        self.btn_open_out.config(state="disabled")
        self._clear_log(); self._log("Starting validation...\n\n")
        self.status_bar.config(text="Validating...")
        self.prog.start(12); self.phase = "preflight"
        t = threading.Thread(
            target=self._worker_preflight,
            args=(self.input_path.get().strip(), self.sheet_name.get().strip() or "Products", self.prev_path.get().strip()),
            daemon=True
        )
        t.start()
        self.after(50, self._poll_validation_only)

    def _worker_preflight(self, inp_path: str, sheet: str, prev_path_str: str):
        try:
            try:
                df = pd.read_excel(inp_path, sheet_name=sheet, dtype=str)
            except Exception as e:
                self.q.put(("__VALIDATION_FAIL__", {
                    "detail": f"Error 104: Blank/Unreadable Import Sheet\nCannot read sheet '{sheet}' in '{inp_path}'.\n\n{e}",
                    "codes": ["104"],
                    "broken_titles": []
                }))
                return
            df = df.fillna("")
            total = int(df["Title*"].astype(str).str.strip().ne("").sum()) if "Title*" in df.columns else 0

            codes = set()
            sections = []
            broken_titles_set = set()

            # Error 104: empty import sheet
            if df.empty or total == 0:
                codes.add("104")
                sections.append("Error 104: Blank/Empty Import\n- The input sheet has no products with non-empty Title*.")

            # Error 105: missing mandatory
            miss_msgs = []
            if "Title*" not in df.columns or "Vendor*" not in df.columns or "Variant Price*" not in df.columns:
                codes.add("105")
                missing_cols = [c for c in ["Title*","Vendor*","Variant Price*"] if c not in df.columns]
                miss_msgs.append(f"- Missing required column(s): {', '.join(missing_cols)}")
            else:
                miss_t = df.index[df["Title*"].astype(str).str.strip() == ""].tolist()
                miss_v = df.index[df["Vendor*"].astype(str).str.strip() == ""].tolist()
                miss_p = df.index[df["Variant Price*"].astype(str).str.strip() == ""].tolist()
                if miss_t: miss_msgs.append(f"- Missing Title* on rows: {', '.join(str(i+2) for i in miss_t)}")
                if miss_v: miss_msgs.append(f"- Missing Vendor* on rows: {', '.join(str(i+2) for i in miss_v)}")
                if miss_p: miss_msgs.append(f"- Missing Variant Price* on rows: {', '.join(str(i+2) for i in miss_p)}")
                if miss_t or miss_v or miss_p:
                    codes.add("105")
            if miss_msgs:
                sections.append("Error 105: Mandatory fields missing\n" + "\n".join(miss_msgs))

            # Error 110: Variant Options Mismatch (Option1 Name/Values must be paired)
            if "Option1 Name" in df.columns or "Option1 Values" in df.columns:
                name_series = df.get("Option1 Name", pd.Series([""] * len(df))).astype(str).str.strip()
                vals_series = df.get("Option1 Values", pd.Series([""] * len(df))).astype(str).str.strip()

                mism_idxs = []
                for i in range(len(df)):
                    has_name = bool(name_series.iat[i])
                    has_vals = bool(vals_series.iat[i])
                    # flag if exactly one is present
                    if has_name ^ has_vals:
                        mism_idxs.append(i)

                if mism_idxs:
                    codes.add("110")
                    lines = []
                    for i in mism_idxs[:60]:  # cap lines for dialog length
                        rowno = i + 2
                        title = df.at[i, "Title*"] if "Title*" in df.columns else ""
                        n = name_series.iat[i]
                        v = vals_series.iat[i]
                        lines.append(f"- Row {rowno}: Title='{title}'  Option1 Name='{n}'  Option1 Values='{v}'")
                    more = f"\n  ... and {len(mism_idxs) - 60} more row(s)" if len(mism_idxs) > 60 else ""
                    sections.append("Error 110: Variant Options Mismatch (Option1)\n" + "\n".join(lines) + more)

            # Error 108: invalid price tokens (non-numeric, zero, or negative)
            if "Variant Price*" in df.columns:
                invalid_idxs = []
                col = df["Variant Price*"].astype(str)
                for i, val in col.items():
                    s = str(val).strip()
                    # Skip blanks here (already handled by Error 105 missing mandatory)
                    if not s:
                        continue
                    if not is_valid_positive_price_token(s):
                        invalid_idxs.append(i)

                if invalid_idxs:
                    codes.add("108")
                    lines = []
                    for i in invalid_idxs[:60]:  # show first 60 rows to keep dialog small
                        title = df.at[i, "Title*"] if "Title*" in df.columns else ""
                        lines.append(f"- Row {i + 2}: {title} — price='{str(df.at[i, 'Variant Price*']).strip()}'")
                    more = f"\n  ... and {len(invalid_idxs) - 60} more row(s)" if len(invalid_idxs) > 60 else ""
                    sections.append("Error 108: Invalid Price\n" + "\n".join(lines) + more)

            # Error 106: missing SEO Title/Description on any row
            present_seo_cols = [c for c in ["SEO Title", "SEO Description"] if c in df.columns]
            if present_seo_cols:
                cond = False
                for c in present_seo_cols:
                    series = df[c].astype(str).str.strip()
                    cond = series.eq("") if cond is False else (cond | series.eq(""))
                if cond is not False:
                    idxs = list(df.index[cond])
                    if idxs:
                        codes.add("106")
                        lines = []
                        for i in idxs[:40]:
                            title = df.at[i, "Title*"] if "Title*" in df.columns else ""
                            rowno = i + 2
                            lines.append(f"- Row {rowno}: {title}")
                        more = f"\n  ... and {len(idxs)-40} more row(s)" if len(idxs) > 40 else ""
                        sections.append("Error 106: Missing SEO Title/Description on rows\n" + "\n".join(lines) + more)

                # Error 111: SEO Length Limits (Title > ~60 or Description > ~320)
                if "SEO Title" in df.columns or "SEO Description" in df.columns:
                    st = df.get("SEO Title", pd.Series([""] * len(df))).astype(str)
                    sd = df.get("SEO Description", pd.Series([""] * len(df))).astype(str)

                    over_idxs = []
                    for i in range(len(df)):
                        title_len = len(st.iat[i].strip())
                        desc_len = len(sd.iat[i].strip())
                        if (title_len > 60) or (desc_len > 320):
                            over_idxs.append((i, title_len, desc_len))

                    if over_idxs:
                        codes.add("111")
                        lines = []
                        for (i, tl, dl) in over_idxs[:60]:  # cap detail lines
                            rowno = i + 2
                            title = df.at[i, "Title*"] if "Title*" in df.columns else ""
                            lines.append(
                                f"- Row {rowno}: Title='{title}'  SEO Title len={tl}  SEO Description len={dl}")
                        more = f"\n  ... and {len(over_idxs) - 60} more row(s)" if len(over_idxs) > 60 else ""
                        sections.append("Error 111: SEO Length Limits (Title > 60 or Description > 320)\n" + "\n".join(
                            lines) + more)

            # Error 107: Title* present but Body (HTML) blank
            if "Title*" in df.columns and "Body (HTML)" in df.columns:
                title_nonempty = df["Title*"].astype(str).str.strip().ne("")
                body_blank = df["Body (HTML)"].astype(str).str.strip().eq("")
                idxs = list(df.index[title_nonempty & body_blank])
                if idxs:
                    codes.add("107")
                    lines = []
                    for i in idxs[:40]:
                        t = df.at[i, "Title*"]
                        lines.append(f"- Row {i+2}: {t}")
                    more = f"\n  ... and {len(idxs)-40} more row(s)" if len(idxs) > 40 else ""
                    sections.append("Error 107: Missing Body (HTML) on rows\n" + "\n".join(lines) + more)

            # Error 102: duplicate titles inside template
            dup_inside = []
            if "Title*" in df.columns:
                tnorm = df["Title*"].astype(str).str.strip().str.lower()
                vc = tnorm.value_counts()
                dups = vc[vc > 1]
                if not dups.empty:
                    seen_map = {}
                    for t in df["Title*"].astype(str):
                        k = t.strip().lower()
                        if k and k not in seen_map:
                            seen_map[k] = t
                    for k,cnt in dups.items():
                        dup_inside.append(f"- {seen_map.get(k,k)} (x{int(cnt)})")
            if dup_inside:
                codes.add("102")
                sections.append("Error 102: Duplicate Titles in Template\n" + "\n".join(dup_inside))

            # Error 112: Very Short/Placeholder Description
            if "Body (HTML)" in df.columns:
                idxs_placeholder = []
                body_series = df["Body (HTML)"].astype(str)
                title_series = df.get("Title*", pd.Series([""] * len(df))).astype(str)

                for i in range(len(df)):
                    body = body_series.iat[i]
                    title = title_series.iat[i].strip()
                    if body.strip():  # only evaluate if something is present; blank is Error 107
                        if _looks_like_placeholder_body(body):
                            idxs_placeholder.append(i)

                if idxs_placeholder:
                    codes.add("112")
                    lines = []
                    for i in idxs_placeholder[:60]:  # cap detail lines for dialog
                        rowno = i + 2
                        t = title_series.iat[i] if i < len(title_series) else ""
                        lines.append(f"- Row {rowno}: {t}")
                    more = f"\n  ... and {len(idxs_placeholder) - 60} more row(s)" if len(idxs_placeholder) > 60 else ""
                    sections.append("Error 112: Very Short/Placeholder Description\n" + "\n".join(lines) + more)

            # Error 102 also: duplicates against previous export titles
            dup_against_export = []
            prev_titles_set = set()
            prev_path = Path(prev_path_str) if prev_path_str else None
            if prev_path and prev_path.exists():
                try:
                    if prev_path.suffix.lower() in {".xlsx",".xls"}:
                        p = pd.read_excel(prev_path, dtype=str).fillna("")
                    else:
                        p = pd.read_csv(prev_path, dtype=str).fillna("")
                    if "Title" in p.columns:
                        prev_titles_set = set(p["Title"].astype(str).str.strip().str.lower().tolist())
                except Exception:
                    pass
            if prev_titles_set and "Title*" in df.columns:
                for t in df["Title*"].astype(str):
                    k = t.strip().lower()
                    if k and k in prev_titles_set:
                        dup_against_export.append(f"- {t}")
            if dup_against_export:
                codes.add("102")
                sections.append("Error 102: Titles already exist in Previous Export\n" + "\n".join(sorted(set(dup_against_export))[:50]))

            # Error 109: bad handle format
            if "Handle (optional)" in df.columns:
                bad_idxs = []
                col = df["Handle (optional)"].astype(str)
                for i, val in col.items():
                    s = str(val).strip()
                    if not s:
                        continue  # optional; blank is fine
                    if not is_valid_handle(s):
                        bad_idxs.append(i)

                if bad_idxs:
                    codes.add("109")
                    lines = []
                    for i in bad_idxs[:60]:  # cap to avoid huge dialog
                        title = df.at[i, "Title*"] if "Title*" in df.columns else ""
                        bad = str(df.at[i, "Handle (optional)"]).strip()
                        lines.append(f"- Row {i + 2}: {title} — handle='{bad}'")
                    more = f"\n  ... and {len(bad_idxs) - 60} more row(s)" if len(bad_idxs) > 60 else ""
                    sections.append("Error 109: Bad Handle Format\n" + "\n".join(lines) + more)

            # Error 101: broken images
            broken_lines = []
            if "Title*" in df.columns:
                titles_series = df["Title*"].astype(str)
            else:
                titles_series = pd.Series([""]*len(df))

            for n in range(1,9):
                col = f"Image URL {n}"
                if col in df.columns:
                    for idx,url in df[col].astype(str).items():
                        if url.strip():
                            title = titles_series.iloc[idx] if idx < len(titles_series) else ""
                            ok, note = check_image_url(url)
                            if not ok:
                                broken_lines.append(f"- [{n}] {title} => {url} ({note})")
                                if title.strip():
                                    broken_titles_set.add(title.strip())

            if broken_lines:
                codes.add("101")
                sections.append("Error 101: Broken Image Link\n" + "\n".join(broken_lines[:200]))

            # Error 103 / 104 for previous export file
            if prev_path and prev_path.exists():
                try:
                    if prev_path.suffix.lower() in {".xlsx",".xls"}:
                        p = pd.read_excel(prev_path, dtype=str)
                    else:
                        p = pd.read_csv(prev_path, dtype=str)
                    if p is None or p.empty:
                        codes.add("104")
                        sections.append("Error 104: Blank/Empty Previous Export\n- The selected previous export file has no rows.")
                    else:
                        highest = load_prev_highest_base(prev_path)
                        if highest == 0:
                            codes.add("103")
                            sections.append("Error 103: Unable to find Highest SKU\n- 'Variant SKU' column missing or contains no valid 6-digit base like 110357/110357-01.")
                except Exception as e:
                    codes.add("103")
                    sections.append(f"Error 103: Unable to read Previous Export\n- {e}")
            elif prev_path_str:
                codes.add("104")
                sections.append("Error 104: Previous Export not found\n- The selected file path does not exist.")

            if codes:
                header = [f"Products found (non-empty Title*): {total}"]
                detail = "\n\n".join(header + sections)
                self.q.put(("__VALIDATION_FAIL__", {
                    "detail": detail,
                    "codes": list(codes),
                    "broken_titles": sorted(broken_titles_set)
                }))
                return

            self.q.put(("__VALIDATION_OK__", {"detail": f"Validation passed.\nProducts found: {total}"}))
        except Exception as e:
            self.q.put(("__VALIDATION_FAIL__", {
                "detail": f"Unexpected error during validation:\n\n{e}",
                "codes": ["104"],
                "broken_titles": []
            }))

    def _poll_validation_only(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if isinstance(msg, tuple) and msg and msg[0] in {"__VALIDATION_OK__","__VALIDATION_FAIL__"}:
                    token, payload = msg
                    self.prog.stop(); self.phase="idle"
                    self.btn_validate.config(state="normal")
                    self.btn_run.config(state="normal")
                    self.btn_open_out.config(state="disabled")
                    if token == "__VALIDATION_OK__":
                        self.last_validation = {
                            "ran": True,
                            "has_errors": False,
                            "summary": payload.get("detail","OK"),
                            "codes": set(),
                            "broken_titles": set()
                        }
                        self.status_bar.config(text="Validation passed.")
                        messagebox.showinfo(APP_TITLE, "Validation passed. You can Run now.")
                    else:
                        codes = set(payload.get("codes", []))
                        detail = payload.get("detail", "Issues found")
                        detail = detail + "\n\n" + ("—" * 60) + "\n" + build_fix_tips(codes)
                        broken_titles = set(payload.get("broken_titles", []))
                        self.last_validation = {
                            "ran": True,
                            "has_errors": True,
                            "summary": detail,
                            "codes": codes,
                            "broken_titles": broken_titles
                        }
                        self.status_bar.config(text="Validation found issues.")
                        self._show_error_dialog(detail)
                    return
                else:
                    self._log(msg if isinstance(msg,str) else str(msg))
        except queue.Empty:
            self.after(50, self._poll_validation_only)

    # ----- run -----
    def _run_only(self):
        if not self.last_validation["ran"]:
            messagebox.showinfo(APP_TITLE, "Please run Validate first.")
            return

        if self.last_validation["has_errors"]:
            codes = self.last_validation.get("codes", set())
            if not (self.proceed_despite_errors.get() and codes and codes.issubset({"101"})):
                messagebox.showwarning(
                    APP_TITLE,
                    "Validation found issues that cannot be bypassed.\n\n"
                    "You may only proceed when the ONLY error is 'Error 101: Broken Image Link' and the checkbox is ticked.\n"
                    "Fix other errors and Validate again."
                )
                return

        label = simpledialog.askstring(APP_TITLE, "Enter Custom Label for the output filename (required):", parent=self)
        if label is None or not str(label).strip():
            messagebox.showinfo(APP_TITLE, "Run cancelled. Custom Label is required.")
            return
        self._run_custom_label = sanitize_filename_part(label)

        script = Path(self.script_path.get())
        if not script.exists():
            alt = filedialog.askopenfilename(title="Locate final-script.py", filetypes=[("Python","*.py"),("All","*.*")])
            if not alt:
                messagebox.showerror(APP_TITLE, "Cannot continue without final-script.py")
                return
            self.script_path.set(alt)

        if not self.input_path.get().strip():
            messagebox.showerror(APP_TITLE, "Choose an Input Excel file."); return

        outdir = self.out_dir.get().strip() or str(Path.cwd())
        Path(outdir).mkdir(parents=True, exist_ok=True)
        self._current_outdir = outdir

        prev = self.prev_path.get().strip()
        args = [sys.executable, str(self.script_path.get()), "--input", self.input_path.get().strip(),
                "--outdir", outdir, "--sheet", self.sheet_name.get().strip() or "Products"]
        if prev: args += ["--prev", prev]

        self.btn_validate.config(state="disabled")
        self.btn_run.config(state="disabled")
        self.btn_open_out.config(state="disabled")
        self._clear_log(); self._log("Launching builder...\n\n")
        self.status_bar.config(text="Building files...")
        self.prog.start(12); self.phase="script"
        t = threading.Thread(target=self._worker, args=(args,), daemon=True)
        t.start(); self.after(50, self._poll_queue)

    def _worker(self, args):
        try:
            proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, universal_newlines=True)
            self.proc = proc
            for line in proc.stdout:
                self.q.put(line)
            rc = proc.wait()
            self._last_exit_code = rc
        except Exception as e:
            self._last_exit_code = -1
            self.q.put(f"\nERROR: {e}\n")
        finally:
            self.q.put("__DONE__")

    def _poll_queue(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg == "__DONE__":
                    self._finish_run(); return
                self._log(msg)
        except queue.Empty:
            self.after(50, self._poll_queue)

    def _find_shopify_import_csv(self, outdir: str):
        """
        Try to locate a Shopify CSV in outdir when 'shopify_import.csv' isn't present.
        Returns a Path or None.
        """
        try:
            d = Path(outdir)
            if not d.exists():
                return None
            # 1) Exact expected name
            p = d / "shopify_import.csv"
            if p.exists():
                return p
            # 2) Any file that looks like the raw/exported CSV
            for pat in ["shopify_import*.csv", "Shopify Product Import*.csv", "*shopify*.csv"]:
                matches = sorted(d.glob(pat), key=lambda x: x.stat().st_mtime, reverse=True)
                if matches:
                    return matches[0]
        except Exception:
            pass
        return None

    def _finish_run(self):
        self.prog.stop()
        self.phase = "idle"
        ok = (self._last_exit_code == 0)

        # Ensure variable exists in scope
        new_path = None

        if ok:
            try:
                if not self._current_outdir:
                    raise RuntimeError("Output folder not set (self._current_outdir is empty).")

                # Try to find the CSV (fallback to glob if the exact name isn't there)
                found_csv = self._find_shopify_import_csv(self._current_outdir)
                if not found_csv:
                    raise FileNotFoundError(
                        f"'shopify_import.csv' not found at: {Path(self._current_outdir) / 'shopify_import.csv'}"
                    )

                codes = self.last_validation.get("codes", set()) or set()
                broken_titles = self.last_validation.get("broken_titles", set()) or set()
                chosen_status = (self.status_choice.get() or "active").strip().lower()

                # Apply status with special case for broken images (Error 101)
                if codes.issubset({"101"}) and self.proceed_despite_errors.get() and broken_titles:
                    self._apply_status_with_broken_images(self._current_outdir, chosen_status, broken_titles)
                    self._log(
                        f"\nApplied Status='{chosen_status}' to all, but set products with broken images to 'draft' "
                        f"({len(broken_titles)} title(s))."
                    )
                else:
                    # Make sure we operate on the file we actually found
                    # If your _force_status_in_csv only accepts a folder, leave as-is;
                    # otherwise you can adapt it to accept a path.
                    self._force_status_in_csv(self._current_outdir, chosen_status)
                    self._log(f"\nApplied Status='{chosen_status}' to {found_csv.name}")

                # Rename the file we actually found
                new_path = self._rename_shopify_import(self._current_outdir, self._run_custom_label)
                if new_path:
                    self._log(f"\nRenamed Shopify import file to:\n{new_path}\n")
                else:
                    self._log("\nWARNING: Could not find 'shopify_import.csv' to rename.\n")

            except Exception as e:
                err_msg = (
                    "Post-processing failed.\n\n"
                    f"Reason: {e}\n\n"
                    f"Output folder: {self._current_outdir or '(unknown)'}"
                )
                self._log(f"\nWARNING: Post-processing error: {e}\n")
                self._show_error_dialog(err_msg)

        # Re-enable controls and finish status
        self.btn_validate.config(state="normal")
        self.btn_run.config(state="normal")
        self.btn_open_out.config(state="normal" if ok else "disabled")
        self.status_bar.config(text="Done." if ok else "Run finished with errors.")
        self._log(f"\nExit code: {self._last_exit_code}\n")
        if ok:
            self._log("Open the output folder to see generated files.\n")

    def _rename_shopify_import(self, outdir: str, label: str):
        from datetime import datetime
        import random

        src = Path(outdir) / "shopify_import.csv"
        if not src.exists():
            return None

        date_str = datetime.now().strftime("%d-%m-%Y")
        label_clean = sanitize_filename_part(label) or "Batch"
        for _ in range(20):
            rnd = random.randint(10000, 99999)
            dst = Path(outdir) / f"Shopify Product Import - {date_str} - {label_clean} - {rnd}.csv"
            if not dst.exists():
                try:
                    src.rename(dst)
                    return str(dst)
                except Exception:
                    continue
        return None

    def _force_status_in_csv(self, outdir: str, status_value: str):
        if pd is None: raise RuntimeError("pandas required to edit output CSV (pip install pandas)")
        if status_value not in {"active","draft"}: return
        out_csv = Path(outdir) / "shopify_import.csv"
        if not out_csv.exists(): raise FileNotFoundError(out_csv)
        df = pd.read_csv(out_csv, dtype=str).fillna("")
        if "Status" not in df.columns: df["Status"] = status_value
        else: df["Status"] = status_value
        df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    def _apply_status_with_broken_images(self, outdir: str, chosen_status: str, broken_titles_set):
        if pd is None: raise RuntimeError("pandas required to edit output CSV (pip install pandas)")
        out_csv = Path(outdir) / "shopify_import.csv"
        if not out_csv.exists(): raise FileNotFoundError(out_csv)

        df = pd.read_csv(out_csv, dtype=str).fillna("")
        df["Status"] = chosen_status if chosen_status in {"active","draft"} else "active"

        broken_norm = {str(t).strip().lower() for t in (broken_titles_set or []) if str(t).strip()}
        if not broken_norm:
            df.to_csv(out_csv, index=False, encoding="utf-8-sig")
            return

        title_norm = df.get("Title", pd.Series([""]*len(df))).astype(str).str.strip().str.lower()
        handles_to_draft = set(df.loc[title_norm.isin(broken_norm), "Handle"].dropna().astype(str).tolist())
        if handles_to_draft:
            df.loc[df["Handle"].astype(str).isin(handles_to_draft), "Status"] = "draft"
        df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    # ----- misc -----
    def _open_change_password(self):
        cfg = load_config()
        dlg = tk.Toplevel(self)
        dlg.title("Change Password"); dlg.resizable(False, False)
        dlg.configure(bg=UI.WHITE)
        frm = ttk.Frame(dlg, padding=UI.LG); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="e", padx=UI.SM, pady=UI.SM)
        uvar = tk.StringVar(value=self.username)
        ttk.Entry(frm, textvariable=uvar, width=24).grid(row=0, column=1, padx=UI.SM, pady=UI.SM)
        ttk.Label(frm, text="Current password:").grid(row=1, column=0, sticky="e", padx=UI.SM, pady=UI.SM)
        pvar = tk.StringVar(); ttk.Entry(frm, textvariable=pvar, show="•", width=24).grid(row=1, column=1, padx=UI.SM, pady=UI.SM)
        ttk.Label(frm, text="New password:").grid(row=2, column=0, sticky="e", padx=UI.SM, pady=UI.SM)
        nvar = tk.StringVar(); ttk.Entry(frm, textvariable=nvar, show="•", width=24).grid(row=2, column=1, padx=UI.SM, pady=UI.SM)
        ttk.Label(frm, text="Confirm new:").grid(row=3, column=0, sticky="e", padx=UI.SM, pady=UI.SM)
        cvar = tk.StringVar(); ttk.Entry(frm, textvariable=cvar, show="•", width=24).grid(row=3, column=1, padx=UI.SM, pady=UI.SM)

        def save_pw():
            u = uvar.get().strip(); p = pvar.get().strip(); n = nvar.get().strip(); c = cvar.get().strip()
            if not u or not p or not n:
                messagebox.showerror(APP_TITLE, "Fill all fields."); return
            if n != c:
                messagebox.showerror(APP_TITLE, "New passwords do not match."); return
            users = cfg.get("users", [])
            for rec in users:
                if rec.get("username")==u and rec.get("password")==p:
                    rec["password"] = n
                    save_config(cfg)
                    messagebox.showinfo(APP_TITLE, "Password updated."); dlg.destroy(); return
            messagebox.showerror(APP_TITLE, "Invalid username or current password.")

        ttk.Button(frm, text="Save", style="Accent.TButton", command=save_pw).grid(row=4, column=1, sticky="e", padx=UI.SM, pady=(UI.MD,0))
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _logout(self):
        if messagebox.askyesno(APP_TITLE, "Log out?"):
            self.destroy()
            self.master.show_login()

    def _open_outdir(self):
        d = self.out_dir.get().strip() or str(Path.cwd())
        try: os.startfile(d)
        except Exception: messagebox.showinfo(APP_TITLE, d)

    def _clear_log(self): self.txt.delete("1.0","end")
    def _log(self, s:str):
        if not s.endswith("\n"): s = s + "\n"
        self.txt.insert("end", s); self.txt.see("end")


# ===== Main App (container that swaps frames) =====
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x800")
        self.minsize(1100, 720)
        set_ttk_theme(self)

        # anti-tearing for animations on Windows
        try:
            self.tk.call('tk', 'scaling', 1.0)
        except Exception:
            pass

        self._frame = None
        self.show_login()

    def show_login(self):
        if self._frame: self._frame.destroy()
        self._frame = LoginFrame(self, on_login=self._on_login)

    def _on_login(self, username: str):
        if self._frame: self._frame.destroy()
        self._frame = DashboardFrame(self, username=username)


def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
