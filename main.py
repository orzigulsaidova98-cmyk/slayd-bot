#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI SlideBot PRO  -  Universal AI presentation generator for Telegram (single file)
=================================================================================

Istalgan mavzu -> professional PPTX + PDF taqdimot.  Tillar: O'zbekcha, Русский, English.

Talablar (requirements.txt):
    aiogram>=3.7
    aiohttp>=3.9
    python-pptx>=0.6.23
    reportlab>=4.0
    Pillow>=10.0

Muhit o'zgaruvchilari (ENV):
    BOT_TOKEN          - Telegram bot token (majburiy)
    GEMINI_API_KEY     - Gemini API kaliti (majburiy). Bir nechta bo'lsa: GEMINI_API_KEYS=key1,key2,key3
    GEMINI_MODEL       - ixtiyoriy, default: gemini-2.5-flash
    PEXELS_API_KEY     - ixtiyoriy (rasmlar sifati yaxshiroq bo'ladi). Bo'lmasa Wikimedia Commons ishlatiladi
    ADMIN_IDS          - ixtiyoriy, vergul bilan: 12345,67890  (limitsiz + /stats)
    DAILY_LIMIT        - bir foydalanuvchiga kuniga nechta taqdimot (default 10)
    MAX_PARALLEL       - bir vaqtda nechta taqdimot yaratilishi (default 2)
    PORT               - Render uchun (avtomatik beriladi): kichik veb-server ishga tushadi
    DB_PATH            - ixtiyoriy, SQLite fayl yo'li (default: slidebot.db)

Shrift (kirill harflari uchun):
    Dastur tizimdagi DejaVu / Noto / Liberation / Arial shriftlarini avtomatik topadi.
    Topilmasa, bot yoniga  fonts/  papkasini yarating va ichiga
    DejaVuSans.ttf + DejaVuSans-Bold.ttf  (yoki arial.ttf + arialbd.ttf) qo'ying.
"""
from __future__ import annotations

import asyncio
import html
import io
import json
import logging
import math
import os
import random
import re
import shutil
import sqlite3
import tempfile
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

import aiohttp
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Pt
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdfcanvas

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.enums import ChatAction, ParseMode
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("slidebot")

# =============================================================================
# CONFIG
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_KEYS = [
    k.strip()
    for k in (os.getenv("GEMINI_API_KEYS") or os.getenv("GEMINI_API_KEY") or "").split(",")
    if k.strip()
]
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()
PEXELS_KEY = os.getenv("PEXELS_API_KEY", "").strip()
ADMIN_IDS = {int(x) for x in re.findall(r"\d+", os.getenv("ADMIN_IDS", ""))}
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT", "10") or 10)
MAX_PARALLEL = max(1, int(os.getenv("MAX_PARALLEL", "2") or 2))
PORT = int(os.getenv("PORT", "0") or 0)
DB_PATH = os.getenv("DB_PATH", str(BASE_DIR / "slidebot.db"))
USER_AGENT = "AISlideBotPRO/1.0 (Telegram presentation bot)"

SLIDE_COUNTS = [6, 8, 10, 12, 15]
DEFAULT_COUNT = 10
LANGS = ["uz", "ru", "en"]
LANG_FLAG = {"uz": "🇺🇿 UZ", "ru": "🇷🇺 RU", "en": "🇬🇧 EN"}

HTTP: Optional[aiohttp.ClientSession] = None
GEN_SEM: Optional[asyncio.Semaphore] = None
_key_counter = 0

# =============================================================================
# FONTS  (kirill + o' g' harflari uchun Unicode shrift kerak)
# =============================================================================
FR, FB = "SB-Regular", "SB-Bold"
FONT_CMAP: set = set()
FONT_PAIRS = [
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf"),
    ("NotoSans-Regular.ttf", "NotoSans-Bold.ttf"),
    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf"),
    ("arial.ttf", "arialbd.ttf"),
    ("Arial.ttf", "Arial Bold.ttf"),
    ("Arial.ttf", "Arial Bold.ttf"),
]
FONT_DIRS = [
    BASE_DIR / "fonts",
    Path("/usr/share/fonts/truetype/dejavu"),
    Path("/usr/share/fonts/dejavu"),
    Path("/usr/share/fonts/TTF"),
    Path("/usr/share/fonts/truetype/noto"),
    Path("/usr/share/fonts/opentype/noto"),
    Path("/usr/share/fonts/truetype/liberation"),
    Path("/usr/share/fonts/truetype/liberation2"),
    Path("/usr/share/fonts/liberation"),
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("C:/Windows/Fonts"),
]
FONT_DOWNLOADS = [
    (
        "DejaVuSans.ttf",
        [
            "https://cdn.jsdelivr.net/gh/dejavu-fonts/dejavu-fonts@version_2_37/ttf/DejaVuSans.ttf",
            "https://raw.githubusercontent.com/dejavu-fonts/dejavu-fonts/version_2_37/ttf/DejaVuSans.ttf",
        ],
    ),
    (
        "DejaVuSans-Bold.ttf",
        [
            "https://cdn.jsdelivr.net/gh/dejavu-fonts/dejavu-fonts@version_2_37/ttf/DejaVuSans-Bold.ttf",
            "https://raw.githubusercontent.com/dejavu-fonts/dejavu-fonts/version_2_37/ttf/DejaVuSans-Bold.ttf",
        ],
    ),
]


def _find_font_pair() -> Optional[tuple]:
    for reg, bold in FONT_PAIRS:
        for d in FONT_DIRS:
            a, b = d / reg, d / bold
            if a.is_file() and b.is_file():
                try:
                    f = TTFont("probe", str(a))
                    if ord("Ж") in f.face.charToGlyph and ord("o") in f.face.charToGlyph:
                        return str(a), str(b)
                except Exception:  # noqa: BLE001
                    continue
    return None


def _download_fonts() -> None:
    import urllib.request

    target = BASE_DIR / "fonts"
    target.mkdir(exist_ok=True)
    for name, urls in FONT_DOWNLOADS:
        dest = target / name
        if dest.is_file():
            continue
        for u in urls:
            try:
                req = urllib.request.Request(u, headers={"User-Agent": USER_AGENT})
                with urllib.request.urlopen(req, timeout=25) as resp:
                    data = resp.read()
                if len(data) > 100_000 and data[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO"):
                    dest.write_bytes(data)
                    log.info("Font yuklandi: %s", name)
                    break
            except Exception as e:  # noqa: BLE001
                log.warning("Font yuklab bo'lmadi (%s): %s", u, e)


def setup_fonts() -> None:
    """Shriftlarni reportlab'ga ro'yxatdan o'tkazadi (kirill harflarini qo'llab-quvvatlashi shart)."""
    global FONT_CMAP
    pair = _find_font_pair()
    if not pair:
        try:
            _download_fonts()
        except Exception as e:  # noqa: BLE001
            log.warning("Font yuklash xatosi: %s", e)
        pair = _find_font_pair()
    if pair:
        reg, bold = pair
        pdfmetrics.registerFont(TTFont(FR, reg))
        pdfmetrics.registerFont(TTFont(FB, bold))
        FONT_CMAP = set(pdfmetrics.getFont(FR).face.charToGlyph.keys()) & set(
            pdfmetrics.getFont(FB).face.charToGlyph.keys()
        )
        log.info("Shrift: %s", reg)
        return
    # oxirgi chora: reportlab bilan birga keladigan Vera (kirill yo'q!)
    import reportlab

    vdir = Path(reportlab.__file__).parent / "fonts"
    pdfmetrics.registerFont(TTFont(FR, str(vdir / "Vera.ttf")))
    pdfmetrics.registerFont(TTFont(FB, str(vdir / "VeraBd.ttf")))
    FONT_CMAP = set(pdfmetrics.getFont(FR).face.charToGlyph.keys())
    log.error(
        "DIQQAT: kirill harflarini qo'llab-quvvatlovchi shrift topilmadi! "
        "fonts/ papkasiga DejaVuSans.ttf va DejaVuSans-Bold.ttf qo'ying."
    )


def sanitize(s: Any) -> str:
    """Matnni tozalaydi: boshqaruv belgilari, shriftda yo'q belgilar (emoji va h.k.) olib tashlanadi."""
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufeff\u200b-\u200f\u202a-\u202e]", "", s)
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")
    if not FONT_CMAP:
        return s
    out = []
    for ch in s:
        o = ord(ch)
        if ch == "\n" or o in FONT_CMAP:
            out.append(ch)
        elif ch in "ʻʼ‘’`´ʹ":
            out.append("'")
        elif ch in "“”„«»":
            out.append('"')
        elif ch in "–—―‒":
            out.append("-")
        elif ch in "\u00a0\u2009\u202f":
            out.append(" ")
    return "".join(out)


def sw(s: str, size: float, bold: bool = False) -> float:
    return pdfmetrics.stringWidth(s, FB if bold else FR, size)


def wrap_lines(s: str, size: float, bold: bool, width: float) -> list:
    lines: list = []
    for para in s.split("\n"):
        words = para.split()
        if not words:
            lines.append("")
            continue
        cur = ""
        for w in words:
            while sw(w, size, bold) > width and len(w) > 1:
                k = len(w)
                while k > 1 and sw(w[:k], size, bold) > width:
                    k -= 1
                if cur:
                    lines.append(cur)
                    cur = ""
                lines.append(w[:k])
                w = w[k:]
            if not w:
                continue
            t = f"{cur} {w}" if cur else w
            if sw(t, size, bold) <= width:
                cur = t
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
    return lines


def fit(s: str, w: float, h: float, max_size: float, min_size: float = 12, bold: bool = False, lh: float = 1.22):
    """Matn berilgan qutiga sig'adigan eng katta shrift o'lchamini topadi. (size, lines, text) qaytaradi."""
    s = sanitize(s).strip()
    size = int(max_size)
    min_size = int(min_size)
    while size >= min_size:
        lines = wrap_lines(s, size, bold, w)
        if len(lines) * size * lh <= h + 0.5:
            return size, lines, s
        size -= 1
    size = min_size
    lines = wrap_lines(s, size, bold, w)
    maxl = max(1, int(h // (size * lh)))
    if len(lines) > maxl:
        lines = lines[:maxl]
        last = lines[-1]
        while last and sw(last + "...", size, bold) > w:
            last = last[:-1]
        lines[-1] = last.rstrip() + "..."
        return size, lines, " ".join(lines)
    return size, lines, s


# =============================================================================
# COLORS / THEMES
# =============================================================================
def hex_rgb(h: str) -> tuple:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgb_hex(c: tuple) -> str:
    return "#%02X%02X%02X" % tuple(max(0, min(255, int(round(v)))) for v in c)


def mix(a: str, b: str, t: float) -> str:
    ca, cb = hex_rgb(a), hex_rgb(b)
    return rgb_hex(tuple(ca[i] * (1 - t) + cb[i] * t for i in range(3)))


def contrast(bg: str) -> str:
    r, g, b = hex_rgb(bg)
    return "#1A1A1A" if (0.299 * r + 0.587 * g + 0.114 * b) > 150 else "#FFFFFF"


THEMES = {
    "industry": dict(dark=False, bg="#F5F6F8", card="#FFFFFF", primary="#1F2933", accent="#F2711C", text="#1F2933",
                     muted="#667381", border="#DFE3E8", pal=["#F2711C", "#1F2933", "#3E7CB1", "#9AA5B1", "#E0A458", "#5C6B7A"]),
    "nature": dict(dark=False, bg="#F4F9F1", card="#FFFFFF", primary="#1B4332", accent="#52B788", text="#1B2E24",
                   muted="#5F7A6B", border="#DAEADB", pal=["#52B788", "#1B4332", "#95D5B2", "#D9A441", "#2D6A4F", "#8AB17D"]),
    "tech": dict(dark=True, bg="#0B1220", card="#141F36", primary="#0F1A30", accent="#38BDF8", text="#E8EEF9",
                 muted="#93A4C0", border="#243252", pal=["#38BDF8", "#A78BFA", "#34D399", "#F472B6", "#FBBF24", "#60A5FA"]),
    "business": dict(dark=False, bg="#F8F7F3", card="#FFFFFF", primary="#0F2A47", accent="#D69E2E", text="#14263B",
                     muted="#64748B", border="#E6E2D6", pal=["#D69E2E", "#0F2A47", "#3B82A0", "#94A3B8", "#B45309", "#1E40AF"]),
    "creative": dict(dark=False, bg="#FFF7F5", card="#FFFFFF", primary="#2E1065", accent="#F43F5E", text="#2A1744",
                     muted="#7C6A8F", border="#F1DCE2", pal=["#F43F5E", "#7C3AED", "#F59E0B", "#2E1065", "#EC4899", "#06B6D4"]),
    "science": dict(dark=False, bg="#F3F8FB", card="#FFFFFF", primary="#0B3C5D", accent="#00A6A6", text="#0B2A3C",
                    muted="#5B7587", border="#D7E7F0", pal=["#00A6A6", "#0B3C5D", "#328CC1", "#D9B310", "#1D2731", "#7FC8F8"]),
    "education": dict(dark=False, bg="#F6F8FE", card="#FFFFFF", primary="#1E3A8A", accent="#F59E0B", text="#172554",
                      muted="#64748B", border="#DCE3F6", pal=["#F59E0B", "#1E3A8A", "#3B82F6", "#10B981", "#EF4444", "#8B5CF6"]),
    "food": dict(dark=False, bg="#FFF9F1", card="#FFFFFF", primary="#5A2A10", accent="#E85D04", text="#3B1D0B",
                 muted="#8A6A55", border="#F1E1CD", pal=["#E85D04", "#5A2A10", "#FAA307", "#6A994E", "#BC4749", "#9C6644"]),
    "health": dict(dark=False, bg="#F2FAF9", card="#FFFFFF", primary="#0F5C58", accent="#14B8A6", text="#0B3B38",
                   muted="#5E8582", border="#D3EBE8", pal=["#14B8A6", "#0F5C58", "#38BDF8", "#F59E0B", "#EF4444", "#6366F1"]),
    "history": dict(dark=False, bg="#FAF5EA", card="#FFFDF7", primary="#4A2C17", accent="#B7791F", text="#3A2413",
                    muted="#85694F", border="#E9DCC0", pal=["#B7791F", "#4A2C17", "#9C4221", "#6B8E5A", "#2C5282", "#C05621"]),
    "earth": dict(dark=False, bg="#F1F7F8", card="#FFFFFF", primary="#12455A", accent="#F28F3B", text="#0F3340",
                  muted="#58798A", border="#D5E6EA", pal=["#F28F3B", "#12455A", "#4FA3A5", "#8CB369", "#BC4B51", "#5B8E7D"]),
}
FIELD_THEME = {
    "manufacturing": "industry", "daily_life": "business", "technology": "tech", "business": "business",
    "art": "creative", "nature": "nature", "science": "science", "math": "science", "education": "education",
    "history": "history", "geography": "earth", "health": "health", "food": "food", "sport": "health",
}


class Theme:
    def __init__(self, name: str) -> None:
        d = THEMES.get(name) or THEMES["education"]
        self.name = name if name in THEMES else "education"
        for k, v in d.items():
            setattr(self, k, v)
        self.heading = self.text if self.dark else self.primary
        self.on_primary = contrast(self.primary)
        self.on_accent = contrast(self.accent)
        self.on_primary_soft = mix(self.on_primary, self.primary, 0.25)


# =============================================================================
# RENDERERS  (bir xil layout -> PPTX va PDF).  Koordinatalar: 1280 x 720 piksel, yuqori-chap boshlang'ich
# =============================================================================
W, H = 1280, 720
PX = 9525  # 1 px = 9525 EMU  (1280 px = 13.333 in)


class PptxRenderer:
    def __init__(self) -> None:
        self.prs = Presentation()
        self.prs.slide_width = Emu(W * PX)
        self.prs.slide_height = Emu(H * PX)
        self.slide = None

    @staticmethod
    def _rgb(c: str) -> RGBColor:
        return RGBColor(*hex_rgb(c))

    def new_slide(self, bg: str) -> None:
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        f = self.slide.background.fill
        f.solid()
        f.fore_color.rgb = self._rgb(bg)

    def _style(self, shp, fill, stroke, sw_):
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = self._rgb(fill)
        else:
            shp.fill.background()
        if stroke:
            shp.line.color.rgb = self._rgb(stroke)
            shp.line.width = Pt(sw_ * 0.75)
        else:
            shp.line.fill.background()
        try:
            shp.shadow.inherit = False
        except Exception:  # noqa: BLE001
            pass

    def rect(self, x, y, w, h, fill=None, stroke=None, sw=1.0, r=0):
        r = min(r, w / 2, h / 2)
        kind = MSO_SHAPE.ROUNDED_RECTANGLE if r > 0 else MSO_SHAPE.RECTANGLE
        shp = self.slide.shapes.add_shape(kind, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
        if r > 0:
            shp.adjustments[0] = min(0.5, r / max(1.0, min(w, h)))
        self._style(shp, fill, stroke, sw)

    def oval(self, x, y, w, h, fill=None, stroke=None, sw=1.0):
        shp = self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
        self._style(shp, fill, stroke, sw)

    def line(self, x1, y1, x2, y2, color, w=2.0, dash=False):
        c = self.slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Emu(int(x1 * PX)), Emu(int(y1 * PX)), Emu(int(x2 * PX)), Emu(int(y2 * PX))
        )
        c.line.color.rgb = self._rgb(color)
        c.line.width = Pt(w * 0.75)
        if dash:
            c.line.dash_style = MSO_LINE.DASH

    def poly(self, pts, fill=None, stroke=None, sw=1.0):
        pts = [(int(px * PX), int(py * PX)) for px, py in pts]
        fb = self.slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1.0)
        fb.add_line_segments(pts[1:], close=True)
        shp = fb.convert_to_shape()
        self._style(shp, fill, stroke, sw)

    def text(self, x, y, w, h, s, size, color, bold=False, align="l", valign="t", lh=1.22):
        s = sanitize(s)
        tb = self.slide.shapes.add_textbox(Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[valign]
        al = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        for i, line in enumerate(s.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = al
            run = p.add_run()
            run.text = line
            f = run.font
            f.name = "Arial"
            f.size = Pt(round(size * 0.75, 1))
            f.bold = bool(bold)
            f.color.rgb = self._rgb(color)

    def image(self, path, x, y, w, h):
        self.slide.shapes.add_picture(path, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))

    def notes(self, s):
        s = sanitize(s).strip()
        if s:
            self.slide.notes_slide.notes_text_frame.text = s

    def finish(self, path):
        self.prs.core_properties.title = "AI SlideBot PRO"
        self.prs.core_properties.author = "AI SlideBot PRO"
        self.prs.save(path)


class PdfRenderer:
    def __init__(self, path: str) -> None:
        self.c = pdfcanvas.Canvas(path, pagesize=(W, H))
        self.c.setTitle("AI SlideBot PRO")
        self.c.setAuthor("AI SlideBot PRO")
        self.started = False

    def new_slide(self, bg):
        if self.started:
            self.c.showPage()
        self.started = True
        self.c.setFillColor(HexColor(bg))
        self.c.rect(0, 0, W, H, stroke=0, fill=1)

    def _paint(self, fill, stroke, sw_):
        if fill:
            self.c.setFillColor(HexColor(fill))
        if stroke:
            self.c.setStrokeColor(HexColor(stroke))
            self.c.setLineWidth(sw_)
        return 1 if fill else 0, 1 if stroke else 0

    def rect(self, x, y, w, h, fill=None, stroke=None, sw=1.0, r=0):
        f, s = self._paint(fill, stroke, sw)
        if not f and not s:
            return
        r = min(r, w / 2, h / 2)
        if r > 0:
            self.c.roundRect(x, H - y - h, w, h, r, stroke=s, fill=f)
        else:
            self.c.rect(x, H - y - h, w, h, stroke=s, fill=f)

    def oval(self, x, y, w, h, fill=None, stroke=None, sw=1.0):
        f, s = self._paint(fill, stroke, sw)
        if f or s:
            self.c.ellipse(x, H - y - h, x + w, H - y, stroke=s, fill=f)

    def line(self, x1, y1, x2, y2, color, w=2.0, dash=False):
        self.c.setStrokeColor(HexColor(color))
        self.c.setLineWidth(w)
        self.c.setDash(6, 4) if dash else self.c.setDash()
        self.c.line(x1, H - y1, x2, H - y2)
        self.c.setDash()

    def poly(self, pts, fill=None, stroke=None, sw=1.0):
        f, s = self._paint(fill, stroke, sw)
        if not f and not s:
            return
        p = self.c.beginPath()
        p.moveTo(pts[0][0], H - pts[0][1])
        for px, py in pts[1:]:
            p.lineTo(px, H - py)
        p.close()
        self.c.drawPath(p, stroke=s, fill=f)

    def text(self, x, y, w, h, s, size, color, bold=False, align="l", valign="t", lh=1.22):
        s = sanitize(s)
        lines = wrap_lines(s, size, bold, w)
        total = len(lines) * size * lh
        top = y if valign == "t" else (y + (h - total) / 2 if valign == "m" else y + h - total)
        self.c.setFillColor(HexColor(color))
        self.c.setFont(FB if bold else FR, size)
        for i, line in enumerate(lines):
            base = top + i * size * lh + size * (0.5 * (lh - 1) + 0.82)
            yy = H - base
            if align == "c":
                self.c.drawCentredString(x + w / 2, yy, line)
            elif align == "r":
                self.c.drawRightString(x + w, yy, line)
            else:
                self.c.drawString(x, yy, line)

    def image(self, path, x, y, w, h):
        self.c.drawImage(path, x, H - y - h, w, h, mask="auto")

    def notes(self, s):
        pass

    def finish(self, path=None):
        self.c.showPage()
        self.c.save()


# =============================================================================
# IMAGES
# =============================================================================
IMG_SCALE = 1.5


def make_art(w: int, h: int, th: Theme, seed: int) -> Image.Image:
    """Rasm topilmasa — mavzu ranglarida chiroyli abstrakt fon."""
    rnd = random.Random(seed)
    c1 = hex_rgb(th.primary)
    c2 = hex_rgb(mix(th.accent, th.primary, 0.55))
    base = Image.new("RGB", (w, h))
    gd = ImageDraw.Draw(base)
    step = 3
    for yy in range(0, h, step):
        t = yy / max(1, h)
        gd.rectangle([0, yy, w, yy + step], fill=tuple(int(c1[i] * (1 - t) + c2[i] * t) for i in range(3)))
    img = base.convert("RGBA")
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    m = min(w, h)
    for _ in range(9):
        rr = rnd.randint(int(m * 0.12), int(m * 0.42))
        cx, cy = rnd.randint(0, w), rnd.randint(0, h)
        col = hex_rgb(rnd.choice(th.pal))
        d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], fill=col + (rnd.randint(45, 105),))
    for _ in range(3):
        x0 = rnd.randint(-w // 3, w)
        bw = rnd.randint(int(w * 0.08), int(w * 0.22))
        col = hex_rgb(rnd.choice(th.pal))
        d.polygon([(x0, 0), (x0 + bw, 0), (x0 + bw - h // 2, h), (x0 - h // 2, h)], fill=col + (55,))
    for _ in range(3):
        rr = rnd.randint(int(m * 0.2), int(m * 0.5))
        cx, cy = rnd.randint(0, w), rnd.randint(0, h)
        d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], outline=(255, 255, 255, 70), width=max(2, w // 150))
    return Image.alpha_composite(img, layer).convert("RGB")


class ImageStore:
    def __init__(self, workdir: str, th: Theme) -> None:
        self.dir = Path(workdir)
        self.th = th
        self.cache: dict = {}

    def get(self, slide: dict, w: float, h: float, radius: float = 0, allow_art: bool = False) -> Optional[str]:
        key = (slide.get("_i"), int(w), int(h), int(radius))
        if key in self.cache:
            return self.cache[key]
        tw_, th_ = int(w * IMG_SCALE), int(h * IMG_SCALE)
        img = None
        raw = slide.get("image_raw")
        if raw:
            try:
                img = Image.open(io.BytesIO(raw))
                img.load()
                img = img.convert("RGB")
            except Exception as e:  # noqa: BLE001
                log.warning("Rasm ochilmadi: %s", e)
                img = None
        if img is None and allow_art:
            img = make_art(tw_, th_, self.th, seed=(slide.get("_i") or 0) * 7 + 3)
        if img is None:
            self.cache[key] = None
            return None
        img = ImageOps.fit(img, (tw_, th_), method=Image.LANCZOS, centering=(0.5, 0.4))
        name = f"img_{len(self.cache)}_{int(w)}x{int(h)}"
        if radius > 0:
            ss = 2
            mask = Image.new("L", (tw_ * ss, th_ * ss), 0)
            ImageDraw.Draw(mask).rounded_rectangle(
                [0, 0, tw_ * ss - 1, th_ * ss - 1], radius=int(radius * IMG_SCALE * ss), fill=255
            )
            mask = mask.resize((tw_, th_), Image.LANCZOS)
            img = img.convert("RGBA")
            img.putalpha(mask)
            path = str(self.dir / (name + ".png"))
            img.save(path, "PNG", optimize=True)
        else:
            path = str(self.dir / (name + ".jpg"))
            img.save(path, "JPEG", quality=86, optimize=True)
        self.cache[key] = path
        return path


async def fetch_image_bytes(query: str, cache: dict) -> Optional[bytes]:
    """Pexels (kalit bo'lsa) yoki Wikimedia Commons'dan rasm topadi."""
    q = re.sub(r"\s+", " ", query).strip()[:80]
    if not q or HTTP is None:
        return None
    if q in cache:
        return cache[q]
    data: Optional[bytes] = None
    urls: list = []
    try:
        if PEXELS_KEY:
            async with HTTP.get(
                "https://api.pexels.com/v1/search",
                params={"query": q, "per_page": 5, "orientation": "landscape"},
                headers={"Authorization": PEXELS_KEY},
                timeout=aiohttp.ClientTimeout(total=15),
            ) as resp:
                if resp.status == 200:
                    js = await resp.json(content_type=None)
                    for ph in js.get("photos") or []:
                        src = ph.get("src") or {}
                        u = src.get("large") or src.get("medium")
                        if u:
                            urls.append(u)
        if not urls:
            async with HTTP.get(
                "https://commons.wikimedia.org/w/api.php",
                params={
                    "action": "query", "format": "json", "generator": "search",
                    "gsrsearch": f"{q} filetype:bitmap", "gsrnamespace": 6, "gsrlimit": 8,
                    "prop": "imageinfo", "iiprop": "url|mime|size", "iiurlwidth": 1280,
                },
                timeout=aiohttp.ClientTimeout(total=15),
            ) as resp:
                if resp.status == 200:
                    js = await resp.json(content_type=None)
                    pages = sorted((js.get("query") or {}).get("pages", {}).values(), key=lambda p: p.get("index", 99))
                    for pg in pages:
                        info = (pg.get("imageinfo") or [{}])[0]
                        if info.get("mime") not in ("image/jpeg", "image/png"):
                            continue
                        if (info.get("width") or 0) < 700 or (info.get("height") or 0) < 450:
                            continue
                        u = info.get("thumburl") or info.get("url")
                        if u:
                            urls.append(u)
        for u in urls[:3]:
            try:
                async with HTTP.get(u, timeout=aiohttp.ClientTimeout(total=25)) as resp:
                    if resp.status != 200:
                        continue
                    b = await resp.content.read(8_000_000)
                    if len(b) > 20_000:
                        data = b
                        break
            except Exception as e:  # noqa: BLE001
                log.info("Rasm yuklash xatosi: %s", e)
    except Exception as e:  # noqa: BLE001
        log.info("Rasm qidirish xatosi (%s): %s", q, e)
    cache[q] = data
    return data


# =============================================================================
# LAYOUT HELPERS
# =============================================================================
CX0, CW = 64, 1152
CY0, CH = 160, 504


def put_text(r, x, y, w, h, s, max_size, color, bold=False, align="l", valign="t", min_size=13, lh=1.22) -> int:
    size, _lines, s2 = fit(s, w, h, max_size, min_size, bold, lh)
    if s2.strip():
        r.text(x, y, w, h, s2, size, color, bold, align, valign, lh)
    return size


def card(r, th: Theme, x, y, w, h, fill=None, radius=16):
    r.rect(x, y, w, h, fill=fill or th.card, stroke=th.border, sw=1.5, r=radius)


def header(r, th: Theme, deck: dict, title: str, idx: int, total: int):
    put_text(r, CX0, 32, CW, 86, title, 40, th.heading, bold=True, valign="m", min_size=22)
    r.rect(CX0, 126, 84, 6, fill=th.accent, r=3)
    put_text(r, CX0, 686, 850, 20, deck["title"], 13, th.muted, min_size=10)
    put_text(r, 1100, 686, 116, 20, f"{idx}/{total}", 13, th.muted, align="r", min_size=10)


def number_badge(r, th: Theme, x, y, d, n, fill=None):
    fill = fill or th.accent
    r.oval(x, y, d, d, fill=fill)
    put_text(r, x, y, d, d, str(n), d * 0.46, contrast(fill), bold=True, align="c", valign="m", min_size=10)


def card_list(r, th: Theme, x, y, w, h, items, max_size=26, min_size=14):
    n = len(items)
    gap = 14
    slot = min(112.0, (h - (n - 1) * gap) / n)
    total = n * slot + (n - 1) * gap
    y0 = y + (h - total) / 2
    pad = 28
    size = max_size
    for it in items:
        sz, _l, _s = fit(it, w - pad - 20, slot - 16, max_size, min_size)
        size = min(size, sz)
    for k, it in enumerate(items):
        yy = y0 + k * (slot + gap)
        card(r, th, x, yy, w, slot, radius=14)
        r.rect(x, yy + 12, 6, slot - 24, fill=th.accent, r=3)
        put_text(r, x + pad, yy + 8, w - pad - 20, slot - 16, it, size, th.text, valign="m", min_size=size)


def dot_list(r, th: Theme, x, y, w, h, items, color, dot_color, max_size=22, min_size=13):
    n = len(items)
    slot = min(88.0, h / n)
    size = max_size
    for it in items:
        sz, _l, _s = fit(it, w - 34, slot - 8, max_size, min_size)
        size = min(size, sz)
    for k, it in enumerate(items):
        yy = y + k * slot
        r.oval(x, yy + slot / 2 - 7, 14, 14, fill=dot_color)
        put_text(r, x + 30, yy + 4, w - 30, slot - 8, it, size, color, valign="m", min_size=size)


def nice_max(v: float) -> float:
    if v <= 0:
        return 1.0
    e = 10 ** math.floor(math.log10(v))
    f = v / e
    for nf in (1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10):
        if f <= nf + 1e-9:
            return nf * e
    return 10 * e


def fmt_num(v: float) -> str:
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}".replace(",", " ")
    if abs(v) >= 100:
        return f"{v:,.0f}".replace(",", " ")
    return f"{v:.2f}".rstrip("0").rstrip(".")


def draw_chart(r, th: Theme, kind: str, labels: list, values: list, unit: str, x, y, w, h):
    n = len(values)
    unit = sanitize(unit)[:30]
    if kind == "pie":
        tot = sum(max(0.0, v) for v in values) or 1.0
        d = min(h - 50, w * 0.52)
        cx, cy = x + 24 + d / 2, y + h / 2
        ang = -90.0
        for i, v in enumerate(values):
            frac = max(0.0, v) / tot
            if frac <= 0:
                continue
            col = th.pal[i % len(th.pal)]
            if frac > 0.9995:
                r.oval(cx - d / 2, cy - d / 2, d, d, fill=col, stroke=th.card, sw=2)
                continue
            sweep = frac * 360
            steps = max(3, int(sweep / 4))
            pts = [(cx, cy)]
            for k in range(steps + 1):
                a = math.radians(ang + sweep * k / steps)
                pts.append((cx + d / 2 * math.cos(a), cy + d / 2 * math.sin(a)))
            r.poly(pts, fill=col, stroke=th.card, sw=2)
            ang += sweep
        lx = cx + d / 2 + 36
        lw = x + w - lx - 12
        row = min(46.0, (h - 20) / n)
        ly0 = y + (h - row * n) / 2
        for i, v in enumerate(values):
            yy = ly0 + i * row
            r.rect(lx, yy + row / 2 - 9, 18, 18, fill=th.pal[i % len(th.pal)], r=4)
            pct = max(0.0, v) / tot * 100
            ptxt = f"{pct:.0f}%" if pct >= 1 or pct == 0 else f"{pct:.1f}%"
            put_text(r, lx + 30, yy, lw - 30 - 66, row, labels[i], 18, th.text, valign="m", min_size=11)
            put_text(r, lx + lw - 64, yy, 64, row, ptxt, 18, th.heading, bold=True, align="r", valign="m", min_size=11)
        return
    pl, pr, pt_, pb = 78, 26, 46, 74
    px, py, pw, ph = x + pl, y + pt_, w - pl - pr, h - pt_ - pb
    lo, hi = min(min(values), 0.0), max(max(values), 0.0)
    if hi - lo <= 0:
        hi = lo + 1
    raw_step = (hi - lo) / 5
    e = 10 ** math.floor(math.log10(raw_step))
    step = next(k * e for k in (1, 2, 2.5, 5, 10) if raw_step <= k * e + 1e-12)
    vmin = math.floor(lo / step + 1e-9) * step
    vmax = math.ceil(hi / step - 1e-9) * step
    if vmax <= vmin:
        vmax = vmin + step
    ticks = int(round((vmax - vmin) / step))

    def Y(v):
        return py + ph - (v - vmin) / (vmax - vmin) * ph

    for i in range(ticks + 1):
        v = vmin + (vmax - vmin) * i / ticks
        yy = Y(v)
        r.line(px, yy, px + pw, yy, th.border, 1.5)
        put_text(r, x + 2, yy - 10, pl - 12, 20, fmt_num(v), 14, th.muted, align="r", valign="m", min_size=10)
    if vmin < 0:
        r.line(px, Y(0), px + pw, Y(0), th.muted, 2)
    slot = pw / n
    if unit:
        put_text(r, x + 2, y + 4, w - 4, 24, unit, 15, th.muted, min_size=10)
    centers = [px + slot * (i + 0.5) for i in range(n)]
    if kind == "line":
        pts = [(centers[i], Y(values[i])) for i in range(n)]
        for i in range(n - 1):
            r.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], th.accent, 5)
        for i, (qx, qy) in enumerate(pts):
            r.oval(qx - 8, qy - 8, 16, 16, fill=th.card, stroke=th.accent, sw=4)
            put_text(r, qx - 60, qy - 36, 120, 22, fmt_num(values[i]), 16, th.text, bold=True, align="c", valign="m", min_size=10)
    else:
        bw = min(84.0, slot * 0.62)
        for i, v in enumerate(values):
            top, bot = Y(max(v, 0)), Y(min(v, 0))
            r.rect(centers[i] - bw / 2, top, bw, max(2.0, bot - top), fill=th.pal[i % len(th.pal)], r=6)
            vy = top - 28 if v >= 0 else bot + 6
            put_text(r, centers[i] - 70, vy, 140, 22, fmt_num(v), 16, th.text, bold=True, align="c", valign="m", min_size=10)
    for i in range(n):
        put_text(r, centers[i] - slot / 2 + 3, py + ph + 12, slot - 6, 54, labels[i], 15, th.text, align="c", min_size=10)


# =============================================================================
# SLAYD MAKETLARI
# =============================================================================
THANKS = {"uz": "E'tiboringiz uchun rahmat!", "ru": "Спасибо за внимание!", "en": "Thank you!"}
CREDIT = {"uz": "Rasmlar: ", "ru": "Изображения: ", "en": "Images: "}


def L_title(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.primary)
    iw = 560
    img = store.get(s, iw, H, 0, allow_art=True)
    if img:
        r.image(img, W - iw, 0, iw, H)
    r.rect(W - iw - 10, 0, 10, H, fill=th.accent)
    r.rect(70, 72, 196, 36, fill=th.accent, r=18)
    put_text(r, 70, 72, 196, 36, "AI SlideBot PRO", 15, th.on_accent, bold=True, align="c", valign="m", min_size=10)
    put_text(r, 70, 140, 570, 300, s.get("title") or deck["title"], 58, th.on_primary, bold=True, valign="m", min_size=26)
    r.rect(70, 462, 110, 6, fill=th.accent, r=3)
    sub = s.get("subtitle") or deck.get("subtitle") or ""
    if sub:
        put_text(r, 70, 490, 570, 150, sub, 26, th.on_primary_soft, min_size=14)


def L_agenda(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    items = s["items"]
    n = len(items)
    cols = 2 if n > 3 else 1
    rows = math.ceil(n / cols)
    gap = 20
    cw = (CW - gap) / 2 if cols == 2 else 900.0
    x0 = CX0 if cols == 2 else CX0 + (CW - cw) / 2
    ch = min(112.0, (CH - 30 - (rows - 1) * gap) / rows)
    total_h = rows * ch + (rows - 1) * gap
    y0 = CY0 + (CH - total_h) / 2
    for k, it in enumerate(items):
        col, row = k // rows, k % rows
        x, y = x0 + col * (cw + gap), y0 + row * (ch + gap)
        card(r, th, x, y, cw, ch)
        number_badge(r, th, x + 22, y + (ch - 54) / 2, 54, k + 1)
        put_text(r, x + 96, y + 8, cw - 116, ch - 16, it, 26, th.text, valign="m", min_size=14)


def L_bullets(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    img = store.get(s, 446, 490, 20, allow_art=False) if s.get("image_raw") else None
    if img:
        card_list(r, th, CX0, CY0 + 4, 660, 496, s["bullets"], max_size=26)
        r.image(img, CX0 + CW - 446, CY0 + 7, 446, 490)
    else:
        card_list(r, th, CX0, CY0 + 4, CW, 496, s["bullets"], max_size=28)


def L_highlight(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    img = store.get(s, 476, 490, 20, allow_art=True)
    lw = 640
    has_b = bool(s.get("bullets"))
    sh = 270 if has_b else 496
    card(r, th, CX0, CY0 + 4, lw, sh, radius=18)
    r.rect(CX0, CY0 + 24, 8, sh - 40, fill=th.accent, r=4)
    put_text(r, CX0 + 40, CY0 + 20, lw - 70, sh - 32, s["statement"], 38, th.heading, bold=True, valign="m", min_size=18)
    if has_b:
        dot_list(r, th, CX0 + 6, CY0 + 292, lw - 6, 208, s["bullets"], th.text, th.accent, max_size=22)
    if img:
        r.image(img, CX0 + CW - 476, CY0 + 7, 476, 490)


def L_process(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    steps = s["steps"]
    n = len(steps)
    if n <= 4:
        gap = 24
        cw = (CW - (n - 1) * gap) / n
        top = 258
        ch = 350
        xs = [CX0 + i * (cw + gap) for i in range(n)]
        r.line(xs[0] + cw / 2, top, xs[-1] + cw / 2, top, th.accent, 4)
        for i, st in enumerate(steps):
            x = xs[i]
            card(r, th, x, top, cw, ch, radius=18)
            put_text(r, x + 16, top + 58, cw - 32, 92, st["title"], 26, th.heading, bold=True, align="c", valign="m", min_size=15)
            r.rect(x + cw / 2 - 28, top + 160, 56, 5, fill=th.accent, r=2)
            put_text(r, x + 20, top + 182, cw - 40, ch - 200, st["text"], 21, th.text, align="c", min_size=13)
        for i in range(n):
            number_badge(r, th, xs[i] + cw / 2 - 38, top - 38, 76, i + 1)
    else:
        cols, gapx, gapy = 3, 24, 22
        cw = (CW - 2 * gapx) / 3
        ch = 236
        for i, st in enumerate(steps):
            row, col = i // cols, i % cols
            in_row = min(cols, n - row * cols)
            off = (cols - in_row) * (cw + gapx) / 2
            x = CX0 + off + col * (cw + gapx)
            y = CY0 + 6 + row * (ch + gapy)
            card(r, th, x, y, cw, ch, radius=18)
            number_badge(r, th, x + 20, y + 20, 58, i + 1)
            put_text(r, x + 94, y + 14, cw - 112, 72, st["title"], 24, th.heading, bold=True, valign="m", min_size=14)
            put_text(r, x + 22, y + 96, cw - 44, ch - 108, st["text"], 20, th.text, min_size=12)


def L_cycle(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    steps = s["steps"]
    n = len(steps)
    cx, cy = 640.0, 412.0
    rx, ry = 350.0, 190.0
    bw, bh = 280.0, 108.0
    centers = []
    for i in range(n):
        a = math.radians(-90 + 360 * i / n)
        centers.append((cx + rx * math.cos(a), cy + ry * math.sin(a)))
    for i in range(n):
        a, b = centers[i], centers[(i + 1) % n]
        r.line(a[0], a[1], b[0], b[1], th.accent, 4)
    d = 190
    r.oval(cx - d / 2, cy - d / 2, d, d, fill=th.primary)
    put_text(r, cx - d / 2 + 16, cy - d / 2 + 16, d - 32, d - 32, s.get("center") or deck["title"], 24,
             th.on_primary, bold=True, align="c", valign="m", min_size=12)
    for i, st in enumerate(steps):
        x, y = centers[i][0] - bw / 2, centers[i][1] - bh / 2
        card(r, th, x, y, bw, bh, radius=16)
        number_badge(r, th, x + 12, y + (bh - 36) / 2, 36, i + 1)
        if st["text"]:
            put_text(r, x + 58, y + 8, bw - 70, 40, st["title"], 19, th.heading, bold=True, valign="m", min_size=12)
            put_text(r, x + 58, y + 50, bw - 70, bh - 58, st["text"], 15, th.text, min_size=10)
        else:
            put_text(r, x + 58, y + 8, bw - 70, bh - 16, st["title"], 20, th.heading, bold=True, valign="m", min_size=12)
    for i in range(n):
        a, b = centers[i], centers[(i + 1) % n]
        mx, my = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
        dx, dy = b[0] - a[0], b[1] - a[1]
        ln = math.hypot(dx, dy) or 1
        ux, uy = dx / ln, dy / ln
        px_, py_ = -uy, ux
        tip = (mx + ux * 13, my + uy * 13)
        r.poly([tip, (mx - ux * 9 + px_ * 10, my - uy * 9 + py_ * 10), (mx - ux * 9 - px_ * 10, my - uy * 9 - py_ * 10)],
               fill=th.accent, stroke=th.bg, sw=1.5)


def L_compare(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    cw = (CW - 32) / 2
    for k, (side, col) in enumerate(((s["left"], th.pal[1] if th.dark else th.primary), (s["right"], th.accent))):
        x = CX0 + k * (cw + 32)
        y = CY0 + 4
        card(r, th, x, y, cw, 496, radius=18)
        r.rect(x, y, cw, 72, fill=col, r=18)
        r.rect(x, y + 40, cw, 32, fill=col)
        put_text(r, x + 28, y + 6, cw - 56, 60, side["title"], 28, contrast(col), bold=True, valign="m", min_size=15)
        dot_list(r, th, x + 30, y + 94, cw - 60, 384, side["items"], th.text, col if k == 0 else th.accent, max_size=23)


def L_stats(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    stats = s["stats"]
    n = len(stats)
    gap = 24
    cw = (CW - (n - 1) * gap) / n
    y, ch = CY0 + 24, 440
    for i, st in enumerate(stats):
        x = CX0 + i * (cw + gap)
        card(r, th, x, y, cw, ch, radius=20)
        r.rect(x + 24, y + 28, 64, 8, fill=th.pal[i % len(th.pal)], r=4)
        put_text(r, x + 20, y + 54, cw - 40, 130, st["value"], 72, th.heading, bold=True, valign="m", min_size=26)
        put_text(r, x + 24, y + 196, cw - 48, 70, st["label"], 26, th.heading, bold=True, min_size=14)
        put_text(r, x + 24, y + 274, cw - 48, ch - 292, st["text"], 20, th.muted, min_size=12)


def L_chart(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    ch_ = s["chart"]
    wide = 700 if s.get("bullets") else CW
    card(r, th, CX0, CY0 + 4, wide, 496, radius=18)
    draw_chart(r, th, ch_["kind"], ch_["labels"], ch_["values"], ch_.get("unit", ""), CX0 + 14, CY0 + 14, wide - 28, 476)
    if s.get("bullets"):
        card_list(r, th, CX0 + 700 + 28, CY0 + 4, CW - 728, 496, s["bullets"], max_size=22, min_size=12)


def L_timeline(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    ev = s["events"]
    n = len(ev)
    x0, x1 = 120.0, 1160.0
    spacing = (x1 - x0) / (n - 1) if n > 1 else 0
    ly = 414
    r.rect(x0 - 40, ly - 3, (x1 - x0) + 80, 6, fill=th.border, r=3)
    cwid = min(270.0, max(150.0, spacing * 1.9 - 16)) if n > 1 else 400.0
    chh = 196
    for i, e in enumerate(ev):
        cx = x0 + i * spacing if n > 1 else 640.0
        up = i % 2 == 0
        x = min(max(CX0, cx - cwid / 2), CX0 + CW - cwid)
        y = ly - 34 - chh if up else ly + 34
        r.line(cx, ly, cx, (y + chh) if up else y, th.accent, 3)
        card(r, th, x, y, cwid, chh, radius=14)
        r.rect(x, y + 14, 6, chh - 28, fill=th.pal[i % len(th.pal)], r=3)
        put_text(r, x + 20, y + 10, cwid - 34, 50, e["label"], 22, th.heading, bold=True, valign="m", min_size=12)
        put_text(r, x + 20, y + 64, cwid - 34, chh - 74, e["text"], 17, th.text, min_size=11)
        r.oval(cx - 13, ly - 13, 26, 26, fill=th.accent, stroke=th.bg, sw=4)


def L_formula(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.bg)
    header(r, th, deck, s["title"], idx, total)
    fs = s["formulas"]
    n = len(fs)
    cols = 1 if n <= 2 else 2
    rows = math.ceil(n / cols)
    gap = 24
    cw = (CW - (cols - 1) * gap) / cols
    ch = (CH - 8 - (rows - 1) * gap) / rows
    ch = min(ch, 250.0)
    tot_h = rows * ch + (rows - 1) * gap
    y0 = CY0 + (CH - tot_h) / 2
    for i, f in enumerate(fs):
        col, row = i % cols, i // cols
        x, y = CX0 + col * (cw + gap), y0 + row * (ch + gap)
        card(r, th, x, y, cw, ch, radius=18)
        r.rect(x, y + 16, 8, ch - 32, fill=th.pal[i % len(th.pal)], r=4)
        put_text(r, x + 34, y + 14, cw - 60, 30, f["name"], 20, th.muted, min_size=11)
        put_text(r, x + 34, y + 48, cw - 60, ch * 0.46, f["expr"], 44, th.heading, bold=True, valign="m", min_size=18)
        if f["note"]:
            put_text(r, x + 34, y + 54 + ch * 0.46, cw - 60, ch - 66 - ch * 0.46, f["note"], 19, th.text, min_size=11)


def L_conclusion(r, th, deck, s, store, idx, total, lang):
    r.new_slide(th.primary)
    put_text(r, 70, 50, 740, 90, s["title"], 46, th.on_primary, bold=True, valign="m", min_size=24)
    r.rect(70, 146, 90, 6, fill=th.accent, r=3)
    pts = s["points"]
    n = len(pts)
    gap = 14
    rh = min(104.0, (470 - (n - 1) * gap) / n)
    row_fill = mix(th.primary, "#FFFFFF", 0.10)
    size = 26
    for p in pts:
        sz, _l, _s = fit(p, 600, rh - 14, 26, 14)
        size = min(size, sz)
    for i, p in enumerate(pts):
        y = 180 + i * (rh + gap)
        r.rect(70, y, 740, rh, fill=row_fill, r=14)
        number_badge(r, th, 90, y + (rh - 46) / 2, 46, i + 1)
        put_text(r, 158, y + 7, 634, rh - 14, p, size, th.on_primary, valign="m", min_size=size)
    r.rect(858, 150, 352, 440, fill=th.accent, r=26)
    put_text(r, 884, 180, 300, 380, THANKS.get(lang, THANKS["en"]), 46, th.on_accent, bold=True, align="c", valign="m", min_size=22)
    if deck.get("credits"):
        put_text(r, 70, 672, 1140, 24, CREDIT.get(lang, "") + deck["credits"], 13, th.on_primary_soft, min_size=9)


LAYOUTS = {
    "title": L_title, "agenda": L_agenda, "bullets": L_bullets, "highlight": L_highlight, "process": L_process,
    "cycle": L_cycle, "compare": L_compare, "stats": L_stats, "chart": L_chart, "timeline": L_timeline,
    "formula": L_formula, "conclusion": L_conclusion,
}


def layout_deck(r, th: Theme, deck: dict, store: ImageStore, lang: str) -> None:
    total = len(deck["slides"])
    for i, s in enumerate(deck["slides"]):
        s["_i"] = i
        LAYOUTS[s["type"]](r, th, deck, s, store, i + 1, total, lang)
        r.notes(s.get("notes", ""))


def render_deck(deck: dict, workdir: str, lang: str, base_name: str) -> tuple:
    """PPTX va PDF fayllarini yaratadi. (pptx_path, pdf_path) qaytaradi."""
    th = Theme(deck["theme"])
    store = ImageStore(workdir, th)
    pptx_path = str(Path(workdir) / f"{base_name}.pptx")
    pdf_path = str(Path(workdir) / f"{base_name}.pdf")
    rp = PptxRenderer()
    layout_deck(rp, th, deck, store, lang)
    rp.finish(pptx_path)
    rf = PdfRenderer(pdf_path)
    layout_deck(rf, th, deck, store, lang)
    rf.finish()
    return pptx_path, pdf_path


# =============================================================================
# GEMINI
# =============================================================================
class GeminiError(Exception):
    pass


class BlockedError(Exception):
    pass


LANG_NAME = {
    "uz": "Uzbek (Latin script; write o' and g' with a plain apostrophe, never with special modifier letters)",
    "ru": "Russian",
    "en": "English",
}

SYSTEM_PROMPT = """You are the engine of "AI SlideBot PRO", a UNIVERSAL presentation generator.
You turn ANY safe topic (school subjects, university courses, professions, manufacturing, daily life, technology, business, art, culture, nature, hobbies, ...) into a well-structured, visually varied slide deck. You are NOT limited to academic subjects.

Return ONLY one JSON object (no markdown fences, no commentary) shaped like:
{
 "status": "ok" | "clarify" | "refuse",
 "message": "only for clarify/refuse: ONE short question or a short polite reason, in the target language",
 "field": "manufacturing|daily_life|technology|business|art|nature|science|math|education|history|geography|health|food|sport|other",
 "theme": "industry|nature|tech|business|creative|science|education|food|health|history|earth",
 "title": "deck title", "subtitle": "short subtitle",
 "slides": [ ...slide objects... ]
}

THEME choice: manufacturing/processes -> industry; nature/animals/weather -> nature; IT/AI/gadgets -> tech; business/finance/marketing -> business; art/design/music/film -> creative; physics/chemistry/biology/math -> science; school/education -> education; cooking/food -> food; medicine/sport -> health; history/culture -> history; geography/travel -> earth.

SLIDE TYPES (field "type"):
- title:      {type,title,subtitle,image_query}
- agenda:     {type,title,items:[3-8 short strings]}
- bullets:    {type,title,bullets:[3-5 strings],image_query?,notes}
- highlight:  {type,title,statement:"one strong sentence",bullets:[0-3 strings],image_query?,notes}
- process:    {type,title,steps:[{title,text} x 3-6],notes}          (ordered stages / how it is made / how to do it)
- cycle:      {type,title,center:"short label",steps:[{title,text} x 3-6],notes}   (cyclic processes)
- compare:    {type,title,left:{title,items:[2-5]},right:{title,items:[2-5]},notes}  (pros/cons, A vs B, before/after)
- stats:      {type,title,stats:[{value:"short, e.g. 1500°C or 70%",label,text} x 2-4],notes}
- chart:      {type,title,chart:{kind:"bar"|"pie"|"line",labels:[...2-8],values:[numbers],unit:""},bullets:[0-3 insights],notes}
- timeline:   {type,title,events:[{label,text} x 3-7],notes}
- formula:    {type,title,formulas:[{name,expr,note} x 1-4],notes}   (ONLY for math/physics/chemistry/engineering topics where real formulas matter)
- conclusion: {type,title,points:[3-5 strings],notes}

STRUCTURE RULES:
- Exactly N slides (N is given). Slide 1 is "title", the last slide is "conclusion".
- First detect the field of the topic, then choose structure, sections, visuals and theme that fit it.
- For product / "how is X made" topics a good (not strict) order: introduction, what it is, raw materials, production stages (process), how it works (cycle or bullets), technologies/equipment, quality control, pros & cons (compare), applications, conclusion. Adapt freely to the topic.
- For business use stats/chart/compare/timeline; for nature use cycle/timeline/stats; for history use timeline; for science/math use definitions, formulas (only if relevant) and examples.
- NEVER add formulas, equations or scientific jargon to topics that do not need them.
- Vary layouts: when N>=8 use at least 3 different non-bullet types; never more than 2 "bullets" slides in a row.
- Keep text SHORT: titles <= 7 words; bullets <= 16 words; step title <= 4 words, step text <= 18 words; stat value <= 8 characters.
- "notes": 1-2 sentence speaker notes for content slides.
- All visible text must be in the target language. "image_query" is ALWAYS English: a concrete, photographable scene or object (2-5 words), no brands, no person names, e.g. "glass furnace factory", "coffee beans roasting". Give image_query to the title slide and to about half of the bullets/highlight slides; omit it elsewhere.

ACCURACY: be factual. Use numbers only if they are well known; otherwise say "approximately" or stay qualitative. Never invent precise statistics. Chart data must be real or clearly typical/approximate (mention "approx." in unit or insights).

SAFETY: use status "refuse" only for harmful topics (weapons/explosives manufacturing, illegal drug synthesis, malware, self-harm instructions, sexual content, hate, harassment of real people, fraud). Ordinary manufacturing or technology topics (glass, paper, bread, cars, phones, internet, AI...) are SAFE. NEVER refuse because a topic is "not a school subject".
If the topic is truly too vague to build a deck (e.g. one unclear word), use status "clarify" with ONE short question; otherwise always produce the deck (status "ok").
"""


def parse_json_text(text: str) -> Optional[dict]:
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"\s*```$", "", t)
    for cand in (t, t[t.find("{"): t.rfind("}") + 1] if "{" in t and "}" in t else ""):
        if not cand:
            continue
        try:
            obj = json.loads(cand)
            if isinstance(obj, dict):
                return obj
        except Exception:  # noqa: BLE001
            continue
    return None


async def gemini_json(system: str, user: str) -> dict:
    global _key_counter
    if not GEMINI_KEYS:
        raise GeminiError("GEMINI_API_KEY sozlanmagan")
    if HTTP is None:
        raise GeminiError("HTTP session yo'q")
    body = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "user", "parts": [{"text": user}]}],
        "generationConfig": {"temperature": 0.7, "responseMimeType": "application/json", "maxOutputTokens": 16384},
    }
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    last: Any = None
    for attempt in range(6):
        key = GEMINI_KEYS[_key_counter % len(GEMINI_KEYS)]
        _key_counter += 1
        try:
            async with HTTP.post(
                url,
                json=body,
                headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                timeout=aiohttp.ClientTimeout(total=170),
            ) as resp:
                status = resp.status
                raw = await resp.text()
        except (aiohttp.ClientError, asyncio.TimeoutError) as e:
            last = e
            await asyncio.sleep(1.5 * (attempt + 1))
            continue
        if status in (429, 500, 502, 503, 504):
            last = f"HTTP {status}"
            await asyncio.sleep(2.0 * (attempt + 1))
            continue
        if status != 200:
            raise GeminiError(f"HTTP {status}: {raw[:300]}")
        try:
            data = json.loads(raw)
        except Exception as e:  # noqa: BLE001
            last = e
            continue
        fb = data.get("promptFeedback") or {}
        if fb.get("blockReason"):
            raise BlockedError(str(fb.get("blockReason")))
        cands = data.get("candidates") or []
        if not cands:
            raise BlockedError("no candidates")
        cand = cands[0]
        parts = (cand.get("content") or {}).get("parts") or []
        text = "".join(p.get("text", "") for p in parts if isinstance(p, dict) and not p.get("thought"))
        if not text.strip() and cand.get("finishReason") in ("SAFETY", "PROHIBITED_CONTENT", "BLOCKLIST", "SPII", "RECITATION"):
            raise BlockedError(str(cand.get("finishReason")))
        obj = parse_json_text(text)
        if obj is None:
            last = "JSON parse xatosi"
            continue
        return obj
    raise GeminiError(str(last))


# ---------- normalizatsiya ----------
def _t(v: Any, n: int = 160) -> str:
    if v is None or isinstance(v, (list, dict)):
        return ""
    v = re.sub(r"\s+", " ", str(v)).strip()
    return v if len(v) <= n else v[: n - 1].rstrip() + "…"


def _items(v: Any, nmax: int, each: int = 170) -> list:
    if isinstance(v, str):
        v = [x for x in re.split(r"\n+", v) if x.strip()]
    out = []
    if isinstance(v, list):
        for x in v:
            if isinstance(x, dict):
                x = x.get("text") or x.get("title") or x.get("label") or " ".join(str(z) for z in x.values())
            t = _t(re.sub(r"^[\-\*•\d\.\)\s]+", "", str(x)) if isinstance(x, str) else x, each)
            if t:
                out.append(t)
    return out[:nmax]


def _steps(v: Any, nmax: int) -> list:
    out = []
    if isinstance(v, list):
        for x in v:
            if isinstance(x, dict):
                t = _t(x.get("title") or x.get("label") or x.get("name"), 60)
                tx = _t(x.get("text") or x.get("description") or x.get("desc"), 200)
                if not t and tx:
                    t, tx = _t(tx, 40), ""
            else:
                t, tx = _t(x, 60), ""
            if t:
                out.append({"title": t, "text": tx})
    return out[:nmax]


def _num(v: Any) -> Optional[float]:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    try:
        s = str(v).replace("\u00a0", "").replace(" ", "").replace(",", ".").replace("%", "")
        f = float(s)
        return f if math.isfinite(f) else None
    except Exception:  # noqa: BLE001
        return None


def _fallback_bullets(s: dict) -> list:
    out: list = []
    for k, v in s.items():
        if k in ("type", "title", "notes", "image_query"):
            continue
        if isinstance(v, str):
            out.append(_t(v, 200))
        elif isinstance(v, list):
            out += _items(v, 6, 200)
        elif isinstance(v, dict):
            out += _items(list(v.values()), 3, 200)
    return [x for x in out if x][:6]


def _norm_slide(s: dict) -> Optional[dict]:
    t = str(s.get("type", "bullets")).lower().strip()
    title = _t(s.get("title"), 90)
    base = {"title": title, "notes": _t(s.get("notes"), 400), "image_query": _t(s.get("image_query"), 80)}
    if t == "title":
        return {**base, "type": "title", "subtitle": _t(s.get("subtitle"), 150)}
    if t == "agenda":
        it = _items(s.get("items") or s.get("bullets"), 8, 70)
        return {**base, "type": "agenda", "items": it} if len(it) >= 2 else None
    if t == "highlight":
        st = _t(s.get("statement") or s.get("text"), 220)
        b = _items(s.get("bullets"), 3, 140)
        if not st and b:
            st = b.pop(0)
        return {**base, "type": "highlight", "statement": st, "bullets": b} if st else None
    if t in ("process", "cycle"):
        st = _steps(s.get("steps"), 6)
        if len(st) < 2:
            b = _fallback_bullets(s)
            return {**base, "type": "bullets", "bullets": b} if b else None
        if t == "cycle" and len(st) >= 3:
            return {**base, "type": "cycle", "steps": st, "center": _t(s.get("center"), 40)}
        return {**base, "type": "process", "steps": st}
    if t == "compare":
        sides = []
        for k in ("left", "right"):
            d = s.get(k) if isinstance(s.get(k), dict) else {}
            sides.append({"title": _t(d.get("title"), 50), "items": _items(d.get("items") or d.get("bullets"), 5, 130)})
        if sides[0]["items"] and sides[1]["items"]:
            return {**base, "type": "compare", "left": sides[0], "right": sides[1]}
        b = _fallback_bullets(s)
        return {**base, "type": "bullets", "bullets": b} if b else None
    if t == "stats":
        out = []
        for x in s.get("stats") if isinstance(s.get("stats"), list) else []:
            if isinstance(x, dict) and _t(x.get("value"), 14):
                out.append({"value": _t(x.get("value"), 14), "label": _t(x.get("label"), 60), "text": _t(x.get("text"), 130)})
        out = out[:4]
        if len(out) >= 2:
            return {**base, "type": "stats", "stats": out}
        b = _fallback_bullets(s)
        return {**base, "type": "bullets", "bullets": b} if b else None
    if t == "chart":
        c = s.get("chart") if isinstance(s.get("chart"), dict) else {}
        labels = c.get("labels") if isinstance(c.get("labels"), list) else []
        vals = c.get("values") if isinstance(c.get("values"), list) else []
        pairs = [(_t(a, 40), _num(b)) for a, b in zip(labels, vals)]
        pairs = [(a, b) for a, b in pairs if a and b is not None][:8]
        kind = str(c.get("kind", "bar")).lower()
        kind = kind if kind in ("bar", "pie", "line") else "bar"
        if kind == "pie" and (any(b < 0 for _a, b in pairs) or sum(b for _a, b in pairs) <= 0):
            kind = "bar"
        if len(pairs) >= 2:
            return {**base, "type": "chart", "bullets": _items(s.get("bullets"), 3, 140),
                    "chart": {"kind": kind, "labels": [a for a, _b in pairs], "values": [b for _a, b in pairs],
                              "unit": _t(c.get("unit"), 30)}}
        b = _items(s.get("bullets"), 5, 200)
        return {**base, "type": "bullets", "bullets": b} if b else None
    if t == "timeline":
        ev = []
        for x in s.get("events") if isinstance(s.get("events"), list) else []:
            if isinstance(x, dict) and (_t(x.get("label")) or _t(x.get("text"))):
                ev.append({"label": _t(x.get("label") or x.get("title") or x.get("date"), 40), "text": _t(x.get("text"), 150)})
        ev = ev[:7]
        if len(ev) >= 2:
            return {**base, "type": "timeline", "events": ev}
        b = _fallback_bullets(s)
        return {**base, "type": "bullets", "bullets": b} if b else None
    if t == "formula":
        fs = []
        for x in s.get("formulas") if isinstance(s.get("formulas"), list) else []:
            if isinstance(x, dict) and _t(x.get("expr") or x.get("formula"), 70):
                fs.append({"name": _t(x.get("name"), 60), "expr": _t(x.get("expr") or x.get("formula"), 70), "note": _t(x.get("note"), 140)})
        fs = fs[:4]
        if fs:
            return {**base, "type": "formula", "formulas": fs}
        b = _fallback_bullets(s)
        return {**base, "type": "bullets", "bullets": b} if b else None
    if t == "conclusion":
        p = _items(s.get("points") or s.get("bullets") or s.get("items"), 5, 170)
        return {**base, "type": "conclusion", "points": p} if p else None
    b = _items(s.get("bullets") or s.get("items") or s.get("points"), 6, 200) or _fallback_bullets(s)
    return {**base, "type": "bullets", "bullets": b[:6]} if b else None


def normalize_deck(obj: dict, count: int) -> dict:
    title = _t(obj.get("title"), 90) or "Presentation"
    subtitle = _t(obj.get("subtitle"), 150)
    theme = str(obj.get("theme", "")).lower().strip()
    if theme not in THEMES:
        theme = FIELD_THEME.get(str(obj.get("field", "")).lower().strip(), "education")
    slides = []
    for s in obj.get("slides") or []:
        if isinstance(s, dict):
            ns = _norm_slide(s)
            if ns:
                slides.append(ns)
    if not slides:
        raise GeminiError("bo'sh slaydlar")
    for s in slides:
        if not s.get("title") and s["type"] not in ("title",):
            s["title"] = title
    if slides[0]["type"] != "title":
        slides.insert(0, {"type": "title", "title": title, "subtitle": subtitle, "image_query": "", "notes": ""})
    if slides[-1]["type"] != "conclusion":
        pts = [x["title"] for x in slides[1:] if x.get("title")][:5]
        slides.append({"type": "conclusion", "title": "✓", "points": pts or [title], "notes": "", "image_query": ""})
    if len(slides) > count + 2:
        slides = slides[: count + 1] + [slides[-1]]
    slides[0]["title"] = slides[0].get("title") or title
    slides[0]["subtitle"] = slides[0].get("subtitle") or subtitle
    return {"title": title, "subtitle": subtitle, "theme": theme, "slides": slides, "credits": ""}


async def generate_plan(topic: str, count: int, lang: str, force: bool) -> tuple:
    """('ok', deck) | ('clarify', msg) | ('refuse', msg)"""
    user = (
        f"TOPIC: {topic}\nN (exact slide count): {count}\nTARGET LANGUAGE: {LANG_NAME[lang]}\n"
        + ("Do NOT ask for clarification: choose the most reasonable interpretation and produce the deck.\n" if force else "")
        + "Return the JSON object now."
    )
    for _try in range(2):
        obj = await gemini_json(SYSTEM_PROMPT, user)
        status = str(obj.get("status", "ok")).lower()
        msg = _t(obj.get("message"), 300)
        if status == "clarify" and not force and msg:
            return "clarify", msg
        if status == "refuse":
            return "refuse", msg
        try:
            return "ok", normalize_deck(obj, count)
        except GeminiError:
            continue
    raise GeminiError("Gemini yaroqli taqdimot qaytarmadi")


async def attach_images(deck: dict) -> None:
    cache: dict = {}
    sem = asyncio.Semaphore(4)

    async def one(s: dict) -> None:
        q = s.get("image_query")
        if not q or s["type"] not in ("title", "bullets", "highlight"):
            return
        async with sem:
            raw = await fetch_image_bytes(q, cache)
        if raw:
            s["image_raw"] = raw

    await asyncio.gather(*(one(s) for s in deck["slides"]), return_exceptions=True)
    got = [s for s in deck["slides"] if s.get("image_raw")]
    srcs = []
    if got:
        if PEXELS_KEY:
            srcs.append("Pexels")
        srcs.append("Wikimedia Commons")
    deck["credits"] = ", ".join(srcs)


def safe_name(s: str) -> str:
    n = re.sub(r"[^\w\-]+", "_", s, flags=re.UNICODE).strip("_")[:40]
    return n or "presentation"


# =============================================================================
# DATABASE (SQLite)
# =============================================================================
DB: Optional[sqlite3.Connection] = None


def db_init() -> None:
    global DB
    DB = sqlite3.connect(DB_PATH, check_same_thread=False)
    DB.executescript(
        """
        CREATE TABLE IF NOT EXISTS users(uid INTEGER PRIMARY KEY, lang TEXT, name TEXT, created INTEGER);
        CREATE TABLE IF NOT EXISTS usage(uid INTEGER, day TEXT, n INTEGER, PRIMARY KEY(uid, day));
        CREATE TABLE IF NOT EXISTS history(id INTEGER PRIMARY KEY AUTOINCREMENT, uid INTEGER, topic TEXT,
                                           slides INTEGER, lang TEXT, ts INTEGER);
        """
    )
    DB.commit()


def today() -> str:
    return datetime.now(timezone(timedelta(hours=5))).strftime("%Y-%m-%d")


def get_lang(uid: int) -> Optional[str]:
    row = DB.execute("SELECT lang FROM users WHERE uid=?", (uid,)).fetchone()
    return row[0] if row and row[0] in LANGS else None


def set_lang(uid: int, lang: str, name: str = "") -> None:
    DB.execute(
        "INSERT INTO users(uid, lang, name, created) VALUES(?,?,?,?) "
        "ON CONFLICT(uid) DO UPDATE SET lang=excluded.lang, name=excluded.name",
        (uid, lang, name[:80], int(time.time())),
    )
    DB.commit()


def used_today(uid: int) -> int:
    row = DB.execute("SELECT n FROM usage WHERE uid=? AND day=?", (uid, today())).fetchone()
    return row[0] if row else 0


def add_usage(uid: int, topic: str, slides: int, lang: str) -> None:
    DB.execute(
        "INSERT INTO usage(uid, day, n) VALUES(?,?,1) ON CONFLICT(uid, day) DO UPDATE SET n=n+1", (uid, today())
    )
    DB.execute("INSERT INTO history(uid, topic, slides, lang, ts) VALUES(?,?,?,?,?)", (uid, topic[:200], slides, lang, int(time.time())))
    DB.commit()


def stats_text() -> str:
    users = DB.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    total = DB.execute("SELECT COUNT(*) FROM history").fetchone()[0]
    tod = DB.execute("SELECT COALESCE(SUM(n),0) FROM usage WHERE day=?", (today(),)).fetchone()[0]
    last = DB.execute("SELECT topic FROM history ORDER BY id DESC LIMIT 5").fetchall()
    return (
        f"📊 <b>Stats</b>\n👥 Users: {users}\n📑 Presentations: {total}\n📅 Today: {tod}\n\n"
        + "\n".join("• " + html.escape(x[0]) for x in last)
    )


# =============================================================================
# TEXTS (uz / ru / en)
# =============================================================================
T = {
    "uz": {
        "choose_lang": "🌐 Tilni tanlang / Выберите язык / Choose your language",
        "lang_set": "✅ Til tanlandi: O'zbekcha",
        "welcome": (
            "👋 <b>AI SlideBot PRO</b>ga xush kelibsiz!\n\n"
            "Men <b>istalgan mavzu</b>da professional taqdimot (<b>PPTX + PDF</b>) yarataman: maktab fanlari, kasb, "
            "ishlab chiqarish, texnologiya, biznes, san'at, tabiat, kundalik hayot va boshqalar.\n\n"
            "✍️ Shunchaki mavzuni yozing, masalan:\n"
            "• <i>Shisha qanday ishlab chiqariladi?</i>\n• <i>Internet qanday ishlaydi?</i>\n• <i>Brend qanday yaratiladi?</i>\n"
            "• <i>Vulqonlar qanday paydo bo'ladi?</i>"
        ),
        "help": (
            "ℹ️ <b>Yordam</b>\n\n1. Mavzuni yozing.\n2. Slaydlar soni va tilni tanlang.\n3. «🚀 Yaratish» tugmasini bosing.\n\n"
            "Natija: PPTX (tahrirlash mumkin) va PDF fayl.\n"
            "Buyruqlar: /start, /lang, /new, /cancel, /help"
        ),
        "btn_new": "📊 Yangi taqdimot", "btn_lang": "🌐 Til", "btn_help": "ℹ️ Yordam",
        "ask_topic": "✍️ Taqdimot mavzusini yozing (istalgan mavzu):",
        "topic_short": "Mavzu juda qisqa. Iltimos, aniqroq yozing.",
        "topic_long": "Mavzu juda uzun. Iltimos, 400 belgidan qisqaroq yozing.",
        "topic_label": "Mavzu", "slides_label": "Slaydlar soni", "pres_lang_label": "Taqdimot tili",
        "setup_hint": "Sozlamalarni tanlang va «🚀 Yaratish» tugmasini bosing.",
        "btn_go": "🚀 Yaratish", "btn_cancel": "✖️ Bekor qilish",
        "cancelled": "❌ Bekor qilindi.", "expired": "Bu so'rov eskirgan. Mavzuni qaytadan yozing.",
        "busy": "⏳ Sizning oldingi taqdimotingiz hali tayyorlanmoqda. Iltimos, kuting.",
        "limit": "⚠️ Bugungi limit tugadi (kuniga {n} ta taqdimot). Ertaga qayta urinib ko'ring.",
        "st_queue": "⏳ Navbatda turibsiz, tez orada boshlanadi...",
        "st_analyze": "🧠 Mavzu tahlil qilinmoqda va taqdimot rejasi tuzilmoqda...",
        "st_images": "🖼 Mavzuga mos rasmlar tanlanmoqda...",
        "st_build": "🎨 Dizayn va fayllar (PPTX + PDF) yaratilmoqda...",
        "clarify_hint": "✍️ Javobingizni yozing — shundan keyin davom etaman.",
        "refuse": "🚫 Bu mavzuda taqdimot yarata olmayman{reason}\nBoshqa mavzuni yozib ko'ring.",
        "done": (
            "✅ <b>Taqdimot tayyor!</b>\n\n📊 {title}\n📑 {n} ta slayd · PPTX + PDF\n\n"
            "⚠️ AI xato qilishi mumkin — muhim ma'lumotlarni tekshirib oling. Rasmlar ochiq manbalardan olinadi."
        ),
        "err_blocked": "⚠️ Bu so'rov xavfsizlik filtri tomonidan to'xtatildi. Mavzuni boshqacha yozib ko'ring.",
        "err_gemini": "⚠️ AI xizmati hozir band yoki javob bermadi. Birozdan keyin qayta urinib ko'ring.",
        "err_generic": "⚠️ Kutilmagan xatolik yuz berdi. Iltimos, qayta urinib ko'ring.",
    },
    "ru": {
        "choose_lang": "🌐 Tilni tanlang / Выберите язык / Choose your language",
        "lang_set": "✅ Язык выбран: Русский",
        "welcome": (
            "👋 Добро пожаловать в <b>AI SlideBot PRO</b>!\n\n"
            "Я создаю профессиональные презентации (<b>PPTX + PDF</b>) на <b>любую тему</b>: школьные предметы, профессии, "
            "производство, технологии, бизнес, искусство, природа, повседневная жизнь и многое другое.\n\n"
            "✍️ Просто напишите тему, например:\n"
            "• <i>Как производят стекло?</i>\n• <i>Как работает интернет?</i>\n• <i>Как создать бренд?</i>\n"
            "• <i>Как образуются вулканы?</i>"
        ),
        "help": (
            "ℹ️ <b>Помощь</b>\n\n1. Напишите тему.\n2. Выберите количество слайдов и язык.\n3. Нажмите «🚀 Создать».\n\n"
            "Результат: PPTX (можно редактировать) и PDF.\n"
            "Команды: /start, /lang, /new, /cancel, /help"
        ),
        "btn_new": "📊 Новая презентация", "btn_lang": "🌐 Язык", "btn_help": "ℹ️ Помощь",
        "ask_topic": "✍️ Напишите тему презентации (любую):",
        "topic_short": "Тема слишком короткая. Пожалуйста, уточните.",
        "topic_long": "Тема слишком длинная. Пожалуйста, уложитесь в 400 символов.",
        "topic_label": "Тема", "slides_label": "Количество слайдов", "pres_lang_label": "Язык презентации",
        "setup_hint": "Выберите настройки и нажмите «🚀 Создать».",
        "btn_go": "🚀 Создать", "btn_cancel": "✖️ Отмена",
        "cancelled": "❌ Отменено.", "expired": "Этот запрос устарел. Напишите тему заново.",
        "busy": "⏳ Предыдущая презентация ещё создаётся. Пожалуйста, подождите.",
        "limit": "⚠️ Дневной лимит исчерпан ({n} презентаций в день). Попробуйте завтра.",
        "st_queue": "⏳ Вы в очереди, скоро начнём...",
        "st_analyze": "🧠 Анализирую тему и составляю план презентации...",
        "st_images": "🖼 Подбираю подходящие изображения...",
        "st_build": "🎨 Создаю дизайн и файлы (PPTX + PDF)...",
        "clarify_hint": "✍️ Напишите ответ — после этого я продолжу.",
        "refuse": "🚫 Я не могу создать презентацию на эту тему{reason}\nПопробуйте другую тему.",
        "done": (
            "✅ <b>Презентация готова!</b>\n\n📊 {title}\n📑 Слайдов: {n} · PPTX + PDF\n\n"
            "⚠️ ИИ может ошибаться — проверяйте важные данные. Изображения взяты из открытых источников."
        ),
        "err_blocked": "⚠️ Запрос остановлен фильтром безопасности. Попробуйте сформулировать тему иначе.",
        "err_gemini": "⚠️ Сервис ИИ сейчас занят или не ответил. Попробуйте чуть позже.",
        "err_generic": "⚠️ Произошла непредвиденная ошибка. Пожалуйста, попробуйте ещё раз.",
    },
    "en": {
        "choose_lang": "🌐 Tilni tanlang / Выберите язык / Choose your language",
        "lang_set": "✅ Language set: English",
        "welcome": (
            "👋 Welcome to <b>AI SlideBot PRO</b>!\n\n"
            "I create professional presentations (<b>PPTX + PDF</b>) on <b>any topic</b>: school subjects, professions, "
            "manufacturing, technology, business, art, nature, everyday life and more.\n\n"
            "✍️ Just type a topic, for example:\n"
            "• <i>How is glass made?</i>\n• <i>How does the internet work?</i>\n• <i>How to build a brand?</i>\n"
            "• <i>How do volcanoes form?</i>"
        ),
        "help": (
            "ℹ️ <b>Help</b>\n\n1. Type a topic.\n2. Choose slide count and language.\n3. Tap «🚀 Create».\n\n"
            "Result: an editable PPTX and a PDF.\n"
            "Commands: /start, /lang, /new, /cancel, /help"
        ),
        "btn_new": "📊 New presentation", "btn_lang": "🌐 Language", "btn_help": "ℹ️ Help",
        "ask_topic": "✍️ Type the presentation topic (any topic):",
        "topic_short": "The topic is too short. Please be more specific.",
        "topic_long": "The topic is too long. Please keep it under 400 characters.",
        "topic_label": "Topic", "slides_label": "Slides", "pres_lang_label": "Presentation language",
        "setup_hint": "Choose your settings and tap «🚀 Create».",
        "btn_go": "🚀 Create", "btn_cancel": "✖️ Cancel",
        "cancelled": "❌ Cancelled.", "expired": "This request has expired. Please type the topic again.",
        "busy": "⏳ Your previous presentation is still being created. Please wait.",
        "limit": "⚠️ Daily limit reached ({n} presentations per day). Please try again tomorrow.",
        "st_queue": "⏳ You're in the queue, starting soon...",
        "st_analyze": "🧠 Analyzing the topic and planning the presentation...",
        "st_images": "🖼 Choosing suitable images...",
        "st_build": "🎨 Building the design and files (PPTX + PDF)...",
        "clarify_hint": "✍️ Type your answer and I will continue.",
        "refuse": "🚫 I can't create a presentation on this topic{reason}\nPlease try another topic.",
        "done": (
            "✅ <b>Your presentation is ready!</b>\n\n📊 {title}\n📑 {n} slides · PPTX + PDF\n\n"
            "⚠️ AI can make mistakes — please verify important facts. Images come from open sources."
        ),
        "err_blocked": "⚠️ This request was stopped by a safety filter. Try rephrasing the topic.",
        "err_gemini": "⚠️ The AI service is busy or did not respond. Please try again in a moment.",
        "err_generic": "⚠️ An unexpected error occurred. Please try again.",
    },
}
BTN_NEW = {T[l]["btn_new"] for l in LANGS}
BTN_LANG = {T[l]["btn_lang"] for l in LANGS}
BTN_HELP = {T[l]["btn_help"] for l in LANGS}

# =============================================================================
# TELEGRAM BOT
# =============================================================================
router = Router()
PENDING: dict = {}
ACTIVE: set = set()


def kb_main(lang: str) -> ReplyKeyboardMarkup:
    t = T[lang]
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=t["btn_new"])], [KeyboardButton(text=t["btn_lang"]), KeyboardButton(text=t["btn_help"])]],
        resize_keyboard=True,
    )


def kb_lang() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text=LANG_FLAG[l], callback_data=f"lang:{l}") for l in LANGS]]
    )


def kb_setup(p: dict, t: dict) -> InlineKeyboardMarkup:
    cnt_row = [
        InlineKeyboardButton(text=(f"✅ {c}" if c == p["count"] else str(c)), callback_data=f"cnt:{c}") for c in SLIDE_COUNTS
    ]
    lang_row = [
        InlineKeyboardButton(text=(f"✅ {LANG_FLAG[l]}" if l == p["plang"] else LANG_FLAG[l]), callback_data=f"pl:{l}")
        for l in LANGS
    ]
    return InlineKeyboardMarkup(
        inline_keyboard=[
            cnt_row,
            lang_row,
            [InlineKeyboardButton(text=t["btn_go"], callback_data="go"), InlineKeyboardButton(text=t["btn_cancel"], callback_data="cancel")],
        ]
    )


def setup_text(p: dict, t: dict) -> str:
    return (
        f"📝 <b>{t['topic_label']}:</b> {html.escape(p['topic'])}\n\n"
        f"📑 <b>{t['slides_label']}:</b> {p['count']}\n"
        f"🌐 <b>{t['pres_lang_label']}:</b> {LANG_FLAG[p['plang']]}\n\n"
        f"{t['setup_hint']}"
    )


def cleanup_pending() -> None:
    if len(PENDING) > 300:
        now = time.time()
        for k in [k for k, v in PENDING.items() if now - v.get("ts", 0) > 7200]:
            PENDING.pop(k, None)


async def safe_edit(msg: Message, text: str, markup: Optional[InlineKeyboardMarkup] = None) -> None:
    try:
        await msg.edit_text(text, reply_markup=markup)
    except TelegramBadRequest:
        pass
    except Exception as e:  # noqa: BLE001
        log.info("edit xatosi: %s", e)


@router.message(CommandStart())
async def h_start(m: Message) -> None:
    lang = get_lang(m.from_user.id)
    if not lang:
        await m.answer(T["uz"]["choose_lang"], reply_markup=kb_lang())
        return
    await m.answer(T[lang]["welcome"], reply_markup=kb_main(lang))


@router.message(Command("lang"))
async def h_lang(m: Message) -> None:
    await m.answer(T["uz"]["choose_lang"], reply_markup=kb_lang())


@router.message(Command("help"))
async def h_help(m: Message) -> None:
    lang = get_lang(m.from_user.id) or "uz"
    await m.answer(T[lang]["help"], reply_markup=kb_main(lang))


@router.message(Command("new"))
async def h_new(m: Message) -> None:
    lang = get_lang(m.from_user.id)
    if not lang:
        await m.answer(T["uz"]["choose_lang"], reply_markup=kb_lang())
        return
    PENDING.pop(m.from_user.id, None)
    await m.answer(T[lang]["ask_topic"], reply_markup=kb_main(lang))


@router.message(Command("cancel"))
async def h_cancel(m: Message) -> None:
    lang = get_lang(m.from_user.id) or "uz"
    PENDING.pop(m.from_user.id, None)
    await m.answer(T[lang]["cancelled"], reply_markup=kb_main(lang))


@router.message(Command("stats"))
async def h_stats(m: Message) -> None:
    if m.from_user.id in ADMIN_IDS:
        await m.answer(stats_text())


@router.callback_query(F.data.startswith("lang:"))
async def cb_lang(cb: CallbackQuery) -> None:
    lang = cb.data.split(":", 1)[1]
    if lang not in LANGS:
        await cb.answer()
        return
    set_lang(cb.from_user.id, lang, cb.from_user.full_name or "")
    await cb.answer()
    await safe_edit(cb.message, T[lang]["lang_set"])
    await cb.message.answer(T[lang]["welcome"], reply_markup=kb_main(lang))


async def _pending_for(cb: CallbackQuery) -> Optional[dict]:
    lang = get_lang(cb.from_user.id) or "uz"
    p = PENDING.get(cb.from_user.id)
    if not p or cb.message is None or cb.message.message_id != p.get("msg_id"):
        await cb.answer(T[lang]["expired"], show_alert=True)
        return None
    return p


@router.callback_query(F.data.startswith("cnt:"))
async def cb_count(cb: CallbackQuery) -> None:
    p = await _pending_for(cb)
    if not p:
        return
    try:
        c = int(cb.data.split(":", 1)[1])
    except ValueError:
        await cb.answer()
        return
    if c in SLIDE_COUNTS:
        p["count"] = c
    t = T[get_lang(cb.from_user.id) or "uz"]
    await cb.answer()
    await safe_edit(cb.message, setup_text(p, t), kb_setup(p, t))


@router.callback_query(F.data.startswith("pl:"))
async def cb_plang(cb: CallbackQuery) -> None:
    p = await _pending_for(cb)
    if not p:
        return
    pl = cb.data.split(":", 1)[1]
    if pl in LANGS:
        p["plang"] = pl
    t = T[get_lang(cb.from_user.id) or "uz"]
    await cb.answer()
    await safe_edit(cb.message, setup_text(p, t), kb_setup(p, t))


@router.callback_query(F.data == "cancel")
async def cb_cancel(cb: CallbackQuery) -> None:
    lang = get_lang(cb.from_user.id) or "uz"
    PENDING.pop(cb.from_user.id, None)
    await cb.answer()
    await safe_edit(cb.message, T[lang]["cancelled"])


@router.callback_query(F.data == "go")
async def cb_go(cb: CallbackQuery, bot: Bot) -> None:
    p = await _pending_for(cb)
    if not p:
        return
    await cb.answer()
    try:
        await cb.message.edit_reply_markup(reply_markup=None)
    except Exception:  # noqa: BLE001
        pass
    await start_generation(bot, cb.message.chat.id, cb.from_user.id, p)


@router.message(F.text)
async def h_text(m: Message, bot: Bot) -> None:
    uid = m.from_user.id
    text = (m.text or "").strip()
    lang = get_lang(uid)
    if not lang:
        await m.answer(T["uz"]["choose_lang"], reply_markup=kb_lang())
        return
    t = T[lang]
    if text in BTN_NEW:
        PENDING.pop(uid, None)
        await m.answer(t["ask_topic"])
        return
    if text in BTN_LANG:
        await m.answer(t["choose_lang"], reply_markup=kb_lang())
        return
    if text in BTN_HELP:
        await m.answer(t["help"])
        return
    if text.startswith("/"):
        return
    p = PENDING.get(uid)
    if p and p.get("await") and time.time() - p.get("ts", 0) < 3600:
        p["topic"] = f"{p['topic']}. {text}"[:600]
        p["await"] = False
        p["ts"] = time.time()
        await start_generation(bot, m.chat.id, uid, p)
        return
    if len(text) < 3:
        await m.answer(t["topic_short"])
        return
    if len(text) > 400:
        await m.answer(t["topic_long"])
        return
    cleanup_pending()
    p = {"topic": text, "count": DEFAULT_COUNT, "plang": lang, "clarifies": 0, "await": False, "ts": time.time()}
    sent = await m.answer(setup_text(p, t), reply_markup=kb_setup(p, t))
    p["msg_id"] = sent.message_id
    PENDING[uid] = p


async def start_generation(bot: Bot, chat_id: int, uid: int, p: dict) -> None:
    lang = get_lang(uid) or "uz"
    t = T[lang]
    if uid in ACTIVE:
        await bot.send_message(chat_id, t["busy"])
        return
    if uid not in ADMIN_IDS and used_today(uid) >= DAILY_LIMIT:
        await bot.send_message(chat_id, t["limit"].format(n=DAILY_LIMIT))
        return
    ACTIVE.add(uid)
    workdir = tempfile.mkdtemp(prefix="slidebot_")
    status: Optional[Message] = None
    try:
        status = await bot.send_message(chat_id, t["st_queue"] if GEN_SEM.locked() else t["st_analyze"])
        async with GEN_SEM:
            await safe_edit(status, t["st_analyze"])
            kind, payload = await generate_plan(p["topic"], p["count"], p["plang"], force=p["clarifies"] >= 2)
            if kind == "clarify":
                p["clarifies"] += 1
                p["await"] = True
                p["ts"] = time.time()
                await safe_edit(status, f"❓ {html.escape(payload)}\n\n{t['clarify_hint']}")
                return
            if kind == "refuse":
                reason = f": {html.escape(payload)}" if payload else "."
                await safe_edit(status, t["refuse"].format(reason=reason))
                PENDING.pop(uid, None)
                return
            deck = payload
            await safe_edit(status, t["st_images"])
            await attach_images(deck)
            await safe_edit(status, t["st_build"])
            base = safe_name(deck["title"] or p["topic"])
            pptx_path, pdf_path = await asyncio.to_thread(render_deck, deck, workdir, p["plang"], base)
        await bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
        await bot.send_document(chat_id, FSInputFile(pptx_path, filename=f"{base}.pptx"))
        await bot.send_document(
            chat_id,
            FSInputFile(pdf_path, filename=f"{base}.pdf"),
            caption=t["done"].format(title=html.escape(deck["title"]), n=len(deck["slides"])),
        )
        add_usage(uid, p["topic"], len(deck["slides"]), p["plang"])
        PENDING.pop(uid, None)
        try:
            await status.delete()
        except Exception:  # noqa: BLE001
            pass
    except BlockedError:
        await _fail(bot, chat_id, status, t["err_blocked"])
    except GeminiError as e:
        log.error("Gemini xatosi: %s", e)
        await _fail(bot, chat_id, status, t["err_gemini"])
    except Exception:  # noqa: BLE001
        log.exception("Taqdimot yaratishda xato")
        await _fail(bot, chat_id, status, t["err_generic"])
    finally:
        ACTIVE.discard(uid)
        shutil.rmtree(workdir, ignore_errors=True)


async def _fail(bot: Bot, chat_id: int, status: Optional[Message], text: str) -> None:
    if status is not None:
        try:
            await status.edit_text(text)
            return
        except Exception:  # noqa: BLE001
            pass
    try:
        await bot.send_message(chat_id, text)
    except Exception:  # noqa: BLE001
        pass


# =============================================================================
# MAIN
# =============================================================================
async def start_web() -> None:
    from aiohttp import web

    async def ok(_request):
        return web.Response(text="AI SlideBot PRO is running")

    app = web.Application()
    app.router.add_get("/", ok)
    app.router.add_get("/health", ok)
    runner = web.AppRunner(app)
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", PORT).start()
    log.info("Veb-server %s portda ishga tushdi", PORT)


async def main() -> None:
    global HTTP, GEN_SEM
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN ENV o'zgaruvchisi kiritilmagan!")
    if not GEMINI_KEYS:
        log.warning("GEMINI_API_KEY kiritilmagan - taqdimotlar yaratilmaydi!")
    db_init()
    await asyncio.to_thread(setup_fonts)
    GEN_SEM = asyncio.Semaphore(MAX_PARALLEL)
    HTTP = aiohttp.ClientSession(headers={"User-Agent": USER_AGENT})
    if PORT:
        await start_web()
    bot = Bot(
        token=BOT_TOKEN,
        session=AiohttpSession(timeout=240),
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    dp = Dispatcher()
    dp.include_router(router)
    try:
        await bot.delete_webhook(drop_pending_updates=True)
        me = await bot.get_me()
        log.info("Bot ishga tushdi: @%s | model: %s", me.username, GEMINI_MODEL)
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        await HTTP.close()
        await bot.session.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
