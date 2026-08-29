from tkinter import *
from tkinter import messagebox, filedialog, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta, timezone
import sqlite3
import math
import random
try:
    import easyocr
except ImportError:
    easyocr = None
import csv
import re
import secrets
import os
import shutil

# ---------------- HIGH-DPI / SCREEN SCALING FIX ---------------- #
# Without this, Windows treats the app as DPI-unaware: it still reports
# "logical" (96 DPI) pixel sizes to Tk, so on a scaled display (125%/150%/
# 200%, common on modern laptops) every widget and font renders far smaller
# than it should relative to the physical screen — the exact "everything
# looks zoomed out with empty space around it" symptom. Declaring
# per-monitor DPI awareness here (before any Tk window exists) tells
# Windows to hand Tk the real pixel dimensions instead of scaling the
# whole window as a blurry bitmap after the fact. No-op on non-Windows.
if os.name == "nt":
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor v2
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()       # older Windows fallback
    except Exception:
        pass

# PDF/Excel export are optional — the app must still run fine without them.
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ---------------- GRAPHICS & UI SCALING SETUP ---------------- #
plt.rcParams['figure.dpi'] = 100
plt.rcParams['font.sans-serif'] = 'Segoe UI'
plt.rcParams['font.family'] = 'sans-serif'

# ---------------- THEME (theme.py) ---------------- #
# The palette *data* (light/dark dicts, fixed accent + auth colors) lives
# in theme.py. The mutable "current theme" names below stay here in
# app.py because ~100+ page-rendering functions in this file read
# BG_CANVAS / TEXT_MAIN / etc. as bare module-level globals — see
# theme.py's docstring for why apply_theme() is a thin wrapper rather
# than living entirely in theme.py.
from theme import (
    LIGHT_THEME, DARK_THEME, apply_theme as _apply_theme_palette,
    COLOR_INCOME, COLOR_EXPENSE, COLOR_SAVINGS, COLOR_ACCENT, COLOR_INFO,
    COLOR_DANGER, COLOR_PURPLE,
    AUTH_BG, AUTH_PANEL, AUTH_CARD, AUTH_INPUT, AUTH_BORDER, AUTH_TEXT,
    AUTH_MUTED, AUTH_EMERALD, AUTH_CYAN, AUTH_PURPLE,
)

BG_CANVAS = LIGHT_THEME["BG_CANVAS"]
BG_SIDEBAR = LIGHT_THEME["BG_SIDEBAR"]
CARD_BG = LIGHT_THEME["CARD_BG"]
TEXT_MAIN = LIGHT_THEME["TEXT_MAIN"]
TEXT_MUTED = LIGHT_THEME["TEXT_MUTED"]
BORDER_COLOR = LIGHT_THEME["BORDER_COLOR"]


def apply_theme(is_dark):
    """Swaps the module-level color palette used across every screen.
    Since every load_* function reads these names fresh at build time,
    calling this then rebuilding the workspace re-themes the whole app."""
    _apply_theme_palette(is_dark, globals())


# ---------------- DATABASE (database.py) ---------------- #
# All schema/connection/migration logic lives in database.py (no
# Tkinter dependency there by design). `auth_conn`/`auth_cursor` are
# imported directly since they're never reassigned after creation —
# only `.execute()`/`.commit()` are called on them. `conn`/`cursor`
# stay defined here as app.py's own mutable globals: they get rebound
# via `global conn, cursor; conn, cursor = get_user_db(username)` in
# show_auth_gate / build_login_screen / build_signup_screen / logout /
# restore_database, all of which are methods on this module's
# CooperAppMaster class.
from database import auth_conn, auth_cursor, get_user_db, _safe_db_filename

conn = None
cursor = None

# ---------------- AUTH HELPERS ---------------- #
APP_DATA_DIR = os.path.join(os.path.expanduser("~"), ".cooper_vault")
PROFILE_PICTURES_DIR = os.path.join(APP_DATA_DIR, "profile_pictures")
os.makedirs(PROFILE_PICTURES_DIR, exist_ok=True)

# ---------------- AUTHENTICATION (authentication.py) ---------------- #
# Password hashing + registration-field validation are pure functions
# with no UI/DB dependency — see authentication.py's docstring for why
# the do_login/do_signup/logout *flow* (which builds Toplevel widgets
# and touches the database) stays here rather than moving too.
from authentication import (
    generate_salt, hash_with_salt,
    validate_username_format, password_strength_checks,
    password_strength_label, validate_password_strength,
    validate_email_format,
)


class CooperAppMaster(Tk):
    def __init__(self):
        super().__init__()
        self.title("Cooper Finance Tracker")


        # Tell Tk's own layout engine the real physical DPI of this screen.
        # Without this, even a DPI-aware process still renders point-sized
        # fonts using Tk's 96dpi assumption, which is what makes everything
        # look shrunken/zoomed-out with dead space around it on a modern
        # (125%-200% scaled) display. 72pt = 1 logical inch is the baseline
        # Tk expects; winfo_fpixels('1i') reports how many *real* pixels
        # actually make up an inch on this screen, so this ratio corrects
        # every font and every "pixel" measurement app-wide in one call.
        try:
            true_dpi = self.winfo_fpixels("1i")
            if true_dpi > 0:
                self.tk.call("tk", "scaling", true_dpi / 72.0)
        except Exception:
            pass

        # Size the window relative to the actual screen instead of a fixed
        # 1400x850 (which is huge on a small laptop and tiny/lost-in-space
        # on a large or high-res monitor), then center it.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        target_w = max(1200, min(int(screen_w * 0.86), 1600))
        target_h = max(720, min(int(screen_h * 0.86), 960))
        pos_x = max(0, (screen_w - target_w) // 2)
        pos_y = max(0, (screen_h - target_h) // 3)
        self.geometry(f"{target_w}x{target_h}+{pos_x}+{pos_y}")
        self.minsize(1180, 700)

        self.configure(bg=BG_CANVAS)
        
        # UI Component Styling Adjustments
        self.ttk_style = ttk.Style()
        self.ttk_style.theme_use("clam")
        self.refresh_ttk_style()

        self.sidebar_buttons = {}
        self.current_user_id = None
        self.current_username = None
        self.current_display_name = None
        self.current_profile_picture = None

        self.show_auth_gate()

    # ==================================================================
    # AUTHENTICATION
    # ==================================================================
    def show_auth_gate(self):
        """Decides whether to show Sign Up (no account yet), auto-login
        (Remember Me was checked last time), or the Login screen."""
        global conn, cursor
        auth_cursor.execute("SELECT COUNT(*) FROM users")
        if auth_cursor.fetchone()[0] == 0:
            self.build_signup_screen()
            return

        auth_cursor.execute(
            "SELECT id, username, display_name, profile_picture FROM users WHERE remember_me=1 LIMIT 1"
        )
        remembered = auth_cursor.fetchone()
        if remembered:
            self.current_user_id, self.current_username, self.current_display_name, self.current_profile_picture = remembered
            auth_cursor.execute("SELECT dark_mode FROM users WHERE id=?", (self.current_user_id,))
            apply_theme(bool(auth_cursor.fetchone()[0]))
            # Remember-Me auto-login: open THIS user's private database
            # before entering the workspace, exactly like a normal login.
            conn, cursor = get_user_db(self.current_username)
            self.build_workspace()
            return

        self.build_login_screen()

    def _clear_root(self):
        for widget in self.winfo_children():
            widget.destroy()

    def _rounded_rect(self, canvas, x1, y1, x2, y2, radius=14, **kwargs):
        """
        Draws a rounded rectangle on a Canvas using a smoothed polygon —
        Tkinter has no native rounded-rect primitive, so this is the
        standard workaround. Used for pill-style chips and soft panels
        that would otherwise be hard, sharp-cornered rectangles.
        """
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return canvas.create_polygon(points, smooth=True, **kwargs)

    def _paint_vertical_gradient(self, canvas, top_rgb, bottom_rgb, tag="gradient", steps=80):
        """
        Fills `canvas` with a smooth top-to-bottom gradient by drawing thin
        horizontal color bands — Tkinter has no native gradient fill, so
        this is the standard way to fake one. Call again on <Configure> to
        keep it filling the canvas as it resizes.
        """
        canvas.delete(tag)
        w = canvas.winfo_width() or 1
        h = canvas.winfo_height() or 1
        for i in range(steps):
            t = i / steps
            r = int(top_rgb[0] + (bottom_rgb[0] - top_rgb[0]) * t)
            g = int(top_rgb[1] + (bottom_rgb[1] - top_rgb[1]) * t)
            b = int(top_rgb[2] + (bottom_rgb[2] - top_rgb[2]) * t)
            color = f"#{r:02x}{g:02x}{b:02x}"
            y0 = int(h * i / steps)
            y1 = int(h * (i + 1) / steps) + 1
            canvas.create_rectangle(0, y0, w, y1, fill=color, outline="", tags=tag)
        canvas.tag_lower(tag)

    def _shadow_card(self, parent, padx=44, pady=40, accent=None, card_bg=None, border_color=None, panel_bg=None):
        """
        A soft drop-shadow illusion for a card: several progressively
        lighter offset layers stack behind the card so the shadow reads
        as a soft blur rather than one hard-edged sliver, and an optional
        thin accent bar sits along the top edge for a bit of brand color.
        Real box-shadows don't exist in Tkinter, but layering several
        1-2px-stepped tints gets convincingly close at normal viewing
        size. Pass `card_bg`/`border_color`/`panel_bg` to theme this for
        a surface other than the app's normal light/dark toggle (e.g.
        the always-dark auth screens). Returns the inner (foreground)
        frame — pack/grid content into it exactly as before.
        """
        card_bg = card_bg or BG_CANVAS
        border_color = border_color or BORDER_COLOR
        panel_bg = panel_bg or BG_CANVAS

        shadow_outer = Frame(parent, bg=panel_bg)

        shadow_layers = []
        for i, factor in enumerate((0.965, 0.94, 0.90, 0.84)):
            layer = Frame(shadow_outer, bg=self._shade_hex_color(border_color, factor))
            layer.place(x=2 + i * 2, y=3 + i * 2)
            shadow_layers.append(layer)

        card = Frame(shadow_outer, bg=card_bg, highlightthickness=1, highlightbackground=border_color, padx=padx, pady=pady)
        card.place(x=0, y=0)

        if accent:
            Frame(shadow_outer, bg=accent, height=4).place(x=0, y=0, relwidth=1)

        def sync_shadow_size(event=None):
            w = card.winfo_reqwidth()
            h = card.winfo_reqheight()
            for i, layer in enumerate(shadow_layers):
                layer.place_configure(width=w, height=h)
            shadow_outer.config(width=w + 10, height=h + 11)

        card.bind("<Configure>", sync_shadow_size)
        return shadow_outer, card

    def _rounded_button(self, parent, text, command, bg=None, fg="white",
                         width=280, height=48, font_size=13, radius=12, icon=None, panel_bg=None):
        """
        A premium, fully rounded call-to-action button drawn on a Canvas —
        plain tk.Button can't do rounded corners, but a canvas rounded
        rectangle with bound click/hover/press handlers looks and behaves
        like a real modern button (hover lightens, press darkens, cursor
        changes) and scales cleanly with the window. Pack the returned
        frame with fill=X; height is fixed, width stretches to match.
        `panel_bg` should match whatever surface this sits on (a dark
        card vs. the normal light canvas) so the corners mask cleanly.
        """
        bg = bg or COLOR_ACCENT
        panel_bg = panel_bg or BG_CANVAS
        hover_bg = self._shade_hex_color(bg, 0.90)
        press_bg = self._shade_hex_color(bg, 0.78)
        state = {"bg": bg, "enabled": True}

        wrap = Frame(parent, bg=panel_bg)
        canvas = Canvas(wrap, height=height, highlightthickness=0, bd=0, bg=panel_bg, cursor="hand2")
        canvas.pack(fill=X)

        def draw(color):
            canvas.delete("all")
            w = canvas.winfo_width() or width
            h = canvas.winfo_height() or height
            self._rounded_rect(canvas, 1, 1, w - 1, h - 1, radius=radius, fill=color, outline="")
            label = f"{icon}   {text}" if icon else text
            canvas.create_text(w / 2, h / 2, text=label, fill=fg, font=("Segoe UI", font_size, "bold"))

        def on_enter(e):
            if state["enabled"]:
                draw(hover_bg)

        def on_leave(e):
            draw(state["bg"])

        def on_press(e):
            if state["enabled"]:
                draw(press_bg)

        def on_release(e):
            if state["enabled"]:
                draw(hover_bg)
                command()

        canvas.bind("<Configure>", lambda e: draw(state["bg"]))
        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<ButtonRelease-1>", on_release)

        def set_enabled(is_enabled):
            state["enabled"] = is_enabled
            draw(state["bg"] if is_enabled else self._shade_hex_color(bg, 1.35))

        wrap.set_enabled = set_enabled
        wrap.canvas = canvas
        return wrap

    def _gradient_button(self, parent, text, command, c1=(0x10, 0xB9, 0x81), c2=(0x06, 0xB6, 0xD4),
                          fg="white", height=48, font_size=12, radius=14, panel_bg=None):
        """
        A true two-tone gradient rounded button (emerald -> cyan by
        default) — Canvas polygons can't fill with a gradient directly,
        so this paints thin vertical color-blend bands across the full
        rectangle, then masks the four corners back to `panel_bg` and
        redraws each corner as a quarter-circle in the nearest band
        color. The net result is a genuinely rounded button with a real
        left-to-right color gradient, not just a flat tint.
        """
        panel_bg = panel_bg or AUTH_BG
        state = {"mode": "normal", "enabled": True}

        wrap = Frame(parent, bg=panel_bg)
        canvas = Canvas(wrap, height=height, highlightthickness=0, bd=0, bg=panel_bg, cursor="hand2")
        canvas.pack(fill=X)

        def _blend(a, b, t):
            return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))

        def _adjust(rgb, mode):
            if mode == "press":
                return tuple(max(0, int(c * 0.80)) for c in rgb)
            if mode == "hover":
                return tuple(min(255, int(c * 1.10)) for c in rgb)
            return rgb

        def draw():
            canvas.delete("all")
            w = canvas.winfo_width() or 280
            h = canvas.winfo_height() or height
            mode = state["mode"]
            steps = max(24, w // 6)
            for i in range(steps):
                t = i / steps
                rgb = _adjust(_blend(c1, c2, t), mode)
                color = "#%02x%02x%02x" % rgb
                x0 = int(w * i / steps)
                x1 = int(w * (i + 1) / steps) + 1
                canvas.create_rectangle(x0, 0, x1, h, fill=color, outline="", tags="grad")
            r = radius
            left_rgb = _adjust(c1, mode)
            right_rgb = _adjust(c2, mode)
            left_color = "#%02x%02x%02x" % left_rgb
            right_color = "#%02x%02x%02x" % right_rgb
            # Corner blocks recolored to the panel background, then the
            # rounded corner itself redrawn as a solid pieslice so the
            # gradient rectangle reads as a rounded pill.
            canvas.create_rectangle(0, 0, r, r, fill=panel_bg, outline="", tags="grad")
            canvas.create_rectangle(w - r, 0, w, r, fill=panel_bg, outline="", tags="grad")
            canvas.create_rectangle(0, h - r, r, h, fill=panel_bg, outline="", tags="grad")
            canvas.create_rectangle(w - r, h - r, w, h, fill=panel_bg, outline="", tags="grad")
            canvas.create_arc(0, 0, 2 * r, 2 * r, start=90, extent=90, fill=left_color, outline="", style="pieslice", tags="grad")
            canvas.create_arc(w - 2 * r, 0, w, 2 * r, start=0, extent=90, fill=right_color, outline="", style="pieslice", tags="grad")
            canvas.create_arc(0, h - 2 * r, 2 * r, h, start=180, extent=90, fill=left_color, outline="", style="pieslice", tags="grad")
            canvas.create_arc(w - 2 * r, h - 2 * r, w, h, start=270, extent=90, fill=right_color, outline="", style="pieslice", tags="grad")
            canvas.create_text(w / 2, h / 2, text=text, fill=fg, font=("Segoe UI", font_size, "bold"), tags="grad")

        def set_mode(mode):
            if state["enabled"] or mode == "normal":
                state["mode"] = mode
                draw()

        canvas.bind("<Configure>", lambda e: draw())
        canvas.bind("<Enter>", lambda e: set_mode("hover"))
        canvas.bind("<Leave>", lambda e: set_mode("normal"))
        canvas.bind("<ButtonPress-1>", lambda e: set_mode("press"))
        canvas.bind("<ButtonRelease-1>", lambda e: (set_mode("hover"), command()))

        wrap.canvas = canvas
        return wrap

    def _draw_cooper_robot(self, canvas, cx, cy, scale=1.0, tag="robot"):
        """
        Cooper, the app's mascot, drawn from plain Canvas primitives so it
        matches everywhere Cooper AI appears (no external image needed).
        Rounded body + head, glowing cyan eyes, a small antenna, and an
        emerald chest light. `tag` lets a caller move/animate the whole
        group together via canvas.move(tag, dx, dy).
        """
        s = scale
        # Soft floor shadow (flattened oval) — grounds the "floating" feel
        canvas.create_oval(cx - 46 * s, cy + 92 * s, cx + 46 * s, cy + 104 * s,
                            fill="#050914", outline="", tags=(tag, "shadow"))

        # Body
        self._rounded_rect(canvas, cx - 50 * s, cy - 30 * s, cx + 50 * s, cy + 84 * s,
                            radius=26 * s, fill="#1E293B", outline=AUTH_EMERALD, width=1, tags=tag)
        # Chest panel + emerald light
        self._rounded_rect(canvas, cx - 24 * s, cy + 8 * s, cx + 24 * s, cy + 46 * s,
                            radius=10 * s, fill="#0F172A", outline="", tags=tag)
        canvas.create_oval(cx - 7 * s, cy + 20 * s, cx + 7 * s, cy + 34 * s,
                            fill=AUTH_EMERALD, outline="", tags=tag)

        # Head
        self._rounded_rect(canvas, cx - 44 * s, cy - 92 * s, cx + 44 * s, cy - 22 * s,
                            radius=22 * s, fill="#1E293B", outline=AUTH_CYAN, width=1, tags=tag)
        # Antenna
        canvas.create_line(cx, cy - 92 * s, cx, cy - 112 * s, fill=AUTH_CYAN, width=2, tags=tag)
        canvas.create_oval(cx - 5 * s, cy - 122 * s, cx + 5 * s, cy - 112 * s,
                            fill=AUTH_CYAN, outline="", tags=tag)
        # Eye visor
        self._rounded_rect(canvas, cx - 32 * s, cy - 68 * s, cx + 32 * s, cy - 42 * s,
                            radius=12 * s, fill="#0B1120", outline="", tags=tag)
        # Glowing eyes: a soft halo behind a bright core
        for ex in (-14, 14):
            hx = cx + ex * s
            canvas.create_oval(hx - 9 * s, cy - 60 * s, hx + 9 * s, cy - 44 * s,
                                fill="#0E7490", outline="", tags=tag)
            canvas.create_oval(hx - 5 * s, cy - 58 * s, hx + 5 * s, cy - 48 * s,
                                fill=AUTH_CYAN, outline="", tags=tag)
        # Little arms
        canvas.create_oval(cx - 62 * s, cy + 6 * s, cx - 46 * s, cy + 40 * s,
                            fill="#1E293B", outline=AUTH_EMERALD, width=1, tags=tag)
        canvas.create_oval(cx + 46 * s, cy + 6 * s, cx + 62 * s, cy + 40 * s,
                            fill="#1E293B", outline=AUTH_EMERALD, width=1, tags=tag)

    def _render_auth_hero(self, hero, hero_w):
        """
        Paints the full dark FinTech landing scene onto `hero`: gradient
        background, ambient glow rings, drifting particles, floating
        finance glyphs, the Cooper mascot, brand copy, and a stack of
        feature mini-cards. Returns nothing — call again on <Configure>
        to keep it filling the canvas as it resizes.
        """
        hero.delete("scene")
        w = hero.winfo_width() or hero_w
        h = hero.winfo_height() or 820

        # Deep navy background with a very subtle top-to-bottom shift
        self._paint_vertical_gradient(hero, (0x0B, 0x11, 0x20), (0x0F, 0x17, 0x2A), tag="scene", steps=48)
        hero.tag_lower("scene")

        # Ambient glow rings behind the mascot (thin concentric outlines)
        rcx, rcy = w * 0.5, h * 0.40
        for i, (radius, color) in enumerate([(150, AUTH_EMERALD), (200, AUTH_CYAN), (250, AUTH_PURPLE)]):
            hero.create_oval(rcx - radius, rcy - radius, rcx + radius, rcy + radius,
                              outline=color, width=1, tags="scene")

        # Static particle field — fixed seed so it doesn't jitter on redraw
        rng = random.Random(42)
        for _ in range(26):
            px = rng.uniform(0.04, 0.96) * w
            py = rng.uniform(0.06, 0.94) * h
            pr = rng.uniform(1.2, 2.6)
            tint = rng.choice([AUTH_EMERALD, AUTH_CYAN, AUTH_PURPLE, "#334155"])
            hero.create_oval(px - pr, py - pr, px + pr, py + pr, fill=tint, outline="", tags="scene")

        # Floating finance glyphs, scattered around the mascot
        glyphs = [
            ("₹", 0.16, 0.20), ("💳", 0.82, 0.16), ("🪙", 0.12, 0.52),
            ("📈", 0.84, 0.50), ("🔒", 0.20, 0.78), ("👛", 0.80, 0.80),
        ]
        for glyph, fx, fy in glyphs:
            hero.create_text(w * fx, h * fy, text=glyph, font=("Segoe UI", 15), fill="#475569", tags="scene")

        # Brand block
        hero.create_text(w * 0.5, h * 0.10, text="🦊 COOPER VAULT", font=("Segoe UI", 15, "bold"),
                          fill=AUTH_TEXT, tags="scene")
        hero.create_text(w * 0.5, h * 0.135, text="AI-POWERED PERSONAL FINANCE", font=("Segoe UI", 7, "bold"),
                          fill=AUTH_EMERALD, tags="scene")

        # Cooper mascot (drawn as its own tag group so it can float/animate)
        hero.delete("robot")
        self._draw_cooper_robot(hero, rcx, rcy, scale=0.95, tag="robot")

        hero.create_text(w * 0.5, h * 0.66, text="Understand your money.", font=("Segoe UI", 13, "bold"),
                          fill=AUTH_TEXT, tags="scene")
        hero.create_text(w * 0.5, h * 0.685, text="Control your future.", font=("Segoe UI", 13, "bold"),
                          fill=AUTH_TEXT, tags="scene")

        # Feature mini-cards
        cards = [
            ("✦", "AI Insights", "Smart financial recommendations"),
            ("↗", "Smart Analytics", "Track spending effortlessly"),
            ("🔒", "Secure Vault", "Your data stays protected"),
        ]
        card_w = min(230, w - 48)
        card_h = 50
        cx0 = (w - card_w) / 2
        cy0 = h * 0.75
        for i, (icon, head, sub) in enumerate(cards):
            y1 = cy0 + i * (card_h + 10)
            self._rounded_rect(hero, cx0, y1, cx0 + card_w, y1 + card_h, radius=10,
                                fill=AUTH_PANEL, outline=AUTH_BORDER, width=1, tags="scene")
            hero.create_text(cx0 + 22, y1 + card_h / 2, text=icon, font=("Segoe UI", 12), fill=AUTH_EMERALD, tags="scene")
            hero.create_text(cx0 + 42, y1 + 16, text=head, font=("Segoe UI", 9, "bold"), fill=AUTH_TEXT, anchor="w", tags="scene")
            hero.create_text(cx0 + 42, y1 + 34, text=sub, font=("Segoe UI", 7), fill=AUTH_MUTED, anchor="w", tags="scene")

        # Cards, glyphs, and mascot are drawn in the right order already
        # (mascot last => on top); no extra stacking calls needed.

    def _start_hero_float_animation(self, hero):
        """
        A lightweight idle animation: Cooper drifts a few pixels up and
        down on a sine wave. Cancels any previous loop first so switching
        between Login / Sign Up / Forgot Password doesn't stack up
        duplicate `after()` timers.
        """
        if getattr(self, "_auth_anim_job", None):
            try:
                self.after_cancel(self._auth_anim_job)
            except Exception:
                pass
            self._auth_anim_job = None

        state = {"t": 0.0, "last_dy": 0.0}

        def tick():
            if not hero.winfo_exists():
                return
            state["t"] += 0.09
            dy = math.sin(state["t"]) * 6.0
            try:
                hero.move("robot", 0, dy - state["last_dy"])
            except Exception:
                return
            state["last_dy"] = dy
            self._auth_anim_job = self.after(60, tick)

        tick()

    def _auth_card(self, title, subtitle):
        """
        Shared shell for Sign Up / Login / Forgot Password: a full dark
        FinTech landing scene on the left (Cooper mascot, particles,
        glow rings, feature cards) and a compact dark login card on the
        right — split roughly 58/42. Always dark regardless of the
        workspace's light/dark theme toggle. Returns the Frame each
        screen packs its own fields into — same contract as before, so
        build_signup_screen / build_login_screen / build_forgot_password_screen
        don't need structural changes.
        """
        self._clear_root()
        self.configure(bg=AUTH_BG)

        root_split = Frame(self, bg=AUTH_BG)
        root_split.pack(fill=BOTH, expand=True)

        # A weighted 3:2 grid keeps the hero/form ratio at ~60/40 at any
        # window size (rather than the hero eating all leftover space via
        # pack(expand=True), which was pinning the card flush against the
        # right edge with no margin).
        hero_visible = self.winfo_screenwidth() >= 980
        if hero_visible:
            root_split.grid_columnconfigure(0, weight=3, uniform="auth")
            root_split.grid_columnconfigure(1, weight=2, uniform="auth")
        else:
            root_split.grid_columnconfigure(0, weight=1)
        root_split.grid_rowconfigure(0, weight=1)

        # ---------------- LEFT: DARK CANVAS HERO SCENE (~60%) ---------------- #
        if hero_visible:
            hero = Canvas(root_split, highlightthickness=0, bd=0, bg=AUTH_BG)
            hero.grid(row=0, column=0, sticky="nsew")

            def render_hero(event=None):
                self._render_auth_hero(hero, hero.winfo_width())

            hero.bind("<Configure>", render_hero)
            self.after(30, lambda: self._start_hero_float_animation(hero))

        # ---------------- RIGHT: DARK LOGIN CARD (~40%) ---------------- #
        form_col = 1 if hero_visible else 0
        form_area = Frame(root_split, bg=AUTH_BG)
        form_area.grid(row=0, column=form_col, sticky="nsew")

        shadow_outer, card = self._shadow_card(
            form_area, padx=34, pady=32, accent=AUTH_EMERALD,
            card_bg=AUTH_CARD, border_color=AUTH_BORDER, panel_bg=AUTH_BG
        )
        # Centered within the form column, with a visible margin on every
        # side so the card never touches the column's own edges either.
        shadow_outer.place(relx=0.46, rely=0.48, anchor="center")

        if not hero_visible:
            Label(card, text="🦊 COOPER VAULT", font=("Segoe UI", 14, "bold"), fg=AUTH_TEXT, bg=AUTH_CARD).pack(anchor="w")
            Label(card, text="AI-POWERED PERSONAL FINANCE", font=("Segoe UI", 7, "bold"), fg=AUTH_EMERALD, bg=AUTH_CARD).pack(anchor="w", pady=(0, 14))

        Label(card, text=title, font=("Segoe UI", 18, "bold"), fg=AUTH_TEXT, bg=AUTH_CARD).pack(anchor="w")
        Label(card, text=subtitle, font=("Segoe UI", 9), fg=AUTH_MUTED, bg=AUTH_CARD, wraplength=320, justify=LEFT).pack(anchor="w", pady=(4, 16))

        return card

    def _auth_entry(self, parent, label_text, show=None, icon="✉"):
        """
        Boxed dark input for the auth screens: a filled pill-style field
        (icon + Entry) with a hairline border that lights up emerald on
        focus — matches the "[ icon  Enter email ]" style requested for
        the login card, distinct from the underline inputs used
        elsewhere in the app (e.g. the Profile page).
        """
        Label(parent, text=label_text, font=("Segoe UI", 8, "bold"), fg=AUTH_MUTED, bg=AUTH_CARD).pack(anchor="w", pady=(10, 4))

        box = Frame(parent, bg=AUTH_INPUT, highlightthickness=1, highlightbackground=AUTH_BORDER)
        box.pack(fill=X)

        Label(box, text=icon, font=("Segoe UI", 10), fg=AUTH_MUTED, bg=AUTH_INPUT).pack(side=LEFT, padx=(10, 4))

        entry = Entry(
            box, bg=AUTH_INPUT, fg=AUTH_TEXT, bd=0, font=("Segoe UI", 10),
            highlightthickness=0, show=show, insertbackground=AUTH_TEXT
        )
        entry.pack(side=LEFT, fill=X, expand=True, ipady=7, padx=(0, 10))

        entry.bind("<FocusIn>", lambda e: box.config(highlightbackground=AUTH_EMERALD))
        entry.bind("<FocusOut>", lambda e: box.config(highlightbackground=AUTH_BORDER))

        return entry

    def _auth_password_entry(self, parent, label_text):
        """
        Same boxed dark style as `_auth_entry`, plus a 👁 show/hide toggle.
        """
        Label(parent, text=label_text, font=("Segoe UI", 8, "bold"), fg=AUTH_MUTED, bg=AUTH_CARD).pack(anchor="w", pady=(10, 4))

        box = Frame(parent, bg=AUTH_INPUT, highlightthickness=1, highlightbackground=AUTH_BORDER)
        box.pack(fill=X)

        Label(box, text="🔒", font=("Segoe UI", 10), fg=AUTH_MUTED, bg=AUTH_INPUT).pack(side=LEFT, padx=(10, 4))

        entry = Entry(
            box, bg=AUTH_INPUT, fg=AUTH_TEXT, bd=0, font=("Segoe UI", 10),
            highlightthickness=0, show="•", insertbackground=AUTH_TEXT
        )
        entry.pack(side=LEFT, fill=X, expand=True, ipady=7)

        show_pw = BooleanVar(value=False)

        def toggle_pw():
            entry.config(show="" if show_pw.get() else "•")

        Checkbutton(
            box, text="👁", variable=show_pw, command=toggle_pw, bg=AUTH_INPUT,
            activebackground=AUTH_INPUT, bd=0, fg=AUTH_MUTED, selectcolor=AUTH_INPUT
        ).pack(side=LEFT, padx=(2, 8))

        entry.bind("<FocusIn>", lambda e: box.config(highlightbackground=AUTH_EMERALD))
        entry.bind("<FocusOut>", lambda e: box.config(highlightbackground=AUTH_BORDER))

        return entry

    def _labeled_entry(self, parent, label_text, show=None, width=280):
        """
        A modern underlined input: no boxed border, just a thin line under
        the field that lights up in the accent color on focus — a cleaner,
        more current look than a flat bordered box.
        """
        Label(parent, text=label_text, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(12, 5))

        field_wrap = Frame(parent, bg=BG_CANVAS)
        field_wrap.pack(fill=X)

        entry = Entry(
            field_wrap, bg=BG_CANVAS, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 12),
            highlightthickness=0, show=show, insertbackground=TEXT_MAIN
        )
        entry.pack(fill=X, ipady=7)

        underline = Frame(field_wrap, bg=BORDER_COLOR, height=2)
        underline.pack(fill=X)

        entry.bind("<FocusIn>", lambda e: underline.config(bg=COLOR_ACCENT))
        entry.bind("<FocusOut>", lambda e: underline.config(bg=BORDER_COLOR))

        return entry

    def _labeled_password_entry(self, parent, label_text):
        """
        Same underline style as _labeled_entry, plus a 👁 show/hide toggle —
        used by every password field so they all look and behave alike
        instead of each screen hand-rolling its own version.
        """
        Label(parent, text=label_text, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(12, 5))

        field_wrap = Frame(parent, bg=BG_CANVAS)
        field_wrap.pack(fill=X)

        row = Frame(field_wrap, bg=BG_CANVAS)
        row.pack(fill=X)

        entry = Entry(
            row, bg=BG_CANVAS, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 12),
            highlightthickness=0, show="•", insertbackground=TEXT_MAIN
        )
        entry.pack(side=LEFT, fill=X, expand=True, ipady=7)

        show_pw = BooleanVar(value=False)

        def toggle_pw():
            entry.config(show="" if show_pw.get() else "•")

        Checkbutton(
            row, text="👁", variable=show_pw, command=toggle_pw, bg=BG_CANVAS,
            activebackground=BG_CANVAS, bd=0, fg=TEXT_MUTED, selectcolor=BG_CANVAS
        ).pack(side=LEFT, padx=(6, 0))

        underline = Frame(field_wrap, bg=BORDER_COLOR, height=2)
        underline.pack(fill=X)

        entry.bind("<FocusIn>", lambda e: underline.config(bg=COLOR_ACCENT))
        entry.bind("<FocusOut>", lambda e: underline.config(bg=BORDER_COLOR))

        return entry


    def _pick_profile_picture(self, preview_label, state):
        filepath = filedialog.askopenfilename(
            title="Choose a profile picture",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif")]
        )
        if not filepath:
            return
        try:
            ext = os.path.splitext(filepath)[1] or ".png"
            dest_name = f"profile_{secrets.token_hex(6)}{ext}"
            dest_path = os.path.join(PROFILE_PICTURES_DIR, dest_name)
            shutil.copy2(filepath, dest_path)
            state["profile_picture"] = dest_path
            preview_label.config(text=f"✓ {os.path.basename(filepath)}", fg=COLOR_INCOME)
        except Exception as e:
            messagebox.showerror("Upload Failed", f"Could not use this image:\n{e}")

    # ------------------------------------------------------------
    # SIGN UP (first run — no accounts exist yet)
    # ------------------------------------------------------------
    def build_signup_screen(self):
        card = self._auth_card("Create Your Account", "Set up Cooper Vault — this only happens once.")

        name_entry = self._auth_entry(card, "Full Name", icon="🧑")
        username_entry = self._auth_entry(card, "Username", icon="👤")
        password_entry = self._auth_password_entry(card, "Password")

        # Compact live strength readout — updates on every keystroke,
        # sits right under the password field, no extra layout space.
        strength_label = Label(card, text="", font=("Segoe UI", 8, "bold"), bg=AUTH_CARD)
        strength_label.pack(anchor="w", pady=(4, 0))

        strength_colors = {"weak": COLOR_DANGER, "medium": "#F59E0B", "strong": AUTH_EMERALD}

        def update_strength(event=None):
            label, key = password_strength_label(password_entry.get())
            if not label:
                strength_label.config(text="")
                return
            strength_label.config(text=f"Password strength: {label}", fg=strength_colors[key])

        password_entry.bind("<KeyRelease>", update_strength)

        # Confirm Password now uses the same helper as Password, so it
        # gets its own 👁 show/hide toggle too (previously only the main
        # Password field had one).
        confirm_entry = self._auth_password_entry(card, "Confirm Password")

        Label(card, text="Security Question (for password recovery)", font=("Segoe UI", 8, "bold"), fg=AUTH_TEXT, bg=AUTH_CARD).pack(anchor="w", pady=(10, 4))
        question_var = StringVar(value="What is your favourite subject?")
        ttk.Combobox(
            card, textvariable=question_var, state="readonly", font=("Segoe UI", 9),
            values=[
                "What is your favourite subject?",
                "What is your pet's name?",
                "What city were you born in?",
                "What is your favourite food?",
            ]
        ).pack(fill=X)

        answer_entry = self._auth_entry(card, "Answer", icon="💬")

        Label(card, text="Profile Picture (optional)", font=("Segoe UI", 8, "bold"), fg=AUTH_TEXT, bg=AUTH_CARD).pack(anchor="w", pady=(10, 4))
        pic_row = Frame(card, bg=AUTH_CARD)
        pic_row.pack(fill=X)
        state = {"profile_picture": None}
        preview_label = Label(pic_row, text="No file chosen", font=("Segoe UI", 8), fg=AUTH_MUTED, bg=AUTH_CARD)
        preview_label.pack(side=LEFT)
        Button(pic_row, text="Choose...", bg=AUTH_INPUT, fg=AUTH_TEXT, bd=0, font=("Segoe UI", 8),
               padx=10, pady=4, command=lambda: self._pick_profile_picture(preview_label, state)).pack(side=RIGHT)

        def do_signup():
            global conn, cursor
            name = name_entry.get().strip()
            username = username_entry.get().strip()
            password = password_entry.get()
            confirm = confirm_entry.get()
            answer = answer_entry.get().strip()

            # ---- 1. Empty-field check ----
            if not name or not username or not password or not confirm:
                messagebox.showerror("Missing Info", "Please fill in all required fields.")
                return
            if not answer:
                messagebox.showerror("Missing Info", "Please answer the security question — it's needed to recover your password later.")
                return

            # ---- 2. Username format ----
            ok, msg = validate_username_format(username)
            if not ok:
                messagebox.showerror("Invalid Username", msg)
                return

            # ---- 3. Username uniqueness (case-insensitive) ----
            # SQLite's UNIQUE constraint on `username` is case-SENSITIVE
            # by default, so "Shloka" and "shloka" would otherwise be
            # treated as two different accounts — this explicit
            # case-insensitive lookup, run BEFORE any insert, is what
            # actually prevents that.
            auth_cursor.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?)", (username,))
            if auth_cursor.fetchone():
                messagebox.showerror("Username Taken", "Username already exists.\n\nPlease choose another username.")
                return

            # ---- 4. Password strength ----
            ok, msg = validate_password_strength(password)
            if not ok:
                messagebox.showerror("Weak Password", msg)
                return

            # ---- 5. Confirm password match ----
            if password != confirm:
                messagebox.showerror("Password Mismatch", "Passwords do not match.")
                return

            # ---- 6. Create account (only after every check above passes) ----
            salt = generate_salt()
            pw_hash = hash_with_salt(password, salt)
            answer_salt = generate_salt()
            answer_hash = hash_with_salt(answer.lower(), answer_salt)

            try:
                auth_cursor.execute(
                    """INSERT INTO users
                       (username, display_name, password_hash, salt, security_question,
                        security_answer_hash, profile_picture, remember_me, created_date)
                       VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?)""",
                    (username, name, pw_hash, salt, question_var.get(),
                     answer_salt + "$" + answer_hash, state["profile_picture"],
                     datetime.today().strftime("%d-%m-%Y"))
                )
                auth_conn.commit()
            except sqlite3.IntegrityError:
                # Belt-and-suspenders: the case-insensitive check above
                # already catches this in the normal case, but two
                # simultaneous registrations could still race past it.
                messagebox.showerror("Username Taken", "Username already exists.\n\nPlease choose another username.")
                return
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not create your account:\n{e}")
                return

            auth_cursor.execute("SELECT id FROM users WHERE username=?", (username,))
            self.current_user_id = auth_cursor.fetchone()[0]
            self.current_username = username
            self.current_display_name = name
            self.current_profile_picture = state["profile_picture"]

            # Brand-new account -> brand-new, empty private database.
            conn, cursor = get_user_db(username)

            messagebox.showinfo("Account Created", f"Welcome to Cooper Vault, {name}!")
            self.build_workspace()

        btn_create_account = self._gradient_button(card, "Create Account", do_signup, panel_bg=AUTH_CARD)
        btn_create_account.pack(fill=X, pady=(20, 0))

        Button(
            card, text="Already have an account? Log in", bg=AUTH_CARD, fg=AUTH_MUTED, bd=0,
            font=("Segoe UI", 8, "underline"), command=lambda: self.build_login_screen()
        ).pack(anchor="w", pady=(12, 0))

    # ------------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------------
    def build_login_screen(self):
        card = self._auth_card("Welcome Back", "Log in to your Cooper Vault.")

        username_entry = self._auth_entry(card, "Username", icon="👤")
        password_entry = self._auth_password_entry(card, "Password")

        options_row = Frame(card, bg=AUTH_CARD)
        options_row.pack(fill=X, pady=(10, 0))

        remember_var = BooleanVar(value=False)
        Checkbutton(options_row, text="Remember Me", variable=remember_var, bg=AUTH_CARD,
                    activebackground=AUTH_CARD, font=("Segoe UI", 8), fg=AUTH_MUTED, selectcolor=AUTH_CARD).pack(side=LEFT)

        Button(
            options_row, text="Forgot Password?", bg=AUTH_CARD, fg=AUTH_CYAN, bd=0,
            font=("Segoe UI", 8, "underline"), command=lambda: self.build_forgot_password_screen()
        ).pack(side=RIGHT)

        error_label = Label(card, text="", font=("Segoe UI", 8), fg="#F87171", bg=AUTH_CARD)
        error_label.pack(anchor="w", pady=(8, 0))

        def do_login(event=None):
            global conn, cursor
            username = username_entry.get().strip()
            password = password_entry.get()

            auth_cursor.execute(
                "SELECT id, display_name, salt, password_hash, profile_picture FROM users WHERE username=?",
                (username,)
            )
            row = auth_cursor.fetchone()
            if not row or hash_with_salt(password, row[2]) != row[3]:
                error_label.config(text="Incorrect username or password.")
                return

            user_id, display_name, salt, pw_hash, profile_picture = row

            auth_cursor.execute("UPDATE users SET remember_me=0")
            if remember_var.get():
                auth_cursor.execute("UPDATE users SET remember_me=1 WHERE id=?", (user_id,))
            auth_conn.commit()

            self.current_user_id = user_id
            self.current_username = username
            self.current_display_name = display_name
            self.current_profile_picture = profile_picture
            auth_cursor.execute("SELECT dark_mode FROM users WHERE id=?", (user_id,))
            apply_theme(bool(auth_cursor.fetchone()[0]))

            # Open THIS user's own private financial database — this is
            # the actual fix for cross-account data leakage: `conn` /
            # `cursor` now point at data/cooper_<username>.db, and every
            # existing page just uses those bare names unchanged.
            conn, cursor = get_user_db(username)

            self.build_workspace()

        password_entry.bind("<Return>", do_login)

        # Gradient emerald -> cyan sign-in button, per the premium
        # FinTech look requested for the login screen specifically.
        btn_login = self._gradient_button(card, "Sign In", do_login, panel_bg=AUTH_CARD)
        btn_login.pack(fill=X, pady=(18, 0))

        divider_row = Frame(card, bg=AUTH_CARD)
        divider_row.pack(fill=X, pady=(18, 12))
        Frame(divider_row, bg=AUTH_BORDER, height=1).pack(side=LEFT, fill=X, expand=True)
        Label(divider_row, text="  OR  ", font=("Segoe UI", 7, "bold"), fg=AUTH_MUTED, bg=AUTH_CARD).pack(side=LEFT)
        Frame(divider_row, bg=AUTH_BORDER, height=1).pack(side=LEFT, fill=X, expand=True)

        Button(
            card, text="Create an account", bg=AUTH_CARD, fg=AUTH_EMERALD, bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2", command=lambda: self.build_signup_screen()
        ).pack()

    # ------------------------------------------------------------
    # FORGOT PASSWORD
    # ------------------------------------------------------------
    def build_forgot_password_screen(self):
        card = self._auth_card("Reset Password", "Answer your security question to reset your password.")

        username_entry = self._auth_entry(card, "Username", icon="👤")
        question_label = Label(card, text="", font=("Segoe UI", 8, "bold"), fg=AUTH_TEXT, bg=AUTH_CARD, wraplength=280, justify=LEFT)
        question_label.pack(anchor="w", pady=(12, 0))

        answer_entry_holder = {"entry": None}
        new_pw_holder = {"entry": None}
        confirm_pw_holder = {"entry": None}
        found_user = {"id": None, "salt": None, "answer_hash": None}

        step2_frame = Frame(card, bg=AUTH_CARD)

        def find_question():
            username = username_entry.get().strip()
            auth_cursor.execute("SELECT id, security_question, security_answer_hash FROM users WHERE username=?", (username,))
            row = auth_cursor.fetchone()
            if not row:
                messagebox.showerror("Not Found", "No account with that username.")
                return

            found_user["id"], question, answer_hash = row
            found_user["answer_hash"] = answer_hash
            question_label.config(text=f"Q: {question}")

            for w in step2_frame.winfo_children():
                w.destroy()
            step2_frame.pack(fill=X)

            answer_entry_holder["entry"] = self._auth_entry(step2_frame, "Your Answer", icon="💬")
            new_pw_holder["entry"] = self._auth_entry(step2_frame, "New Password", show="•", icon="🔒")
            confirm_pw_holder["entry"] = self._auth_entry(step2_frame, "Confirm New Password", show="•", icon="🔒")

            def do_reset():
                answer = answer_entry_holder["entry"].get().strip().lower()
                new_pw = new_pw_holder["entry"].get()
                confirm_pw = confirm_pw_holder["entry"].get()

                stored_salt, stored_hash = found_user["answer_hash"].split("$", 1)
                if hash_with_salt(answer, stored_salt) != stored_hash:
                    messagebox.showerror("Incorrect Answer", "That answer doesn't match our records.")
                    return
                if len(new_pw) < 4 or new_pw != confirm_pw:
                    messagebox.showerror("Invalid Password", "Passwords must match and be at least 4 characters.")
                    return

                new_salt = generate_salt()
                auth_cursor.execute(
                    "UPDATE users SET password_hash=?, salt=? WHERE id=?",
                    (hash_with_salt(new_pw, new_salt), new_salt, found_user["id"])
                )
                auth_conn.commit()
                messagebox.showinfo("Password Reset", "Your password has been updated. Please log in.")
                self.build_login_screen()

            btn_reset_pw = self._gradient_button(step2_frame, "Reset Password", do_reset, panel_bg=AUTH_CARD)
            btn_reset_pw.pack(fill=X, pady=(18, 0))

        btn_find_account = self._rounded_button(
            card, "Find Account", find_question, bg=AUTH_INPUT, fg=AUTH_TEXT, icon="🔍", panel_bg=AUTH_CARD
        )
        btn_find_account.pack(fill=X, pady=(16, 0))

        step2_frame.pack(fill=X)

        Button(
            card, text="← Back to Login", bg=AUTH_CARD, fg=AUTH_MUTED, bd=0,
            font=("Segoe UI", 8, "underline"), command=lambda: self.build_login_screen()
        ).pack(anchor="w", pady=(16, 0))

    def logout(self):
        global conn, cursor
        if self.current_user_id is not None:
            try:
                auth_cursor.execute("UPDATE users SET remember_me=0 WHERE id=?", (self.current_user_id,))
                auth_conn.commit()
            except sqlite3.Error:
                pass

        # Close the current user's private financial database safely so
        # there's no leftover connection pointing at their data once
        # another account logs in.
        if conn is not None:
            try:
                conn.commit()
                conn.close()
            except sqlite3.Error:
                pass
        conn = None
        cursor = None

        self.current_user_id = None
        self.current_username = None
        self.current_display_name = None
        self.current_profile_picture = None
        self.build_login_screen()

    def highlight_sidebar(self, active_page):
        # Reset every nav row to its resting state
        for text, ref in self.nav_buttons.items():
            btn, bar, row = ref["btn"], ref["bar"], ref["row"]
            row.config(bg=BG_SIDEBAR)
            btn.config(bg=BG_SIDEBAR, fg=TEXT_MAIN, font=("Segoe UI", 10, "normal"))
            bar.config(bg=BG_SIDEBAR)

        # Highlight the selected row: a small emerald left indicator + a
        # subtle tinted card background + bold accent text — a compact
        # active state rather than a big solid block.
        if active_page in self.nav_buttons:
            ref = self.nav_buttons[active_page]
            btn, bar, row = ref["btn"], ref["bar"], ref["row"]
            active_tint = self._shade_hex_color(COLOR_INCOME, 1.82)
            row.config(bg=active_tint)
            btn.config(bg=active_tint, fg=COLOR_INCOME, font=("Segoe UI", 10, "bold"))
            bar.config(bg=COLOR_INCOME)

    def clean_view(self):
        for widget in content_frame.winfo_children():
            widget.destroy()
        if hasattr(self, "workspace_canvas"):
            self.workspace_canvas.yview_moveto(0)

    # ---------------- VISUAL POLISH HELPERS ---------------- #
    def _shade_hex_color(self, hex_color, factor):
        """
        Scales a #RRGGBB color toward black (factor < 1) or white
        (factor > 1). Used to derive a hover shade from a button's own
        background so every button gets sensible hover feedback without
        hand-picking a second color everywhere it's used.
        """
        hex_color = (hex_color or "").lstrip("#")
        if len(hex_color) != 6:
            return f"#{hex_color}" if hex_color else BG_CANVAS
        try:
            r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            return f"#{hex_color}"
        if factor <= 1:
            r, g, b = r * factor, g * factor, b * factor
        else:
            r = r + (255 - r) * (factor - 1)
            g = g + (255 - g) * (factor - 1)
            b = b + (255 - b) * (factor - 1)
        r, g, b = (max(0, min(255, int(c))) for c in (r, g, b))
        return f"#{r:02x}{g:02x}{b:02x}"

    def add_hover_effect(self, button, factor=0.88):
        """
        Plain tk.Button widgets have no built-in rollover feedback unless
        `activebackground` is set, so colored buttons otherwise flash the
        platform's default gray on hover. This binds a subtle darken/
        lighten of the button's own background on mouse enter/leave so
        hovering feels intentional. Safe no-op if bg isn't a hex color.
        """
        try:
            base_bg = button.cget("bg")
            hover_bg = self._shade_hex_color(base_bg, factor)
        except Exception:
            return button

        def on_enter(event):
            try:
                button.config(bg=hover_bg)
            except Exception:
                pass

        def on_leave(event):
            try:
                button.config(bg=base_bg)
            except Exception:
                pass

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        return button

    def refresh_ttk_style(self):
        """
        Re-applies ttk widget styling (Treeview, and any ttk.Combobox
        used in forms/popups) from the *current* theme globals. The
        style object is created once, but the color globals swap on
        every light/dark toggle — without this, tables would keep
        their light-mode colors even after switching to dark mode.
        Safe to call anytime; call again after any apply_theme().
        """
        self.ttk_style.configure(
            "Treeview", background=CARD_BG, foreground=TEXT_MAIN, rowheight=30,
            fieldbackground=CARD_BG, borderwidth=0, font=("Segoe UI", 9)
        )
        self.ttk_style.map(
            "Treeview",
            background=[("selected", self._shade_hex_color(COLOR_ACCENT, 1.5 if BG_CANVAS == "#FFFFFF" else 0.6))],
            foreground=[("selected", "white")],
        )
        self.ttk_style.configure(
            "Treeview.Heading", background=BG_SIDEBAR, foreground=TEXT_MUTED,
            borderwidth=0, font=("Segoe UI", 9, "bold")
        )
        self.ttk_style.configure(
            "TCombobox", fieldbackground=CARD_BG, background=CARD_BG, foreground=TEXT_MAIN,
            arrowcolor=TEXT_MUTED, bordercolor=BORDER_COLOR, font=("Segoe UI", 9)
        )
        self.ttk_style.map(
            "TCombobox",
            fieldbackground=[("readonly", CARD_BG)],
            foreground=[("readonly", TEXT_MAIN)],
        )

    def build_workspace(self):
        self._clear_root()
        self.configure(bg=BG_CANVAS)
        self.refresh_ttk_style()

        # LEFT NAVIGATION SIDEBAR PANEL (Matching pg1_2.png)
        sidebar = Frame(self, bg=BG_SIDEBAR, width=232, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)

        # Logo Context Area
        logo_f = Frame(sidebar, bg=BG_SIDEBAR, pady=18, padx=18)
        logo_f.pack(fill=X)
        Label(logo_f, text="🦊 COOPER", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_SIDEBAR, anchor="w").pack(fill=X)
        Label(logo_f, text="AI FINANCE", font=("Segoe UI", 7, "bold"), fg=TEXT_MUTED, bg=BG_SIDEBAR, anchor="w").pack(fill=X)

        # ------------------------------------------------------------
        # SCROLLABLE NAVIGATION LIST
        # A dedicated canvas + inner frame so a long menu never gets cut
        # off at the window's bottom edge — it scrolls instead, while
        # Profile and Logout stay pinned in the fixed footer below.
        # ------------------------------------------------------------
        nav_container = Frame(sidebar, bg=BG_SIDEBAR)
        nav_container.pack(fill=BOTH, expand=True)

        nav_canvas = Canvas(nav_container, bg=BG_SIDEBAR, highlightthickness=0, bd=0)
        nav_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        nav_scrollbar = Scrollbar(nav_container, orient="vertical", command=nav_canvas.yview, width=8)
        nav_scrollbar.pack(side=RIGHT, fill=Y)
        nav_canvas.configure(yscrollcommand=nav_scrollbar.set)

        nav_frame = Frame(nav_canvas, bg=BG_SIDEBAR)
        nav_window_id = nav_canvas.create_window((0, 0), window=nav_frame, anchor="nw")

        def _resize_nav_frame(event):
            nav_canvas.itemconfig(nav_window_id, width=event.width)
        nav_canvas.bind("<Configure>", _resize_nav_frame)

        nav_frame.bind(
            "<Configure>",
            lambda e: nav_canvas.configure(scrollregion=nav_canvas.bbox("all"))
        )

        def _on_nav_mousewheel(event):
            nav_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_nav_mousewheel(event):
            nav_canvas.bind_all("<MouseWheel>", _on_nav_mousewheel)

        def _unbind_nav_mousewheel(event):
            nav_canvas.unbind_all("<MouseWheel>")

        nav_canvas.bind("<Enter>", _bind_nav_mousewheel)
        nav_canvas.bind("<Leave>", _unbind_nav_mousewheel)

        # Navigation Menu Buttons — everything except Profile, which lives
        # in the pinned footer below so it's always reachable.
        menu_items = [
            ("🏠 Dashboard", self.load_overview),
            ("💳 Expenses", self.load_expenses),
            ("💰 Income", self.load_income),
            ("💼 Budget", self.load_budgets),
            ("🎯 Savings Goals", self.load_savings_goals),
            ("🔁 Recurring", self.load_recurring),
            ("⏰ Reminders", self.load_reminders),
            ("🧾 Receipt Scanner", self.load_receipt_scanner),
            ("📊 Reports", self.load_reports),
            ("🤖 Cooper AI", self.load_cooper_ai),
            ("📩 Message Reader", self.load_sms_scanner),
            ("📅 Bills & EMI", self.load_bills_emi),
            ("📓 Notepad & Tasks", self.load_notes_todo),
            ("👤 Profile", self.load_profile),
        ]

        self.nav_buttons = {}
        for text, command in menu_items:
            row = Frame(nav_frame, bg=BG_SIDEBAR)
            row.pack(fill=X, padx=(0, 8), pady=1)

            bar = Frame(row, bg=BG_SIDEBAR, width=4)
            bar.pack(side=LEFT, fill=Y)

            b = Button(row, text=f"  {text}", font=("Segoe UI", 10, "normal"),
                      fg=TEXT_MAIN, bg=BG_SIDEBAR, activebackground=BG_CANVAS, activeforeground=COLOR_ACCENT,
                      bd=0, anchor="w", cursor="hand2", padx=12, pady=8, command=command)
            b.pack(side=LEFT, fill=X, expand=True)
            self.nav_buttons[text] = {"btn": b, "bar": bar, "row": row}

        # User footer: pinned, always visible — Profile shortcut + who's
        # logged in + Logout, regardless of how long the nav list scrolls.
        divider = Frame(sidebar, bg=BORDER_COLOR, height=1)
        divider.pack(side=BOTTOM, fill=X)

        user_footer = Frame(sidebar, bg=BG_SIDEBAR, padx=14, pady=10)
        user_footer.pack(side=BOTTOM, fill=X)

        Label(
            user_footer,
            text=f"👋 {self.current_display_name or self.current_username or 'User'}",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_SIDEBAR,
            anchor="w"
        ).pack(fill=X)

        btn_logout = Button(
            user_footer,
            text="🚪 Log Out",
            bg=BG_SIDEBAR,
            fg=COLOR_DANGER,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            cursor="hand2",
            anchor="w",
            command=self.logout
        )
        btn_logout.pack(fill=X, pady=(3, 0))
        self.add_hover_effect(btn_logout, factor=0.92)

        # RIGHT VIEW MAIN WORKSPACE CONTAINER
        # Wrapped in a scrollable canvas so pages that run longer than the
        # window (Reports, Cooper AI, Reminders, Notepad & Tasks, etc.) are
        # always fully reachable instead of being cut off at the bottom.
        right_container = Frame(self, bg=BG_CANVAS)
        right_container.pack(side=RIGHT, expand=True, fill=BOTH)

        workspace_canvas = Canvas(right_container, bg=BG_CANVAS, highlightthickness=0)
        workspace_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.workspace_canvas = workspace_canvas

        workspace_scrollbar = Scrollbar(right_container, orient="vertical", command=workspace_canvas.yview)
        workspace_scrollbar.pack(side=RIGHT, fill=Y)
        workspace_canvas.configure(yscrollcommand=workspace_scrollbar.set)

        global content_frame
        content_frame = Frame(workspace_canvas, bg=BG_CANVAS, padx=30, pady=22)
        content_window_id = workspace_canvas.create_window((0, 0), window=content_frame, anchor="nw")

        def _resize_content_frame(event):
            workspace_canvas.itemconfig(content_window_id, width=event.width)

        workspace_canvas.bind("<Configure>", _resize_content_frame)

        content_frame.bind(
            "<Configure>",
            lambda e: workspace_canvas.configure(scrollregion=workspace_canvas.bbox("all"))
        )

        # Only capture the mouse wheel while the pointer is actually over the
        # main content area, so it doesn't fight with popup scroll areas
        # (which bind their own wheel handler while open).
        def _on_workspace_mousewheel(event):
            workspace_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_workspace_mousewheel(event):
            workspace_canvas.bind_all("<MouseWheel>", _on_workspace_mousewheel)

        def _unbind_workspace_mousewheel(event):
            workspace_canvas.unbind_all("<MouseWheel>")

        workspace_canvas.bind("<Enter>", _bind_workspace_mousewheel)
        workspace_canvas.bind("<Leave>", _unbind_workspace_mousewheel)
        
        # Catch up on any recurring income/expenses that came due while the
        # app was closed, before the first screen is drawn.
        self.process_due_recurring(silent=True)

        self.load_overview()

        self.highlight_sidebar("🏠 Dashboard")

    # ---------------- INTERACTIVE PAGES ---------------- #
    
    # PANEL 1: DASHBOARD OVERVIEW (Ref: pg1_2.png)
    def load_overview(self):
        self.highlight_sidebar("🏠 Dashboard")
        self.clean_view()

        # ---------------- DASHBOARD LIVE DATA ---------------- #
        total_income, total_expenses, current_balance = self.get_dashboard_totals()
        upcoming_bills_amount, overdue_bills_count, next_due_bill_text = self.get_bill_dashboard_summary()
        total_savings = self.get_total_savings()
        budget_remaining = self.get_budget_remaining_current_month()
        monthly_balance = self.get_current_month_balance()

        month_labels, income_series, expense_series = self.get_monthly_income_expense_data()

        # ---------------- HEADER ---------------- #
        header = Frame(content_frame, bg=BG_CANVAS)
        header.pack(fill=X, pady=(0, 25))

        # Single timestamp drives BOTH the greeting text and the clock snapshot,
        # so the two can never fall out of sync with each other.
        # Fixed IST offset (UTC+5:30, no DST in India) so the greeting/clock
        # are correct for Indian users even if the machine's system clock
        # is set to a different timezone (e.g. server running in UTC).
        now = datetime.now()
        current_hour = now.hour

        if 5 <= current_hour < 12:
            greet_text, icon_kind, icon_color = "Good Morning ☀️", "sun", "#FBBF24"
        elif 12 <= current_hour < 17:
            greet_text, icon_kind, icon_color = "Good Afternoon 🌤️", "sun", "#F97316"
        elif 17 <= current_hour < 22:
            greet_text, icon_kind, icon_color = "Good Evening 🌇", "moon", "#6366F1"
        else:
            greet_text, icon_kind, icon_color = "Good Night 🌙", "moon", "#4338CA"

        header_top = Frame(header, bg=BG_CANVAS)
        header_top.pack(fill=X)

        # ---- left: greeting + hand-drawn icon (avoids emoji-font glyph issues) ----
        greet_row = Frame(header_top, bg=BG_CANVAS)
        greet_row.pack(side=LEFT, anchor="w")

        icon_canvas = Canvas(greet_row, width=30, height=30, bg=BG_CANVAS, highlightthickness=0)
        icon_canvas.pack(side=LEFT, padx=(0, 8))

        if icon_kind == "sun":
            icon_canvas.create_oval(8, 8, 22, 22, fill=icon_color, outline="")
            for ang in range(0, 360, 45):
                x1 = 15 + 9 * math.cos(math.radians(ang))
                y1 = 15 + 9 * math.sin(math.radians(ang))
                x2 = 15 + 14 * math.cos(math.radians(ang))
                y2 = 15 + 14 * math.sin(math.radians(ang))
                icon_canvas.create_line(x1, y1, x2, y2, fill=icon_color, width=2, capstyle=ROUND)
        else:
            icon_canvas.create_oval(6, 6, 24, 24, fill=icon_color, outline="")
            icon_canvas.create_oval(11, 3, 27, 21, fill=BG_CANVAS, outline="")

        Label(
            greet_row,
            text=greet_text,
            font=("Segoe UI", 22, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(side=LEFT)

        # ---- right: live snapshot clock, built from the SAME `now` as the greeting ----
        clock_col = Frame(header_top, bg=BG_CANVAS)
        clock_col.pack(side=RIGHT, anchor="e")

        Label(
            clock_col,
            text=now.strftime("%I %M %p") + "  •  " + now.strftime("%d %B %Y"),
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="e")

        Label(
            header,
            text="Your financial overview",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 0))

        # ---------------- TOP METRICS ---------------- #
        ribbon = Frame(content_frame, bg=BG_CANVAS)
        ribbon.pack(fill=X, pady=(0, 28))

        budget_has_data = budget_remaining != 0 or (
            datetime.now().strftime("%m-%Y") in [
                row[0] for row in cursor.execute(
                    "SELECT DISTINCT month_year FROM budgets"
                ).fetchall()
            ]
        )

        metrics = [
            ("TOTAL INCOME", f"₹{total_income:,.0f}", "All time", COLOR_INCOME),
            ("TOTAL EXPENSE", f"₹{total_expenses:,.0f}", "All time", COLOR_EXPENSE),
            ("UPCOMING BILLS", f"₹{upcoming_bills_amount:,.0f}", f"{overdue_bills_count} overdue", COLOR_DANGER if overdue_bills_count > 0 else COLOR_ACCENT),
            ("SAVINGS", f"₹{total_savings:,.0f}", "All goals", COLOR_SAVINGS),
            ("BUDGET REMAINING", f"₹{budget_remaining:,.0f}" if budget_has_data else "₹0", "This month" if budget_has_data else "No budget set", COLOR_DANGER if budget_remaining < 0 else COLOR_ACCENT),
            ("MONTHLY BALANCE", f"₹{monthly_balance:,.0f}", "This month", COLOR_INCOME if monthly_balance >= 0 else COLOR_DANGER),
        ]

        for idx in range(3):
            ribbon.columnconfigure(idx, weight=1)

        for i, (title, value, badge, color) in enumerate(metrics):
            row, col = divmod(i, 3)

            card = Frame(
                ribbon,
                bg=BG_CANVAS,
                padx=18,
                pady=16,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR
            )
            card.grid(row=row, column=col, padx=8, pady=(0, 12), sticky="ew")

            accent_bar = Frame(card, bg=color, height=4)
            accent_bar.pack(fill=X, side=TOP, pady=(0, 12))

            Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

            value_row = Frame(card, bg=BG_CANVAS)
            value_row.pack(fill=X, pady=(6, 0))

            Label(
                value_row,
                text=value,
                font=("Segoe UI", 18, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(side=LEFT)

            Label(
                value_row,
                text=f"  {badge}",
                font=("Segoe UI", 9, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(side=LEFT, pady=(4, 0))

        # ---------------- MAIN GRID ---------------- #
        deck = Frame(content_frame, bg=BG_CANVAS)
        deck.pack(fill=BOTH, expand=True)
        # minsize on column 1 guarantees the right column (Recent
        # Transactions / Upcoming Bills Snapshot) always keeps enough
        # width for its headings, even if the left column's matplotlib
        # chart wants more room than its 4/9 weight share on a narrower
        # window — without this, grid was free to compress column 1
        # below its content's natural width, clipping "Upcoming Bills
        # Snapshot" at the edge of its cell.
        deck.columnconfigure(0, weight=4)
        deck.columnconfigure(1, weight=5, minsize=300)
        deck.rowconfigure(0, weight=1)
        deck.rowconfigure(1, weight=1)

        # =========================================================
        # LEFT TOP: SPENDING BY CATEGORY
        # =========================================================
        pie_card = Frame(
            deck,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=15,
            pady=15
        )
        pie_card.grid(row=0, column=0, padx=8, pady=8, sticky="nsew")

        # ---------- month filter setup ----------
        # Always offer the last 12 months so the user can browse forward/back,
        # not just the months that happen to already have expenses logged.
        today = datetime.now().replace(day=1)
        month_display_to_key = {}
        for i in range(12):
            year = today.year
            month = today.month - i
            while month <= 0:
                month += 12
                year -= 1
            m_key = f"{month:02d}-{year}"
            disp = datetime(year, month, 1).strftime("%b %Y")
            month_display_to_key[disp] = m_key

        current_month_display = datetime.now().strftime("%b %Y")

        if not hasattr(self, "dashboard_month_filter") or self.dashboard_month_filter not in month_display_to_key:
            self.dashboard_month_filter = current_month_display

        pie_title_row = Frame(pie_card, bg=BG_CANVAS)
        pie_title_row.pack(fill=X, pady=(0, 2))

        Label(
            pie_title_row,
            text="Spending by Category",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(side=LEFT)

        month_var = StringVar(value=self.dashboard_month_filter)
        month_combo = ttk.Combobox(
            pie_title_row,
            textvariable=month_var,
            values=list(month_display_to_key.keys()),
            state="readonly",
            width=11,
            font=("Segoe UI", 9)
        )
        month_combo.pack(side=RIGHT)

        def on_month_change(event):
            self.dashboard_month_filter = month_var.get()
            self.load_overview()

        month_combo.bind("<<ComboboxSelected>>", on_month_change)

        Label(
            pie_card,
            text="Hover or click a slice to see that category's total for the selected month",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 10))

        # ---------- fetch category spending for the selected month ----------
        selected_month_key = month_display_to_key[self.dashboard_month_filter]
        chart_data = self.get_category_spending_data(selected_month_key)

        pie_wrap = Frame(pie_card, bg=BG_CANVAS)
        pie_wrap.pack(fill=BOTH, expand=True)

        pie_left = Frame(pie_wrap, bg=BG_CANVAS)
        pie_left.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 12))

        pie_right = Frame(pie_wrap, bg=BG_CANVAS, width=270)
        pie_right.pack(side=RIGHT, fill=Y, padx=(10, 0))
        pie_right.pack_propagate(False)

        if not chart_data or chart_data[0][0] == "No Data":
            Label(
                pie_left,
                text="No expense data available for this period",
                font=("Segoe UI", 11),
                fg=TEXT_MUTED,
                bg=BG_CANVAS,
                wraplength=260,
                justify=CENTER
            ).pack(expand=True)

        else:
            labels = [item[0] for item in chart_data]
            values = [item[1] for item in chart_data]

            total_spent = sum(values)

            pie_colors = [
                "#10B981", "#06B6D4", "#8B5CF6", "#F97316",
                "#F472B6", "#FACC15", "#22C55E", "#38BDF8",
                "#F87171", "#14B8A6"
            ]
            slice_colors = [pie_colors[i % len(pie_colors)] for i in range(len(values))]

            fig, ax = plt.subplots(figsize=(5.2, 3.2), facecolor=BG_CANVAS)
            ax.set_facecolor(BG_CANVAS)

            wedges, _ = ax.pie(
                values,
                colors=slice_colors,
                startangle=90,
                wedgeprops=dict(width=0.42, edgecolor="white", linewidth=1.5)
            )

            # center text
            ax.text(
                0, 0.05,
                f"₹{total_spent:,.0f}",
                ha="center",
                va="center",
                fontsize=15,
                fontweight="bold",
                color=TEXT_MAIN
            )
            ax.text(
                0, -0.18,
                "Total Spent",
                ha="center",
                va="center",
                fontsize=10,
                color=TEXT_MUTED
            )

            ax.axis("equal")

            # tooltip bubble shown on hover / click, styled like a native web tooltip
            tooltip = ax.annotate(
                "",
                xy=(0, 0),
                xytext=(18, 18),
                textcoords="offset points",
                fontsize=9,
                fontweight="bold",
                color="white",
                ha="left",
                va="center",
                bbox=dict(boxstyle="round,pad=0.55", fc="#111827", ec="none", alpha=0.95),
                zorder=10,
                visible=False
            )

            pie_canvas = FigureCanvasTkAgg(fig, master=pie_left)
            pie_canvas.draw()
            pie_canvas.get_tk_widget().pack(fill=BOTH, expand=True)
            pie_canvas.get_tk_widget().configure(cursor="hand2")

            # ---------- right side breakdown (legend) ----------
            Label(
                pie_right,
                text="Category Breakdown",
                font=("Segoe UI", 11, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(4, 12))

            legend_rows = []

            def highlight_legend(idx):
                for j, row in enumerate(legend_rows):
                    active_bg = "#F3F4F6" if j == idx else BG_CANVAS
                    row["item"].configure(bg=active_bg)
                    row["txt"].configure(bg=active_bg)
                    row["name_lbl"].configure(bg=active_bg)
                    row["amt_lbl"].configure(bg=active_bg)

            for i, (label, value) in enumerate(chart_data):
                item = Frame(pie_right, bg=BG_CANVAS)
                item.pack(fill=X, pady=3, ipady=3)

                color_box = Frame(
                    item,
                    bg=slice_colors[i],
                    width=12,
                    height=12
                )
                color_box.pack(side=LEFT, padx=(4, 8), pady=4)
                color_box.pack_propagate(False)

                txt = Frame(item, bg=BG_CANVAS)
                txt.pack(side=LEFT, fill=X, expand=True)

                # wraplength + no truncation, so long category names are never cut off
                name_lbl = Label(
                    txt,
                    text=label,
                    font=("Segoe UI", 10, "bold"),
                    fg=TEXT_MAIN,
                    bg=BG_CANVAS,
                    anchor="w",
                    justify=LEFT,
                    wraplength=180
                )
                name_lbl.pack(anchor="w", fill=X)

                pct = (value / total_spent * 100) if total_spent else 0
                amt_lbl = Label(
                    txt,
                    text=f"₹{value:,.0f}   ({pct:.1f}%)",
                    font=("Segoe UI", 9),
                    fg=TEXT_MUTED,
                    bg=BG_CANVAS,
                    anchor="w"
                )
                amt_lbl.pack(anchor="w")

                legend_rows.append({"item": item, "txt": txt, "name_lbl": name_lbl, "amt_lbl": amt_lbl})

            # ---------- interactivity: hover / click a slice ----------
            pinned = {"index": None}

            def show_slice(i, event):
                pct = (values[i] / total_spent * 100) if total_spent else 0
                tooltip.set_text(f"{labels[i]}\n₹{values[i]:,.0f}  ({pct:.1f}%)")
                if event.xdata is not None and event.ydata is not None:
                    tooltip.xy = (event.xdata, event.ydata)
                tooltip.set_visible(True)

                for j, w in enumerate(wedges):
                    if j == i:
                        w.set_linewidth(3)
                        w.set_edgecolor("#111827")
                    else:
                        w.set_linewidth(1.5)
                        w.set_edgecolor("white")

                highlight_legend(i)
                pie_canvas.draw_idle()

            def clear_slice():
                tooltip.set_visible(False)
                for w in wedges:
                    w.set_linewidth(1.5)
                    w.set_edgecolor("white")
                highlight_legend(None)
                pie_canvas.draw_idle()

            def find_wedge(event):
                for i, w in enumerate(wedges):
                    contains, _ = w.contains(event)
                    if contains:
                        return i
                return None

            def on_move(event):
                if pinned["index"] is not None:
                    return
                if event.inaxes != ax:
                    clear_slice()
                    return
                i = find_wedge(event)
                if i is None:
                    clear_slice()
                else:
                    show_slice(i, event)

            def on_click(event):
                if event.inaxes != ax:
                    pinned["index"] = None
                    clear_slice()
                    return
                i = find_wedge(event)
                if i is None:
                    pinned["index"] = None
                    clear_slice()
                elif pinned["index"] == i:
                    pinned["index"] = None
                    clear_slice()
                else:
                    pinned["index"] = i
                    show_slice(i, event)

            def on_leave(event):
                if pinned["index"] is None:
                    clear_slice()

            pie_canvas.mpl_connect("motion_notify_event", on_move)
            pie_canvas.mpl_connect("button_press_event", on_click)
            pie_canvas.mpl_connect("figure_leave_event", on_leave)

        # =========================================================
        # LEFT BOTTOM: INCOME VS EXPENSE
        # =========================================================
        bar_card = Frame(
            deck,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=18,
            pady=16
        )
        bar_card.grid(row=1, column=0, padx=8, pady=8, sticky="nsew")

        Label(
            bar_card,
            text="Income vs Expenses (Month-wise)",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 10))

        has_monthly_data = month_labels != ["No Data"] and (any(income_series) or any(expense_series))

        if not has_monthly_data:
            Label(
                bar_card,
                text="No monthly data available",
                font=("Segoe UI", 11),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(expand=True)
        else:
            fig2, ax2 = plt.subplots(figsize=(4.6, 2.8), facecolor=BG_CANVAS)
            ax2.set_facecolor(BG_CANVAS)

            months = month_labels
            income_vals = income_series
            expense_vals = expense_series

            x = list(range(len(months)))
            width = 0.35

            ax2.bar(
                [i - width/2 for i in x],
                income_vals,
                width=width,
                color=COLOR_INCOME,
                label="Income"
            )
            ax2.bar(
                [i + width/2 for i in x],
                expense_vals,
                width=width,
                color=COLOR_EXPENSE,
                label="Expense"
            )

            ax2.set_xticks(x)
            ax2.set_xticklabels(months)
            ax2.set_ylabel("Amount (₹)", color=TEXT_MUTED, fontsize=9)

            ax2.spines['top'].set_visible(False)
            ax2.spines['right'].set_visible(False)
            ax2.spines['left'].set_color(BORDER_COLOR)
            ax2.spines['bottom'].set_color(BORDER_COLOR)
            ax2.tick_params(colors=TEXT_MUTED, labelsize=9)

            ax2.grid(axis="y", linestyle="--", alpha=0.18)
            ax2.legend(loc="upper right", frameon=False)

            canvas2 = FigureCanvasTkAgg(fig2, master=bar_card)
            canvas2.get_tk_widget().pack(fill=BOTH, expand=True)

        # =========================================================
        # RIGHT FULL: RECENT TRANSACTIONS + BILL SNAPSHOT
        # =========================================================
        right_col = Frame(deck, bg=BG_CANVAS)
        right_col.grid(row=0, column=1, rowspan=2, padx=8, pady=8, sticky="nsew")
        right_col.columnconfigure(0, weight=1)

        # ---------------- RECENT TRANSACTIONS CARD ---------------- #
        tx_card = Frame(
            right_col,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=20,
            pady=15
        )
        tx_card.pack(fill=BOTH, expand=True)

        tx_h = Frame(tx_card, bg=BG_CANVAS)
        tx_h.pack(fill=X, pady=(0, 15))

        Label(
            tx_h,
            text="Recent Transactions",
            font=("Segoe UI", 14, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(side=LEFT)

        Label(
            tx_h,
            text="Live preview",
            font=("Segoe UI", 10, "bold"),
            fg=COLOR_ACCENT,
            bg=BG_CANVAS
        ).pack(side=RIGHT, pady=3)

        # Pull latest expenses from DB instead of static list
        cursor.execute("""
            SELECT name, category, amount, date
            FROM expenses
            ORDER BY id DESC
            LIMIT 5
        """)
        recent_rows = cursor.fetchall()

        if not recent_rows:
            Label(
                tx_card,
                text="No recent transactions yet.",
                font=("Segoe UI", 11),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=10)
        else:
            for name, category, amount, date in recent_rows:
                row = Frame(tx_card, bg=BG_CANVAS, pady=10)
                row.pack(fill=X)

                lbl_f = Frame(row, bg=BG_CANVAS)
                lbl_f.pack(side=LEFT)

                Label(
                    lbl_f,
                    text=name,
                    font=("Segoe UI", 11, "bold"),
                    fg=TEXT_MAIN,
                    bg=BG_CANVAS
                ).pack(anchor="w")

                Label(
                    lbl_f,
                    text=f"{date} • {category}",
                    font=("Segoe UI", 9),
                    fg=TEXT_MUTED,
                    bg=BG_CANVAS
                ).pack(anchor="w")

                Label(
                    row,
                    text=f"-₹{amount:,.0f}",
                    font=("Segoe UI", 12, "bold"),
                    fg=COLOR_EXPENSE,
                    bg=BG_CANVAS
                ).pack(side=RIGHT, pady=5)

                sep = Frame(tx_card, bg=BORDER_COLOR, height=1)
                sep.pack(fill=X)

        # ---------------- UPCOMING BILL SNAPSHOT CARD ---------------- #
        bill_card = Frame(
            right_col,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=18,
            pady=15
        )
        bill_card.pack(fill=X, pady=(16, 0))

        Label(
            bill_card,
            text="Upcoming Bills Snapshot",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            bill_card,
            text=next_due_bill_text,
            font=("Segoe UI", 11),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(8, 4))

        if overdue_bills_count > 0:
            Label(
                bill_card,
                text=f"⚠ {overdue_bills_count} overdue bill(s) need attention",
                font=("Segoe UI", 10, "bold"),
                fg=COLOR_DANGER,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(2, 0))
        else:
            Label(
                bill_card,
                text="All bills are under control.",
                font=("Segoe UI", 10),
                fg=COLOR_INCOME,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(2, 0))

    # PANEL 2: EXPENSES MANAGER (Ref: pg 2_2.png)
    def load_expenses(self):
        self.highlight_sidebar("💳 Expenses")
        self.clean_view()

        if not hasattr(self, "expense_sort_state"):
            self.expense_sort_state = {"column": "date", "reverse": True}

        # ---------------- HEADER ---------------- #
        header = Frame(content_frame, bg=BG_CANVAS)
        header.pack(fill=X, pady=(0, 10))
        Label(header, text="Expenses", font=("Segoe UI", 22, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")

        meta_label = Label(content_frame, font=("Segoe UI", 11), fg=TEXT_MUTED, bg=BG_CANVAS)
        meta_label.pack(anchor="w", pady=(0, 15))

        # ---------------- FILTER BAR ---------------- #
        filter_bar = Frame(content_frame, bg=BG_CANVAS)
        filter_bar.pack(fill=X, pady=10)

        search_var = StringVar()
        search_entry = Entry(
            filter_bar, textvariable=search_var, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0,
            highlightthickness=1, highlightbackground=BORDER_COLOR, font=("Segoe UI", 10), width=24
        )
        search_entry.pack(side=LEFT, ipady=6, ipadx=10)

        placeholder = "🔍 Search transactions..."

        def on_focus_in(event):
            if search_var.get() == placeholder:
                search_entry.delete(0, END)
                search_entry.configure(fg=TEXT_MAIN)

        def on_focus_out(event):
            if not search_var.get().strip():
                search_entry.configure(fg=TEXT_MUTED)
                search_var.set(placeholder)

        search_entry.insert(0, placeholder)
        search_entry.configure(fg=TEXT_MUTED)
        search_entry.bind("<FocusIn>", on_focus_in)
        search_entry.bind("<FocusOut>", on_focus_out)

        Label(filter_bar, text="Category", bg=BG_CANVAS, fg=TEXT_MUTED, font=("Segoe UI", 10)).pack(side=LEFT, padx=(20, 5))

        cursor.execute("SELECT DISTINCT category FROM expenses WHERE category IS NOT NULL AND category != '' ORDER BY category")
        raw_categories = [row[0] for row in cursor.fetchall()]
        seen_cat_lower = {}
        for c in raw_categories:
            seen_cat_lower.setdefault(c.strip().lower(), c.strip())
        categories = ["All"] + sorted(seen_cat_lower.values())

        self.expense_category_var = StringVar(value="All")
        category_dropdown = ttk.Combobox(
            filter_bar, textvariable=self.expense_category_var, values=categories,
            state="readonly", width=18, font=("Segoe UI", 10)
        )
        category_dropdown.pack(side=LEFT, padx=5)

        btn_add_expense = Button(filter_bar, text="➕ Add Expense", bg=TEXT_MAIN, fg=BG_CANVAS, bd=0, font=("Segoe UI", 10, "bold"), padx=12, command=self.add_expense_popup)
        btn_add_expense.pack(side=RIGHT, padx=5)
        self.add_hover_effect(btn_add_expense)

        btn_import_csv = Button(filter_bar, text="📥 Import CSV", bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10, "bold"), padx=12, command=self.import_csv_logic)
        btn_import_csv.pack(side=RIGHT, padx=5)
        self.add_hover_effect(btn_import_csv)

        btn_export_expenses = Button(filter_bar, text="📤 Export CSV", bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10, "bold"), padx=12, command=self.export_expenses_csv)
        btn_export_expenses.pack(side=RIGHT, padx=5)
        self.add_hover_effect(btn_export_expenses)

        # ---------------- TABLE ---------------- #
        table_f = Frame(content_frame, bg=BG_CANVAS)
        table_f.pack(fill=BOTH, expand=True, pady=15)

        columns = ("ID", "Name", "Category", "Amount", "Date")
        tree = ttk.Treeview(table_f, columns=columns, show="headings")
        col_labels = {"ID": "#", "Name": "EXPENSE", "Category": "CATEGORY", "Amount": "AMOUNT", "Date": "DATE"}
        col_widths = {"ID": 50, "Name": 300, "Category": 150, "Amount": 150, "Date": 150}
        col_anchor = {"ID": CENTER, "Name": W, "Category": CENTER, "Amount": E, "Date": CENTER}
        sort_keys = {"ID": "id", "Name": "name", "Category": "category", "Amount": "amount", "Date": "date"}

        def sort_by(col_key):
            if self.expense_sort_state["column"] == col_key:
                self.expense_sort_state["reverse"] = not self.expense_sort_state["reverse"]
            else:
                self.expense_sort_state["column"] = col_key
                self.expense_sort_state["reverse"] = False
            refresh_table()

        for col in columns:
            tree.heading(col, text=col_labels[col], command=lambda c=sort_keys[col]: sort_by(c))
            tree.column(col, width=col_widths[col], anchor=col_anchor[col])

        tree.pack(fill=BOTH, expand=True)

        row_id_map = {}

        def refresh_table():
            tree.delete(*tree.get_children())
            row_id_map.clear()

            raw_search = search_var.get().strip()
            search_term = "" if raw_search == placeholder else raw_search.lower()
            category_filter = self.expense_category_var.get()

            sort_col = self.expense_sort_state["column"]
            order = "DESC" if self.expense_sort_state["reverse"] else "ASC"
            query = f"SELECT id, name, category, amount, date FROM expenses ORDER BY {sort_col} {order}"
            cursor.execute(query)
            rows = cursor.fetchall()

            total_shown = 0
            count_shown = 0

            for row in rows:
                rid, name, category, amount, date = row

                if category_filter != "All" and (category or "").strip().lower() != category_filter.strip().lower():
                    continue
                if search_term and search_term not in (name or "").lower() and search_term not in (category or "").lower():
                    continue

                iid = tree.insert("", END, values=(rid, name, category, f"₹{amount:,.2f}", date))
                row_id_map[iid] = rid
                total_shown += 1
                count_shown += amount

            meta_label.config(text=f"{total_shown} transaction{'s' if total_shown != 1 else ''} · ₹{count_shown:,.2f} shown")

        search_var.trace_add("write", lambda *args: refresh_table())
        category_dropdown.bind("<<ComboboxSelected>>", lambda e: refresh_table())

        def on_double_click(event):
            selected = tree.selection()
            if not selected:
                return
            expense_id = row_id_map.get(selected[0])
            if expense_id is not None:
                self.add_expense_popup(expense_id=expense_id)

        tree.bind("<Double-1>", on_double_click)

        def delete_selected_expense():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select an expense to delete.")
                return

            expense_id = row_id_map.get(selected[0])
            if expense_id is None:
                return

            confirm = messagebox.askyesno("Delete Expense", "Are you sure you want to delete this expense?")
            if confirm:
                try:
                    cursor.execute("DELETE FROM expenses WHERE id=?", (expense_id,))
                    conn.commit()
                except sqlite3.Error as e:
                    messagebox.showerror("Database Error", f"Could not delete this expense:\n{e}")
                    return
                refresh_table()
                messagebox.showinfo("Deleted", "Expense deleted successfully!")

        refresh_table()

        # ---------------- ACTIONS ---------------- #
        action_row = Frame(content_frame, bg=BG_CANVAS)
        action_row.pack(fill=X, pady=(0, 10))

        Label(
            action_row,
            text="Tip: double-click a row to edit it",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(side=LEFT)

        Button(
            action_row,
            text="🗑 Delete Selected",
            bg=COLOR_DANGER,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=delete_selected_expense
        ).pack(side=RIGHT, padx=6)
    def get_income_statistics(self):
        """
        Returns (this_month_total, average_monthly, growth_pct) computed
        live from the income table. growth_pct compares this month to
        last month; average_monthly is the mean across every month that
        has at least one income entry.
        """
        cursor.execute("SELECT amount, date FROM income")
        rows = cursor.fetchall()

        month_totals = {}
        for amount, dt in rows:
            try:
                month_key = datetime.strptime(dt, "%d-%m-%Y").strftime("%m-%Y")
            except (TypeError, ValueError):
                continue
            month_totals[month_key] = month_totals.get(month_key, 0) + amount

        this_month_key = datetime.now().strftime("%m-%Y")
        this_month_total = month_totals.get(this_month_key, 0)

        average_monthly = (sum(month_totals.values()) / len(month_totals)) if month_totals else 0

        last_month_date = datetime.now().replace(day=1) - timedelta(days=1)
        last_month_key = last_month_date.strftime("%m-%Y")
        last_month_total = month_totals.get(last_month_key, 0)

        if last_month_total > 0:
            growth_pct = (this_month_total - last_month_total) / last_month_total * 100
        else:
            growth_pct = 0 if this_month_total == 0 else 100

        return this_month_total, average_monthly, growth_pct

    def load_income(self):
        self.highlight_sidebar("💰 Income")
        self.clean_view()

        if not hasattr(self, "income_sort_state"):
            self.income_sort_state = {"column": "date", "reverse": True}

        Label(content_frame, text="Income Manager", font=("Segoe UI", 20, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")
        Label(content_frame, text="Track all your income sources", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(2, 14))

        # ---------------- FILTER + ACTION BAR ---------------- #
        top = Frame(content_frame, bg=BG_CANVAS)
        top.pack(fill=X)

        btn_add_income = Button(top, text="+ Add Income", bg=COLOR_INCOME, fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=12, pady=6, command=self.add_income_popup)
        btn_add_income.pack(side=RIGHT)
        self.add_hover_effect(btn_add_income)

        btn_export_income = Button(top, text="📤 Export CSV", bg=BG_SIDEBAR, fg=TEXT_MAIN, font=("Segoe UI", 9, "bold"), bd=0, padx=12, pady=6, command=self.export_income_csv)
        btn_export_income.pack(side=RIGHT, padx=(0, 8))
        self.add_hover_effect(btn_export_income)

        search_var = StringVar()
        search_entry = Entry(top, textvariable=search_var, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR, font=("Segoe UI", 9), width=22)
        search_entry.pack(side=LEFT, ipady=5, ipadx=8)

        placeholder = "🔍 Search income..."

        def on_focus_in(event):
            if search_var.get() == placeholder:
                search_entry.delete(0, END)
                search_entry.configure(fg=TEXT_MAIN)

        def on_focus_out(event):
            if not search_var.get().strip():
                search_entry.configure(fg=TEXT_MUTED)
                search_var.set(placeholder)

        search_entry.insert(0, placeholder)
        search_entry.configure(fg=TEXT_MUTED)
        search_entry.bind("<FocusIn>", on_focus_in)
        search_entry.bind("<FocusOut>", on_focus_out)

        Label(top, text="Source", bg=BG_CANVAS, fg=TEXT_MUTED, font=("Segoe UI", 9)).pack(side=LEFT, padx=(16, 5))
        cursor.execute("SELECT DISTINCT source FROM income WHERE source IS NOT NULL AND source != '' ORDER BY source")
        raw_sources = [row[0] for row in cursor.fetchall()]
        seen_lower = {}
        for s in raw_sources:
            seen_lower.setdefault(s.strip().lower(), s.strip())
        source_options = ["All"] + sorted(seen_lower.values())
        self.income_source_var = StringVar(value="All")
        source_dropdown = ttk.Combobox(top, textvariable=self.income_source_var, values=source_options, state="readonly", width=14, font=("Segoe UI", 9))
        source_dropdown.pack(side=LEFT, padx=5)

        # ---------------- LIVE MONTHLY STAT CARDS ---------------- #
        cursor.execute("SELECT IFNULL(SUM(amount),0) FROM income")
        total = cursor.fetchone()[0]
        this_month_total, average_monthly, growth_pct = self.get_income_statistics()

        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=14)

        cards = [
            ("Total Income", f"₹{total:,.0f}", COLOR_INCOME),
            ("This Month", f"₹{this_month_total:,.0f}", COLOR_ACCENT),
            ("Average / Month", f"₹{average_monthly:,.0f}", COLOR_SAVINGS),
            ("Growth vs Last Month", f"{'+' if growth_pct >= 0 else ''}{growth_pct:,.0f}%", COLOR_INCOME if growth_pct >= 0 else COLOR_DANGER),
        ]

        for title, value, color in cards:
            card = Frame(summary, bg=BG_CANVAS, padx=14, pady=12, highlightbackground=BORDER_COLOR, highlightthickness=1)
            card.pack(side=LEFT, expand=True, fill=X, padx=5)
            Label(card, text=title, bg=BG_CANVAS, fg=TEXT_MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            Label(card, text=value, bg=BG_CANVAS, fg=color, font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(4, 0))

        # ---------------- TABLE ---------------- #
        columns = ("ID", "Source", "Amount", "Date")
        tree = ttk.Treeview(content_frame, columns=columns, show="headings", height=11)
        col_labels = {"ID": "#", "Source": "SOURCE", "Amount": "AMOUNT", "Date": "DATE"}
        col_widths = {"ID": 50, "Source": 350, "Amount": 180, "Date": 180}
        col_anchor = {"ID": CENTER, "Source": W, "Amount": CENTER, "Date": CENTER}
        sort_keys = {"ID": "id", "Source": "source", "Amount": "amount", "Date": "date"}

        def sort_by(col_key):
            if self.income_sort_state["column"] == col_key:
                self.income_sort_state["reverse"] = not self.income_sort_state["reverse"]
            else:
                self.income_sort_state["column"] = col_key
                self.income_sort_state["reverse"] = False
            refresh_table()

        for col in columns:
            tree.heading(col, text=col_labels[col], command=lambda c=sort_keys[col]: sort_by(c))
            tree.column(col, width=col_widths[col], anchor=col_anchor[col])

        row_id_map = {}

        def refresh_table():
            tree.delete(*tree.get_children())
            row_id_map.clear()

            raw_search = search_var.get().strip()
            search_term = "" if raw_search == placeholder else raw_search.lower()
            source_filter = self.income_source_var.get()

            sort_col = self.income_sort_state["column"]
            order = "DESC" if self.income_sort_state["reverse"] else "ASC"
            cursor.execute(f"SELECT id, source, amount, date FROM income ORDER BY {sort_col} {order}")

            for rid, source, amount, date in cursor.fetchall():
                if source_filter != "All" and (source or "").strip().lower() != source_filter.strip().lower():
                    continue
                if search_term and search_term not in (source or "").lower():
                    continue
                iid = tree.insert("", END, values=(rid, source, f"₹{amount:,.2f}", date))
                row_id_map[iid] = rid

        search_var.trace_add("write", lambda *args: refresh_table())
        source_dropdown.bind("<<ComboboxSelected>>", lambda e: refresh_table())

        def on_double_click(event):
            selected = tree.selection()
            if not selected:
                return
            income_id = row_id_map.get(selected[0])
            if income_id is not None:
                self.add_income_popup(income_id=income_id)

        tree.bind("<Double-1>", on_double_click)

        refresh_table()
        # fill=X only (not expand vertically) so the table ends shortly
        # after the last row instead of stretching to fill leftover
        # blank space when there are few records.
        tree.pack(fill=X, pady=10)

        # ---------------- ACTIONS ---------------- #
        action_row = Frame(content_frame, bg=BG_CANVAS)
        action_row.pack(fill=X, pady=(0, 10))

        Label(action_row, text="Tip: double-click a row to edit it", font=("Segoe UI", 8), fg=TEXT_MUTED, bg=BG_CANVAS).pack(side=LEFT)

        def delete_selected_income():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select an income entry to delete.")
                return
            income_id = row_id_map.get(selected[0])
            if income_id is None:
                return
            confirm = messagebox.askyesno("Delete Income", "Are you sure you want to delete this income entry?")
            if confirm:
                try:
                    cursor.execute("DELETE FROM income WHERE id=?", (income_id,))
                    conn.commit()
                except sqlite3.Error as e:
                    messagebox.showerror("Database Error", f"Could not delete this income entry:\n{e}")
                    return
                refresh_table()
                messagebox.showinfo("Deleted", "Income entry deleted successfully!")

        Button(
            action_row,
            text="🗑 Delete Selected",
            bg=COLOR_DANGER,
            fg="white",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
            command=delete_selected_income
        ).pack(side=RIGHT, padx=6)

    def load_savings_goals(self):
        self.highlight_sidebar("🎯 Savings Goals")
        self.clean_view()

        # ---------------- PAGE HEADER ---------------- #
        Label(
            content_frame,
            text="Savings Goals",
            font=("Segoe UI", 20, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            content_frame,
            text="Track your goals, contributions and progress in one place",
            font=("Segoe UI", 10),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 14))

        # ---------------- FETCH GOAL DATA ---------------- #
        cursor.execute("""
            SELECT id, goal_name, target_amount, saved_amount, target_date
            FROM savings_goals
            ORDER BY id DESC
        """)
        goals = cursor.fetchall()

        total_goals = len(goals)
        total_target = sum(row[2] for row in goals) if goals else 0
        total_saved = sum(row[3] for row in goals) if goals else 0
        overall_progress = (total_saved / total_target * 100) if total_target > 0 else 0

        # ---------------- SUMMARY METRICS ROW ---------------- #
        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=(0, 25))

        summary_cards = [
            ("TOTAL GOALS", str(total_goals), COLOR_ACCENT),
            ("TOTAL TARGET", f"₹{total_target:,.0f}", TEXT_MAIN),
            ("TOTAL SAVED", f"₹{total_saved:,.0f}", COLOR_INCOME),
            ("PROGRESS", f"{overall_progress:.1f}%", COLOR_SAVINGS)
        ]

        for title, value, color in summary_cards:
            card = Frame(
                summary,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=14,
                pady=12
            )
            card.pack(side=LEFT, expand=True, fill=X, padx=5)

            Label(
                card,
                text=title,
                font=("Segoe UI", 8, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                card,
                text=value,
                font=("Segoe UI", 16, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(4, 0))

        # ---------------- ACTION BAR ---------------- #
        action_bar = Frame(content_frame, bg=BG_CANVAS)
        action_bar.pack(fill=X, pady=(0, 18))

        btn_add_goal = Button(
            action_bar,
            text="➕ Add Goal",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            command=self.add_goal_popup
        )
        btn_add_goal.pack(side=RIGHT)
        self.add_hover_effect(btn_add_goal)

        Button(
            action_bar,
            text="📤 Export CSV",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            command=self.export_goals_csv
        ).pack(side=RIGHT, padx=(0, 8))

        # ---------------- GOALS LIST SECTION ---------------- #
        goals_wrap = Frame(content_frame, bg=BG_CANVAS)
        goals_wrap.pack(fill=BOTH, expand=True)

        if not goals:
            empty_card = Frame(
                goals_wrap,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=25,
                pady=35
            )
            empty_card.pack(fill=X, pady=10)

            Label(
                empty_card,
                text="No savings goals yet",
                font=("Segoe UI", 16, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack()

            Label(
                empty_card,
                text="Create your first goal to start tracking savings progress.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(pady=(8, 0))

            return

        # Render each goal as a clean premium card
        for goal_id, goal_name, target_amount, saved_amount, deadline in goals:
            remaining = max(target_amount - saved_amount, 0)
            progress = (saved_amount / target_amount * 100) if target_amount > 0 else 0

            goal_card = Frame(
                goals_wrap,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=14,
                pady=12
            )
            goal_card.pack(fill=X, pady=6)

            # ---- Top row ----
            top = Frame(goal_card, bg=BG_CANVAS)
            top.pack(fill=X)

            left = Frame(top, bg=BG_CANVAS)
            left.pack(side=LEFT, fill=X, expand=True)

            Label(
                left,
                text=goal_name + ("  🎉 Goal Reached!" if progress >= 100 else ""),
                font=("Segoe UI", 12, "bold"),
                fg=TEXT_MAIN if progress < 100 else COLOR_INCOME,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                left,
                text=f"Deadline: {deadline if deadline else 'Not set'}",
                font=("Segoe UI", 9),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(3, 0))

            right = Frame(top, bg=BG_CANVAS)
            right.pack(side=RIGHT)

            Button(
                right,
                text="➕ Add Money",
                bg=COLOR_INCOME,
                fg="white",
                bd=0,
                font=("Segoe UI", 9, "bold"),
                padx=12,
                pady=6,
                command=lambda gid=goal_id: self.add_money_popup(gid)
            ).pack(side=LEFT, padx=4)

            Button(
                right,
                text="✏ Edit",
                bg=BG_SIDEBAR,
                fg=TEXT_MAIN,
                bd=0,
                font=("Segoe UI", 9, "bold"),
                padx=12,
                pady=6,
                command=lambda gid=goal_id: self.add_goal_popup(goal_id=gid)
            ).pack(side=LEFT, padx=4)

            Button(
                right,
                text="🗑 Delete",
                bg=BG_SIDEBAR,
                fg=TEXT_MAIN,
                bd=0,
                font=("Segoe UI", 9, "bold"),
                padx=12,
                pady=6,
                command=lambda gid=goal_id: self.delete_goal(gid)
            ).pack(side=LEFT, padx=4)

            # ---- Stats row ----
            stats = Frame(goal_card, bg=BG_CANVAS)
            stats.pack(fill=X, pady=(10, 6))

            stat_items = [
                ("Target", f"₹{target_amount:,.0f}", TEXT_MAIN),
                ("Saved", f"₹{saved_amount:,.0f}", COLOR_INCOME),
                ("Remaining", f"₹{remaining:,.0f}", COLOR_EXPENSE),
                ("Progress", f"{progress:.1f}%", COLOR_SAVINGS)
            ]

            for label_txt, value_txt, color in stat_items:
                box = Frame(stats, bg=BG_CANVAS)
                box.pack(side=LEFT, expand=True, fill=X)

                Label(
                    box,
                    text=label_txt,
                    font=("Segoe UI", 9, "bold"),
                    fg=TEXT_MUTED,
                    bg=BG_CANVAS
                ).pack(anchor="w")

                Label(
                    box,
                    text=value_txt,
                    font=("Segoe UI", 12, "bold"),
                    fg=color,
                    bg=BG_CANVAS
                ).pack(anchor="w", pady=(3, 0))

            # ---- Celebration banner when goal is fully funded ----
            if progress >= 100:
                celebrate = Frame(goal_card, bg="#DCFCE7")
                celebrate.pack(fill=X, pady=(2, 10))
                Label(
                    celebrate,
                    text=f"🎉 Goal reached! You've fully funded '{goal_name}'.",
                    font=("Segoe UI", 10, "bold"),
                    fg="#15803D",
                    bg="#DCFCE7",
                    pady=8
                ).pack()

            # ---- Progress bar: thin (8-10px), Emerald/Cyan accent ----
            bar_color = "#15803D" if progress >= 100 else COLOR_ACCENT
            style_name = f"Goal{goal_id}.Horizontal.TProgressbar"
            self.ttk_style.configure(
                style_name, background=bar_color, troughcolor=BG_SIDEBAR,
                bordercolor=BG_SIDEBAR, lightcolor=bar_color, darkcolor=bar_color
            )
            ttk.Progressbar(
                goal_card,
                orient="horizontal",
                mode="determinate",
                maximum=100,
                value=min(progress, 100),
                style=style_name
            ).pack(fill=X, ipady=3, pady=(4, 0))

    def add_money_popup(self, goal_id):
        cursor.execute("""
            SELECT goal_name, target_amount, saved_amount
            FROM savings_goals
            WHERE id=?
        """, (goal_id,))
        row = cursor.fetchone()

        if not row:
            messagebox.showerror("Error", "Goal not found.")
            return

        goal_name, target_amount, saved_amount = row

        pop = Toplevel(self)
        pop.title("Add Money")
        pop.geometry("460x360")
        pop.configure(bg=BG_CANVAS)
        pop.resizable(False, False)
        pop.transient(self)
        pop.grab_set()

        main = Frame(pop, bg=BG_CANVAS, padx=28, pady=24)
        main.pack(fill=BOTH, expand=True)

        Label(
            main,
            text=f"Add Money to {goal_name}",
            font=("Segoe UI", 15, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 18))

        Label(
            main,
            text=f"Saved: ₹{saved_amount:,.0f} / Target: ₹{target_amount:,.0f}",
            font=("Segoe UI", 10),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 20))

        Label(main, text="Amount to Add (₹)", font=("Segoe UI", 10, "bold"),
            fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(0, 5))

        amount_entry = Entry(
            main,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 11),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        amount_entry.pack(fill=X, ipady=9, pady=(0, 24))

        btn_row = Frame(main, bg=BG_CANVAS)
        btn_row.pack(fill=X, side=BOTTOM, pady=(15, 0))

        def save_money():
            amount_text = amount_entry.get().strip()

            if not amount_text:
                messagebox.showwarning("Missing Amount", "Please enter an amount.")
                return

            try:
                add_amount = float(amount_text)
            except (TypeError, ValueError):
                messagebox.showwarning("Invalid Amount", "Please enter a valid numeric amount.")
                return

            if add_amount <= 0:
                messagebox.showwarning("Invalid Amount", "Amount must be greater than 0.")
                return

            new_saved = saved_amount + add_amount

            cursor.execute("""
                UPDATE savings_goals
                SET saved_amount=?
                WHERE id=?
            """, (new_saved, goal_id))
            conn.commit()

            pop.destroy()

            if new_saved >= target_amount > 0 and saved_amount < target_amount:
                messagebox.showinfo("🎉 Goal Reached!", f"Congratulations! '{goal_name}' has hit its target of ₹{target_amount:,.0f}!")
            else:
                messagebox.showinfo("Updated", f"₹{add_amount:,.0f} added to '{goal_name}'.")

            self.load_savings_goals()

        Button(
            btn_row,
            text="Cancel",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10,
            command=pop.destroy
        ).pack(side=LEFT)

        Button(
            btn_row,
            text="Add Money",
            bg=COLOR_INCOME,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10,
            command=save_money
        ).pack(side=RIGHT)

    def delete_goal(self, goal_id):
        cursor.execute("SELECT goal_name FROM savings_goals WHERE id=?", (goal_id,))
        row = cursor.fetchone()

        if not row:
            messagebox.showerror("Error", "Goal not found.")
            return

        goal_name = row[0]

        confirm = messagebox.askyesno(
            "Delete Goal",
            f"Are you sure you want to delete '{goal_name}'?"
        )

        if not confirm:
            return

        cursor.execute("DELETE FROM savings_goals WHERE id=?", (goal_id,))
        conn.commit()

        messagebox.showinfo("Deleted", f"'{goal_name}' deleted successfully.")
        self.load_savings_goals()

    def merge_duplicate_bills(self):
        """
        Auto-merges accidental duplicate bill/EMI entries: rows that share
        the same title (case-insensitive) AND the same due date get combined
        into a single row with the amounts summed. If any duplicate is
        still Unpaid, the merged row stays Unpaid (so nothing gets silently
        marked paid). Bills with the same title but a DIFFERENT due date
        (e.g. month-to-month EMI history) are left alone on purpose.
        """
        cursor.execute("SELECT id, title, amount, category, due_date, status FROM bills_emi")
        rows = cursor.fetchall()

        groups = {}
        for bid, title, amount, category, due_date, status in rows:
            key = ((title or "").strip().lower(), (due_date or "").strip())
            groups.setdefault(key, []).append(
                {"id": bid, "title": title, "amount": amount, "category": category, "due_date": due_date, "status": status}
            )

        merged_any = False
        for key, entries in groups.items():
            if len(entries) < 2:
                continue

            merged_any = True
            keep = min(entries, key=lambda e: e["id"])
            total_amount = sum(e["amount"] for e in entries)
            final_status = "Unpaid" if any(e["status"] != "Paid" for e in entries) else "Paid"
            drop_ids = [e["id"] for e in entries if e["id"] != keep["id"]]

            cursor.execute(
                "UPDATE bills_emi SET amount=?, status=? WHERE id=?",
                (total_amount, final_status, keep["id"])
            )
            if drop_ids:
                cursor.executemany("DELETE FROM bills_emi WHERE id=?", [(d,) for d in drop_ids])

        if merged_any:
            conn.commit()

    def get_bill_due_state(self, due_date, status):
        """
        Returns one of:
        'paid', 'overdue', 'due_soon', 'upcoming', 'unknown'
        """
        if status == "Paid":
            return "paid"

        if not due_date:
            return "unknown"

        try:
            due_dt = datetime.strptime(due_date, "%d-%m-%Y").date()
            today = datetime.today().date()
            diff = (due_dt - today).days

            if diff < 0:
                return "overdue"
            elif diff <= 5:
                return "due_soon"
            else:
                return "upcoming"
        except (TypeError, ValueError):
            return "unknown"
        
    def get_bill_dashboard_summary(self):
        """
        Returns:
        upcoming_unpaid_amount, overdue_count, next_due_bill_text
        """
        cursor.execute("""
            SELECT title, amount, due_date, status
            FROM bills_emi
            WHERE status='Unpaid'
        """)
        rows = cursor.fetchall()

        upcoming_unpaid_amount = 0
        overdue_count = 0
        next_due_bill = None
        next_due_date_obj = None

        for bill_name, amount, due_date, status in rows:
            state = self.get_bill_due_state(due_date, status)

            if state in ("upcoming", "due_soon", "overdue"):
                upcoming_unpaid_amount += amount

            if state == "overdue":
                overdue_count += 1

            try:
                due_dt = datetime.strptime(due_date, "%d-%m-%Y").date()
                if next_due_date_obj is None or due_dt < next_due_date_obj:
                    next_due_date_obj = due_dt
                    next_due_bill = (bill_name, amount, due_date, state)
            except (TypeError, ValueError):
                pass

        if next_due_bill:
            bill_name, amount, due_date, state = next_due_bill
            next_due_bill_text = f"{bill_name} • ₹{amount:,.0f} • Due {due_date}"
        else:
            next_due_bill_text = "No unpaid bills 🎉"

        return upcoming_unpaid_amount, overdue_count, next_due_bill_text

    def get_dashboard_totals(self):
        cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM income")
        total_income = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses")
        total_expenses = cursor.fetchone()[0] or 0

        balance = total_income - total_expenses
        return total_income, total_expenses, balance

    def get_total_savings(self):
        """Sum of saved_amount across every savings goal."""
        cursor.execute("SELECT COALESCE(SUM(saved_amount), 0) FROM savings_goals")
        return cursor.fetchone()[0] or 0

    def get_budget_remaining_current_month(self):
        """
        Budget Remaining = (sum of all budgets set for the current month)
        minus (amount actually spent this month in those categories).
        Returns 0 if no budgets have been set for the current month.
        Reuses get_budget_status_rows so the dashboard card and the
        Budget page always agree on the same numbers.
        """
        rows = self.get_budget_status_rows()
        if not rows:
            return 0
        return sum(r["remaining"] for r in rows)

    def get_current_month_balance(self):
        """Income earned this calendar month minus expenses logged this month."""
        month_key = datetime.now().strftime("%m-%Y")

        def month_total(table):
            cursor.execute(f"SELECT amount, date FROM {table}")
            total = 0
            for amount, dt in cursor.fetchall():
                try:
                    if datetime.strptime(dt, "%d-%m-%Y").strftime("%m-%Y") == month_key:
                        total += amount
                except (TypeError, ValueError):
                    continue
            return total

        return month_total("income") - month_total("expenses")
    
    def get_category_spending_data(self, month=None):
        """
        Returns [(category, total_amount), ...] sorted by highest spend first.

        month: optional 'MM-YYYY' string. When given, only expenses logged in
        that month are included (used by the dashboard's per-month pie chart).
        When None, all-time totals are returned.
        """
        cursor.execute("SELECT category, amount, date FROM expenses")
        rows = cursor.fetchall()

        category_map = {}
        for category, amount, dt in rows:
            if month:
                try:
                    if datetime.strptime(dt, "%d-%m-%Y").strftime("%m-%Y") != month:
                        continue
                except (TypeError, ValueError):
                    continue

            cat = (category or "Other").strip().title()
            category_map[cat] = category_map.get(cat, 0) + amount

        if not category_map:
            return [("No Data", 1)]

        return sorted(category_map.items(), key=lambda x: x[1], reverse=True)
    
    def get_monthly_income_expense_data(self):
        income_map = {}
        expense_map = {}

        # Income monthly totals
        cursor.execute("SELECT amount, date FROM income")
        income_rows = cursor.fetchall()

        for amount, dt in income_rows:
            try:
                month_key = datetime.strptime(dt, "%d-%m-%Y").strftime("%b %Y")
            except (TypeError, ValueError):
                continue
            income_map[month_key] = income_map.get(month_key, 0) + amount

        # Expense monthly totals
        cursor.execute("SELECT amount, date FROM expenses")
        expense_rows = cursor.fetchall()

        for amount, dt in expense_rows:
            try:
                month_key = datetime.strptime(dt, "%d-%m-%Y").strftime("%b %Y")
            except (TypeError, ValueError):
                continue
            expense_map[month_key] = expense_map.get(month_key, 0) + amount

        # Merge months from both
        all_months = set(income_map.keys()) | set(expense_map.keys())

        def month_sort_key(m):
            return datetime.strptime(m, "%b %Y")

        sorted_months = sorted(all_months, key=month_sort_key)

        # Keep only last 6 months for dashboard neatness
        sorted_months = sorted_months[-6:]

        labels = []
        income_values = []
        expense_values = []

        for m in sorted_months:
            labels.append(m)
            income_values.append(income_map.get(m, 0))
            expense_values.append(expense_map.get(m, 0))

        if not labels:
            labels = ["No Data"]
            income_values = [0]
            expense_values = [0]

        return labels, income_values, expense_values
    
    def get_recent_paid_bills(self, limit=3):
        """
        Returns recently paid bills for dashboard activity feed.
        """
        cursor.execute("""
            SELECT title, amount, due_date
            FROM bills_emi
            WHERE status='Paid'
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))
        return cursor.fetchall()

    def load_bills_emi(self):

        self.highlight_sidebar("📅 Bills & EMI")

        self.clean_view()

        self.merge_duplicate_bills()

        # ---------------- HEADER ---------------- #
        header = Frame(content_frame, bg=BG_CANVAS)
        header.pack(fill=X, pady=(0, 18))

        Label(
            header,
            text="Bills & EMI",
            font=("Segoe UI", 22, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            header,
            text="Track recurring bills, subscriptions and EMI payments",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(3, 0))

        # ---------------- ACTION BAR ---------------- #
        action_bar = Frame(content_frame, bg=BG_CANVAS)
        action_bar.pack(fill=X, pady=(0, 22))

        btn_add_bill = Button(
            action_bar,
            text="➕ Add Bill / EMI",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=16,
            pady=8,
            command=self.add_bill_popup
        )
        btn_add_bill.pack(side=RIGHT)
        self.add_hover_effect(btn_add_bill)

        # ---------------- SUMMARY CARDS ---------------- #
        cursor.execute("""
            SELECT
                COUNT(*),
                IFNULL(SUM(amount),0),
                IFNULL(SUM(CASE WHEN status='Paid' THEN amount ELSE 0 END),0),
                IFNULL(SUM(CASE WHEN status='Unpaid' THEN amount ELSE 0 END),0)
            FROM bills_emi
        """)
        total_count, total_amount, paid_amount, unpaid_amount = cursor.fetchone()

        # extra smart bill stats
        cursor.execute("""
            SELECT due_date, status
            FROM bills_emi
        """)
        all_bill_rows = cursor.fetchall()

        overdue_count = 0
        due_soon_count = 0

        for due_date, status in all_bill_rows:
            state = self.get_bill_due_state(due_date, status)
            if state == "overdue":
                overdue_count += 1
            elif state == "due_soon":
                due_soon_count += 1

        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=(0, 24))

        cards = [
            ("TOTAL BILLS", f"{total_count}", COLOR_ACCENT),
            ("UNPAID", f"₹{unpaid_amount:,.0f}", COLOR_EXPENSE),
            ("OVERDUE", f"{overdue_count}", COLOR_DANGER),
            ("DUE SOON", f"{due_soon_count}", "#D97706")
        ]

        for title, value, color in cards:
            card = Frame(
                summary,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=18,
                pady=16
            )
            card.pack(side=LEFT, expand=True, fill=X, padx=6)

            Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                card,
                text=value,
                font=("Segoe UI", 18, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(8, 0))

        # ---------------- LIST SECTION ---------------- #
        section = Frame(content_frame, bg=BG_CANVAS)
        section.pack(fill=BOTH, expand=True)

        Label(
            section,
            text="Your Bills & EMIs",
            font=("Segoe UI", 14, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 12))

        cursor.execute("""
            SELECT id, title, amount, category, due_date, status
            FROM bills_emi
            ORDER BY id DESC
        """)
        bills = cursor.fetchall()

        if not bills:
            empty_box = Frame(
                section,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=25,
                pady=25
            )
            empty_box.pack(fill=X, pady=10)

            Label(
                empty_box,
                text="No bills or EMI records yet",
                font=("Segoe UI", 14, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                empty_box,
                text="Add your first bill, subscription or EMI to start tracking due payments.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(5, 0))
            return

        for bill_id, title, amount, category, due_date, status in bills:
            card = Frame(
                section,
                bg=BG_CANVAS,
                highlightthickness=1,
                highlightbackground=BORDER_COLOR,
                padx=20,
                pady=18
            )
            card.pack(fill=X, pady=8)

            # top row
            top = Frame(card, bg=BG_CANVAS)
            top.pack(fill=X)

            Label(
                top,
                text=title,
                font=("Segoe UI", 14, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(side=LEFT)

            bill_state = self.get_bill_due_state(due_date, status)

            if bill_state == "paid":
                badge_text = "Paid"
                badge_bg = COLOR_INCOME
            elif bill_state == "overdue":
                badge_text = "Overdue"
                badge_bg = COLOR_DANGER
            elif bill_state == "due_soon":
                badge_text = "Due Soon"
                badge_bg = "#D97706"
            elif bill_state == "upcoming":
                badge_text = "Upcoming"
                badge_bg = COLOR_ACCENT
            else:
                badge_text = status
                badge_bg = TEXT_MUTED

            Label(
                top,
                text=badge_text,
                font=("Segoe UI", 9, "bold"),
                fg="white",
                bg=badge_bg,
                padx=10,
                pady=4
            ).pack(side=RIGHT)

            # middle info row
            info = Frame(card, bg=BG_CANVAS)
            info.pack(fill=X, pady=(10, 10))

            Label(
                info,
                text=f"Amount: ₹{amount:,.0f}",
                font=("Segoe UI", 10, "bold"),
                fg=COLOR_EXPENSE,
                bg=BG_CANVAS
            ).pack(side=LEFT, padx=(0, 18))

            Label(
                info,
                text=f"Category: {category if category else '—'}",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(side=LEFT, padx=(0, 18))

            Label(
                info,
                text=f"Due: {due_date if due_date else 'Not set'}",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(side=LEFT)

            # action buttons
            btn_row = Frame(card, bg=BG_CANVAS)
            btn_row.pack(fill=X)

            if status != "Paid":
                Button(
                    btn_row,
                    text="✔ Mark Paid",
                    bg=COLOR_INCOME,
                    fg="white",
                    bd=0,
                    font=("Segoe UI", 10, "bold"),
                    padx=14,
                    pady=7,
                    command=lambda bid=bill_id: self.mark_bill_paid(bid)
                ).pack(side=LEFT)
            else:
                Button(
                    btn_row,
                    text="↺ Mark Unpaid",
                    bg=BG_SIDEBAR,
                    fg=TEXT_MAIN,
                    bd=0,
                    font=("Segoe UI", 10, "bold"),
                    padx=14,
                    pady=7,
                    command=lambda bid=bill_id: self.mark_bill_unpaid(bid)
                ).pack(side=LEFT)

            Button(
                btn_row,
                text="✏ Edit",
                bg=BG_SIDEBAR,
                fg=TEXT_MAIN,
                bd=0,
                font=("Segoe UI", 10, "bold"),
                padx=14,
                pady=7,
                command=lambda bid=bill_id: self.add_bill_popup(bid)
            ).pack(side=LEFT, padx=8)

            Button(
                btn_row,
                text="🗑 Delete",
                bg=BG_SIDEBAR,
                fg=TEXT_MAIN,
                bd=0,
                font=("Segoe UI", 10, "bold"),
                padx=14,
                pady=7,
                command=lambda bid=bill_id: self.delete_bill(bid)
            ).pack(side=LEFT, padx=8)

    def add_bill_popup(self, bill_id=None):
        editing = bill_id is not None
        existing = None
        if editing:
            cursor.execute("SELECT title, amount, category, due_date FROM bills_emi WHERE id=?", (bill_id,))
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This bill/EMI no longer exists.")
                return

        # Uses the same pinned-bottom-bar popup layout as Add Recurring:
        # the Save/Cancel bar is packed BEFORE the scrollable form area,
        # so it always keeps its space and is never pushed off-screen —
        # this popup previously packed its button bar in a plain
        # (non-scrolling) Toplevel AFTER an expand=True content frame,
        # which is what let the Save button get squeezed out of view
        # whenever the form content was taller than the fixed 560px
        # popup height.
        pop, body, button_bar = self.create_form_popup(
            "Edit Bill / EMI" if editing else "Add Bill / EMI", 500, 560
        )

        Label(
            body,
            text="Edit Bill / EMI" if editing else "Add Bill / EMI",
            font=("Segoe UI", 15, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22, pady=(18, 16))

        # Title
        Label(body, text="Bill / EMI Title", font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=22)

        title_entry = Entry(
            body,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        title_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # Amount
        Label(body, text="Amount (₹)", font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=22)

        amount_entry = Entry(
            body,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        amount_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # Category
        Label(body, text="Category", font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=22)

        category_entry = Entry(
            body,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        category_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # Due date
        Label(body, text="Due Date (DD-MM-YYYY)", font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=22)

        due_row = Frame(body, bg=BG_CANVAS)
        due_row.pack(fill=X, padx=22, pady=(4, 14))

        due_entry = Entry(
            due_row,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        due_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        self.attach_date_picker(due_entry).pack(side=LEFT, padx=(6, 0), ipady=4)

        if editing:
            ex_title, ex_amount, ex_category, ex_due = existing
            title_entry.insert(0, ex_title or "")
            amount_entry.insert(0, str(ex_amount))
            category_entry.insert(0, ex_category or "Bills")
            due_entry.insert(0, ex_due or datetime.today().strftime("%d-%m-%Y"))
        else:
            category_entry.insert(0, "Bills")
            due_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))

        Label(
            body,
            text="Examples: Netflix, Laptop EMI, WiFi Bill, Phone Recharge",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22, pady=(0, 16))

        def save_bill():
            title = title_entry.get().strip()
            category = category_entry.get().strip()
            due_date = due_entry.get().strip()

            if not title:
                messagebox.showwarning("Missing Data", "Please enter a bill/EMI title.")
                return

            ok, amount = self.validate_amount_input(amount_entry.get(), "Amount")
            if not ok:
                return

            ok, due_date = self.validate_date_input(due_date)
            if not ok:
                return

            try:
                if editing:
                    cursor.execute(
                        "UPDATE bills_emi SET title=?, amount=?, category=?, due_date=? WHERE id=?",
                        (title, amount, category, due_date, bill_id)
                    )
                else:
                    cursor.execute("""
                        INSERT INTO bills_emi (title, amount, category, due_date, status)
                        VALUES (?, ?, ?, ?, 'Unpaid')
                    """, (title, amount, category, due_date))
                conn.commit()
                self.merge_duplicate_bills()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this bill:\n{e}")
                return

            pop.destroy()
            self.load_bills_emi()

        # ---------------- BUTTONS (always visible, pinned bottom bar) ----------------

        Button(
            button_bar,
            text="Cancel",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10,
            command=pop.destroy
        ).pack(side=LEFT)

        Button(
            button_bar,
            text="Save Changes" if editing else "Save Bill",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=22,
            pady=10,
            command=save_bill
        ).pack(side=RIGHT)

    
    def mark_bill_paid(self, bill_id):
        cursor.execute("SELECT title, status FROM bills_emi WHERE id=?", (bill_id,))
        row = cursor.fetchone()

        if not row:
            messagebox.showerror("Error", "Bill not found.")
            return

        title, status = row

        if status == "Paid":
            messagebox.showinfo("Already Paid", f"'{title}' is already marked as paid.")
            return

        cursor.execute("""
            UPDATE bills_emi
            SET status='Paid'
            WHERE id=?
        """, (bill_id,))
        conn.commit()

        messagebox.showinfo("Updated", f"'{title}' marked as paid.")
        self.load_bills_emi()

    def mark_bill_unpaid(self, bill_id):
        cursor.execute("SELECT title, status FROM bills_emi WHERE id=?", (bill_id,))
        row = cursor.fetchone()

        if not row:
            messagebox.showerror("Error", "Bill not found.")
            return

        title, status = row

        if status != "Paid":
            messagebox.showinfo("Already Unpaid", f"'{title}' is already marked as unpaid.")
            return

        cursor.execute("""
            UPDATE bills_emi
            SET status='Unpaid'
            WHERE id=?
        """, (bill_id,))
        conn.commit()

        messagebox.showinfo("Updated", f"'{title}' marked as unpaid.")
        self.load_bills_emi()

    def delete_bill(self, bill_id):
        cursor.execute("SELECT title FROM bills_emi WHERE id=?", (bill_id,))
        row = cursor.fetchone()

        if not row:
            messagebox.showerror("Error", "Bill not found.")
            return

        title = row[0]

        confirm = messagebox.askyesno(
            "Delete Bill",
            f"Are you sure you want to delete '{title}'?"
        )

        if not confirm:
            return

        cursor.execute("DELETE FROM bills_emi WHERE id=?", (bill_id,))
        conn.commit()

        messagebox.showinfo("Deleted", f"'{title}' deleted successfully.")
        self.load_bills_emi()
            
    def load_recurring(self):

        self.highlight_sidebar("🔁 Recurring")

        self.clean_view()

        self.process_due_recurring(silent=True)

        Label(
            content_frame,
            text="Recurring Transactions",
            font=("Segoe UI", 22, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            content_frame,
            text="Manage repeated income and expenses like subscriptions, rent, bills and salary",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 20))

        

    # Top bar
        top_bar = Frame(content_frame, bg=BG_CANVAS)
        top_bar.pack(fill=X, pady=(0, 20))

        Button(
            top_bar,
            text="Process Due Transactions",
            bg=COLOR_INCOME,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            command=self.process_due_recurring
        ).pack(side=RIGHT, padx=(10, 0))

        Button(
            top_bar,
            text="+ Add Recurring",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            command=self.add_recurring_popup
        ).pack(side=RIGHT)

        Button(
            top_bar,
            text="📤 Export CSV",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            command=self.export_recurring_csv
        ).pack(side=RIGHT, padx=(10, 0))



    # Summary
        cursor.execute("SELECT COUNT(*) FROM recurring_transactions")
        total_recurring = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM recurring_transactions WHERE txn_type='expense'")
        recurring_expense_count = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM recurring_transactions WHERE txn_type='income'")
        recurring_income_count = cursor.fetchone()[0]

        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=(0, 20))

        cards = [
            ("TOTAL RECURRING", str(total_recurring), COLOR_ACCENT),
            ("RECURRING EXPENSES", str(recurring_expense_count), COLOR_EXPENSE),
            ("RECURRING INCOME", str(recurring_income_count), COLOR_INCOME)
        ]

        for title, value, color in cards:
            card = Frame(
                summary,
                bg=BG_CANVAS,
                padx=20,
                pady=18,
                highlightbackground=BORDER_COLOR,
                highlightthickness=1
            )
            card.pack(side=LEFT, expand=True, fill=X, padx=6)

            Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                card,
                text=value,
                font=("Segoe UI", 18, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(6, 0))

    # Table
        table_frame = Frame(content_frame, bg=BG_CANVAS)
        table_frame.pack(fill=BOTH, expand=True, pady=10)

        tree = ttk.Treeview(
            table_frame,
            columns=("ID", "Title", "Type", "Category", "Amount", "Frequency", "Next Due"),
            show="headings"
        )

        tree.heading("ID", text="#")
        tree.heading("Title", text="TITLE")
        tree.heading("Type", text="TYPE")
        tree.heading("Category", text="CATEGORY")
        tree.heading("Amount", text="AMOUNT")
        tree.heading("Frequency", text="FREQUENCY")
        tree.heading("Next Due", text="NEXT DUE DATE")

        tree.column("ID", width=50, anchor=CENTER)
        tree.column("Title", width=260, anchor=W)
        tree.column("Type", width=100, anchor=CENTER)
        tree.column("Category", width=140, anchor=CENTER)
        tree.column("Amount", width=130, anchor=E)
        tree.column("Frequency", width=120, anchor=CENTER)
        tree.column("Next Due", width=160, anchor=CENTER)

        cursor.execute("""
            SELECT id, title, txn_type, category, amount, frequency, next_due_date
            FROM recurring_transactions
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()

        for idx, row in enumerate(rows, 1):
            tree.insert(
                "",
                END,
                values=(
                    idx,
                    f"  {row[1]}",
                    row[2].title(),
                    row[3] if row[3] else "-",
                    f"₹{row[4]:,.2f}",
                    row[5].title(),
                    row[6]
                )
            )

        tree.pack(fill=BOTH, expand=True)

        sub_tree, sub_card = self.build_subscriptions_section(content_frame)
        self.build_subscription_action_bar(sub_card, sub_tree)

        rem_tree, reminder_card = self.build_reminders_section(content_frame)
        self.build_reminder_action_bar(reminder_card, rem_tree)

    def build_reminders_section(self, parent):
        reminder_card = Frame(
            parent,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=15,
            pady=15
        )
        reminder_card.pack(fill=BOTH, expand=True, pady=12)

        Label(
            reminder_card,
            text="Reminders",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 10))

        rem_tree = ttk.Treeview(
            reminder_card,
            columns=("ID", "Title", "Due Date", "Status"),
            show="headings",
            height=7
        )
        rem_tree.heading("ID", text="#")
        rem_tree.heading("Title", text="REMINDER")
        rem_tree.heading("Due Date", text="DUE DATE")
        rem_tree.heading("Status", text="STATUS")

        rem_tree.column("ID", width=50, anchor=CENTER)
        rem_tree.column("Title", width=260, anchor=W)
        rem_tree.column("Due Date", width=140, anchor=CENTER)
        rem_tree.column("Status", width=120, anchor=CENTER)

        cursor.execute("SELECT id, title, due_date, status FROM reminders ORDER BY id DESC")
        rows = cursor.fetchall()

        for idx, row in enumerate(rows, 1):
            rem_tree.insert("", END, values=(idx, row[1], row[2], row[3]))

        rem_tree.pack(fill=BOTH, expand=True)

        return rem_tree, reminder_card
    
    def build_reminder_action_bar(self, parent, rem_tree):
        action_row = Frame(parent, bg=BG_CANVAS)
        action_row.pack(fill=X, pady=(10, 0))

        def mark_selected_reminder_done():
            selected = rem_tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a reminder to mark as done.")
                return

            item = rem_tree.item(selected[0])
            values = item["values"]
            if not values:
                return

            display_id = values[0]

            cursor.execute("SELECT id, title, due_date, status FROM reminders ORDER BY id DESC")
            rows = cursor.fetchall()

            if display_id < 1 or display_id > len(rows):
                messagebox.showerror("Action Error", "Could not map selected reminder.")
                return

            actual_db_id = rows[display_id - 1][0]

            cursor.execute("UPDATE reminders SET status='Done' WHERE id=?", (actual_db_id,))
            conn.commit()
            self.load_recurring()

        def delete_selected_reminder():
            selected = rem_tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a reminder to delete.")
                return

            item = rem_tree.item(selected[0])
            values = item["values"]
            if not values:
                return

            display_id = values[0]

            cursor.execute("SELECT id, title, due_date, status FROM reminders ORDER BY id DESC")
            rows = cursor.fetchall()

            if display_id < 1 or display_id > len(rows):
                messagebox.showerror("Delete Error", "Could not map selected reminder.")
                return

            actual_db_id = rows[display_id - 1][0]

            confirm = messagebox.askyesno(
                "Delete Reminder",
                "Are you sure you want to delete the selected reminder?"
            )
            if not confirm:
                return

            cursor.execute("DELETE FROM reminders WHERE id=?", (actual_db_id,))
            conn.commit()
            self.load_recurring()

        Button(
            action_row,
            text="✔ Mark Done",
            bg=COLOR_INCOME,
            fg=BG_CANVAS,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=mark_selected_reminder_done
        ).pack(side=RIGHT, padx=6)

        Button(
            action_row,
            text="🗑 Delete Selected",
            bg=COLOR_DANGER,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=delete_selected_reminder
        ).pack(side=RIGHT, padx=6)

    def load_reminders(self):

        self.highlight_sidebar("⏰ Reminders")

        self.clean_view()

        Label(
            content_frame,
            text="Smart Reminders",
            font=("Segoe UI", 22, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            content_frame,
            text="Upcoming recurring payments, subscriptions, salary reminders and overdue alerts",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 20))

        today = datetime.today().date()

        cursor.execute("""
            SELECT id, title, amount, category, txn_type, frequency, next_due_date
            FROM recurring_transactions
            ORDER BY next_due_date ASC
        """)
        rows = cursor.fetchall()

        overdue = []
        due_today = []
        upcoming = []

        for row in rows:
            rid, title, amount, category, txn_type, frequency, next_due = row

            try:
                due_date = datetime.strptime(next_due, "%d-%m-%Y").date()
            except (TypeError, ValueError):
                continue

            diff = (due_date - today).days

            item = {
                "id": rid,
                "title": title,
                "amount": amount,
                "category": category if category else "Other",
                "txn_type": txn_type,
                "frequency": frequency,
                "next_due": next_due,
                "diff": diff
            }

            if diff < 0:
                overdue.append(item)
            elif diff == 0:
                due_today.append(item)
            elif 0 < diff <= 7:
                upcoming.append(item)

    # ---------------- SUMMARY CARDS ---------------- #
        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=(0, 20))

        cards = [
            ("OVERDUE", str(len(overdue)), COLOR_EXPENSE),
            ("DUE TODAY", str(len(due_today)), COLOR_ACCENT),
            ("NEXT 7 DAYS", str(len(upcoming)), COLOR_INCOME)
        ]

        for title, value, color in cards:
            card = Frame(
                summary,
                bg=BG_CANVAS,
                padx=20,
                pady=18,
                highlightbackground=BORDER_COLOR,
                highlightthickness=1
            )
            card.pack(side=LEFT, expand=True, fill=X, padx=6)

            Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

            Label(
                card,
                text=value,
                font=("Segoe UI", 18, "bold"),
                fg=color,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(6, 0))

    # ---------------- ACTION BAR ---------------- #
        action_bar = Frame(content_frame, bg=BG_CANVAS)
        action_bar.pack(fill=X, pady=(0, 15))

        Button(
            action_bar,
            text="Go to Recurring",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            command=self.load_recurring
        ).pack(side=RIGHT)

    # ---------------- REMINDER SECTIONS ---------------- #
        container = Frame(content_frame, bg=BG_CANVAS)
        container.pack(fill=BOTH, expand=True)

        def section_title(parent, text):
            Label(
                parent,
                text=text,
                font=("Segoe UI", 14, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(12, 10))

        def reminder_card(parent, item, status_color, status_text):
            card = Frame(
                parent,
                bg=BG_CANVAS,
                highlightbackground=BORDER_COLOR,
                highlightthickness=1,
                padx=18,
                pady=16
            )
            card.pack(fill=X, pady=6)

            top = Frame(card, bg=BG_CANVAS)
            top.pack(fill=X)

            Label(
                top,
                text=item["title"],
                font=("Segoe UI", 13, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(side=LEFT)

            Label(
                top,
                text=status_text,
                font=("Segoe UI", 9, "bold"),
                fg=status_color,
                bg=BG_CANVAS
            ).pack(side=RIGHT)

            meta_text = f"{item['txn_type'].title()} • {item['category']} • {item['frequency'].title()}"
            Label(
                card,
                text=meta_text,
                font=("Segoe UI", 9),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(4, 0))

            bottom = Frame(card, bg=BG_CANVAS)
            bottom.pack(fill=X, pady=(10, 0))

            Label(
                bottom,
                text=f"Amount: ₹{item['amount']:,.2f}",
                font=("Segoe UI", 10, "bold"),
                fg=TEXT_MAIN,
                bg=BG_CANVAS
            ).pack(side=LEFT)

            Label(
                bottom,
                text=f"Due: {item['next_due']}",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(side=RIGHT)

    # OVERDUE
        section_title(container, "Overdue")
        if overdue:
            for item in overdue:
                reminder_card(
                    container,
                    item,
                    COLOR_EXPENSE,
                    f"{abs(item['diff'])} day(s) overdue"
                )
        else:
            Label(
                container,
                text="No overdue recurring transactions 🎉",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

    # DUE TODAY
        section_title(container, "Due Today")
        if due_today:
            for item in due_today:
                reminder_card(
                    container,
                    item,
                    COLOR_ACCENT,
                    "Due today"
                )
        else:
            Label(
                container,
                text="Nothing due today.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

    # UPCOMING
        section_title(container, "Next 7 Days")
        if upcoming:
            for item in upcoming:
                reminder_card(
                    container,
                    item,
                    COLOR_INCOME,
                    f"Due in {item['diff']} day(s)"
                )
        else:
            Label(
                container,
                text="No upcoming recurring transactions in the next 7 days.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w")

    # ---------------- PERSONAL REMINDERS ---------------- #
        section_title(container, "Personal Reminders")

        reminder_card_frame = Frame(
            container,
            bg=BG_CANVAS,
            highlightbackground=BORDER_COLOR,
            highlightthickness=1,
            padx=16,
            pady=16
        )
        reminder_card_frame.pack(fill=BOTH, expand=True, pady=(8, 12))

        # Top bar
        top_bar = Frame(reminder_card_frame, bg=BG_CANVAS)
        top_bar.pack(fill=X, pady=(0, 10))

        Label(
            top_bar,
            text="Your reminders — bills, deadlines, tasks",
            font=("Segoe UI", 12, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(side=LEFT)

        btn_add_reminder = Button(
            top_bar,
            text="+ Add Reminder",
            bg=COLOR_PURPLE,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=7,
            command=self.add_reminder_popup
        )
        btn_add_reminder.pack(side=RIGHT)
        self.add_hover_effect(btn_add_reminder)

        rem_tree = ttk.Treeview(
            reminder_card_frame,
            columns=("ID", "Title", "Due Date", "Priority", "Notes", "Status"),
            show="headings",
            height=6
        )
        rem_tree.heading("ID", text="#")
        rem_tree.heading("Title", text="TITLE")
        rem_tree.heading("Due Date", text="DUE DATE")
        rem_tree.heading("Priority", text="PRIORITY")
        rem_tree.heading("Notes", text="NOTES")
        rem_tree.heading("Status", text="STATUS")

        rem_tree.column("ID", width=40, anchor=CENTER)
        rem_tree.column("Title", width=220, anchor=W)
        rem_tree.column("Due Date", width=110, anchor=CENTER)
        rem_tree.column("Priority", width=90, anchor=CENTER)
        rem_tree.column("Notes", width=200, anchor=W)
        rem_tree.column("Status", width=100, anchor=CENTER)

        cursor.execute("SELECT id, title, due_date, priority, notes, status FROM reminders ORDER BY id DESC")
        reminder_rows = cursor.fetchall()

        pending_count = completed_count = overdue_count = 0
        for row in reminder_rows:
            state = self.get_reminder_state(row[2], row[5])
            if state == "pending":
                pending_count += 1
            elif state == "completed":
                completed_count += 1
            elif state == "overdue":
                overdue_count += 1

        Label(
            reminder_card_frame,
            text=f"Pending: {pending_count}   •   Overdue: {overdue_count}   •   Completed: {completed_count}",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 10))

        for idx, row in enumerate(reminder_rows, 1):
            rid, r_title, r_due, r_priority, r_notes, r_status = row
            state = self.get_reminder_state(r_due, r_status)
            display_status = {"completed": "Completed", "overdue": "Overdue", "pending": "Pending"}[state]
            tag = state
            rem_tree.insert(
                "", END,
                values=(idx, r_title, r_due, r_priority or "Medium", (r_notes or "")[:40], display_status),
                tags=(tag,)
            )

        rem_tree.tag_configure("overdue", foreground=COLOR_DANGER)
        rem_tree.tag_configure("completed", foreground=COLOR_INCOME)
        rem_tree.tag_configure("pending", foreground=TEXT_MAIN)

        if not reminder_rows:
            Label(
                reminder_card_frame,
                text="No personal reminders yet. Click \"+ Add Reminder\" to create one.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=(0, 12))
        else:
            rem_tree.pack(fill=X, pady=(0, 12))

        Label(
            reminder_card_frame,
            text="Tip: double-click a reminder to edit it",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 8))

        action_row = Frame(reminder_card_frame, bg=BG_CANVAS)
        action_row.pack(fill=X)

        def _get_selected_manual_reminder_db_id():
            selected = rem_tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a reminder first.")
                return None

            values = rem_tree.item(selected[0])["values"]
            if not values:
                return None

            display_id = values[0]

            cursor.execute("SELECT id FROM reminders ORDER BY id DESC")
            rows = cursor.fetchall()

            if display_id < 1 or display_id > len(rows):
                return None

            return rows[display_id - 1][0]

        def edit_selected_reminder(event=None):
            actual_id = _get_selected_manual_reminder_db_id()
            if actual_id is None:
                return
            self.add_reminder_popup(reminder_id=actual_id)

        rem_tree.bind("<Double-1>", edit_selected_reminder)

        def mark_manual_reminder_done():
            actual_id = _get_selected_manual_reminder_db_id()
            if actual_id is None:
                return
            self.mark_reminder_complete(actual_id)

        def delete_manual_reminder():
            actual_id = _get_selected_manual_reminder_db_id()
            if actual_id is None:
                return

            if not messagebox.askyesno("Delete Reminder", "Delete selected reminder?"):
                return

            self.delete_reminder(actual_id)

        Button(
            action_row,
            text="✏ Edit",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=edit_selected_reminder
        ).pack(side=RIGHT, padx=6)

        Button(
            action_row,
            text="✔ Mark Complete",
            bg="#16A34A",
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=mark_manual_reminder_done
        ).pack(side=RIGHT, padx=6)

        Button(
            action_row,
            text="🗑 Delete",
            bg=COLOR_DANGER,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=delete_manual_reminder
        ).pack(side=RIGHT, padx=6)

    def suggest_category_from_merchant(self, merchant):
        if not merchant:
            return "Others"

        m = merchant.lower()

        food_keywords = ["swiggy", "zomato", "dominos", "pizza", "cafe", "restaurant", "burger", "kfc", "mcdonald"]
        travel_keywords = ["ola", "uber", "metro", "rail", "irctc", "fuel", "petrol", "diesel"]
        shopping_keywords = ["amazon", "flipkart", "myntra", "mall", "store", "shop"]
        health_keywords = ["apollo", "pharmacy", "hospital", "clinic", "medical", "medico"]
        bills_keywords = ["electricity", "water", "broadband", "wifi", "jio", "airtel", "bsnl"]
        entertainment_keywords = ["netflix", "spotify", "prime", "pvr", "bookmyshow"]
        education_keywords = ["udemy", "coursera", "college", "academy", "course"]

        if any(k in m for k in food_keywords):
            return "Food"
        if any(k in m for k in travel_keywords):
            return "Travel"
        if any(k in m for k in shopping_keywords):
            return "Shopping"
        if any(k in m for k in health_keywords):
            return "Health"
        if any(k in m for k in bills_keywords):
            return "Bills"
        if any(k in m for k in entertainment_keywords):
            return "Entertainment"
        if any(k in m for k in education_keywords):
            return "Education"

        return "Others"
    
    def extract_receipt_data(self, lines):
        """
        lines = OCR output list of strings
        returns merchant, amount, date
        """
        merchant = "Unknown Merchant"
        amount = 0.0
        date = datetime.today().strftime("%d-%m-%Y")

        if not lines:
            return merchant, amount, date

        # Merchant = usually first meaningful line
        for line in lines:
            clean = line.strip()
            if len(clean) >= 3:
                merchant = clean[:60]
                break

        # Date detection — done first so date digits (day/month/year) can be
        # excluded from amount detection below. Without this, a date like
        # "15/07/2026" could contribute "2026" as a false amount candidate
        # and win as the (incorrect) max value.
        text_blob = " ".join(lines)

        date_patterns = [
            r'\b\d{2}[/-]\d{2}[/-]\d{4}\b',
            r'\b\d{2}[/-]\d{2}[/-]\d{2}\b',
            r'\b\d{4}[/-]\d{2}[/-]\d{2}\b'
        ]

        found_date = None
        matched_date_text = None
        for pat in date_patterns:
            m = re.search(pat, text_blob)
            if m:
                found_date = m.group(0)
                matched_date_text = m.group(0)
                break

        if found_date:
            # normalize to dd-mm-yyyy if possible
            raw = found_date.replace("/", "-")
            parts = raw.split("-")

            try:
                if len(parts[0]) == 4:
                    # yyyy-mm-dd -> dd-mm-yyyy
                    yyyy, mm, dd = parts
                    date = f"{dd}-{mm}-{yyyy}"
                elif len(parts[2]) == 2:
                    # dd-mm-yy -> dd-mm-20yy
                    dd, mm, yy = parts
                    date = f"{dd}-{mm}-20{yy}"
                else:
                    date = raw
            except (TypeError, ValueError):
                date = datetime.today().strftime("%d-%m-%Y")

        # Amount detection — search the text with the date substring removed
        # so date digits can never be mistaken for the receipt total.
        amount_search_text = text_blob.replace(matched_date_text, " ") if matched_date_text else text_blob

        # Prefer a number that appears on a line that actually looks like a
        # total ("total", "amount", "grand total", "net payable", a currency
        # symbol, etc.) — this avoids picking up bill numbers, phone numbers,
        # or item counts, which are often larger than the real total.
        total_keyword_pattern = re.compile(r'(total|amount|amt|grand total|net payable|to pay|payable)', re.IGNORECASE)
        currency_number_pattern = re.compile(r'(?:₹|rs\.?|inr)\s*(\d+(?:[.,]\d{2})?)', re.IGNORECASE)

        amount = 0.0
        for line in lines:
            if matched_date_text and matched_date_text in line:
                continue
            if total_keyword_pattern.search(line):
                nums = re.findall(r'\d+(?:[.,]\d{2})?', line)
                nums = [float(n.replace(",", "")) for n in nums if float(n.replace(",", "")) > 0]
                if nums:
                    amount = max(nums)
                    break

        if amount == 0.0:
            currency_matches = currency_number_pattern.findall(amount_search_text)
            if currency_matches:
                amount = max(float(v.replace(",", "")) for v in currency_matches)

        if amount == 0.0:
            # Last resort: largest plausible number anywhere (excluding the date)
            amount_candidates = []
            amount_pattern = re.findall(r'\d+(?:[.,]\d{2})?', amount_search_text)
            for a in amount_pattern:
                try:
                    val = float(a.replace(",", ""))
                    if val > 10:
                        amount_candidates.append(val)
                except (TypeError, ValueError):
                    pass
            if amount_candidates:
                amount = max(amount_candidates)

        return merchant, amount, date

    def load_receipt_scanner(self):

        self.highlight_sidebar("🧾 Receipt Scanner")

        self.clean_view()

        Label(
            content_frame,
            text="Receipt Scanner",
            font=("Segoe UI", 22, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            content_frame,
            text="Scan receipt images and convert them into expense entries",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 20))

        # ---------------- MAIN WRAPPER ---------------- #
        wrapper = Frame(content_frame, bg=BG_CANVAS)
        wrapper.pack(fill=BOTH, expand=True)

        # LEFT PANEL
        left_panel = Frame(
            wrapper,
            bg=BG_CANVAS,
            highlightbackground=BORDER_COLOR,
            highlightthickness=1,
            padx=20,
            pady=20
        )
        left_panel.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))

        # RIGHT PANEL
        right_panel = Frame(
            wrapper,
            bg=BG_CANVAS,
            highlightbackground=BORDER_COLOR,
            highlightthickness=1,
            padx=20,
            pady=20,
            width=380
        )
        right_panel.pack(side=RIGHT, fill=Y)
        right_panel.pack_propagate(False)

        Label(
            left_panel,
            text="OCR Receipt Preview",
            font=("Segoe UI", 14, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 12))

        receipt_output = Text(
            left_panel,
            font=("Consolas", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            wrap=WORD,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        receipt_output.pack(fill=BOTH, expand=True)

        # ---------- extracted vars ----------
        self.receipt_merchant_var = StringVar(value="Not scanned yet")
        self.receipt_amount_var = StringVar(value="₹0.00")
        self.receipt_date_var = StringVar(value="-")
        self.receipt_category_var = StringVar(value="Others")

        Label(
            right_panel,
            text="Extracted Details",
            font=("Segoe UI", 14, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 15))

        def info_row(parent, title, var):
            row = Frame(parent, bg=BG_CANVAS)
            row.pack(fill=X, pady=8)

            Label(
                row,
                text=title,
                font=("Segoe UI", 10, "bold"),
                fg=TEXT_MUTED,
                bg=BG_CANVAS,
                width=10,
                anchor="w"
            ).pack(side=LEFT)

            Label(
                row,
                textvariable=var,
                font=("Segoe UI", 11),
                fg=TEXT_MAIN,
                bg=BG_CANVAS,
                anchor="w"
            ).pack(side=LEFT)

        info_row(right_panel, "Merchant", self.receipt_merchant_var)
        info_row(right_panel, "Amount", self.receipt_amount_var)
        info_row(right_panel, "Date", self.receipt_date_var)
        info_row(right_panel, "Category", self.receipt_category_var)

        # ---------- editable category ----------
        Label(
            right_panel,
            text="Edit Category",
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(18, 5))

        category_box = ttk.Combobox(
            right_panel,
            values=["Food", "Travel", "Shopping", "Entertainment", "Health", "Bills", "Education", "Others"],
            state="readonly",
            font=("Segoe UI", 10)
        )
        category_box.pack(fill=X, ipady=4)
        category_box.set("Others")

        # ---------- internal state ----------
        self.last_receipt_amount = 0.0
        self.last_receipt_date = datetime.today().strftime("%d-%m-%Y")
        self.last_receipt_merchant = "Unknown Merchant"

    # ---------- scan function ----------
        def scan_receipt_ui():
            filepath = filedialog.askopenfilename(
                title="Select Receipt Image",
                filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
            )

            if not filepath:
                return

            receipt_output.delete("1.0", END)
            receipt_output.insert(END, "⏳ Reading receipt...\n\n")
            self.update()

            try:
                if easyocr is None:
                    receipt_output.delete("1.0", END)
                    receipt_output.insert(
                        END,
                        "❌ Receipt scanning requires the 'easyocr' package.\n\n"
                        "Install it with:\n    pip install easyocr"
                    )
                    return
                reader = easyocr.Reader(['en'], gpu=False)
                result = reader.readtext(filepath, detail=0)

                merchant, amount, date = self.extract_receipt_data(result)
                category = self.suggest_category_from_merchant(merchant)

                self.last_receipt_merchant = merchant
                self.last_receipt_amount = amount
                self.last_receipt_date = date

                self.receipt_merchant_var.set(merchant)
                self.receipt_amount_var.set(f"₹{amount:,.2f}")
                self.receipt_date_var.set(date)
                self.receipt_category_var.set(category)
                category_box.set(category)

                receipt_output.delete("1.0", END)
                receipt_output.insert(END, "🧾 RECEIPT OCR RESULT\n")
                receipt_output.insert(END, "=" * 50 + "\n\n")

                if result:
                    for line in result:
                        receipt_output.insert(END, f"{line}\n")
                else:
                    receipt_output.insert(END, "No readable text found.")

            except Exception as e:
                receipt_output.delete("1.0", END)
                receipt_output.insert(END, f"❌ Error while scanning receipt:\n\n{str(e)}")

        # ---------- add to expenses ----------
        def add_scanned_receipt_to_expenses():
            merchant = self.last_receipt_merchant
            amount = self.last_receipt_amount
            date = self.last_receipt_date
            category = category_box.get().strip()

            if not merchant or amount <= 0:
                messagebox.showwarning("No Valid Receipt", "Please scan a valid receipt first.")
                return

            try:
                cursor.execute(
                    "INSERT INTO expenses (name, amount, category, date) VALUES (?, ?, ?, ?)",
                    (merchant, amount, category, date)
                )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this expense:\n{e}")
                return

            self.receipt_category_var.set(category)

            messagebox.showinfo(
                "Expense Added",
                f"Receipt added to expenses:\n\n{merchant}\n₹{amount:,.2f}\n{category}"
            )

        # ---------- action buttons ----------
        btn_wrap = Frame(right_panel, bg=BG_CANVAS)
        btn_wrap.pack(fill=X, pady=(25, 0))

        btn_scan = Button(
            btn_wrap,
            text="Scan Receipt",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=10,
            command=scan_receipt_ui
        )
        btn_scan.pack(fill=X, pady=5)
        self.add_hover_effect(btn_scan)

        btn_add_scanned = Button(
            btn_wrap,
            text="Add to Expenses",
            bg=COLOR_INCOME,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=10,
            command=add_scanned_receipt_to_expenses
        )
        btn_add_scanned.pack(fill=X, pady=5)
        self.add_hover_effect(btn_add_scanned)

        Button(
            btn_wrap,
            text="Open Expenses Page",
            bg=TEXT_MAIN,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=10,
            command=self.load_expenses
        ).pack(fill=X, pady=5)

    def get_expense_totals_by_category(self, month_year=None):
        """
        Returns a dict like:
        {
            "Food": 2865.0,
            "Travel": 1070.0,
            ...
        }
        month_year: optional 'MM-YYYY' filter. When given, only expenses
        logged in that month are summed (used by the budget tracker so
        "spent" reflects that month's spend, not all-time spend).
        """
        cursor.execute("SELECT category, amount, date FROM expenses")
        rows = cursor.fetchall()

        totals = {}
        for category, amount, dt in rows:
            if month_year:
                try:
                    if datetime.strptime(dt, "%d-%m-%Y").strftime("%m-%Y") != month_year:
                        continue
                except (TypeError, ValueError):
                    continue
            totals[category] = totals.get(category, 0) + float(amount or 0)

        return totals
    
    def get_budget_status_rows(self, month_year=None):
        """
        Returns a list of dicts for each budget category in the given month
        (month_year format: 'MM-YYYY'; defaults to the current month):
        [
            {
                "category": "Food",
                "budget": 5000,
                "spent": 2865,
                "remaining": 2135,
                "used_pct": 57.3,
                "is_over": False
            },
            ...
        ]
        """
        if month_year is None:
            month_year = datetime.now().strftime("%m-%Y")

        expense_totals = self.get_expense_totals_by_category(month_year)

        cursor.execute("""
            SELECT category, amount
            FROM budgets
            WHERE month_year=?
            ORDER BY category
        """, (month_year,))
        budget_rows = cursor.fetchall()

        results = []
        for category, budget_amount in budget_rows:
            budget_amount = float(budget_amount or 0)
            spent = float(expense_totals.get(category, 0))
            remaining = budget_amount - spent
            used_pct = (spent / budget_amount * 100) if budget_amount > 0 else 0
            is_over = spent > budget_amount

            results.append({
                "category": category,
                "budget": budget_amount,
                "spent": spent,
                "remaining": remaining,
                "used_pct": used_pct,
                "is_over": is_over
            })

        return results

    # PANEL 3: BUDGETS MANAGER (Ref: pg 3_2.png)
    def add_budget_popup(self, category=None, month_year=None):
        """
        category=None -> Add mode (choose category + month).
        category=<str> -> Edit mode for that category's budget in month_year
        (month_year defaults to the current month).
        """
        editing = category is not None
        month_year = month_year or datetime.now().strftime("%m-%Y")

        existing_amount = None
        if editing:
            cursor.execute("SELECT amount FROM budgets WHERE category=? AND month_year=?", (category, month_year))
            row = cursor.fetchone()
            existing_amount = row[0] if row else None

        pop, body = self.create_scroll_popup("Edit Budget" if editing else "Add Budget", 460, 480)

        Label(body, text="Edit Budget" if editing else "Add Budget", font=("Segoe UI", 16, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(20, 15))

        Label(body, text="Category", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(0, 4))

        default_categories = ["Food", "Travel", "Shopping", "Entertainment", "Health", "Education", "Bills"]
        categories = self.get_dropdown_options("expenses", "category", default_categories)

        if editing:
            category_widget = Label(body, text=category, font=("Segoe UI", 11, "bold"), fg=TEXT_MAIN, bg=BG_SIDEBAR, anchor="w", padx=10, pady=8)
            category_widget.pack(fill=X, padx=25)
            category_var = StringVar(value=category)
            other_entry = None
        else:
            category_var = StringVar(value=categories[0])
            category_menu = ttk.Combobox(body, textvariable=category_var, values=categories, state="readonly", font=("Segoe UI", 11))
            category_menu.pack(fill=X, padx=25, ipady=5)

            other_label = Label(body, text="Custom Category", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS)
            other_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)

            def category_changed(event=None):
                if category_var.get() == "Other":
                    other_label.pack(anchor="w", padx=25, pady=(14, 4))
                    other_entry.pack(fill=X, padx=25, ipady=7)
                else:
                    other_label.pack_forget()
                    other_entry.pack_forget()

            category_menu.bind("<<ComboboxSelected>>", category_changed)

        Label(body, text="Monthly Budget Amount (₹)", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(14, 4))
        amount_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        amount_entry.pack(fill=X, padx=25, ipady=7)
        if existing_amount is not None:
            amount_entry.insert(0, str(existing_amount))

        Label(body, text=f"Month: {datetime.strptime(month_year, '%m-%Y').strftime('%B %Y')}", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(14, 0))

        def save():
            final_category = category_var.get()
            if not editing and final_category == "Other":
                final_category = other_entry.get().strip().title()
            if not final_category:
                messagebox.showerror("Missing Category", "Please choose or enter a category.")
                return

            ok, amount = self.validate_amount_input(amount_entry.get(), "Budget amount")
            if not ok:
                return

            try:
                cursor.execute(
                    "INSERT OR REPLACE INTO budgets (category, amount, month_year) VALUES (?, ?, ?)",
                    (final_category, amount, month_year)
                )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this budget:\n{e}")
                return

            pop.destroy()
            self.load_budgets()

        Button(
            body,
            text="Save Changes" if editing else "Save Budget",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 11, "bold"),
            command=save
        ).pack(fill=X, padx=25, pady=25, ipady=8)

    def load_budgets(self):
        self.highlight_sidebar("💼 Budget")
        self.clean_view()

        current_month_key = datetime.now().strftime("%m-%Y")
        month_display = datetime.now().strftime("%B %Y")

        budget_title_row = Frame(content_frame, bg=BG_CANVAS)
        budget_title_row.pack(fill=X)
        Label(budget_title_row, text="Budget Tracker", font=("Segoe UI", 22, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT, anchor="w")
        btn_add_budget = Button(
            budget_title_row,
            text="➕ Add Budget",
            bg=TEXT_MAIN,
            fg=BG_CANVAS,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
            command=self.add_budget_popup
        )
        btn_add_budget.pack(side=RIGHT, padx=(8, 0))
        self.add_hover_effect(btn_add_budget)

        btn_export_budgets = Button(
            budget_title_row,
            text="📤 Export CSV",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
            command=self.export_budgets_csv
        )
        btn_export_budgets.pack(side=RIGHT)
        self.add_hover_effect(btn_export_budgets)
        Label(content_frame, text=f"{month_display} allocation planning summary", font=("Segoe UI", 11), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(2, 20))

        # ---------------- SUMMARY VALUES (current month only) ---------------- #
        budget_rows = self.get_budget_status_rows(current_month_key)
        total_budget = sum(r["budget"] for r in budget_rows)
        total_spent = sum(r["spent"] for r in budget_rows)
        remaining = total_budget - total_spent
        utilisation = (total_spent / total_budget * 100) if total_budget > 0 else 0

        summary_panel = Frame(content_frame, bg=BG_CANVAS)
        summary_panel.pack(fill=X, pady=(0, 25))

        metrics = [
            ("TOTAL BUDGET", f"₹{total_budget:,.0f}", TEXT_MUTED),
            ("TOTAL SPENT", f"₹{total_spent:,.0f}", COLOR_EXPENSE),
            ("REMAINING", f"₹{remaining:,.0f}", COLOR_INCOME if remaining >= 0 else COLOR_DANGER),
            ("UTILISATION", f"{utilisation:.0f}%", COLOR_ACCENT)
        ]

        for title, val, col in metrics:
            f = Frame(summary_panel, bg=BG_CANVAS)
            f.pack(side=LEFT, expand=True, fill=X)
            Label(f, text=title, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
            Label(f, text=val, font=("Segoe UI", 18, "bold"), fg=col, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))

        if not budget_rows:
            empty_card = Frame(content_frame, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=20, pady=30)
            empty_card.pack(fill=X, pady=10)
            Label(empty_card, text=f"No budgets set for {month_display} yet.", font=("Segoe UI", 12, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack()
            Label(empty_card, text="Click '➕ Add Budget' above to set a spending limit for a category.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(pady=(4, 0))
            return

        # ---------------- PER-CATEGORY PROGRESS LIST ---------------- #
        list_card = Frame(content_frame, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=18, pady=16)
        list_card.pack(fill=X, pady=(0, 15))

        Label(list_card, text="Category Breakdown", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 10))

        for r in sorted(budget_rows, key=lambda x: x["used_pct"], reverse=True):
            row_f = Frame(list_card, bg=BG_CANVAS)
            row_f.pack(fill=X, pady=8)

            if r["used_pct"] >= 100:
                bar_color, status_text, status_color = COLOR_DANGER, "Exceeded", COLOR_DANGER
            elif r["used_pct"] >= 90:
                bar_color, status_text, status_color = "#F97316", "Near limit", "#F97316"
            else:
                bar_color, status_text, status_color = COLOR_INCOME, "On track", COLOR_INCOME

            top_row = Frame(row_f, bg=BG_CANVAS)
            top_row.pack(fill=X)
            Label(top_row, text=r["category"], font=("Segoe UI", 11, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)
            Label(top_row, text=f" • {status_text}", font=("Segoe UI", 9, "bold"), fg=status_color, bg=BG_CANVAS).pack(side=LEFT)
            Label(top_row, text=f"₹{r['spent']:,.0f} / ₹{r['budget']:,.0f}", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(side=RIGHT)

            bar_track = Frame(row_f, bg=BORDER_COLOR, height=10)
            bar_track.pack(fill=X, pady=(6, 4))
            bar_track.pack_propagate(False)
            fill_pct = min(r["used_pct"], 100) / 100
            if fill_pct > 0:
                Frame(bar_track, bg=bar_color, height=10).place(relx=0, rely=0, relwidth=fill_pct, relheight=1)

            bottom_row = Frame(row_f, bg=BG_CANVAS)
            bottom_row.pack(fill=X)
            remaining_text = f"₹{r['remaining']:,.0f} remaining" if r["remaining"] >= 0 else f"₹{abs(r['remaining']):,.0f} over budget"
            Label(bottom_row, text=f"{r['used_pct']:.0f}% used  •  {remaining_text}", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS).pack(side=LEFT)

            Button(
                bottom_row, text="Edit", font=("Segoe UI", 9, "bold"), bd=0, bg=BG_CANVAS, fg=COLOR_ACCENT,
                cursor="hand2", command=lambda c=r["category"]: self.add_budget_popup(category=c, month_year=current_month_key)
            ).pack(side=RIGHT, padx=(8, 0))

            def delete_budget(c=r["category"]):
                confirm = messagebox.askyesno("Delete Budget", f"Remove the {month_display} budget for '{c}'?")
                if confirm:
                    cursor.execute("DELETE FROM budgets WHERE category=? AND month_year=?", (c, current_month_key))
                    conn.commit()
                    self.load_budgets()

            Button(
                bottom_row, text="Delete", font=("Segoe UI", 9, "bold"), bd=0, bg=BG_CANVAS, fg=COLOR_DANGER,
                cursor="hand2", command=delete_budget
            ).pack(side=RIGHT)

        # ---------------- BUDGET VS ACTUAL CHART CARD ---------------- #
        chart_f = Frame(
            content_frame,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=15,
            pady=15
        )
        chart_f.pack(fill=BOTH, expand=True, pady=10)

        Label(
            chart_f,
            text="Budget vs Actual Spend",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 10))

        categories = [r["category"] for r in budget_rows]
        budget_values = [r["budget"] for r in budget_rows]
        actual_values = [r["spent"] for r in budget_rows]

        fig, ax = plt.subplots(figsize=(8, 3.6), facecolor=BG_CANVAS)
        ax.set_facecolor(BG_CANVAS)

        x = list(range(len(categories)))
        width = 0.35

        ax.bar([i - width/2 for i in x], budget_values, width=width, color="#22D3EE", label="Budget")
        ax.bar([i + width/2 for i in x], actual_values, width=width, color="#FB923C", label="Actual")

        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=20, ha="right", fontsize=9, color=TEXT_MUTED)

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(BORDER_COLOR)
        ax.spines['bottom'].set_color(BORDER_COLOR)
        ax.tick_params(axis='y', colors=TEXT_MUTED, labelsize=9)
        ax.legend(frameon=False)

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=chart_f)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=BOTH, expand=True)

    def get_report_summary(self):
        """
        Returns:
        total_income, total_expenses, total_savings, savings_rate
        """
        cursor.execute("SELECT IFNULL(SUM(amount), 0) FROM income")
        total_income = float(cursor.fetchone()[0] or 0)

        cursor.execute("SELECT IFNULL(SUM(amount), 0) FROM expenses")
        total_expenses = float(cursor.fetchone()[0] or 0)

        total_savings = total_income - total_expenses
        savings_rate = (total_savings / total_income * 100) if total_income > 0 else 0

        return total_income, total_expenses, total_savings, savings_rate
    
    def get_report_category_spending(self):
        """
        Returns two lists:
        categories, totals
        sorted by highest spending first
        """
        cursor.execute("""
            SELECT category, IFNULL(SUM(amount), 0) as total
            FROM expenses
            GROUP BY category
            ORDER BY total DESC
        """)
        rows = cursor.fetchall()

        categories = []
        totals = []

        for category, total in rows:
            categories.append(category)
            totals.append(float(total or 0))

        return categories, totals

    def get_weekday_spending(self):
        """
        Returns two lists (weekday_labels, weekday_totals) with totals
        summed across Mon..Sun, used for the Reports page's day-of-week
        spending pattern chart.
        """
        cursor.execute("SELECT amount, date FROM expenses")
        rows = cursor.fetchall()

        weekday_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        totals = [0.0] * 7

        for amount, dt in rows:
            try:
                weekday_idx = datetime.strptime(dt, "%d-%m-%Y").weekday()
                totals[weekday_idx] += float(amount or 0)
            except (TypeError, ValueError):
                continue

        return weekday_labels, totals

    def get_report_period_data(self, period_type="all", period_key=None):
        """
        period_type: "all" | "monthly" | "yearly"
        period_key: "MM-YYYY" for monthly, "YYYY" for yearly, ignored for "all"

        Returns a dict with everything the Reports page and its exports need:
        total_income, total_expenses, total_savings, savings_rate,
        category_totals (list of (category, amount) sorted desc),
        top_categories (top 5 of the above),
        highest_expense (name, amount, category, date) or None,
        highest_income (source, amount, date) or None,
        label (human-readable period label for headers/exports)
        """
        def in_period(date_str):
            try:
                d = datetime.strptime(date_str, "%d-%m-%Y")
            except (TypeError, ValueError):
                return False
            if period_type == "monthly":
                return d.strftime("%m-%Y") == period_key
            if period_type == "yearly":
                return d.strftime("%Y") == period_key
            return True

        cursor.execute("SELECT name, amount, category, date FROM expenses")
        expense_rows = [r for r in cursor.fetchall() if in_period(r[3])]

        cursor.execute("SELECT source, amount, date FROM income")
        income_rows = [r for r in cursor.fetchall() if in_period(r[2])]

        total_expenses = sum(r[1] for r in expense_rows)
        total_income = sum(r[1] for r in income_rows)
        total_savings = total_income - total_expenses
        savings_rate = (total_savings / total_income * 100) if total_income > 0 else 0

        category_map = {}
        for name, amount, category, date in expense_rows:
            cat = (category or "Other").strip().title()
            category_map[cat] = category_map.get(cat, 0) + amount
        category_totals = sorted(category_map.items(), key=lambda x: x[1], reverse=True)

        highest_expense = max(expense_rows, key=lambda r: r[1]) if expense_rows else None
        highest_income = max(income_rows, key=lambda r: r[1]) if income_rows else None

        if period_type == "monthly" and period_key:
            label = datetime.strptime(period_key, "%m-%Y").strftime("%B %Y")
        elif period_type == "yearly" and period_key:
            label = period_key
        else:
            label = "All Time"

        return {
            "total_income": total_income,
            "total_expenses": total_expenses,
            "total_savings": total_savings,
            "savings_rate": savings_rate,
            "category_totals": category_totals,
            "top_categories": category_totals[:5],
            "highest_expense": highest_expense,
            "highest_income": highest_income,
            "label": label,
        }

    def get_available_report_years(self):
        """Distinct years (as strings) that have any expense or income activity."""
        years = set()
        cursor.execute("SELECT date FROM expenses UNION SELECT date FROM income")
        for (dt,) in cursor.fetchall():
            try:
                years.add(datetime.strptime(dt, "%d-%m-%Y").strftime("%Y"))
            except (TypeError, ValueError):
                continue
        current_year = str(datetime.now().year)
        years.add(current_year)
        return sorted(years, reverse=True)

    # PANEL 4: REPORTS & ANALYTICS (Ref: pg 4_2.png)
    def load_reports(self):

        self.highlight_sidebar("📊 Reports")

        self.clean_view()

        reports_title_row = Frame(content_frame, bg=BG_CANVAS)
        reports_title_row.pack(fill=X)
        Label(reports_title_row, text="Reports", font=("Segoe UI", 22, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT, anchor="w")

        export_bar = Frame(reports_title_row, bg=BG_CANVAS)
        export_bar.pack(side=RIGHT)

        Button(
            export_bar,
            text="📄 PDF",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            command=lambda: self.export_report_pdf()
        ).pack(side=RIGHT, padx=(6, 0))

        Button(
            export_bar,
            text="📊 Excel",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            command=lambda: self.export_report_excel()
        ).pack(side=RIGHT, padx=(6, 0))

        Button(
            export_bar,
            text="🗂 Full CSV",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            command=self.export_full_report_csv
        ).pack(side=RIGHT, padx=(6, 0))

        Label(content_frame, text="Financial summary & analytics for the period you choose below", font=("Segoe UI", 11), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(2, 15))

        # ---------------- PERIOD SELECTOR ---------------- #
        if not hasattr(self, "report_period_type"):
            self.report_period_type = "All Time"
        if not hasattr(self, "report_period_key"):
            self.report_period_key = None

        period_row = Frame(content_frame, bg=BG_CANVAS)
        period_row.pack(fill=X, pady=(0, 20))

        Label(period_row, text="Report:", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(side=LEFT, padx=(0, 8))

        period_type_var = StringVar(value=self.report_period_type)
        period_type_combo = ttk.Combobox(
            period_row, textvariable=period_type_var,
            values=["All Time", "Monthly", "Yearly"],
            state="readonly", width=12, font=("Segoe UI", 9)
        )
        period_type_combo.pack(side=LEFT, padx=(0, 10))

        sub_period_var = StringVar()
        sub_period_combo = ttk.Combobox(period_row, textvariable=sub_period_var, state="readonly", width=14, font=("Segoe UI", 9))

        def build_sub_period_values():
            if period_type_var.get() == "Monthly":
                months = []
                today = datetime.now().replace(day=1)
                for i in range(24):
                    year = today.year
                    month = today.month - i
                    while month <= 0:
                        month += 12
                        year -= 1
                    months.append((datetime(year, month, 1).strftime("%b %Y"), f"{month:02d}-{year}"))
                return months
            elif period_type_var.get() == "Yearly":
                return [(y, y) for y in self.get_available_report_years()]
            return []

        sub_values = build_sub_period_values()
        sub_period_combo["values"] = [v[0] for v in sub_values]

        if period_type_var.get() in ("Monthly", "Yearly") and sub_values:
            if self.report_period_key and self.report_period_key in [v[1] for v in sub_values]:
                match = next(v[0] for v in sub_values if v[1] == self.report_period_key)
                sub_period_var.set(match)
            else:
                sub_period_var.set(sub_values[0][0])
                self.report_period_key = sub_values[0][1]
            sub_period_combo.pack(side=LEFT)

        def on_period_type_change(event=None):
            self.report_period_type = period_type_var.get()
            self.report_period_key = None
            self.load_reports()

        def on_sub_period_change(event=None):
            values = build_sub_period_values()
            match = next((v[1] for v in values if v[0] == sub_period_var.get()), None)
            self.report_period_key = match
            self.load_reports()

        period_type_combo.bind("<<ComboboxSelected>>", on_period_type_change)
        sub_period_combo.bind("<<ComboboxSelected>>", on_sub_period_change)

        period_map = {"All Time": "all", "Monthly": "monthly", "Yearly": "yearly"}
        report = self.get_report_period_data(period_map[self.report_period_type], self.report_period_key)

        weekday_labels, weekday_totals = self.get_weekday_spending()

        # ---------------- SUMMARY GRID ---------------- #
        grid = Frame(content_frame, bg=BG_CANVAS)
        grid.pack(fill=X, pady=(0, 20))

        summary_items = [
            ("INCOME", f"₹{report['total_income']:,.0f}", COLOR_INCOME),
            ("EXPENSES", f"₹{report['total_expenses']:,.0f}", COLOR_EXPENSE),
            ("SAVINGS", f"₹{report['total_savings']:,.0f}", COLOR_SAVINGS if report['total_savings'] >= 0 else COLOR_DANGER),
            ("SAVINGS RATE", f"{report['savings_rate']:.1f}%", COLOR_INCOME if report['savings_rate'] >= 0 else COLOR_DANGER)
        ]
        for title, val, color in summary_items:
            box = Frame(grid, bg=BG_CANVAS)
            box.pack(side=LEFT, expand=True, fill=X)
            Label(box, text=title, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
            Label(box, text=val, font=("Segoe UI", 18, "bold"), fg=color, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))

        # ---------------- HIGHEST EXPENSE / HIGHEST INCOME ---------------- #
        highlight_row = Frame(content_frame, bg=BG_CANVAS)
        highlight_row.pack(fill=X, pady=(0, 20))

        he_card = Frame(highlight_row, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=16, pady=14)
        he_card.pack(side=LEFT, expand=True, fill=X, padx=(0, 8))
        Label(he_card, text="HIGHEST EXPENSE", font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
        if report["highest_expense"]:
            name, amount, category, date = report["highest_expense"]
            Label(he_card, text=f"₹{amount:,.0f} — {name}", font=("Segoe UI", 13, "bold"), fg=COLOR_EXPENSE, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))
            Label(he_card, text=f"{category or 'Other'} • {date}", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
        else:
            Label(he_card, text="No expenses in this period", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))

        hi_card = Frame(highlight_row, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=16, pady=14)
        hi_card.pack(side=LEFT, expand=True, fill=X, padx=(8, 0))
        Label(hi_card, text="HIGHEST INCOME", font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
        if report["highest_income"]:
            source, amount, date = report["highest_income"]
            Label(hi_card, text=f"₹{amount:,.0f} — {source}", font=("Segoe UI", 13, "bold"), fg=COLOR_INCOME, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))
            Label(hi_card, text=f"{date}", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
        else:
            Label(hi_card, text="No income in this period", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))

        # ---------------- CATEGORY REPORT + TOP SPENDING ---------------- #
        cat_deck = Frame(content_frame, bg=BG_CANVAS)
        cat_deck.pack(fill=X, pady=(0, 20))
        cat_deck.columnconfigure(0, weight=1)
        cat_deck.columnconfigure(1, weight=1)

        cat_card = Frame(cat_deck, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=16, pady=14)
        cat_card.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        Label(cat_card, text=f"Category Report — {report['label']}", font=("Segoe UI", 12, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 8))

        if report["category_totals"]:
            total_spend = sum(v for _, v in report["category_totals"]) or 1
            for cat, amt in report["category_totals"][:8]:
                row = Frame(cat_card, bg=BG_CANVAS)
                row.pack(fill=X, pady=3)
                Label(row, text=cat, font=("Segoe UI", 10), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)
                Label(row, text=f"₹{amt:,.0f} ({amt/total_spend*100:.0f}%)", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(side=RIGHT)
        else:
            Label(cat_card, text="No category data for this period.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")

        top_card = Frame(cat_deck, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=16, pady=14)
        top_card.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
        Label(top_card, text="Top Spending Categories", font=("Segoe UI", 12, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 8))

        if report["top_categories"]:
            medal = ["🥇", "🥈", "🥉", "4.", "5."]
            for i, (cat, amt) in enumerate(report["top_categories"]):
                row = Frame(top_card, bg=BG_CANVAS)
                row.pack(fill=X, pady=3)
                Label(row, text=f"{medal[i]} {cat}", font=("Segoe UI", 10, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)
                Label(row, text=f"₹{amt:,.0f}", font=("Segoe UI", 10, "bold"), fg=COLOR_EXPENSE, bg=BG_CANVAS).pack(side=RIGHT)
        else:
            Label(top_card, text="No spending data for this period.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")

        # Split Chart Deck
        deck = Frame(content_frame, bg=BG_CANVAS)
        deck.pack(fill=BOTH, expand=True, pady=10)
        deck.columnconfigure(0, weight=1)
        deck.columnconfigure(1, weight=1)
        
        # Left Grid Split: 6-Month Comparison Block Component
        g1 = Frame(deck, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=15, pady=15)
        g1.grid(row=0, column=0, padx=8, sticky="nsew")

        Label(
            g1,
            text="Spending by Day of Week",
            font=("Segoe UI", 12, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        Label(
            g1,
            text="Which days you tend to spend the most on (all time)",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 6))

        fig, ax = plt.subplots(figsize=(4.4, 2.8), facecolor=BG_CANVAS)
        ax.set_facecolor(BG_CANVAS)

        if weekday_labels and any(weekday_totals):
            peak_idx = weekday_totals.index(max(weekday_totals))
            bar_colors = [
                COLOR_EXPENSE if i == peak_idx else "#FDBA9C"
                for i in range(len(weekday_totals))
            ]

            ax.bar(weekday_labels, weekday_totals, color=bar_colors, width=0.6)

            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color(BORDER_COLOR)
            ax.spines['bottom'].set_color(BORDER_COLOR)
            ax.tick_params(colors=TEXT_MUTED, labelsize=9)
        else:
            ax.text(
                0.5, 0.5,
                "No expense data yet",
                ha="center",
                va="center",
                fontsize=11,
                color=TEXT_MUTED
            )
            ax.axis("off")

        canvas1 = FigureCanvasTkAgg(fig, master=g1)
        canvas1.get_tk_widget().pack(fill=BOTH, expand=True)
        
        # Right Grid Split: Line Dot Trend Graph
        g2 = Frame(deck, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=15, pady=15)
        g2.grid(row=0, column=1, padx=8, sticky="nsew")

        Label(
            g2,
            text=f"Financial Overview — {report['label']}",
            font=("Segoe UI", 12, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w")

        fig2, ax2 = plt.subplots(figsize=(4.4, 2.8), facecolor=BG_CANVAS)
        ax2.set_facecolor(BG_CANVAS)

        overview_labels = ["Income", "Expenses", "Savings"]
        overview_values = [report['total_income'], report['total_expenses'], max(report['total_savings'], 0)]

        overview_colors = [COLOR_INCOME, COLOR_EXPENSE, COLOR_SAVINGS]

        ax2.bar(overview_labels, overview_values, color=overview_colors, width=0.55)

        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.spines['left'].set_color(BORDER_COLOR)
        ax2.spines['bottom'].set_color(BORDER_COLOR)
        ax2.tick_params(colors=TEXT_MUTED, labelsize=9)

        canvas2 = FigureCanvasTkAgg(fig2, master=g2)
        canvas2.get_tk_widget().pack(fill=BOTH, expand=True)
        
    # PANEL 5: COOPER AI & STREAKS ENGINE (Ref: pg 5_2.png)
    def generate_cooper_insights(self):
        """
        Builds a short list of (icon, title, detail) insights computed
        live from the user's actual data — used on the Cooper AI page.
        """
        insights = []

        now = datetime.now()
        this_month_key = now.strftime("%m-%Y")
        last_month_dt = now.replace(day=1) - timedelta(days=1)
        last_month_key = last_month_dt.strftime("%m-%Y")

        this_month_cats = self.get_category_spending_data(this_month_key)
        this_month_total = sum(v for c, v in this_month_cats if c != "No Data")

        last_month_cats = self.get_category_spending_data(last_month_key)
        last_month_total = sum(v for c, v in last_month_cats if c != "No Data")

        # ---------------- Monthly spending summary ---------------- #
        # (date is stored as DD-MM-YYYY text, so tally via Python rather than SQL strftime)
        cursor.execute("SELECT date FROM expenses")
        this_month_txn_count = 0
        for (dt,) in cursor.fetchall():
            try:
                if datetime.strptime(dt, "%d-%m-%Y").strftime("%m-%Y") == this_month_key:
                    this_month_txn_count += 1
            except (TypeError, ValueError):
                continue

        if this_month_total > 0:
            insights.append((
                "🗓️", "This month's spending summary",
                f"You've logged {this_month_txn_count} expense{'s' if this_month_txn_count != 1 else ''} totalling ₹{this_month_total:,.0f} so far this month."
            ))

        # ---------------- Upcoming bill reminder ---------------- #
        cursor.execute("SELECT title, amount, due_date FROM bills_emi WHERE status IN ('Unpaid')")
        soonest_bill = None
        soonest_days = None
        for title, amount, due_date in cursor.fetchall():
            try:
                due_dt = datetime.strptime(due_date, "%d-%m-%Y").date()
            except (TypeError, ValueError):
                continue
            days_left = (due_dt - now.date()).days
            if days_left <= 7 and (soonest_days is None or days_left < soonest_days):
                soonest_days = days_left
                soonest_bill = (title, amount, days_left)

        if soonest_bill:
            title, amount, days_left = soonest_bill
            if days_left < 0:
                insights.append(("🚨", "Bill overdue", f"'{title}' (₹{amount:,.0f}) was due {abs(days_left)} day{'s' if abs(days_left) != 1 else ''} ago."))
            elif days_left == 0:
                insights.append(("⏰", "Bill due today", f"'{title}' (₹{amount:,.0f}) is due today."))
            else:
                insights.append(("⏰", "Upcoming bill reminder", f"'{title}' (₹{amount:,.0f}) is due in {days_left} day{'s' if days_left != 1 else ''}."))

        if this_month_cats and this_month_cats[0][0] != "No Data":
            top_cat, top_amt = this_month_cats[0]
            insights.append((
                "🏆", "Top category this month",
                f"{top_cat} leads your spending at ₹{top_amt:,.0f} so far."
            ))

        if last_month_total > 0:
            change_pct = ((this_month_total - last_month_total) / last_month_total) * 100
            if change_pct > 5:
                insights.append((
                    "📈", "Spending is trending up",
                    f"You've spent {change_pct:.0f}% more than last month so far."
                ))
            elif change_pct < -5:
                insights.append((
                    "📉", "Nice, spending is down",
                    f"You've spent {abs(change_pct):.0f}% less than last month so far."
                ))
            else:
                insights.append((
                    "⚖️", "Steady spending",
                    "Your spending pace is about the same as last month."
                ))

        total_income, total_expenses, total_savings, savings_rate = self.get_report_summary()
        if total_income > 0:
            if savings_rate >= 20:
                insights.append((
                    "💰", "Healthy savings rate",
                    f"You're saving {savings_rate:.0f}% of your income — that's solid."
                ))
            elif savings_rate >= 0:
                top_cat_name = this_month_cats[0][0] if this_month_cats and this_month_cats[0][0] != "No Data" else None
                suggestion = f" Trimming your {top_cat_name} spending is the fastest way to raise it." if top_cat_name else ""
                insights.append((
                    "🪙", "Savings suggestion",
                    f"You're currently saving {savings_rate:.0f}% of your income — aim for 20% if you can.{suggestion}"
                ))
            else:
                insights.append((
                    "⚠️", "Spending more than you earn",
                    f"Expenses currently exceed income by ₹{abs(total_savings):,.0f}. Consider setting a budget on your top category."
                ))

        cursor.execute("SELECT category, amount FROM budgets")
        for category, budget_amt in cursor.fetchall():
            if not budget_amt:
                continue
            cursor.execute("SELECT COALESCE(SUM(amount),0) FROM expenses WHERE category=?", (category,))
            spent = cursor.fetchone()[0] or 0
            if spent >= budget_amt:
                insights.append((
                    "🚨", f"{category} budget exceeded",
                    f"You've spent ₹{spent:,.0f} of your ₹{budget_amt:,.0f} {category} budget."
                ))
            elif spent >= budget_amt * 0.8:
                insights.append((
                    "🔶", f"{category} budget almost hit",
                    f"You're at {(spent / budget_amt * 100):.0f}% of your {category} budget."
                ))

        if not insights:
            insights.append((
                "👋", "Getting started",
                "Log a few expenses and I'll start surfacing insights here."
            ))

        return insights[:6]

    def answer_cooper_question(self, question):
        """
        Lightweight rule-based Q&A engine that answers questions about the
        user's real data — powers the 'Ask Cooper' box.
        """
        q = question.lower().strip()
        total_income, total_expenses, total_savings, savings_rate = self.get_report_summary()

        if not q:
            return "Ask me something like \"how much did I spend on food\" or \"what's my savings rate\"."

        if "goal" in q:
            cursor.execute("SELECT COUNT(*), COALESCE(SUM(saved_amount),0), COALESCE(SUM(target_amount),0) FROM savings_goals")
            n, saved, target = cursor.fetchone()
            if n:
                return f"You have {n} savings goal(s): ₹{saved:,.0f} saved toward ₹{target:,.0f} total."
            return "You don't have any savings goals set up yet — head to Savings Goals to add one."

        if "save" in q or "saving" in q:
            return f"You've saved ₹{total_savings:,.0f} overall — a savings rate of {savings_rate:.1f}%."

        if "income" in q or "earn" in q:
            return f"Your total recorded income is ₹{total_income:,.0f}."

        if "budget" in q:
            cursor.execute("SELECT COUNT(*) FROM budgets")
            n = cursor.fetchone()[0]
            if n:
                return f"You have {n} budget categor{'y' if n == 1 else 'ies'} set up. Check the Budget page for utilisation."
            return "You haven't set any budgets yet — head to the Budget page to create one."

        if "spend" in q or "expense" in q or "spent" in q or "cost" in q:
            cursor.execute("SELECT DISTINCT category FROM expenses")
            cats = [c[0] for c in cursor.fetchall() if c[0]]
            for cat in cats:
                if cat.lower() in q:
                    cursor.execute("SELECT COALESCE(SUM(amount),0) FROM expenses WHERE category=?", (cat,))
                    amt = cursor.fetchone()[0] or 0
                    return f"You've spent ₹{amt:,.0f} on {cat} in total."
            return f"You've spent ₹{total_expenses:,.0f} in total across all categories."

        return "I can help with questions about your income, expenses, savings, budgets or goals — try asking things like \"how much did I spend on food\" or \"what's my savings rate\"."

    def get_user_streak(self):
        """
        Computes a REAL activity streak: consecutive days (counting back from
        today, or from yesterday if nothing was logged yet today) on which the
        user logged at least one expense or income entry. Persists the best
        streak ever reached in system_streaks so it survives a broken streak.
        Returns (current_streak, best_streak).
        """
        cursor.execute("SELECT date FROM expenses UNION SELECT date FROM income")
        dates = set()
        for (dt,) in cursor.fetchall():
            try:
                dates.add(datetime.strptime(dt, "%d-%m-%Y").date())
            except (TypeError, ValueError):
                continue

        today = datetime.now().date()
        anchor = today if today in dates else (today - timedelta(days=1))

        streak = 0
        cursor_date = anchor
        while cursor_date in dates:
            streak += 1
            cursor_date -= timedelta(days=1)

        cursor.execute("SELECT best_streak FROM system_streaks WHERE id=1")
        row = cursor.fetchone()
        best = max(row[0] if row and row[0] else 0, streak)

        if row:
            cursor.execute("UPDATE system_streaks SET current_streak=?, best_streak=? WHERE id=1", (streak, best))
        else:
            cursor.execute(
                "INSERT INTO system_streaks (id, current_streak, best_streak, daily_goal) VALUES (1, ?, ?, 0)",
                (streak, best)
            )
        conn.commit()

        return streak, best

    def get_financial_tip(self):
        """A rotating, genuinely-useful tip. Rotates by day so it's stable within a session but changes daily."""
        tips = [
            "Try the 50/30/20 rule: 50% needs, 30% wants, 20% savings.",
            "Log expenses the same day you spend — same-day logging is far more accurate than end-of-week catch-up.",
            "Small recurring subscriptions add up fast. Review your Recurring tab once a month.",
            "Set a budget for your top spending category first — it has the biggest impact.",
            "An emergency fund of 3 months' expenses is a great first savings goal.",
            "Round up your daily spends and save the difference — small habits compound.",
            "Review last month's Category Report before setting this month's budget.",
        ]
        return tips[datetime.now().timetuple().tm_yday % len(tips)]

    def load_cooper_ai(self):

        self.highlight_sidebar("🤖 Cooper AI")

        self.clean_view()

        header_row = Frame(content_frame, bg=BG_CANVAS)
        header_row.pack(fill=X, pady=(0, 14))

        mascot_canvas = Canvas(header_row, width=64, height=90, highlightthickness=0, bd=0, bg=BG_CANVAS)
        mascot_canvas.pack(side=LEFT, padx=(0, 12))
        # Medium-sized Cooper mascot as a page identity marker — smaller
        # than the full login-screen hero, per the "support, don't
        # dominate" sizing rule for in-app pages.
        self._draw_cooper_robot(mascot_canvas, 32, 40, scale=0.34, tag="mini_cooper")

        title_col = Frame(header_row, bg=BG_CANVAS)
        title_col.pack(side=LEFT, fill=X, expand=True)
        Label(title_col, text="Cooper AI", font=("Segoe UI", 20, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")
        Label(title_col, text="Your personal financial companion", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")

        current_streak, best_streak = self.get_user_streak()

        # ---------------- GREETING / STREAK CARD ---------------- #
        banner = Frame(content_frame, bg=BG_CANVAS, highlightthickness=1, highlightbackground=COLOR_EXPENSE, padx=16, pady=14)
        banner.pack(fill=X, pady=(0, 16))

        lbl_f = Frame(banner, bg=BG_CANVAS)
        lbl_f.pack(side=LEFT)

        hour = datetime.now().hour
        time_greeting = "Good morning" if 5 <= hour < 12 else "Good afternoon" if 12 <= hour < 17 else "Good evening" if 17 <= hour < 22 else "Good night"

        if current_streak >= 3:
            headline = f"{time_greeting}! You're on a {current_streak}-day logging streak 🎉"
        elif current_streak == 1:
            headline = f"{time_greeting}! Nice, you've logged something today — day 1."
        else:
            headline = f"{time_greeting}! Log an expense or income today to start a streak."

        Label(lbl_f, text=headline, font=("Segoe UI", 12, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")
        Label(
            lbl_f,
            text=f"Best streak so far: {best_streak} day{'s' if best_streak != 1 else ''}.",
            font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS
        ).pack(anchor="w", pady=(2, 0))

        s_f = Frame(banner, bg=BG_CANVAS)
        s_f.pack(side=RIGHT, padx=6)
        Label(s_f, text=str(current_streak), font=("Segoe UI", 19, "bold"), fg=COLOR_EXPENSE, bg=BG_CANVAS).pack()
        Label(s_f, text="🔥 Day Streak", font=("Segoe UI", 8, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack()

        # ---------------- ASK COOPER (interactive Q&A) — its own,
        # compact card, separated from insights/badges below ---------------- #
        ask_card = Frame(content_frame, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=16, pady=14)
        ask_card.pack(fill=X, pady=(0, 16))

        Label(ask_card, text="💬 Ask Cooper", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")
        Label(
            ask_card, text="Ask about your income, expenses, savings, budgets or goals",
            font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 8))

        chat_log = Text(
            ask_card, height=5, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 9),
            wrap="word", state="disabled", padx=10, pady=8
        )
        chat_log.pack(fill=X, pady=(0, 8))
        chat_log.tag_configure("question", font=("Segoe UI", 9, "bold"), foreground=TEXT_MAIN)
        chat_log.tag_configure("answer", foreground=TEXT_MUTED)

        def append_chat(question, answer):
            chat_log.configure(state="normal")
            if question:
                chat_log.insert(END, f"You: {question}\n", "question")
            chat_log.insert(END, f"Cooper: {answer}\n\n", "answer")
            chat_log.configure(state="disabled")
            chat_log.see(END)

        append_chat(None, "Hey! I'm Cooper 🦊 — ask me anything about your finances, like \"what's my savings rate\" or \"how much did I spend on food\".")

        input_row = Frame(ask_card, bg=BG_CANVAS)
        input_row.pack(fill=X)

        ask_entry = Entry(
            input_row, bg="white", fg=TEXT_MAIN, bd=0, highlightthickness=1,
            highlightbackground=BORDER_COLOR, font=("Segoe UI", 9)
        )
        ask_entry.pack(side=LEFT, fill=X, expand=True, ipady=6, ipadx=6, padx=(0, 8))

        def handle_ask(event=None):
            question = ask_entry.get().strip()
            if not question:
                return
            answer = self.answer_cooper_question(question)
            append_chat(question, answer)
            ask_entry.delete(0, END)

        ask_entry.bind("<Return>", handle_ask)

        Button(
            input_row, text="Ask", bg=COLOR_ACCENT, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), padx=16, command=handle_ask
        ).pack(side=RIGHT)

        # ---------------- QUICK QUESTIONS — its own section, so the ask
        # card above doesn't have to hold every control at once ---------------- #
        Label(content_frame, text="Quick Questions", font=("Segoe UI", 12, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 8))

        quick_row = Frame(content_frame, bg=BG_CANVAS)
        quick_row.pack(fill=X, pady=(0, 18))

        def ask_quick(question):
            answer = self.answer_cooper_question(question)
            append_chat(question, answer)

        quick_questions = [
            ("💸 Spending", "how much did I spend this month"),
            ("💰 Savings", "what's my savings rate"),
            ("📊 Budget", "how am I doing on my budget"),
            ("❤️ Financial Health", "how healthy are my finances"),
        ]
        for label, question in quick_questions:
            Button(
                quick_row, text=label, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 9),
                padx=12, pady=6, cursor="hand2", command=lambda q=question: ask_quick(q)
            ).pack(side=LEFT, padx=(0, 8))

        # ---------------- RECENT AI INSIGHTS (data-driven) ---------------- #
        insights_wrap = Frame(content_frame, bg=BG_CANVAS)
        insights_wrap.pack(fill=X, pady=(0, 16))

        Label(insights_wrap, text="✨ Recent AI Insights", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 8))

        for icon, title, detail in self.generate_cooper_insights():
            row = Frame(insights_wrap, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=12, pady=10)
            row.pack(fill=X, pady=3)

            Label(row, text=icon, font=("Segoe UI", 14), bg=BG_CANVAS).pack(side=LEFT, padx=(0, 10))

            txt_col = Frame(row, bg=BG_CANVAS)
            txt_col.pack(side=LEFT, fill=X, expand=True)
            Label(txt_col, text=title, font=("Segoe UI", 10, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS, anchor="w").pack(fill=X, anchor="w")
            Label(txt_col, text=detail, font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS, anchor="w", wraplength=760, justify=LEFT).pack(fill=X, anchor="w")

        # ---------------- BADGE COLLECTION ---------------- #
        b_frame = LabelFrame(content_frame, text="🏅 Badge Collection", font=("Segoe UI", 11, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS, bd=0)
        b_frame.pack(fill=X, pady=(0, 16))

        badges = [("🌱 New Saver", "1-day streak", 1), ("🏅 Saver Starter", "3-day streak", 3), ("🏆 Budget Warrior", "7-day streak", 7)]
        for title, subtitle, threshold in badges:
            earned = best_streak >= threshold
            box_bg = "#FEF3C7" if earned else BG_SIDEBAR
            b_box = Frame(b_frame, bg=box_bg, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=12, pady=10)
            b_box.pack(side=LEFT, expand=True, fill=X, padx=5, pady=8)
            Label(b_box, text=title if earned else f"🔒 {title.split(' ', 1)[1]}", font=("Segoe UI", 10, "bold"), fg=TEXT_MAIN if earned else TEXT_MUTED, bg=box_bg).pack()
            Label(b_box, text=subtitle + (" ✓ earned" if earned else " — not yet"), font=("Segoe UI", 8), fg=TEXT_MUTED, bg=box_bg).pack(pady=(2, 0))

        # ---------------- FINANCIAL TIP OF THE DAY ---------------- #
        tip_card = Frame(content_frame, bg="#ECFEFF", highlightthickness=1, highlightbackground="#A5F3FC", padx=16, pady=12)
        tip_card.pack(fill=X, pady=(0, 10))
        Label(tip_card, text="💡 Financial Tip of the Day", font=("Segoe UI", 11, "bold"), fg="#155E75", bg="#ECFEFF").pack(anchor="w")
        Label(tip_card, text=self.get_financial_tip(), font=("Segoe UI", 9), fg="#0E7490", bg="#ECFEFF", wraplength=800, justify=LEFT).pack(anchor="w", pady=(3, 0))

    # PANEL 6: SMS TRANSACTION SCANNER BACKEND LINK
    def parse_bank_sms(self, body):
        """
        Extracts (txn_type, amount, merchant, date) from a bank/card
        transaction SMS body. txn_type is 'debit' or 'credit' (defaults
        to 'debit' if neither keyword is present). Any piece that can't
        be confidently detected comes back as None (amount/merchant) or
        today's date as a safe fallback.
        """
        text = body or ""

        txn_type = "debit"
        if re.search(r'\bcredited\b|\breceived\b|\bcredit\b', text, re.IGNORECASE):
            txn_type = "credit"
        if re.search(r'\bdebited\b|\bspent\b|\bdebit\b|\bpaid\b', text, re.IGNORECASE):
            txn_type = "debit"

        amount = None
        amt_match = re.search(r'(?:INR|Rs\.?|₹)\s*([\d,]+(?:\.\d{1,2})?)', text, re.IGNORECASE)
        if amt_match:
            try:
                amount = float(amt_match.group(1).replace(',', ''))
            except ValueError:
                amount = None

        merchant = None
        at_match = re.search(r'\bat\s+([A-Za-z0-9&.\-\' ]{3,40})', text)
        if at_match:
            merchant = at_match.group(1).strip().rstrip('.').title()

        date_val = datetime.today().strftime("%d-%m-%Y")
        date_match = re.search(r'\b(\d{2})[/-](\d{2})[/-](\d{2,4})\b', text)
        if date_match:
            dd, mm, yy = date_match.groups()
            if len(yy) == 2:
                yy = "20" + yy
            try:
                datetime.strptime(f"{dd}-{mm}-{yy}", "%d-%m-%Y")
                date_val = f"{dd}-{mm}-{yy}"
            except ValueError:
                pass

        return txn_type, amount, merchant, date_val

    def load_sms_scanner(self):

        self.highlight_sidebar("📩 Message Reader")

        self.clean_view()

        Label(content_frame, text="Message Reader", font=("Segoe UI", 22, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 5))
        Label(
            content_frame,
            text="Reads bank/card transaction SMS and turns unread debit or credit alerts into ledger entries",
            font=("Segoe UI", 11),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 20))

        # ---------------- SUMMARY CARDS ---------------- #
        cursor.execute("SELECT COUNT(*) FROM messages WHERE status='unread'")
        unread_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM messages")
        total_count = cursor.fetchone()[0]

        summary = Frame(content_frame, bg=BG_CANVAS)
        summary.pack(fill=X, pady=(0, 18))
        for title, value, color in [
            ("TOTAL MESSAGES", str(total_count), COLOR_ACCENT),
            ("UNREAD", str(unread_count), COLOR_EXPENSE if unread_count else COLOR_INCOME),
        ]:
            card = Frame(summary, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=18, pady=14)
            card.pack(side=LEFT, expand=True, fill=X, padx=6)
            Label(card, text=title, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
            Label(card, text=value, font=("Segoe UI", 18, "bold"), fg=color, bg=BG_CANVAS).pack(anchor="w", pady=(4, 0))

        # ---------------- ACTION BAR ---------------- #
        action_bar = Frame(content_frame, bg=BG_CANVAS)
        action_bar.pack(fill=X, pady=(0, 15))

        btn_add_message = Button(
            action_bar, text="➕ Add Message", bg=COLOR_ACCENT, fg="white", bd=0,
            font=("Segoe UI", 10, "bold"), padx=14, pady=8,
            command=lambda: self.add_message_popup()
        )
        btn_add_message.pack(side=LEFT)
        self.add_hover_effect(btn_add_message)

        # ---------------- TABLE ---------------- #
        table_f = Frame(content_frame, bg=BG_CANVAS)
        table_f.pack(fill=BOTH, expand=True)

        tree = ttk.Treeview(table_f, columns=("ID", "Sender", "Body", "Status"), show="headings")
        tree.heading("ID", text="#")
        tree.heading("Sender", text="SENDER")
        tree.heading("Body", text="MESSAGE")
        tree.heading("Status", text="STATUS")

        tree.column("ID", width=50, anchor=CENTER)
        tree.column("Sender", width=140, anchor=CENTER)
        tree.column("Body", width=520, anchor=W)
        tree.column("Status", width=110, anchor=CENTER)

        row_id_map = {}

        def refresh_table():
            tree.delete(*tree.get_children())
            row_id_map.clear()
            cursor.execute("SELECT id, sender, body, status FROM messages ORDER BY id DESC")
            for mid, sender, body, status in cursor.fetchall():
                iid = tree.insert("", END, values=(mid, sender, f"  {body}", (status or "").title()))
                row_id_map[iid] = mid

        refresh_table()

        if total_count == 0:
            Label(
                table_f,
                text="No messages yet. Click \"➕ Add Message\" to paste in a bank SMS.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=10)
        else:
            tree.pack(fill=BOTH, expand=True, pady=(0, 15))

        def delete_selected_message():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a message to delete.")
                return
            msg_id = row_id_map.get(selected[0])
            if msg_id is None:
                return
            if not messagebox.askyesno("Delete Message", "Delete this message? This won't remove any expense/income it already created."):
                return
            try:
                cursor.execute("DELETE FROM messages WHERE id=?", (msg_id,))
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not delete this message:\n{e}")
                return
            self.load_sms_scanner()

        def trigger_sync_messages():
            cursor.execute("SELECT id, sender, body FROM messages WHERE status='unread'")
            unread = cursor.fetchall()
            if not unread:
                messagebox.showinfo("Inbox Clean", "All messages have already been processed.")
                return

            added = 0
            skipped = 0
            try:
                for msg_id, sender, body in unread:
                    txn_type, amount, merchant, date_val = self.parse_bank_sms(body)

                    if amount and amount > 0 and txn_type == "debit":
                        name = merchant or (sender or "Card/Bank Debit")
                        category = self.suggest_category_from_merchant(merchant or "")
                        cursor.execute(
                            "INSERT INTO expenses (name, amount, category, date) VALUES (?, ?, ?, ?)",
                            (name, amount, category, date_val)
                        )
                        added += 1
                    elif amount and amount > 0 and txn_type == "credit":
                        source = merchant or (sender or "Bank Credit")
                        cursor.execute(
                            "INSERT INTO income (source, amount, date) VALUES (?, ?, ?)",
                            (source, amount, date_val)
                        )
                        added += 1
                    else:
                        skipped += 1

                cursor.execute("UPDATE messages SET status='read' WHERE status='unread'")
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not process these messages:\n{e}")
                return

            if skipped:
                messagebox.showinfo(
                    "Sync Complete",
                    f"Added {added} transaction(s) to your ledger.\n{skipped} message(s) didn't contain a "
                    "recognizable amount and were marked read without creating an entry."
                )
            else:
                messagebox.showinfo("Sync Complete", f"Added {added} transaction(s) to your ledger.")
            self.load_sms_scanner()

        btn_row = Frame(content_frame, bg=BG_CANVAS)
        btn_row.pack(fill=X)

        btn_sync = Button(
            btn_row, text="🔄 Sync Unread Messages", bg=COLOR_ACCENT, fg="white", bd=0,
            font=("Segoe UI", 11, "bold"), padx=15, pady=9, command=trigger_sync_messages
        )
        btn_sync.pack(side=LEFT)
        self.add_hover_effect(btn_sync)

        btn_delete_msg = Button(
            btn_row, text="🗑 Delete Selected", bg=COLOR_DANGER, fg="white", bd=0,
            font=("Segoe UI", 10, "bold"), padx=14, pady=9, command=delete_selected_message
        )
        btn_delete_msg.pack(side=RIGHT)
        self.add_hover_effect(btn_delete_msg)

    def add_message_popup(self):
        """Manually add a bank/card SMS (paste it in) instead of relying only on demo data."""
        pop, body = self.create_scroll_popup("Add Message", 520, 420)

        Label(body, text="Add Message", font=("Segoe UI", 16, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(20, 15))

        Label(body, text="Sender (e.g. HDFCBK, ICICIB)", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(0, 4))
        sender_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        sender_entry.pack(fill=X, padx=25, ipady=7)

        Label(body, text="Message Text", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 4))
        body_text = Text(body, height=6, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10), wrap=WORD, highlightthickness=1, highlightbackground=BORDER_COLOR)
        body_text.pack(fill=X, padx=25, ipady=6)

        def save_message():
            sender = sender_entry.get().strip() or "Unknown"
            text_val = body_text.get("1.0", END).strip()
            if not text_val:
                messagebox.showerror("Missing Message", "Please paste in the message text.")
                return
            try:
                cursor.execute(
                    "INSERT INTO messages (sender, body, status) VALUES (?, ?, 'unread')",
                    (sender, text_val)
                )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this message:\n{e}")
                return
            pop.destroy()
            self.load_sms_scanner()

        Button(
            body, text="Save Message", bg=COLOR_ACCENT, fg="white", bd=0,
            font=("Segoe UI", 11, "bold"), command=save_message
        ).pack(fill=X, padx=25, pady=25, ipady=8)

    # ---------------- POPUP WINDOW ACTIONS & LOGIC HELPER METHODS ---------------- #

    def build_subscriptions_section(self, parent):
        sub_card = Frame(
            parent,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            padx=15,
            pady=15
        )
        sub_card.pack(fill=BOTH, expand=True, pady=12)

        Label(
            sub_card,
            text="Subscriptions",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", pady=(0, 10))

        sub_tree = ttk.Treeview(
            sub_card,
            columns=("ID", "Title", "Amount", "Frequency", "Next Due", "Status"),
            show="headings",
            height=7
        )

        sub_tree.heading("ID", text="#")
        sub_tree.heading("Title", text="TITLE")
        sub_tree.heading("Amount", text="AMOUNT")
        sub_tree.heading("Frequency", text="FREQUENCY")
        sub_tree.heading("Next Due", text="NEXT DUE")
        sub_tree.heading("Status", text="STATUS")

        sub_tree.column("ID", width=45, anchor=CENTER)
        sub_tree.column("Title", width=220, anchor=W)
        sub_tree.column("Amount", width=110, anchor=E)
        sub_tree.column("Frequency", width=110, anchor=CENTER)
        sub_tree.column("Next Due", width=130, anchor=CENTER)
        sub_tree.column("Status", width=100, anchor=CENTER)

        cursor.execute("""
            SELECT id, title, amount, frequency, next_due_date
            FROM recurring_transactions
            WHERE lower(txn_type)='expense'
              AND (
                    lower(category)='subscription'
                    OR lower(title) LIKE '%netflix%'
                    OR lower(title) LIKE '%spotify%'
                    OR lower(title) LIKE '%prime%'
                    OR lower(title) LIKE '%subscription%'
                  )
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()

        for idx, row in enumerate(rows, 1):
            sub_tree.insert(
                "",
                END,
                values=(
                    idx,
                    row[1],
                    f"₹{row[2]:,.2f}",
                    row[3].title() if row[3] else "-",
                    row[4],
                    "Active"
                )
            )

        sub_tree.pack(fill=BOTH, expand=True)
        return sub_tree, sub_card
    
    def build_subscription_action_bar(self, parent, sub_tree):
        action_row = Frame(parent, bg=BG_CANVAS)
        action_row.pack(fill=X, pady=(10, 0))

        def _get_selected_subscription_db_id():
            selected = sub_tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a subscription first.")
                return None

            item = sub_tree.item(selected[0])
            values = item["values"]
            if not values:
                return None

            display_id = values[0]

            cursor.execute("""
                SELECT id, title, amount, frequency, next_due_date
                FROM recurring_transactions
                WHERE lower(txn_type)='expense'
                  AND (
                        lower(category)='subscription'
                        OR lower(title) LIKE '%netflix%'
                        OR lower(title) LIKE '%spotify%'
                        OR lower(title) LIKE '%prime%'
                        OR lower(title) LIKE '%subscription%'
                      )
                ORDER BY id DESC
            """)
            rows = cursor.fetchall()

            if display_id < 1 or display_id > len(rows):
                messagebox.showerror("Selection Error", "Could not map selected subscription.")
                return None

            return rows[display_id - 1][0]

        def cancel_selected_subscription():
            actual_db_id = _get_selected_subscription_db_id()
            if actual_db_id is None:
                return

            cursor.execute(
                "UPDATE recurring_transactions SET category='Cancelled Subscription' WHERE id=?",
                (actual_db_id,)
            )
            conn.commit()
            self.load_recurring()

        def delete_selected_subscription():
            actual_db_id = _get_selected_subscription_db_id()
            if actual_db_id is None:
                return

            confirm = messagebox.askyesno(
                "Delete Subscription",
                "Are you sure you want to delete the selected subscription?"
            )
            if not confirm:
                return

            cursor.execute("DELETE FROM recurring_transactions WHERE id=?", (actual_db_id,))
            conn.commit()
            self.load_recurring()

        Button(
            action_row,
            text="⛔ Cancel Subscription",
            bg="#F59E0B",
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=cancel_selected_subscription
        ).pack(side=RIGHT, padx=6)

        Button(
            action_row,
            text="🗑 Delete Selected",
            bg=COLOR_DANGER,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            command=delete_selected_subscription
        ).pack(side=RIGHT, padx=6)

    def add_expense_popup(self, expense_id=None):
        """
        expense_id=None -> Add mode (inserts a new row).
        expense_id=<int> -> Edit mode (pre-fills the form and updates that row).
        """
        editing = expense_id is not None
        existing = None
        if editing:
            cursor.execute("SELECT name, amount, category, date, notes FROM expenses WHERE id=?", (expense_id,))
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This expense no longer exists.")
                return

        pop, body = self.create_scroll_popup("Edit Expense" if editing else "Add Expense", 460, 640)

        Label(
            body,
            text="Edit Expense" if editing else "Add Expense",
            font=("Segoe UI", 16, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(20, 15))

        Label(body, text="Expense Name", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(0, 4))
        name_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        name_entry.pack(fill=X, padx=25, ipady=7)

        Label(body, text="Amount (₹)", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 4))
        amount_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        amount_entry.pack(fill=X, padx=25, ipady=7)

        Label(body, text="Category", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 4))

        default_categories = ["Food", "Travel", "Shopping", "Entertainment", "Health", "Education", "Bills"]
        categories = self.get_dropdown_options("expenses", "category", default_categories)

        category_var = StringVar(value=categories[0])
        category_menu = ttk.Combobox(body, textvariable=category_var, values=categories, state="readonly", font=("Segoe UI", 11))
        category_menu.pack(fill=X, padx=25, ipady=5)

        other_label = Label(body, text="Custom Category", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold"))
        other_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)

        def category_changed(event=None):
            if category_var.get() == "Other":
                other_label.pack(anchor="w", padx=25, pady=(14, 4))
                other_entry.pack(fill=X, padx=25, ipady=7)
            else:
                other_label.pack_forget()
                other_entry.pack_forget()

        category_menu.bind("<<ComboboxSelected>>", category_changed)

        Label(body, text="Date (DD-MM-YYYY)", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 4))
        date_row = Frame(body, bg=BG_CANVAS)
        date_row.pack(fill=X, padx=25)
        date_entry = Entry(date_row, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        date_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        date_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))
        self.attach_date_picker(date_entry).pack(side=LEFT, padx=(6, 0), ipady=5)

        Label(body, text="Notes (optional)", fg=TEXT_MUTED, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 4))
        notes_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        notes_entry.pack(fill=X, padx=25, ipady=7)

        # Pre-fill for edit mode
        if editing:
            ex_name, ex_amount, ex_category, ex_date, ex_notes = existing
            name_entry.insert(0, ex_name or "")
            amount_entry.insert(0, str(ex_amount))
            cat_match = next((c for c in categories if c.lower() == (ex_category or "").strip().lower()), None)
            if cat_match:
                category_var.set(cat_match)
            else:
                category_var.set("Other")
                category_changed()
                other_entry.insert(0, ex_category or "")
            date_entry.delete(0, END)
            date_entry.insert(0, ex_date or datetime.today().strftime("%d-%m-%Y"))
            notes_entry.insert(0, ex_notes or "")

        def save():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("Missing Name", "Please enter an expense name.")
                return

            ok, amount = self.validate_amount_input(amount_entry.get(), "Amount")
            if not ok:
                return

            category = category_var.get()
            if category == "Other":
                category = other_entry.get().strip().title()
            if not category:
                messagebox.showerror("Missing Category", "Please choose or enter a category.")
                return

            ok, date_val = self.validate_date_input(date_entry.get())
            if not ok:
                return

            notes = notes_entry.get().strip()

            try:
                if editing:
                    cursor.execute(
                        "UPDATE expenses SET name=?, amount=?, category=?, date=?, notes=? WHERE id=?",
                        (name, amount, category, date_val, notes, expense_id)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO expenses(name, amount, category, date, notes) VALUES(?, ?, ?, ?, ?)",
                        (name, amount, category, date_val, notes)
                    )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this expense:\n{e}")
                return

            pop.destroy()
            self.load_expenses()

        Button(
            body,
            text="Save Changes" if editing else "Save Expense",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 11, "bold"),
            command=save
        ).pack(fill=X, padx=25, pady=25, ipady=8)

    def add_goal_popup(self, goal_id=None):
        """
        goal_id=None -> Add mode (inserts a new goal).
        goal_id=<int> -> Edit mode (pre-fills the form and updates that row).
        """
        editing = goal_id is not None
        existing = None
        if editing:
            cursor.execute(
                "SELECT goal_name, target_amount, saved_amount, target_date FROM savings_goals WHERE id=?",
                (goal_id,)
            )
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This savings goal no longer exists.")
                return

        pop = Toplevel(self)
        pop.title("Edit Savings Goal" if editing else "Add Savings Goal")
        pop.geometry("540x640")
        pop.minsize(540, 640)
        pop.configure(bg=BG_CANVAS)
        pop.resizable(False, False)
        pop.transient(self)
        pop.grab_set()

        Label(
            pop,
            text="Edit Savings Goal" if editing else "Create Savings Goal",
            font=("Segoe UI", 16, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(20, 5))

        Label(
            pop,
            text="Track a target like laptop, emergency fund, trip, etc.",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(0, 15))

    # Goal name
        Label(pop, text="Goal Name", fg=TEXT_MAIN, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(8, 4))
        goal_entry = Entry(pop, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        goal_entry.pack(fill=X, padx=25, ipady=7)

    # Target amount
        Label(pop, text="Target Amount (₹)", fg=TEXT_MAIN, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(15, 4))
        target_entry = Entry(pop, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        target_entry.pack(fill=X, padx=25, ipady=7)

    # Already saved
        Label(pop, text="Already Saved (₹)", fg=TEXT_MAIN, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(15, 4))
        saved_entry = Entry(pop, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        saved_entry.pack(fill=X, padx=25, ipady=7)
        saved_entry.insert(0, "0")

    # Target date (with date picker)
        Label(pop, text="Target Date (DD-MM-YYYY)", fg=TEXT_MAIN, bg=BG_CANVAS, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(15, 4))
        date_row = Frame(pop, bg=BG_CANVAS)
        date_row.pack(fill=X, padx=25)
        date_entry = Entry(date_row, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        date_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        self.attach_date_picker(date_entry).pack(side=LEFT, padx=(6, 0), ipady=5)

        if editing:
            ex_name, ex_target, ex_saved, ex_date = existing
            goal_entry.insert(0, ex_name or "")
            target_entry.delete(0, END)
            target_entry.insert(0, str(ex_target))
            saved_entry.delete(0, END)
            saved_entry.insert(0, str(ex_saved))
            date_entry.insert(0, ex_date or datetime.today().strftime("%d-%m-%Y"))
        else:
            date_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))

        def save_goal():
            goal_name = goal_entry.get().strip()
            if not goal_name:
                messagebox.showerror("Missing Name", "Please enter a goal name.")
                return

            ok, target_amount = self.validate_amount_input(target_entry.get(), "Target amount")
            if not ok:
                return

            raw_saved = saved_entry.get().strip()
            try:
                saved_amount = float(raw_saved) if raw_saved else 0.0
                if saved_amount < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Amount", "Already Saved must be a valid number (0 or more).")
                return

            ok, target_date = self.validate_date_input(date_entry.get())
            if not ok:
                return

            try:
                if editing:
                    cursor.execute(
                        "UPDATE savings_goals SET goal_name=?, target_amount=?, saved_amount=?, target_date=? WHERE id=?",
                        (goal_name, target_amount, saved_amount, target_date, goal_id)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO savings_goals (goal_name, target_amount, saved_amount, target_date) VALUES (?, ?, ?, ?)",
                        (goal_name, target_amount, saved_amount, target_date)
                    )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this goal:\n{e}")
                return

            pop.destroy()

            if saved_amount >= target_amount > 0:
                messagebox.showinfo("🎉 Goal Reached!", f"Congratulations! '{goal_name}' has hit its target!")

            self.load_savings_goals()

        button_frame = Frame(pop, bg=BG_CANVAS)
        button_frame.pack(side=BOTTOM, fill=X, pady=20)

        Button(
            button_frame,
            text="Save Changes" if editing else "Save Goal",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 11, "bold"),
            padx=15,
            pady=8,
            command=save_goal
        ).pack(fill=X, padx=25, pady=30)

    def add_recurring_popup(self):

        pop, body, button_bar = self.create_form_popup("Add Recurring Transaction", 560, 680)

        Label(
            body,
            text="Add Recurring Transaction",
            font=("Segoe UI", 15, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22, pady=(18, 4))

        Label(
            body,
            text="Automatically remember your recurring income and expenses.",
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22, pady=(0, 16))

        # ---------------- TITLE ----------------

        Label(
            body,
            text="Title",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        title_entry = Entry(
            body,
            font=("Segoe UI", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        title_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # ---------------- AMOUNT ----------------

        Label(
            body,
            text="Amount (₹)",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        amount_entry = Entry(
            body,
            font=("Segoe UI", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        amount_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # ---------------- CATEGORY ----------------

        Label(
            body,
            text="Category",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        category_entry = Entry(
            body,
            font=("Segoe UI", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        category_entry.pack(fill=X, padx=22, ipady=7, pady=(4, 12))

        # ---------------- TYPE ----------------

        Label(
            body,
            text="Transaction Type",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        txn_type = StringVar(value="expense")

        OptionMenu(
            body,
            txn_type,
            "expense",
            "income"
        ).pack(fill=X, padx=22, pady=(4, 12))

        # ---------------- FREQUENCY ----------------

        Label(
            body,
            text="Frequency",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        frequency = StringVar(value="monthly")

        OptionMenu(
            body,
            frequency,
            "monthly",
            "weekly"
        ).pack(fill=X, padx=22, pady=(4, 12))

        # ---------------- NEXT DUE ----------------

        Label(
            body,
            text="Next Due Date (DD-MM-YYYY)",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=22)

        due_row = Frame(body, bg=BG_CANVAS)
        due_row.pack(fill=X, padx=22, pady=(4, 18))

        due_entry = Entry(
            due_row,
            font=("Segoe UI", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        due_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        self.attach_date_picker(due_entry).pack(side=LEFT, padx=(6, 0), ipady=4)

        due_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))

        # ---------------- SAVE FUNCTION ----------------

        def save_recurring():

            title = title_entry.get().strip()
            amount = amount_entry.get().strip()
            category = category_entry.get().strip()
            txn = txn_type.get()
            freq = frequency.get()
            next_due = due_entry.get().strip()

            if not title:
                messagebox.showwarning(
                    "Missing Data",
                    "Please enter a title."
                )
                return

            if not amount:
                messagebox.showwarning(
                    "Missing Data",
                    "Please enter an amount."
                )
                return

            try:
                amount = float(amount)
            except (TypeError, ValueError):
                messagebox.showerror(
                    "Invalid Amount",
                    "Amount must be numeric."
                )
                return

            try:
                datetime.strptime(next_due, "%d-%m-%Y")
            except (TypeError, ValueError):
                messagebox.showerror(
                    "Invalid Date",
                    "Date must be DD-MM-YYYY"
                )
                return

            cursor.execute("""
                INSERT INTO recurring_transactions
                (
                    title,
                    amount,
                    category,
                    txn_type,
                    frequency,
                    next_due_date
                )
                VALUES
                (?, ?, ?, ?, ?, ?)
            """,
            (
                title,
                amount,
                category,
                txn,
                freq,
                next_due
            ))

            conn.commit()

            messagebox.showinfo(
                "Success",
                "Recurring transaction added successfully."
            )

            pop.destroy()

            self.load_recurring()

        # ---------------- BUTTONS (always visible, pinned bottom bar) ----------------

        Button(
            button_bar,
            text="Cancel",
            command=pop.destroy,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10
        ).pack(side=LEFT)

        Button(
            button_bar,
            text="Save Recurring",
            command=save_recurring,
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10
        ).pack(side=RIGHT)


    def add_reminder_popup(self, reminder_id=None):
        editing = reminder_id is not None
        existing = None
        if editing:
            cursor.execute(
                "SELECT title, due_date, priority, notes, status FROM reminders WHERE id=?",
                (reminder_id,)
            )
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This reminder no longer exists.")
                return

        pop, body = self.create_scroll_popup(
            "Edit Reminder" if editing else "Add Reminder",
            560,
            760
        )
        pop.transient(self)
        pop.grab_set()

        Label(
            body,
            text="Edit Reminder" if editing else "Add Reminder",
            font=("Segoe UI", 18, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(20, 6))

        Label(
            body,
            text="Create a personal reminder for deadlines, fees, bills or important tasks.",
            font=("Segoe UI", 10),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(0, 16))

        Label(
            body,
            text="Reminder Title *",
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(8, 4))

        title_entry = Entry(
            body,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 11),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        title_entry.pack(fill=X, padx=25, ipady=8)

        Label(
            body,
            text="Due Date (DD-MM-YYYY) *",
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(15, 4))

        due_row = Frame(body, bg=BG_CANVAS)
        due_row.pack(fill=X, padx=25)
        due_entry = Entry(
            due_row,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 11),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        due_entry.pack(side=LEFT, fill=X, expand=True, ipady=8)
        self.attach_date_picker(due_entry).pack(side=LEFT, padx=(6, 0), ipady=5)

        Label(
            body,
            text="Priority",
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(15, 4))

        priority_var = StringVar(value="Medium")
        priority_menu = ttk.Combobox(
            body,
            textvariable=priority_var,
            values=["High", "Medium", "Low"],
            state="readonly",
            font=("Segoe UI", 11)
        )
        priority_menu.pack(fill=X, padx=25, ipady=5)

        Label(
            body,
            text="Notes (optional)",
            font=("Segoe UI", 10, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(15, 4))

        notes_entry = Entry(
            body,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 11),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        notes_entry.pack(fill=X, padx=25, ipady=8)

        if editing:
            ex_title, ex_due, ex_priority, ex_notes, ex_status = existing
            title_entry.insert(0, ex_title or "")
            due_entry.insert(0, ex_due or datetime.today().strftime("%d-%m-%Y"))
            priority_var.set(ex_priority if ex_priority in ("High", "Medium", "Low") else "Medium")
            notes_entry.insert(0, ex_notes or "")
        else:
            due_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))

        def save_reminder():
            title = title_entry.get().strip()
            if not title:
                messagebox.showerror("Missing Title", "Please enter a reminder title.")
                return

            ok, due_date = self.validate_date_input(due_entry.get())
            if not ok:
                return

            priority = priority_var.get()
            notes = notes_entry.get().strip()

            try:
                if editing:
                    cursor.execute(
                        "UPDATE reminders SET title=?, due_date=?, priority=?, notes=? WHERE id=?",
                        (title, due_date, priority, notes, reminder_id)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO reminders (title, due_date, status, priority, notes) VALUES (?, ?, ?, ?, ?)",
                        (title, due_date, "Pending", priority, notes)
                    )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this reminder:\n{e}")
                return

            pop.destroy()
            self.load_reminders()

        Button(
            body,
            text="Save Changes" if editing else "Save Reminder",
            bg=COLOR_PURPLE,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=12,
            command=save_reminder
        ).pack(fill=X, padx=25, pady=25)

    def get_reminder_state(self, due_date, status):
        """Returns 'completed', 'overdue', or 'pending'."""
        if status == "Completed":
            return "completed"
        try:
            due_dt = datetime.strptime(due_date, "%d-%m-%Y").date()
            if due_dt < datetime.today().date():
                return "overdue"
        except (TypeError, ValueError):
            pass
        return "pending"

    def mark_reminder_complete(self, reminder_id):
        cursor.execute("UPDATE reminders SET status='Completed' WHERE id=?", (reminder_id,))
        conn.commit()
        self.load_reminders()

    def mark_reminder_pending(self, reminder_id):
        cursor.execute("UPDATE reminders SET status='Pending' WHERE id=?", (reminder_id,))
        conn.commit()
        self.load_reminders()

    def delete_reminder(self, reminder_id):
        cursor.execute("SELECT title FROM reminders WHERE id=?", (reminder_id,))
        row = cursor.fetchone()
        if not row:
            messagebox.showerror("Error", "Reminder not found.")
            return

        confirm = messagebox.askyesno("Delete Reminder", f"Are you sure you want to delete '{row[0]}'?")
        if not confirm:
            return

        cursor.execute("DELETE FROM reminders WHERE id=?", (reminder_id,))
        conn.commit()
        self.load_reminders()

    def create_form_popup(self, title, width=560, height=640):
        """
        Popup layout with a scrollable form area and a Save/Cancel button
        bar that is PINNED to the bottom and always visible — unlike
        create_scroll_popup (used elsewhere), whose buttons live inside
        the scrollable region. The button bar is packed with side=BOTTOM
        BEFORE the scrollable area, which matters: Tk's pack() assigns
        cavity space in call order, so an expand=True widget packed
        first will greedily claim the leftover cavity and starve
        anything packed after it — packing the fixed bottom bar first
        guarantees it always gets its space, regardless of how tall the
        form content grows.

        Returns (pop, body, button_bar): pack form fields into `body`
        (it scrolls if content exceeds the popup height), pack
        Save/Cancel buttons into `button_bar` (always visible).
        """
        pop = Toplevel(self)
        pop.title(title)
        pop.geometry(f"{width}x{height}")
        pop.configure(bg=BG_CANVAS)

        outer = Frame(pop, bg=BG_CANVAS)
        outer.pack(fill=BOTH, expand=True)

        # Fixed bottom bar — packed FIRST so it always claims its space.
        button_bar = Frame(outer, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=22, pady=14)
        button_bar.pack(side=BOTTOM, fill=X)

        # Scrollable content area gets whatever's left.
        content_container = Frame(outer, bg=BG_CANVAS)
        content_container.pack(side=TOP, fill=BOTH, expand=True)

        canvas = Canvas(content_container, bg=BG_CANVAS, highlightthickness=0)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = Scrollbar(content_container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        body = Frame(canvas, bg=BG_CANVAS)
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def resize_body(event):
            canvas.itemconfig(window_id, width=event.width)
        canvas.bind("<Configure>", resize_body)

        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # Scoped to hover-over-this-popup only, so it doesn't leak into
        # other open windows once this popup closes.
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        return pop, body, button_bar

    def create_scroll_popup(self, title, width=560, height=700):

        pop = Toplevel(self)
        pop.title(title)
        pop.geometry(f"{width}x{height}")
        pop.configure(bg=BG_CANVAS)

        container = Frame(pop, bg=BG_CANVAS)
        container.pack(fill=BOTH, expand=True)

        canvas = Canvas(
            container,
            bg=BG_CANVAS,
            highlightthickness=0
        )
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = Scrollbar(
            container,
            orient="vertical",
            command=canvas.yview
        )
        scrollbar.pack(side=RIGHT, fill=Y)

        canvas.configure(yscrollcommand=scrollbar.set)

        body = Frame(canvas, bg=BG_CANVAS)

        window_id = canvas.create_window(
            (0, 0),
            window=body,
            anchor="nw"
        )

        def resize_body(event):
            canvas.itemconfig(window_id, width=event.width)

        canvas.bind("<Configure>", resize_body)

        body.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        def _mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _mousewheel)

        return pop, body

    def attach_date_picker(self, entry_widget, initial_date=None):
        """
        Attaches a small 📅 button next to `entry_widget` that opens a
        lightweight calendar popup (no external dependency needed) and
        writes the chosen date back into entry_widget as DD-MM-YYYY.
        Call this right after packing/gridding entry_widget; it inserts
        the button using the same geometry manager entry_widget uses.
        """
        try:
            base_date = datetime.strptime(initial_date, "%d-%m-%Y") if initial_date else datetime.today()
        except (TypeError, ValueError):
            base_date = datetime.today()

        state = {"year": base_date.year, "month": base_date.month}

        def open_picker():
            picker = Toplevel(self)
            picker.title("Select Date")
            picker.configure(bg=BG_CANVAS)
            picker.resizable(False, False)
            picker.grab_set()

            header = Frame(picker, bg=BG_CANVAS)
            header.pack(fill=X, padx=12, pady=(12, 6))

            month_label = Label(header, font=("Segoe UI", 11, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS)
            month_label.pack(side=LEFT, expand=True)

            grid_f = Frame(picker, bg=BG_CANVAS)
            grid_f.pack(padx=12, pady=(0, 12))

            def render():
                for w in grid_f.winfo_children():
                    w.destroy()
                import calendar as _cal
                month_label.config(text=datetime(state["year"], state["month"], 1).strftime("%B %Y"))
                for i, wd in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                    Label(grid_f, text=wd, font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS, width=4).grid(row=0, column=i, pady=(0, 4))
                cal = _cal.Calendar(firstweekday=0)
                row = 1
                for week in cal.monthdayscalendar(state["year"], state["month"]):
                    for col, day in enumerate(week):
                        if day == 0:
                            Label(grid_f, text="", bg=BG_CANVAS, width=4).grid(row=row, column=col)
                        else:
                            def pick(d=day):
                                chosen = datetime(state["year"], state["month"], d)
                                entry_widget.delete(0, END)
                                entry_widget.insert(0, chosen.strftime("%d-%m-%Y"))
                                picker.destroy()
                            Button(
                                grid_f, text=str(day), font=("Segoe UI", 9), width=3,
                                bg=BG_CANVAS, fg=TEXT_MAIN, bd=0, cursor="hand2",
                                activebackground=COLOR_ACCENT, activeforeground="white",
                                command=pick
                            ).grid(row=row, column=col, padx=1, pady=1)
                    row += 1

            def change_month(delta):
                m = state["month"] + delta
                y = state["year"]
                if m > 12:
                    m = 1; y += 1
                elif m < 1:
                    m = 12; y -= 1
                state["month"], state["year"] = m, y
                render()

            Button(header, text="◀", font=("Segoe UI", 9, "bold"), bd=0, bg=BG_CANVAS, fg=TEXT_MAIN, cursor="hand2", command=lambda: change_month(-1)).pack(side=LEFT)
            Button(header, text="▶", font=("Segoe UI", 9, "bold"), bd=0, bg=BG_CANVAS, fg=TEXT_MAIN, cursor="hand2", command=lambda: change_month(1)).pack(side=RIGHT)

            render()

        btn = Button(
            entry_widget.master,
            text="📅",
            font=("Segoe UI", 10),
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            cursor="hand2",
            command=open_picker
        )
        return btn

    def get_dropdown_options(self, table, column, defaults):
        """
        Builds a clean dropdown option list from a DB column plus a set of
        sensible defaults, merged and de-duplicated CASE-INSENSITIVELY so
        that e.g. 'Salary' and 'salary' never show up as two separate
        options. Every option is normalized to Title Case for display and
        for what actually gets saved back to the DB. Always ends with
        "Other" so the user can still type something custom.
        """
        cursor.execute(f"SELECT DISTINCT {column} FROM {table} WHERE {column} IS NOT NULL AND {column} != ''")
        db_values = [row[0] for row in cursor.fetchall()]

        normalized = {}
        for val in list(defaults) + db_values:
            clean = (val or "").strip()
            if not clean:
                continue
            key = clean.lower()
            if key not in normalized:
                normalized[key] = clean.title()

        return sorted(normalized.values()) + ["Other"]

    def validate_amount_input(self, raw_text, field_label="Amount"):
        """
        Shared amount validator used across Expense/Income/Budget popups.
        Returns (True, float_value) or (False, None) after showing a
        friendly error message box.
        """
        raw_text = (raw_text or "").strip()
        if not raw_text:
            messagebox.showerror("Missing Amount", f"Please enter {field_label.lower()}.")
            return False, None
        try:
            value = float(raw_text)
        except ValueError:
            messagebox.showerror("Invalid Amount", f"{field_label} must be a number (e.g. 250 or 250.50).")
            return False, None
        if value <= 0:
            messagebox.showerror("Invalid Amount", f"{field_label} must be greater than zero.")
            return False, None
        return True, value

    def validate_date_input(self, raw_text):
        """
        Shared date validator (expects DD-MM-YYYY, matching the format
        used everywhere else in the app). Returns (True, text) or (False, None).
        """
        raw_text = (raw_text or "").strip()
        if not raw_text:
            messagebox.showerror("Missing Date", "Please enter a date.")
            return False, None
        try:
            datetime.strptime(raw_text, "%d-%m-%Y")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter the date as DD-MM-YYYY (e.g. 05-07-2026).")
            return False, None
        return True, raw_text

    def add_income_popup(self, income_id=None):
        """
        income_id=None -> Add mode. income_id=<int> -> Edit mode.
        """
        editing = income_id is not None
        existing = None
        if editing:
            cursor.execute("SELECT source, amount, date, notes FROM income WHERE id=?", (income_id,))
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This income entry no longer exists.")
                return

        pop, body = self.create_scroll_popup("Edit Income" if editing else "Add Income", 560, 700)

        Label(
            body,
            text="Edit Income Entry" if editing else "Add Income Entry",
            font=("Segoe UI", 16, "bold"),
            fg=TEXT_MAIN,
            bg=BG_CANVAS
        ).pack(anchor="w", padx=25, pady=(20, 15))

        # ---- Income Source (dropdown fed from DB + custom) ---- #
        Label(body, text="Income Source", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(0, 4))

        default_sources = ["Salary", "Freelance", "Investment", "Gift", "Business", "Interest", "Rent"]
        sources = self.get_dropdown_options("income", "source", default_sources)

        source_var = StringVar(value=sources[0])
        source_menu = ttk.Combobox(body, textvariable=source_var, values=sources, state="readonly", font=("Segoe UI", 11))
        source_menu.pack(fill=X, padx=25, ipady=5)

        other_label = Label(body, text="Custom Source", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS)
        other_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)

        def source_changed(event=None):
            if source_var.get() == "Other":
                other_label.pack(anchor="w", padx=25, pady=(14, 4))
                other_entry.pack(fill=X, padx=25, ipady=7)
            else:
                other_label.pack_forget()
                other_entry.pack_forget()

        source_menu.bind("<<ComboboxSelected>>", source_changed)

        # ---- Amount ---- #
        Label(body, text="Amount (₹)", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(14, 4))
        amount_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        amount_entry.pack(fill=X, padx=25, ipady=7)

        # ---- Date (with picker) ---- #
        Label(body, text="Date (DD-MM-YYYY)", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(14, 4))
        date_row = Frame(body, bg=BG_CANVAS)
        date_row.pack(fill=X, padx=25)
        date_entry = Entry(date_row, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        date_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        date_entry.insert(0, datetime.today().strftime("%d-%m-%Y"))
        self.attach_date_picker(date_entry).pack(side=LEFT, padx=(6, 0), ipady=5)

        # ---- Notes ---- #
        Label(body, text="Notes (optional)", font=("Segoe UI", 10, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", padx=25, pady=(14, 4))
        notes_entry = Entry(body, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 11), highlightthickness=1, highlightbackground=BORDER_COLOR)
        notes_entry.pack(fill=X, padx=25, ipady=7)

        if editing:
            ex_source, ex_amount, ex_date, ex_notes = existing
            match = next((s for s in sources if s.lower() == (ex_source or "").strip().lower()), None)
            if match:
                source_var.set(match)
            else:
                source_var.set("Other")
                source_changed()
                other_entry.insert(0, ex_source or "")
            amount_entry.insert(0, str(ex_amount))
            date_entry.delete(0, END)
            date_entry.insert(0, ex_date or datetime.today().strftime("%d-%m-%Y"))
            notes_entry.insert(0, ex_notes or "")

        def save_income():
            source = source_var.get()
            if source == "Other":
                source = other_entry.get().strip().title()
            if not source:
                messagebox.showerror("Missing Source", "Please choose or enter an income source.")
                return

            ok, amount = self.validate_amount_input(amount_entry.get(), "Amount")
            if not ok:
                return

            ok, date_val = self.validate_date_input(date_entry.get())
            if not ok:
                return

            notes = notes_entry.get().strip()

            try:
                if editing:
                    cursor.execute(
                        "UPDATE income SET source=?, amount=?, date=?, notes=? WHERE id=?",
                        (source, amount, date_val, notes, income_id)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO income (source, amount, date, notes) VALUES (?, ?, ?, ?)",
                        (source, amount, date_val, notes)
                    )
                conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save this income entry:\n{e}")
                return

            pop.destroy()
            self.load_income()

        Button(
            body,
            text="Save Changes" if editing else "Save Income",
            bg=COLOR_INCOME,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=10,
            command=save_income
        ).pack(fill=X, padx=25, pady=25)

    def process_due_recurring(self, silent=False):
        """
        Walks every recurring transaction and, for each one whose
        next_due_date is today or in the past, generates the matching
        income/expense row(s) and advances next_due_date forward until it's
        in the future again (so re-opening the app after being away for a
        while catches up on every missed occurrence, not just one).

        silent=True is used for automatic background runs (e.g. on app
        startup or whenever a tab that shows this data is opened) so the
        user isn't interrupted with a popup every time. silent=False is
        used for the manual "Process Due Transactions" button.
        """
        today = datetime.today().date()

        cursor.execute("""
            SELECT id, title, amount, category, txn_type, frequency, next_due_date
            FROM recurring_transactions
        """)
        rows = cursor.fetchall()

        processed = 0

        for row in rows:
            rid, title, amount, category, txn_type, frequency, next_due = row

            try:
                due_date = datetime.strptime(next_due, "%d-%m-%Y").date()
            except (TypeError, ValueError):
                continue

            # Catch up on every missed occurrence, not just the first one.
            # Guard against a bad/zero frequency causing an infinite loop.
            safety_counter = 0
            while due_date <= today and safety_counter < 500:
                safety_counter += 1

                occurrence_date = due_date.strftime("%d-%m-%Y")

                # Never create a duplicate: skip if this exact recurring
                # transaction already generated a row for this exact date.
                if txn_type == "expense":
                    cursor.execute(
                        "SELECT 1 FROM expenses WHERE name=? AND category=? AND amount=? AND date=?",
                        (title, category if category else "Other", amount, occurrence_date)
                    )
                    already_exists = cursor.fetchone() is not None
                    if not already_exists:
                        cursor.execute(
                            "INSERT INTO expenses (name, amount, category, date) VALUES (?, ?, ?, ?)",
                            (title, amount, category if category else "Other", occurrence_date)
                        )
                else:
                    cursor.execute(
                        "SELECT 1 FROM income WHERE source=? AND amount=? AND date=?",
                        (title, amount, occurrence_date)
                    )
                    already_exists = cursor.fetchone() is not None
                    if not already_exists:
                        cursor.execute(
                            "INSERT INTO income (source, amount, date) VALUES (?, ?, ?)",
                            (title, amount, occurrence_date)
                        )

                if not already_exists:
                    processed += 1

                # Advance to the next occurrence
                if frequency == "weekly":
                    due_date = due_date + timedelta(days=7)
                else:
                    due_date = due_date + timedelta(days=30)

            cursor.execute(
                "UPDATE recurring_transactions SET next_due_date=? WHERE id=?",
                (due_date.strftime("%d-%m-%Y"), rid)
            )

        conn.commit()

        if silent:
            return processed

        if processed == 0:
            messagebox.showinfo("Recurring Transactions", "No recurring transactions are due today.")
        else:
            messagebox.showinfo("Recurring Transactions", f"{processed} recurring transaction(s) processed successfully.")

        self.load_recurring()

    def import_csv_logic(self):
        fp = filedialog.askopenfilename(filetypes=[("Comma Separated", "*.csv")])
        if not fp:
            return

        imported = 0
        skipped = 0

        try:
            with open(fp, 'r', encoding='utf-8-sig', newline='') as f:
                reader = csv.reader(f)
                next(reader, None)  # skip header row
                for row in reader:
                    if len(row) < 3:
                        skipped += 1
                        continue

                    name = (row[0] or "").strip()
                    category = (row[2] or "Others").strip() or "Others"

                    try:
                        amount = float(row[1])
                    except (ValueError, IndexError):
                        skipped += 1
                        continue

                    if not name or amount <= 0:
                        skipped += 1
                        continue

                    cursor.execute(
                        "INSERT INTO expenses (name, amount, category, date) VALUES (?, ?, ?, ?)",
                        (name, amount, category, datetime.today().strftime("%d-%m-%Y"))
                    )
                    imported += 1

            conn.commit()
        except OSError as e:
            messagebox.showerror("Import Failed", f"Could not open this file:\n{e}")
            return
        except csv.Error as e:
            messagebox.showerror("Import Failed", f"This file doesn't look like a valid CSV:\n{e}")
            return
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Could not save the imported rows:\n{e}")
            return

        if imported == 0:
            messagebox.showwarning(
                "Nothing Imported",
                "No valid rows were found. Expected columns: Name, Amount, Category (a header row is fine)."
            )
            return

        summary = f"Imported {imported} expense{'s' if imported != 1 else ''} successfully."
        if skipped:
            summary += f"\n{skipped} row{'s' if skipped != 1 else ''} skipped (missing or invalid data)."
        messagebox.showinfo("Import Complete", summary)
        self.load_expenses()

    def export_expenses_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Expenses CSV"
        )

        if not filepath:
            return

        cursor.execute("SELECT id, name, amount, category, date FROM expenses ORDER BY id DESC")
        rows = cursor.fetchall()

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Expense Name", "Amount", "Category", "Date"])
            writer.writerows(rows)

        messagebox.showinfo("Export Success", "Expenses exported successfully.")


    def export_income_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Income CSV"
        )

        if not filepath:
            return

        cursor.execute("SELECT id, source, amount, date FROM income ORDER BY id DESC")
        rows = cursor.fetchall()

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Income Source", "Amount", "Date"])
            writer.writerows(rows)

        messagebox.showinfo("Export Success", "Income exported successfully.")


    def export_budgets_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Budgets CSV"
        )

        if not filepath:
            return

        cursor.execute("SELECT category, amount, month_year FROM budgets ORDER BY category")
        rows = cursor.fetchall()

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Category", "Budget Amount", "Month/Year"])
            writer.writerows(rows)

        messagebox.showinfo("Export Success", "Budgets exported successfully.")


    def export_goals_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Savings Goals CSV"
        )

        if not filepath:
            return

        cursor.execute("""
            SELECT id, goal_name, target_amount, saved_amount, target_date
            FROM savings_goals
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Goal Name", "Target Amount", "Saved Amount", "Target Date"])
            writer.writerows(rows)

        messagebox.showinfo("Export Success", "Savings goals exported successfully.")


    def export_recurring_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Recurring Transactions CSV"
        )

        if not filepath:
            return

        cursor.execute("""
            SELECT id, title, amount, category, txn_type, frequency, next_due_date
            FROM recurring_transactions
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Title", "Amount", "Category", "Type", "Frequency", "Next Due Date"])
            writer.writerows(rows)

        messagebox.showinfo("Export Success", "Recurring transactions exported successfully.")


    def export_report_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                "PDF Export Unavailable",
                "PDF export requires the 'reportlab' package.\n\nInstall it with:\npip install reportlab"
            )
            return

        period_type = {"All Time": "all", "Monthly": "monthly", "Yearly": "yearly"}[getattr(self, "report_period_type", "All Time")]
        report = self.get_report_period_data(period_type, getattr(self, "report_period_key", None))

        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Report as PDF",
            initialfile=f"Cooper_Vault_Report_{report['label'].replace(' ', '_')}.pdf"
        )
        if not filepath:
            return

        try:
            doc = SimpleDocTemplate(filepath, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []

            elements.append(Paragraph(f"Cooper Vault — Financial Report", styles["Title"]))
            elements.append(Paragraph(f"Period: {report['label']}", styles["Normal"]))
            elements.append(Paragraph(f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}", styles["Normal"]))
            elements.append(Spacer(1, 14))

            summary_data = [
                ["Total Income", f"Rs. {report['total_income']:,.2f}"],
                ["Total Expenses", f"Rs. {report['total_expenses']:,.2f}"],
                ["Net Savings", f"Rs. {report['total_savings']:,.2f}"],
                ["Savings Rate", f"{report['savings_rate']:.1f}%"],
            ]
            summary_table = Table(summary_data, colWidths=[200, 200])
            summary_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F4F6")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E5E7EB")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]))
            elements.append(summary_table)
            elements.append(Spacer(1, 18))

            if report["highest_expense"]:
                name, amount, category, date = report["highest_expense"]
                elements.append(Paragraph(f"Highest Expense: {name} — Rs. {amount:,.2f} ({category or 'Other'}, {date})", styles["Normal"]))
            if report["highest_income"]:
                source, amount, date = report["highest_income"]
                elements.append(Paragraph(f"Highest Income: {source} — Rs. {amount:,.2f} ({date})", styles["Normal"]))
            elements.append(Spacer(1, 18))

            elements.append(Paragraph("Category Breakdown", styles["Heading2"]))
            if report["category_totals"]:
                cat_data = [["Category", "Amount", "% of Spend"]]
                total_spend = sum(v for _, v in report["category_totals"]) or 1
                for cat, amt in report["category_totals"]:
                    cat_data.append([cat, f"Rs. {amt:,.2f}", f"{amt/total_spend*100:.1f}%"])
                cat_table = Table(cat_data, colWidths=[180, 140, 100])
                cat_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E5E7EB")),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]))
                elements.append(cat_table)
            else:
                elements.append(Paragraph("No category data for this period.", styles["Normal"]))

            doc.build(elements)
            messagebox.showinfo("Export Success", f"PDF report saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not generate the PDF report:\n{e}")

    def export_report_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(
                "Excel Export Unavailable",
                "Excel export requires the 'openpyxl' package.\n\nInstall it with:\npip install openpyxl"
            )
            return

        period_type = {"All Time": "all", "Monthly": "monthly", "Yearly": "yearly"}[getattr(self, "report_period_type", "All Time")]
        report = self.get_report_period_data(period_type, getattr(self, "report_period_key", None))

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Report as Excel",
            initialfile=f"Cooper_Vault_Report_{report['label'].replace(' ', '_')}.xlsx"
        )
        if not filepath:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Summary"

            header_fill = PatternFill(start_color="111827", end_color="111827", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            ws["A1"] = "Cooper Vault — Financial Report"
            ws["A1"].font = Font(bold=True, size=14)
            ws["A2"] = f"Period: {report['label']}"
            ws["A3"] = f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}"

            ws["A5"] = "Metric"
            ws["B5"] = "Value"
            for cell in ("A5", "B5"):
                ws[cell].fill = header_fill
                ws[cell].font = header_font

            summary_rows = [
                ("Total Income", report["total_income"]),
                ("Total Expenses", report["total_expenses"]),
                ("Net Savings", report["total_savings"]),
                ("Savings Rate (%)", round(report["savings_rate"], 1)),
            ]
            for i, (label, value) in enumerate(summary_rows, start=6):
                ws[f"A{i}"] = label
                ws[f"B{i}"] = value

            if report["highest_expense"]:
                name, amount, category, date = report["highest_expense"]
                ws["A11"] = "Highest Expense"
                ws["B11"] = f"{name} — Rs. {amount:,.2f} ({category or 'Other'}, {date})"
            if report["highest_income"]:
                source, amount, date = report["highest_income"]
                ws["A12"] = "Highest Income"
                ws["B12"] = f"{source} — Rs. {amount:,.2f} ({date})"

            for col, width in [("A", 26), ("B", 40)]:
                ws.column_dimensions[col].width = width

            ws2 = wb.create_sheet("Category Breakdown")
            ws2["A1"] = "Category"
            ws2["B1"] = "Amount"
            ws2["C1"] = "% of Spend"
            for cell in ("A1", "B1", "C1"):
                ws2[cell].fill = header_fill
                ws2[cell].font = header_font

            total_spend = sum(v for _, v in report["category_totals"]) or 1
            for i, (cat, amt) in enumerate(report["category_totals"], start=2):
                ws2[f"A{i}"] = cat
                ws2[f"B{i}"] = amt
                ws2[f"C{i}"] = round(amt / total_spend * 100, 1)

            for col, width in [("A", 22), ("B", 16), ("C", 12)]:
                ws2.column_dimensions[col].width = width

            wb.save(filepath)
            messagebox.showinfo("Export Success", f"Excel report saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not generate the Excel report:\n{e}")

    def export_full_report_csv(self):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Full Financial Report"
        )

        if not filepath:
            return

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

        # ---------- EXPENSES ----------
            writer.writerow(["==== EXPENSES ===="])
            writer.writerow(["ID", "Expense Name", "Amount", "Category", "Date"])
            cursor.execute("SELECT id, name, amount, category, date FROM expenses ORDER BY id DESC")
            writer.writerows(cursor.fetchall())
            writer.writerow([])

        # ---------- INCOME ----------
            writer.writerow(["==== INCOME ===="])
            writer.writerow(["ID", "Income Source", "Amount", "Date"])
            cursor.execute("SELECT id, source, amount, date FROM income ORDER BY id DESC")
            writer.writerows(cursor.fetchall())
            writer.writerow([])

        # ---------- BUDGETS ----------
            writer.writerow(["==== BUDGETS ===="])
            writer.writerow(["Category", "Budget Amount", "Month/Year"])
            cursor.execute("SELECT category, amount, month_year FROM budgets ORDER BY category")
            writer.writerows(cursor.fetchall())
            writer.writerow([])

        # ---------- SAVINGS GOALS ----------
            writer.writerow(["==== SAVINGS GOALS ===="])
            writer.writerow(["ID", "Goal Name", "Target Amount", "Saved Amount", "Target Date"])
            cursor.execute("""
                SELECT id, goal_name, target_amount, saved_amount, target_date
                FROM savings_goals
                ORDER BY id DESC
            """)
            writer.writerows(cursor.fetchall())
            writer.writerow([])

        # ---------- RECURRING ----------
            writer.writerow(["==== RECURRING TRANSACTIONS ===="])
            writer.writerow(["ID", "Title", "Amount", "Category", "Type", "Frequency", "Next Due Date"])
            cursor.execute("""
                SELECT id, title, amount, category, txn_type, frequency, next_due_date
                FROM recurring_transactions
                ORDER BY id DESC
            """)
            writer.writerows(cursor.fetchall())

        messagebox.showinfo("Export Success", "Full financial report exported successfully.")

    # ==================================================================
    # PANEL: NOTES & TO-DO LIST
    # ==================================================================
    def load_notes_todo(self):
        self.highlight_sidebar("📓 Notepad & Tasks")
        self.clean_view()

        Label(content_frame, text="Notes & To-Do", font=("Segoe UI", 22, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w")
        Label(content_frame, text="Quick tasks and sticky notes — separate from your finances, just for staying organized", font=("Segoe UI", 11), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w", pady=(2, 20))

        split = Frame(content_frame, bg=BG_CANVAS)
        split.pack(fill=BOTH, expand=True)
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)

        # ------------------------------------------------------------
        # LEFT: TO-DO LIST
        # ------------------------------------------------------------
        todo_col = Frame(split, bg=BG_CANVAS)
        todo_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        todo_header = Frame(todo_col, bg=BG_CANVAS)
        todo_header.pack(fill=X, pady=(0, 10))
        Label(todo_header, text="✅ To-Do List", font=("Segoe UI", 14, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)

        cursor.execute("SELECT COUNT(*), COALESCE(SUM(is_done),0) FROM todos")
        total_todos, done_todos = cursor.fetchone()
        total_todos = total_todos or 0
        done_todos = done_todos or 0

        Label(
            todo_header,
            text=f"{done_todos}/{total_todos} done",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_MUTED,
            bg=BG_CANVAS
        ).pack(side=RIGHT)

        # Quick-add row
        quick_add_row = Frame(todo_col, bg=BG_CANVAS)
        quick_add_row.pack(fill=X, pady=(0, 10))

        new_task_entry = Entry(
            quick_add_row,
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        new_task_entry.pack(side=LEFT, fill=X, expand=True, ipady=8)
        new_task_entry.insert(0, "Add a quick task and press Enter...")
        new_task_entry.config(fg=TEXT_MUTED)

        def clear_placeholder(event):
            if new_task_entry.get() == "Add a quick task and press Enter...":
                new_task_entry.delete(0, END)
                new_task_entry.config(fg=TEXT_MAIN)

        new_task_entry.bind("<FocusIn>", clear_placeholder)

        def quick_add_task(event=None):
            task = new_task_entry.get().strip()
            if not task or task == "Add a quick task and press Enter...":
                return
            cursor.execute(
                "INSERT INTO todos (task, is_done, priority, due_date, created_date) VALUES (?, 0, 'Medium', NULL, ?)",
                (task, datetime.today().strftime("%d-%m-%Y"))
            )
            conn.commit()
            self.load_notes_todo()

        new_task_entry.bind("<Return>", quick_add_task)

        btn_quick_add = Button(
            quick_add_row,
            text="+ Add",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=12,
            command=quick_add_task
        )
        btn_quick_add.pack(side=LEFT, padx=(8, 0))
        self.add_hover_effect(btn_quick_add)

        btn_todo_options = Button(
            quick_add_row,
            text="⚙",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            command=lambda: self.add_todo_popup()
        )
        btn_todo_options.pack(side=LEFT, padx=(6, 0))
        self.add_hover_effect(btn_todo_options)

        # Scrollable task list
        todo_list_frame = Frame(todo_col, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=12, pady=12)
        todo_list_frame.pack(fill=BOTH, expand=True)

        cursor.execute("SELECT id, task, is_done, priority, due_date FROM todos ORDER BY is_done ASC, id DESC")
        todo_rows = cursor.fetchall()

        priority_colors = {"High": COLOR_DANGER, "Medium": "#D97706", "Low": TEXT_MUTED}

        if not todo_rows:
            Label(
                todo_list_frame,
                text="No tasks yet. Add one above to get started.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=10)
        else:
            for tid, task, is_done, priority, due_date in todo_rows:
                row = Frame(todo_list_frame, bg=BG_CANVAS)
                row.pack(fill=X, pady=4)

                check_var = BooleanVar(value=bool(is_done))

                def toggle_done(tid=tid, var=check_var):
                    cursor.execute("UPDATE todos SET is_done=? WHERE id=?", (1 if var.get() else 0, tid))
                    conn.commit()
                    self.load_notes_todo()

                cb = Checkbutton(
                    row,
                    variable=check_var,
                    bg=BG_CANVAS,
                    activebackground=BG_CANVAS,
                    command=toggle_done
                )
                cb.pack(side=LEFT)

                text_col = Frame(row, bg=BG_CANVAS)
                text_col.pack(side=LEFT, fill=X, expand=True, padx=(4, 0))

                task_font = ("Segoe UI", 10, "overstrike") if is_done else ("Segoe UI", 10)
                task_color = TEXT_MUTED if is_done else TEXT_MAIN
                Label(text_col, text=task, font=task_font, fg=task_color, bg=BG_CANVAS, anchor="w").pack(fill=X)

                meta_bits = [f"● {priority or 'Medium'}"]
                if due_date:
                    meta_bits.append(f"Due {due_date}")
                Label(
                    text_col,
                    text="   ".join(meta_bits),
                    font=("Segoe UI", 8, "bold"),
                    fg=priority_colors.get(priority, TEXT_MUTED),
                    bg=BG_CANVAS,
                    anchor="w"
                ).pack(fill=X)

                Button(
                    row,
                    text="🗑",
                    bg=BG_CANVAS,
                    fg=COLOR_DANGER,
                    bd=0,
                    font=("Segoe UI", 10),
                    command=lambda tid=tid: self.delete_todo(tid)
                ).pack(side=RIGHT)

                Button(
                    row,
                    text="✏",
                    bg=BG_CANVAS,
                    fg=TEXT_MUTED,
                    bd=0,
                    font=("Segoe UI", 10),
                    command=lambda tid=tid: self.add_todo_popup(todo_id=tid)
                ).pack(side=RIGHT)

        if done_todos > 0:
            Button(
                todo_col,
                text="Clear completed",
                bg=BG_CANVAS,
                fg=TEXT_MUTED,
                bd=0,
                font=("Segoe UI", 9, "underline"),
                command=self.clear_completed_todos
            ).pack(anchor="e", pady=(8, 0))

        # ------------------------------------------------------------
        # RIGHT: NOTES
        # ------------------------------------------------------------
        notes_col = Frame(split, bg=BG_CANVAS)
        notes_col.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        notes_header = Frame(notes_col, bg=BG_CANVAS)
        notes_header.pack(fill=X, pady=(0, 10))
        Label(notes_header, text="🗒 Notes", font=("Segoe UI", 14, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)

        btn_new_note = Button(
            notes_header,
            text="+ New Note",
            bg=COLOR_PURPLE,
            fg="white",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            command=lambda: self.add_note_popup()
        )
        btn_new_note.pack(side=RIGHT)
        self.add_hover_effect(btn_new_note)

        notes_scroll_frame = Frame(notes_col, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=12, pady=12)
        notes_scroll_frame.pack(fill=BOTH, expand=True)

        cursor.execute("SELECT id, title, content, color, updated_date FROM notes ORDER BY id DESC")
        note_rows = cursor.fetchall()

        if not note_rows:
            Label(
                notes_scroll_frame,
                text="No notes yet. Click \"+ New Note\" to jot something down.",
                font=("Segoe UI", 10),
                fg=TEXT_MUTED,
                bg=BG_CANVAS
            ).pack(anchor="w", pady=10)
        else:
            notes_grid = Frame(notes_scroll_frame, bg=BG_CANVAS)
            notes_grid.pack(fill=BOTH, expand=True)
            notes_grid.columnconfigure(0, weight=1)
            notes_grid.columnconfigure(1, weight=1)

            for idx, (nid, title, content, color, updated_date) in enumerate(note_rows):
                r, c = divmod(idx, 2)
                card = Frame(notes_grid, bg=color or "#FEF3C7", padx=12, pady=10, highlightthickness=1, highlightbackground=BORDER_COLOR)
                card.grid(row=r, column=c, sticky="nsew", padx=6, pady=6)
                notes_grid.rowconfigure(r, weight=1)

                Label(card, text=title, font=("Segoe UI", 10, "bold"), fg="#1F2937", bg=color or "#FEF3C7", anchor="w", wraplength=180).pack(fill=X)

                preview = (content or "")[:120] + ("..." if content and len(content) > 120 else "")
                Label(card, text=preview, font=("Segoe UI", 9), fg="#374151", bg=color or "#FEF3C7", anchor="w", justify=LEFT, wraplength=180).pack(fill=X, pady=(6, 8))

                btn_row = Frame(card, bg=color or "#FEF3C7")
                btn_row.pack(fill=X)
                Label(btn_row, text=updated_date or "", font=("Segoe UI", 7), fg="#6B7280", bg=color or "#FEF3C7").pack(side=LEFT)
                Button(btn_row, text="✏", bg=color or "#FEF3C7", fg="#374151", bd=0, font=("Segoe UI", 9),
                       command=lambda nid=nid: self.add_note_popup(note_id=nid)).pack(side=RIGHT)
                Button(btn_row, text="🗑", bg=color or "#FEF3C7", fg=COLOR_DANGER, bd=0, font=("Segoe UI", 9),
                       command=lambda nid=nid: self.delete_note(nid)).pack(side=RIGHT, padx=(0, 6))

    def delete_todo(self, todo_id):
        cursor.execute("DELETE FROM todos WHERE id=?", (todo_id,))
        conn.commit()
        self.load_notes_todo()

    def clear_completed_todos(self):
        if not messagebox.askyesno("Clear Completed", "Remove all completed tasks?"):
            return
        cursor.execute("DELETE FROM todos WHERE is_done=1")
        conn.commit()
        self.load_notes_todo()

    def add_todo_popup(self, todo_id=None):
        editing = todo_id is not None
        existing = None
        if editing:
            cursor.execute("SELECT task, priority, due_date FROM todos WHERE id=?", (todo_id,))
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This task no longer exists.")
                return

        pop = Toplevel(self)
        pop.title("Edit Task" if editing else "Add Task")
        pop.geometry("460x420")
        pop.configure(bg=BG_CANVAS)
        pop.resizable(False, False)
        pop.transient(self)
        pop.grab_set()

        Label(pop, text="Edit Task" if editing else "Add Task", font=("Segoe UI", 15, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22, pady=(18, 12))

        Label(pop, text="Task", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        task_entry = Entry(pop, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10), highlightthickness=1, highlightbackground=BORDER_COLOR)
        task_entry.pack(fill=X, padx=22, pady=(4, 12), ipady=7)

        Label(pop, text="Priority", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        priority_var = StringVar(value="Medium")
        ttk.Combobox(pop, textvariable=priority_var, values=["High", "Medium", "Low"], state="readonly", font=("Segoe UI", 10)).pack(fill=X, padx=22, pady=(4, 12))

        Label(pop, text="Due Date (optional, DD-MM-YYYY)", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        due_row = Frame(pop, bg=BG_CANVAS)
        due_row.pack(fill=X, padx=22)
        due_entry = Entry(due_row, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10), highlightthickness=1, highlightbackground=BORDER_COLOR)
        due_entry.pack(side=LEFT, fill=X, expand=True, ipady=7)
        self.attach_date_picker(due_entry).pack(side=LEFT, padx=(6, 0), ipady=5)

        if editing:
            ex_task, ex_priority, ex_due = existing
            task_entry.insert(0, ex_task or "")
            priority_var.set(ex_priority if ex_priority in ("High", "Medium", "Low") else "Medium")
            if ex_due:
                due_entry.insert(0, ex_due)

        def save_task():
            task = task_entry.get().strip()
            if not task:
                messagebox.showerror("Missing Task", "Please enter a task description.")
                return

            due_date = due_entry.get().strip()
            if due_date:
                ok, due_date = self.validate_date_input(due_date)
                if not ok:
                    return
            else:
                due_date = None

            if editing:
                cursor.execute(
                    "UPDATE todos SET task=?, priority=?, due_date=? WHERE id=?",
                    (task, priority_var.get(), due_date, todo_id)
                )
            else:
                cursor.execute(
                    "INSERT INTO todos (task, is_done, priority, due_date, created_date) VALUES (?, 0, ?, ?, ?)",
                    (task, priority_var.get(), due_date, datetime.today().strftime("%d-%m-%Y"))
                )
            conn.commit()
            pop.destroy()
            self.load_notes_todo()

        Button(
            pop,
            text="Save Changes" if editing else "Add Task",
            bg=COLOR_ACCENT,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=10,
            command=save_task
        ).pack(fill=X, padx=22, pady=20)

    def delete_note(self, note_id):
        if not messagebox.askyesno("Delete Note", "Delete this note permanently?"):
            return
        cursor.execute("DELETE FROM notes WHERE id=?", (note_id,))
        conn.commit()
        self.load_notes_todo()

    def add_note_popup(self, note_id=None):
        editing = note_id is not None
        existing = None
        if editing:
            cursor.execute("SELECT title, content, color FROM notes WHERE id=?", (note_id,))
            existing = cursor.fetchone()
            if not existing:
                messagebox.showerror("Not Found", "This note no longer exists.")
                return

        pop = Toplevel(self)
        pop.title("Edit Note" if editing else "New Note")
        pop.geometry("480x560")
        pop.configure(bg=BG_CANVAS)
        pop.resizable(False, False)
        pop.transient(self)
        pop.grab_set()

        # Same fix as Add Recurring / Add Bill: the button bar is its own
        # frame packed with side=BOTTOM *before* the expanding content
        # frame is packed, so pack() reserves its space first and it can
        # never get pushed outside the visible popup — previously every
        # widget (including Save) was packed directly onto `pop` in one
        # flat top-down stack with no dedicated bottom area, so a tall
        # Title + Note + Color stack could push Save below the fixed
        # 520px window with no way to reach it.
        main = Frame(pop, bg=BG_CANVAS)
        main.pack(fill=BOTH, expand=True)

        button_bar = Frame(main, bg=BG_CANVAS, highlightthickness=1, highlightbackground=BORDER_COLOR, padx=22, pady=14)
        button_bar.pack(side=BOTTOM, fill=X)

        content = Frame(main, bg=BG_CANVAS)
        content.pack(side=TOP, fill=BOTH, expand=True)

        Label(content, text="Edit Note" if editing else "New Note", font=("Segoe UI", 15, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22, pady=(18, 12))

        Label(content, text="Title", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        title_entry = Entry(content, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10), highlightthickness=1, highlightbackground=BORDER_COLOR)
        title_entry.pack(fill=X, padx=22, pady=(4, 12), ipady=7)

        Label(content, text="Note", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        content_text = Text(content, bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 10), highlightthickness=1, highlightbackground=BORDER_COLOR, height=8, wrap="word")
        content_text.pack(fill=BOTH, expand=True, padx=22, pady=(4, 12))

        Label(content, text="Color", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", padx=22)
        color_options = ["#FEF3C7", "#DBEAFE", "#DCFCE7", "#FCE7F3", "#E0E7FF"]
        color_var = StringVar(value=color_options[0])
        color_row = Frame(content, bg=BG_CANVAS)
        color_row.pack(anchor="w", padx=22, pady=(4, 14))

        swatches = []

        def pick_color(c):
            color_var.set(c)
            for sw, sc in swatches:
                sw.config(highlightthickness=2 if sc == c else 0, highlightbackground=TEXT_MAIN)

        for c in color_options:
            sw = Frame(color_row, bg=c, width=24, height=24, cursor="hand2")
            sw.pack(side=LEFT, padx=4)
            sw.pack_propagate(False)
            sw.bind("<Button-1>", lambda e, c=c: pick_color(c))
            swatches.append((sw, c))

        if editing:
            ex_title, ex_content, ex_color = existing
            title_entry.insert(0, ex_title or "")
            content_text.insert("1.0", ex_content or "")
            if ex_color in color_options:
                pick_color(ex_color)
            else:
                pick_color(color_options[0])
        else:
            pick_color(color_options[0])

        def save_note():
            title = title_entry.get().strip()
            note_content = content_text.get("1.0", END).strip()
            if not title:
                messagebox.showerror("Missing Title", "Please enter a note title.")
                return

            now_str = datetime.today().strftime("%d-%m-%Y")
            if editing:
                cursor.execute(
                    "UPDATE notes SET title=?, content=?, color=?, updated_date=? WHERE id=?",
                    (title, note_content, color_var.get(), now_str, note_id)
                )
            else:
                cursor.execute(
                    "INSERT INTO notes (title, content, color, created_date, updated_date) VALUES (?, ?, ?, ?, ?)",
                    (title, note_content, color_var.get(), now_str, now_str)
                )
            conn.commit()
            pop.destroy()
            self.load_notes_todo()

        # ---------------- BUTTONS (always visible, pinned bottom bar) ----------------

        Button(
            button_bar,
            text="Cancel",
            bg=BG_SIDEBAR,
            fg=TEXT_MAIN,
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10,
            command=pop.destroy
        ).pack(side=LEFT)

        Button(
            button_bar,
            text="Save Changes" if editing else "Save Note",
            bg=COLOR_PURPLE,
            fg="white",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=22,
            pady=10,
            command=save_note
        ).pack(side=RIGHT)


    # ==================================================================
    # PANEL: PROFILE
    # ==================================================================
    def load_profile(self):
        self.highlight_sidebar("👤 Profile")
        self.clean_view()

        auth_cursor.execute(
            "SELECT display_name, email, phone, monthly_income, dark_mode, profile_picture, username, created_date "
            "FROM users WHERE id=?",
            (self.current_user_id,)
        )
        row = auth_cursor.fetchone()
        if not row:
            Label(content_frame, text="Could not load your profile.", fg=COLOR_DANGER, bg=BG_CANVAS).pack(anchor="w")
            return
        name, email, phone, monthly_income, dark_mode, profile_picture, username, created_date = row
        display_name = name or username

        cursor.execute("SELECT COUNT(*) FROM expenses")
        expense_count = cursor.fetchone()[0] or 0
        cursor.execute("SELECT COUNT(*) FROM income")
        income_count = cursor.fetchone()[0] or 0
        current_streak, best_streak = self.get_user_streak()

        # ================================================================
        # HERO BANNER — gradient header, avatar, name, quick stat chips
        # ================================================================
        banner = Canvas(content_frame, height=204, highlightthickness=0, bd=0, cursor="hand2")
        banner.pack(fill=X, pady=(0, 24))

        photo_state = {"profile_picture": profile_picture}

        def change_photo():
            self._pick_profile_picture(pic_status, photo_state)

        def render_banner(event=None):
            self._paint_vertical_gradient(banner, (0x0B, 0x11, 0x20), (0x10, 0x3A, 0x35))
            banner.delete("content")
            w = banner.winfo_width() or 1000
            h = banner.winfo_height() or 204

            # Soft ambient depth blobs, echoing the auth-screen hero
            banner.create_oval(w - 260, -140, w + 60, 220, fill="#134E4A", outline="", tags="content")
            banner.create_oval(-120, h - 160, 180, h + 140, fill="#0B1120", outline="", tags="content")

            # Avatar: soft white ring + accent disc + initials + edit badge
            initials = "".join([p[0].upper() for p in display_name.split() if p])[:2] or "U"
            cx, cy, r_av = 78, h // 2 - 6, 48
            banner.create_oval(cx - r_av, cy - r_av, cx + r_av, cy + r_av, fill="#FFFFFF", outline="", tags="content")
            banner.create_oval(cx - r_av + 5, cy - r_av + 5, cx + r_av - 5, cy + r_av - 5, fill=COLOR_INCOME, outline="", tags="content")
            banner.create_text(cx, cy, text=initials, font=("Segoe UI", 21, "bold"), fill="white", tags="content")

            # Small camera badge, bottom-right of the avatar, hinting the
            # whole avatar is clickable to change the photo.
            bx, by, br = cx + r_av - 12, cy + r_av - 10, 16
            banner.create_oval(bx - br, by - br, bx + br, by + br, fill=COLOR_SAVINGS, outline="white", width=2, tags="content")
            banner.create_text(bx, by, text="📷", font=("Segoe UI", 9), tags="content")

            text_x = cx + r_av + 30
            banner.create_text(text_x, cy - 30, text=display_name, font=("Segoe UI", 20, "bold"), fill="white", anchor="w", tags="content")
            banner.create_text(text_x, cy, text=f"@{username}", font=("Segoe UI", 10), fill="#A7F3D0", anchor="w", tags="content")
            banner.create_text(text_x, cy + 26, text=f"Member since {created_date or '—'}", font=("Segoe UI", 8), fill="#6EE7B7", anchor="w", tags="content")

            # Stat chips (rounded pills), right-aligned, evenly spaced from the right edge
            chips = [
                (f"{expense_count + income_count}", "Transactions"),
                (f"{best_streak}d", "Best Streak"),
                (f"{current_streak}d", "Current Streak"),
            ]
            chip_w, chip_h, gap = 140, 70, 16
            start_x = w - 28 - (chip_w * len(chips) + gap * (len(chips) - 1))
            start_x = max(text_x + 240, start_x)
            for i, (value, caption) in enumerate(chips):
                x1 = start_x + i * (chip_w + gap)
                y1 = h // 2 - chip_h // 2
                # A faint offset layer behind each chip for a touch of lift
                self._rounded_rect(banner, x1 + 2, y1 + 3, x1 + chip_w + 2, y1 + chip_h + 3, radius=15,
                                    fill="#04241F", outline="", tags="content")
                # Tk can't do true alpha blending, so a light solid tint
                # stands in for a "translucent white" chip on the gradient.
                self._rounded_rect(banner, x1, y1, x1 + chip_w, y1 + chip_h, radius=15,
                                    fill="#ECFDF5", outline="", tags="content")
                banner.create_text(x1 + chip_w / 2, y1 + chip_h / 2 - 12, text=value,
                                    font=("Segoe UI", 15, "bold"), fill="#065F46", tags="content")
                banner.create_text(x1 + chip_w / 2, y1 + chip_h / 2 + 15, text=caption,
                                    font=("Segoe UI", 8, "bold"), fill=COLOR_INCOME, tags="content")

            # Clicking anywhere on the avatar opens the photo picker
            banner.tag_bind("content", "<Button-1>", None)
            avatar_hit = banner.create_oval(cx - r_av, cy - r_av, cx + r_av, cy + r_av, fill="", outline="", tags="avatar_hit")
            banner.tag_bind(avatar_hit, "<Button-1>", lambda e: change_photo())

        banner.bind("<Configure>", render_banner)

        # ================================================================
        # BODY — two-column shadow cards
        # ================================================================
        split = Frame(content_frame, bg=BG_CANVAS)
        split.pack(fill=BOTH, expand=True)
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)

        # ------------------------------------------------------------
        # LEFT: PERSONAL INFO
        # ------------------------------------------------------------
        left_outer, left = self._shadow_card(split, padx=28, pady=26, accent=COLOR_ACCENT)
        left_outer.grid(row=0, column=0, sticky="new", padx=(0, 14), pady=(0, 16))

        section_header = Frame(left, bg=BG_CANVAS)
        section_header.pack(fill=X, pady=(0, 16))
        Label(section_header, text="🪪", font=("Segoe UI", 14), bg=BG_CANVAS).pack(side=LEFT, padx=(0, 8))
        Label(section_header, text="Personal Info", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)

        # Photo
        photo_row = Frame(left, bg=BG_CANVAS)
        photo_row.pack(fill=X, pady=(0, 14))

        pic_status = Label(
            photo_row,
            text=("✓ Photo set" if profile_picture else "No photo set"),
            font=("Segoe UI", 9),
            fg=COLOR_INCOME if profile_picture else TEXT_MUTED,
            bg=BG_CANVAS
        )
        pic_status.pack(side=LEFT)

        btn_change_photo = Button(photo_row, text="Change Photo", bg=BG_SIDEBAR, fg=TEXT_MAIN, bd=0, font=("Segoe UI", 9),
                                   padx=10, pady=5, command=change_photo)
        btn_change_photo.pack(side=RIGHT)
        self.add_hover_effect(btn_change_photo)

        Label(left, text="Username", font=("Segoe UI", 9, "bold"), fg=TEXT_MUTED, bg=BG_CANVAS).pack(anchor="w")
        Label(left, text=f"@{username}", font=("Segoe UI", 10), fg=TEXT_MAIN, bg=BG_CANVAS).pack(anchor="w", pady=(0, 8))

        name_entry = self._labeled_entry(left, "Full Name")
        name_entry.insert(0, name or "")

        email_entry = self._labeled_entry(left, "Email")
        email_entry.insert(0, email or "")

        phone_entry = self._labeled_entry(left, "Phone")
        phone_entry.insert(0, phone or "")

        income_entry = self._labeled_entry(left, "Monthly Income (₹)")
        income_entry.insert(0, str(monthly_income) if monthly_income else "")

        save_status = Label(left, text="", font=("Segoe UI", 9), fg=COLOR_INCOME, bg=BG_CANVAS)
        save_status.pack(anchor="w", pady=(8, 0))

        def save_profile():
            new_name = name_entry.get().strip()
            new_email = email_entry.get().strip()
            new_phone = phone_entry.get().strip()
            income_raw = income_entry.get().strip()

            if not new_name:
                messagebox.showerror("Missing Name", "Name cannot be empty.")
                return

            if income_raw:
                ok, income_val = self.validate_amount_input(income_raw, "Monthly Income")
                if not ok:
                    return
            else:
                income_val = None

            try:
                auth_cursor.execute(
                    "UPDATE users SET display_name=?, email=?, phone=?, monthly_income=?, profile_picture=? WHERE id=?",
                    (new_name, new_email, new_phone, income_val, photo_state["profile_picture"], self.current_user_id)
                )
                auth_conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save your profile:\n{e}")
                return

            self.current_display_name = new_name
            self.current_profile_picture = photo_state["profile_picture"]
            save_status.config(text="✓ Saved")
            self.build_workspace()

        btn_save_profile = self._rounded_button(left, "Save Changes", save_profile, icon="💾")
        btn_save_profile.pack(fill=X, pady=(18, 0))

        # ------------------------------------------------------------
        # RIGHT: PREFERENCES + DATA
        # ------------------------------------------------------------
        right = Frame(split, bg=BG_CANVAS)
        right.grid(row=0, column=1, sticky="new", padx=(14, 0))

        theme_outer, theme_card = self._shadow_card(right, padx=28, pady=24, accent=COLOR_PURPLE)
        theme_outer.pack(fill=X, pady=(0, 16))

        theme_header = Frame(theme_card, bg=BG_CANVAS)
        theme_header.pack(fill=X, pady=(0, 12))
        Label(theme_header, text="🎨", font=("Segoe UI", 14), bg=BG_CANVAS).pack(side=LEFT, padx=(0, 8))
        Label(theme_header, text="Appearance", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)

        dark_var = BooleanVar(value=bool(dark_mode))

        def toggle_dark_mode():
            new_val = 1 if dark_var.get() else 0
            try:
                auth_cursor.execute("UPDATE users SET dark_mode=? WHERE id=?", (new_val, self.current_user_id))
                auth_conn.commit()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Could not save theme preference:\n{e}")
                return
            apply_theme(bool(new_val))
            self.build_workspace()

        Checkbutton(
            theme_card, text="Dark Mode", variable=dark_var, command=toggle_dark_mode,
            bg=BG_CANVAS, activebackground=BG_CANVAS, font=("Segoe UI", 10, "bold"), fg=TEXT_MAIN,
            selectcolor=BG_CANVAS
        ).pack(anchor="w")

        Label(theme_card, text="Switches the whole app between light and dark instantly.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_CANVAS, wraplength=320, justify=LEFT).pack(anchor="w", pady=(4, 0))

        data_outer, data_card = self._shadow_card(right, padx=28, pady=24, accent=COLOR_SAVINGS)
        data_outer.pack(fill=X)

        data_header = Frame(data_card, bg=BG_CANVAS)
        data_header.pack(fill=X, pady=(0, 12))
        Label(data_header, text="🗄", font=("Segoe UI", 14), bg=BG_CANVAS).pack(side=LEFT, padx=(0, 8))
        Label(data_header, text="Your Data", font=("Segoe UI", 13, "bold"), fg=TEXT_MAIN, bg=BG_CANVAS).pack(side=LEFT)

        btn_export_data = self._rounded_button(
            data_card, "Export Data (Full CSV)", self.export_full_report_csv,
            bg=BG_SIDEBAR, fg=TEXT_MAIN, icon="📤", height=44, font_size=11
        )
        btn_export_data.pack(fill=X, pady=5)

        btn_backup = self._rounded_button(
            data_card, "Backup Database", self.backup_database,
            bg=BG_SIDEBAR, fg=TEXT_MAIN, icon="💾", height=44, font_size=11
        )
        btn_backup.pack(fill=X, pady=5)

        btn_restore = self._rounded_button(
            data_card, "Restore Database", self.restore_database,
            bg=BG_SIDEBAR, fg=COLOR_DANGER, icon="♻", height=44, font_size=11
        )
        btn_restore.pack(fill=X, pady=5)

        Label(
            data_card,
            text="Backups save your entire database to a file you choose. Restoring replaces ALL current data — use with care.",
            font=("Segoe UI", 8), fg=TEXT_MUTED, bg=BG_CANVAS, wraplength=320, justify=LEFT
        ).pack(anchor="w", pady=(8, 0))


    def backup_database(self):
        # Backs up the CURRENTLY LOGGED-IN user's private database file
        # only — not other accounts' data.
        conn.commit()
        filepath = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db")],
            title="Save Database Backup",
            initialfile=f"cooper_{self.current_username}_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.db"
        )
        if not filepath:
            return
        try:
            shutil.copy2(_safe_db_filename(self.current_username), filepath)
            messagebox.showinfo("Backup Complete", f"Database backed up to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Backup Failed", f"Could not create backup:\n{e}")

    def restore_database(self):
        if not messagebox.askyesno(
            "Restore Database",
            "This will REPLACE your account's current data with the backup file you choose. "
            "Other accounts are not affected. This cannot be undone. Continue?"
        ):
            return

        filepath = filedialog.askopenfilename(
            title="Choose a Database Backup to Restore",
            filetypes=[("SQLite Database", "*.db")]
        )
        if not filepath:
            return

        global conn, cursor
        own_db_path = _safe_db_filename(self.current_username)
        try:
            conn.commit()
            conn.close()
            shutil.copy2(filepath, own_db_path)
            conn, cursor = get_user_db(self.current_username)
        except Exception as e:
            messagebox.showerror("Restore Failed", f"Could not restore this backup:\n{e}")
            return

        messagebox.showinfo("Restore Complete", "Your data has been restored. Please log in again.")
        self.logout()

if __name__ == "__main__":
    app = CooperAppMaster()
    app.mainloop()