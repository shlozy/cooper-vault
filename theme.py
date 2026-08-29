"""
Cooper Vault — shared UI color palette.

This module owns the *data*: the light/dark palettes and the fixed
accent/auth-screen colors. It does not own any mutable "current theme"
state itself — app.py keeps its own module-level BG_CANVAS / TEXT_MAIN /
etc. names (reassigned by apply_theme()) because ~100+ page-rendering
functions in app.py read those names as bare globals. If this module
mutated its own globals() instead, names already imported into app.py's
namespace wouldn't see the change, silently breaking dark-mode
switching. apply_theme() here is written to take the caller's globals()
dict explicitly, so app.py can do:

    from theme import LIGHT_THEME, DARK_THEME, apply_theme as _apply_theme_palette

    def apply_theme(is_dark):
        _apply_theme_palette(is_dark, globals())

...which updates app.py's own globals correctly while keeping the
palette values themselves centralized here.
"""

# ---------------- FIGMA LIGHT DESIGN PALETTE (initial/default values) ---------------- #
BG_CANVAS = "#FFFFFF"       # Main clean white background
BG_SIDEBAR = "#F9FAFB"      # Soft off-white for navigation panel
CARD_BG = "#FFFFFF"         # Card surface (same as canvas in light mode)
TEXT_MAIN = "#1F2937"       # Charcoal black for readable text
TEXT_MUTED = "#6B7280"      # Subtle gray for secondary labels
BORDER_COLOR = "#E5E7EB"    # Light gray borders for clean grid look

LIGHT_THEME = {
    "BG_CANVAS": "#FFFFFF", "BG_SIDEBAR": "#F9FAFB", "CARD_BG": "#FFFFFF",
    "TEXT_MAIN": "#1F2937", "TEXT_MUTED": "#6B7280", "BORDER_COLOR": "#E5E7EB",
}
# "Midnight Emerald" dark palette — deliberately layered so the sidebar,
# page background, and cards each read as a distinct depth (darkest to
# lightest), instead of one flat near-black everywhere.
DARK_THEME = {
    "BG_CANVAS": "#101827", "BG_SIDEBAR": "#0D1525", "CARD_BG": "#151F32",
    "TEXT_MAIN": "#F8FAFC", "TEXT_MUTED": "#94A3B8", "BORDER_COLOR": "#243247",
}


def apply_theme(is_dark, target_globals):
    """
    Copies the chosen palette (dark or light) into `target_globals` — the
    caller's own globals() dict. Since every load_* page function in
    app.py reads BG_CANVAS/TEXT_MAIN/etc. as bare module-level names,
    this must mutate the *caller's* namespace, not this module's, which
    is why it's an explicit parameter rather than this module updating
    its own globals().
    """
    palette = DARK_THEME if is_dark else LIGHT_THEME
    for key, value in palette.items():
        target_globals[key] = value


# ---------------- FIXED ACCENT COLORS (same in light and dark mode) ---------------- #
COLOR_INCOME = "#10B981"    # Mint Emerald Green
COLOR_EXPENSE = "#F97316"   # Tangerine Orange
COLOR_SAVINGS = "#06B6D4"   # Electric Cyan Blue
COLOR_ACCENT = "#10B981"    # Emerald — the app's one primary accent
COLOR_INFO = "#06B6D4"      # Cyan — reserved for informational (non-primary) highlights
COLOR_DANGER = "#DC2626"    # Used for delete/overdue/error actions everywhere
COLOR_PURPLE = "#8B5CF6"    # Used for reminders/notes/Cooper AI accent actions

# ---------------- AUTH SCREENS: FIXED DARK FINTECH PALETTE ---------------- #
# The Login / Sign Up / Forgot Password screens use their own always-dark
# palette regardless of the app's light/dark theme toggle (which only
# affects the workspace once logged in) — a premium landing scene reads
# better dark, and it gives the brand a consistent "front door" look.
AUTH_BG = "#0B1120"          # Deep navy-black canvas background
AUTH_PANEL = "#0F172A"       # Slightly lighter panel base
AUTH_CARD = "#111827"        # Login card surface
AUTH_INPUT = "#1E293B"       # Input field fill
AUTH_BORDER = "#1F2937"      # Hairline borders on dark surfaces
AUTH_TEXT = "#F3F4F6"        # Primary text on dark
AUTH_MUTED = "#94A3B8"       # Secondary text on dark
AUTH_EMERALD = "#10B981"
AUTH_CYAN = "#06B6D4"
AUTH_PURPLE = "#8B5CF6"
