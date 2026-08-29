"""
Cooper Vault — authentication helpers.

Pure functions only: password hashing and registration-field
validation (username format, password strength, email format). No
Tkinter, no database access — these never touch `conn`/`cursor` or any
SQLite table directly, so they're trivially safe to import anywhere
and to unit-test in isolation.

Login/registration/logout/session-handling *flow* (the actual
`do_login`, `do_signup`, `logout` functions that call these helpers
and then touch the database + UI) stays in app.py for now, since those
functions are tightly wound into the Toplevel/Tkinter widgets they
build and the app's page-navigation methods — pulling just the
validation math out here is the safe, low-risk slice of "authentication"
to extract at this stage.
"""
import hashlib
import re
import secrets

def generate_salt():
    return secrets.token_hex(16)


def hash_with_salt(text, salt):
    return hashlib.sha256((salt + text).encode("utf-8")).hexdigest()


# ---------------- REGISTRATION VALIDATION ---------------- #
USERNAME_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9_]{3,19}$")
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PASSWORD_SPECIAL_CHARS = r"!@#$%^&*_\-+=?"
PASSWORD_SPECIAL_PATTERN = re.compile(f"[{PASSWORD_SPECIAL_CHARS}]")


def validate_username_format(username):
    """
    Returns (is_valid, message). Checked BEFORE any database lookup —
    a cheap, purely local check should reject an obviously-malformed
    username before spending a query on it.
    """
    if not username:
        return False, "Please enter a username."
    if " " in username:
        return False, (
            "Invalid username.\n\nUsername must:\n"
            "• Be 4–20 characters long\n"
            "• Start with a letter\n"
            "• Contain only letters, numbers and _\n"
            "• Not contain spaces"
        )
    if not USERNAME_PATTERN.match(username):
        return False, (
            "Invalid username.\n\nUsername must:\n"
            "• Be 4–20 characters long\n"
            "• Start with a letter\n"
            "• Contain only letters, numbers and _\n"
            "• Not contain spaces"
        )
    return True, ""


def password_strength_checks(password):
    """
    Returns an ordered dict of {requirement label: bool met}. Used both
    for the final hard validation on submit and for the live strength
    indicator while typing — one source of truth for both.
    """
    return {
        "8+ characters": len(password) >= 8,
        "Uppercase letter": bool(re.search(r"[A-Z]", password)),
        "Lowercase letter": bool(re.search(r"[a-z]", password)),
        "Number": bool(re.search(r"[0-9]", password)),
        "Special character": bool(PASSWORD_SPECIAL_PATTERN.search(password)),
    }


def password_strength_label(password):
    """Returns (label, color_key) where color_key is 'weak'/'medium'/'strong'."""
    if not password:
        return "", None
    met = sum(password_strength_checks(password).values())
    if met <= 2:
        return "Weak", "weak"
    if met <= 4:
        return "Medium", "medium"
    return "Strong", "strong"


def validate_password_strength(password):
    """
    Returns (is_valid, message). On failure, the message lists exactly
    which requirements are missing (✓/✗) rather than a bare "invalid".
    """
    checks = password_strength_checks(password)
    if all(checks.values()):
        return True, ""
    lines = [f"{'✓' if met else '✗'} {label}" for label, met in checks.items()]
    message = "Password is too weak.\n\nPassword requirements:\n\n" + "\n".join(lines)
    return False, message


def validate_email_format(email):
    """Returns (is_valid, message). Only meaningful when an email field is present."""
    if not email:
        return True, ""  # email is optional wherever it's not a required field
    if not EMAIL_PATTERN.match(email):
        return False, "Please enter a valid email address (e.g. name@example.com)."
    return True, ""
