"""
Cooper Vault — database layer.

Owns everything SQLite-related: the shared auth-only database
(data/users.db, `users` table only) and the per-user private financial
databases (data/cooper_<username>.db). No Tkinter/UI imports here —
this module is safe to import and use completely independently of the
GUI (e.g. for a future CLI tool, tests, or scripts).

Exposes:
    auth_conn, auth_cursor   — the always-open connection to the shared
                                auth database. Never reassigned after
                                creation (only .execute()/.commit()
                                called on it), so it's safe for app.py
                                to import these two names directly.
    get_user_db(username)    — opens (creating on first use) one user's
                                private database and returns
                                (conn, cursor) for it. app.py calls this
                                at login/registration and rebinds its
                                OWN `conn`/`cursor` globals to the
                                result — that reassignment happens in
                                app.py, not here, since app.py is what
                                actually owns those mutable names.
"""
import os
import re
import sqlite3
from datetime import datetime

# ---------------- DATABASE ARCHITECTURE: ONE DB PER USER ---------------- #
# Every account gets its own private SQLite file — data/cooper_<username>.db
# — holding ALL of that user's financial + personal tables (expenses,
# income, budgets, savings_goals, recurring_transactions, reminders,
# bills_emi, system_streaks, messages, todos, notes). The ONLY thing
# that's ever shared across accounts is data/users.db, and it holds
# nothing but login/profile fields — never a single transaction.
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
AUTH_DB_PATH = os.path.join(DATA_DIR, "users.db")
LEGACY_DB_PATH = "cooper_vault.db"


def _add_column_if_missing(cur, table, column, coltype):
    cur.execute(f"PRAGMA table_info({table})")
    existing_cols = [row[1] for row in cur.fetchall()]
    if column not in existing_cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype}")


def _create_auth_schema(cur):
    """The shared authentication database: login + profile fields only."""
    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        display_name TEXT,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        security_question TEXT,
        security_answer_hash TEXT,
        profile_picture TEXT,
        email TEXT,
        phone TEXT,
        monthly_income REAL,
        dark_mode INTEGER DEFAULT 0,
        remember_me INTEGER DEFAULT 0
    )
    """)
    _add_column_if_missing(cur, "users", "created_date", "TEXT")
    # Existing accounts created before this column existed have no join
    # date — backfill with today's date once, rather than leaving it blank.
    cur.execute(
        "UPDATE users SET created_date=? WHERE created_date IS NULL OR created_date=''",
        (datetime.today().strftime("%d-%m-%Y"),)
    )


def _create_user_data_schema(cur):
    """Every personal/financial table, created inside ONE user's private database."""
    cur.execute("""CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, amount REAL, category TEXT, date TEXT)""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS income(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            amount REAL NOT NULL,
            date TEXT NOT NULL
        )
        """)
    cur.execute("""CREATE TABLE IF NOT EXISTS budgets (
        category TEXT, amount REAL, month_year TEXT, PRIMARY KEY(category, month_year))""")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS savings_goals(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        goal_name TEXT NOT NULL,
        target_amount REAL NOT NULL,
        saved_amount REAL NOT NULL DEFAULT 0,
        target_date TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS recurring_transactions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        amount REAL NOT NULL,
        category TEXT,
        txn_type TEXT NOT NULL,      -- "expense" or "income"
        frequency TEXT NOT NULL,     -- "weekly" or "monthly"
        next_due_date TEXT NOT NULL
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS reminders(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        due_date TEXT NOT NULL,
        status TEXT DEFAULT 'Pending'
    )
    """)
    cur.execute("""CREATE TABLE IF NOT EXISTS system_streaks (
        id INTEGER PRIMARY KEY, current_streak INTEGER, best_streak INTEGER, daily_goal REAL)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT, sender TEXT, body TEXT, status TEXT)""")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bills_emi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        amount REAL NOT NULL,
        category TEXT,
        due_date TEXT,
        status TEXT DEFAULT 'Unpaid'
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS todos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        task TEXT NOT NULL,
        is_done INTEGER DEFAULT 0,
        priority TEXT DEFAULT 'Medium',
        due_date TEXT,
        created_date TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        content TEXT,
        color TEXT DEFAULT '#FEF3C7',
        created_date TEXT,
        updated_date TEXT
    )
    """)

    _add_column_if_missing(cur, "expenses", "notes", "TEXT")
    _add_column_if_missing(cur, "income", "notes", "TEXT")
    _add_column_if_missing(cur, "reminders", "priority", "TEXT")
    _add_column_if_missing(cur, "reminders", "notes", "TEXT")

    # Any reminder rows created before the priority column existed default to Medium
    cur.execute("UPDATE reminders SET priority='Medium' WHERE priority IS NULL OR priority=''")

    # Normalize any legacy 'Month YYYY' style budget rows (e.g. "June 2026")
    # to the canonical 'MM-YYYY' format used everywhere else in the app.
    cur.execute("SELECT DISTINCT month_year FROM budgets")
    for (my,) in cur.fetchall():
        if my and not re.match(r"^\d{2}-\d{4}$", my):
            try:
                canonical = datetime.strptime(my, "%B %Y").strftime("%m-%Y")
                cur.execute("UPDATE budgets SET month_year=? WHERE month_year=?", (canonical, my))
            except (TypeError, ValueError):
                pass


def _safe_db_filename(username):
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", username or "user")
    return os.path.join(DATA_DIR, f"cooper_{safe}.db")


def get_user_db(username):
    """
    Opens (creating on first use) the private SQLite database for one
    user, with every personal/financial table ready. Called right after
    a successful login or registration; its return values become the
    module's active `conn` / `cursor` globals, so every existing
    add/edit/delete/report function across the app keeps working
    completely unchanged — they already just call `cursor.execute(...)`
    / `conn.commit()` on whichever database happens to be "current".
    A brand-new user's database starts completely empty — no demo data.
    """
    path = _safe_db_filename(username)
    user_conn = sqlite3.connect(path)
    user_cur = user_conn.cursor()
    _create_user_data_schema(user_cur)
    user_conn.commit()
    return user_conn, user_cur


def _migrate_legacy_single_database():
    """
    One-time upgrade path from the old single-shared-database
    architecture (every account reading/writing one cooper_vault.db).
    Runs at most once — if data/users.db already exists, this is a
    no-op, so restarting the app never re-runs it.

    What it does, exactly:
      1. Creates data/users.db and copies every row from the old
         `users` table into it — no accounts are lost.
      2. The old schema had NO per-user ownership column on expenses /
         income / budgets / etc., so there is no way to know which
         historical row belonged to which account. Rather than invent
         an owner or discard data, EVERY existing account receives its
         own full copy of the old shared financial data as its
         starting point — each user's new private database begins as a
         clone of the old shared one. From that point on, each
         account's data is fully independent (this is a one-time
         starting snapshot, not an ongoing shared source).
      3. The original cooper_vault.db is renamed to
         cooper_vault.db.legacy_backup — it is never deleted, so the
         pre-migration state stays fully recoverable.
      4. On a brand-new install (no cooper_vault.db at all), this just
         creates an empty data/users.db and stops there.
    """
    if os.path.exists(AUTH_DB_PATH):
        return

    if not os.path.exists(LEGACY_DB_PATH):
        fresh_conn = sqlite3.connect(AUTH_DB_PATH)
        _create_auth_schema(fresh_conn.cursor())
        fresh_conn.commit()
        fresh_conn.close()
        return

    legacy_conn = sqlite3.connect(LEGACY_DB_PATH)
    legacy_cur = legacy_conn.cursor()

    new_auth_conn = sqlite3.connect(AUTH_DB_PATH)
    new_auth_cur = new_auth_conn.cursor()
    _create_auth_schema(new_auth_cur)

    legacy_cur.execute("PRAGMA table_info(users)")
    user_cols = [row[1] for row in legacy_cur.fetchall()]
    if user_cols:
        legacy_cur.execute(f"SELECT {', '.join(user_cols)} FROM users")
        legacy_users = legacy_cur.fetchall()
        placeholders = ", ".join("?" for _ in user_cols)
        for row in legacy_users:
            try:
                new_auth_cur.execute(
                    f"INSERT OR IGNORE INTO users ({', '.join(user_cols)}) VALUES ({placeholders})",
                    row
                )
            except sqlite3.Error:
                pass
        new_auth_conn.commit()

        financial_tables = [
            "expenses", "income", "budgets", "savings_goals",
            "recurring_transactions", "reminders", "system_streaks",
            "messages", "bills_emi", "todos", "notes"
        ]
        new_auth_cur.execute("SELECT username FROM users")
        usernames = [r[0] for r in new_auth_cur.fetchall()]
        for username in usernames:
            u_conn, u_cur = get_user_db(username)
            for table in financial_tables:
                try:
                    legacy_cur.execute(f"PRAGMA table_info({table})")
                    cols = [row[1] for row in legacy_cur.fetchall()]
                    if not cols:
                        continue
                    legacy_cur.execute(f"SELECT {', '.join(cols)} FROM {table}")
                    data_rows = legacy_cur.fetchall()
                    ph = ", ".join("?" for _ in cols)
                    for data_row in data_rows:
                        u_cur.execute(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({ph})", data_row)
                except sqlite3.Error:
                    pass
            u_conn.commit()
            u_conn.close()

    new_auth_conn.close()
    legacy_conn.close()

    try:
        os.replace(LEGACY_DB_PATH, LEGACY_DB_PATH + ".legacy_backup")
    except OSError:
        pass


_migrate_legacy_single_database()

# The always-open authentication database — used ONLY for the `users`
# table (login, registration, profile fields, theme/remember-me
# preference). This connection is never swapped and never touches
# financial data.
auth_conn = sqlite3.connect(AUTH_DB_PATH)
auth_cursor = auth_conn.cursor()
_create_auth_schema(auth_cursor)
auth_conn.commit()
