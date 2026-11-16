"""Baza qatlami.

DATABASE_URL bo'lsa — Supabase/Postgres (Render uchun). Bo'lmasa — mahalliy SQLite fayl (sinash uchun).
Kodning hamma joyida `?` belgili SQL yoziladi, Postgres uchun u o'zi `%s` ga aylanadi.
"""
import json
import os
import sqlite3
from datetime import datetime


def _load_env():
    """.env faylni o'qiydi (Render'da u kerak emas: u yerda o'zgaruvchilar panelda turadi)."""
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip("\"'"))


_load_env()
URL = os.getenv("DATABASE_URL", "").strip()
PG = URL.startswith(("postgres://", "postgresql://"))
SQLITE_PATH = os.getenv("DB_PATH", "logistika.db")
TABLES = ["shipments", "events", "subs", "active", "items", "catalog", "trucks"]

_pool = None


def close():
    global _pool
    if _pool is not None:
        _pool.close()
        _pool = None


import atexit  # noqa: E402
atexit.register(close)


def _get_pool():
    global _pool
    if _pool is None:
        from psycopg.rows import dict_row
        from psycopg_pool import ConnectionPool
        _pool = ConnectionPool(
            URL, min_size=1, max_size=4, open=True, timeout=30,
            # prepare_threshold=None: Supabase pooler bilan mos (tayyorlangan so'rovlarsiz)
            kwargs={"row_factory": dict_row, "prepare_threshold": None, "connect_timeout": 15},
            check=ConnectionPool.check_connection)  # uzilgan ulanishni o'zi almashtiradi
    return _pool


class _Conn:
    def __init__(self, raw):
        self._raw = raw

    def execute(self, sql, args=()):
        if PG:
            sql = sql.replace("?", "%s")
        return self._raw.execute(sql, tuple(args))


class db:
    """`with db() as con:` — oxirida o'zi saqlaydi (commit), xato bo'lsa bekor qiladi."""

    def __enter__(self):
        if PG:
            self._cm = _get_pool().connection()
            return _Conn(self._cm.__enter__())
        self._raw = sqlite3.connect(SQLITE_PATH)
        self._raw.row_factory = sqlite3.Row
        return _Conn(self._raw)

    def __exit__(self, et, ev, tb):
        if PG:
            return self._cm.__exit__(et, ev, tb)
        try:
            self._raw.commit() if et is None else self._raw.rollback()
        finally:
            self._raw.close()
        return False


def init_schema():
    pk = "BIGSERIAL PRIMARY KEY" if PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
    real = "DOUBLE PRECISION" if PG else "REAL"
    big = "BIGINT" if PG else "INTEGER"  # Telegram chat ID lari katta (-100...)
    with db() as con:
        con.execute(f"""CREATE TABLE IF NOT EXISTS shipments(
            id {pk}, code TEXT UNIQUE, client TEXT, origin TEXT, destination TEXT, driver TEXT,
            price {real} DEFAULT 0, status TEXT DEFAULT 'Yangi', updated_at TEXT)""")
        con.execute(f"""CREATE TABLE IF NOT EXISTS events(
            id {pk}, code TEXT, who TEXT, kind TEXT, text TEXT, lat {real}, lon {real}, created_at TEXT)""")
        con.execute(f"CREATE TABLE IF NOT EXISTS active(chat_id {big} PRIMARY KEY, code TEXT)")
        con.execute(f"CREATE TABLE IF NOT EXISTS subs(code TEXT, chat_id {big}, PRIMARY KEY(code, chat_id))")
        con.execute(f"""CREATE TABLE IF NOT EXISTS items(
            id {pk}, name TEXT, qty {real}, l {real}, w {real}, h {real}, kg {real}, arrival TEXT,
            status TEXT DEFAULT 'kutilmoqda', created_at TEXT, owner TEXT, mark TEXT, vol {real},
            loaded_date TEXT, note TEXT, topic TEXT, src TEXT)""")
        if not PG:  # eski SQLite fayl bo'lsa, yetishmagan ustunlarni qo'shadi
            cols = {r["name"] for r in con.execute("PRAGMA table_info(items)")}
            for c, t in (("owner", "TEXT"), ("mark", "TEXT"), ("vol", "REAL"), ("loaded_date", "TEXT"),
                         ("note", "TEXT"), ("topic", "TEXT"), ("src", "TEXT")):
                if c not in cols:
                    con.execute(f"ALTER TABLE items ADD COLUMN {c} {t}")
        con.execute("CREATE UNIQUE INDEX IF NOT EXISTS items_src_uq ON items(src)")  # takror yuklashdan himoya
        con.execute(f"""CREATE TABLE IF NOT EXISTS catalog(
            key TEXT PRIMARY KEY, name TEXT, l {real}, w {real}, h {real}, kg {real})""")
        con.execute(f"""CREATE TABLE IF NOT EXISTS trucks(
            chat_id {big} PRIMARY KEY, l {real}, w {real}, h {real}, kg {real})""")
        if PG:  # Supabase API orqali tashqaridan kirishni yopadi (policy yo'q = hech kimga ruxsat yo'q).
            # Bot to'g'ridan-to'g'ri Postgres'ga ulanadi (postgres roli RLS'ni chetlab o'tadi), shuning uchun ishlayveradi.
            for t in TABLES:
                con.execute(f"ALTER TABLE {t} ENABLE ROW LEVEL SECURITY")


def ping():
    """Baza javob beryaptimi? (Supabase bepul loyihasi uxlab qolmasligi uchun ham ishlaydi)"""
    with db() as con:
        con.execute("SELECT 1").fetchone()


def dump_all():
    """Hamma jadvalni JSON qilib beradi (zaxira nusxa)."""
    out = {"saqlangan": datetime.now().isoformat(timespec="seconds")}
    with db() as con:
        for t in TABLES:
            out[t] = [dict(r) for r in con.execute(f"SELECT * FROM {t}")]
    return json.dumps(out, ensure_ascii=False, indent=1, default=str).encode("utf-8")
