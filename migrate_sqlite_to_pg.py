"""BIR MARTALIK ko'chirish: logistika.db (SQLite) -> Supabase Postgres.

Takror ishga tushirsangiz ham dublikat qo'shilmaydi (yuk ID lari saqlanadi, mavjudlari o'tkazib yuboriladi).

    python migrate_sqlite_to_pg.py            # logistika.db ni oladi
    python migrate_sqlite_to_pg.py boshqa.db  # boshqa fayl

DATABASE_URL .env faylda (yoki muhit o'zgaruvchisida) bo'lishi kerak.
"""
import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import database  # noqa: E402

if not database.PG:
    sys.exit("DATABASE_URL topilmadi. .env fayliga Supabase manzilini yozing (README, 3-qadam).")

path = sys.argv[1] if len(sys.argv) > 1 else os.getenv("DB_PATH", "logistika.db")
if not os.path.exists(path):
    sys.exit(f"Fayl topilmadi: {path}")

src = sqlite3.connect(path)
src.row_factory = sqlite3.Row
database.init_schema()

# jadval: (takror tekshiruvi uchun konflikt qoidasi, ID ketma-ketligini tiklash kerakmi)
PLAN = [
    ("shipments", "ON CONFLICT DO NOTHING", True),
    ("items", "ON CONFLICT DO NOTHING", True),
    ("catalog", "ON CONFLICT DO NOTHING", False),
    ("trucks", "ON CONFLICT DO NOTHING", False),
    ("active", "ON CONFLICT DO NOTHING", False),
    ("subs", "ON CONFLICT DO NOTHING", False),
]


def src_has(table):
    return src.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()


total_new = 0
with database.db() as con:
    pg_cols = {}

    def cols_of(table):
        if table not in pg_cols:
            pg_cols[table] = {r["column_name"] for r in con.execute(
                "SELECT column_name FROM information_schema.columns WHERE table_name=? AND table_schema=current_schema()",
                (table,))}
        return pg_cols[table]

    for table, conflict, fix_seq in PLAN:
        if not src_has(table):
            print(f"{table:10} — manba faylda yo'q, o'tkazildi")
            continue
        rows = src.execute(f"SELECT * FROM {table}").fetchall()
        if not rows:
            print(f"{table:10} — bo'sh")
            continue
        cols = [c for c in rows[0].keys() if c in cols_of(table)]
        new = 0
        for r in rows:
            d = dict(r)
            if table == "items" and not d.get("src"):
                d["src"] = f"sqlite:items:{d['id']}"  # takror yuklashni aniqlash uchun belgi
            cur = con.execute(
                f"INSERT INTO {table}({', '.join(cols)}) VALUES({', '.join('?' * len(cols))}) {conflict}",
                [d[c] for c in cols])
            new += cur.rowcount
        if fix_seq:  # keyingi yangi ID lar eskilar bilan to'qnashmasligi uchun
            con.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), "
                        f"(SELECT COALESCE(MAX(id), 1) FROM {table}))")
        total_new += new
        print(f"{table:10} — {len(rows)} qator manbada, {new} ta yangi qo'shildi, {len(rows) - new} ta allaqachon bor edi")

    # kuzatuv yozuvlari: ID kerak emas, takrorni (kod+vaqt+kim+tur+izoh) bo'yicha aniqlaymiz
    if src_has("events"):
        rows = src.execute("SELECT * FROM events ORDER BY id").fetchall()
        new = 0
        for r in rows:
            exists = con.execute(
                "SELECT 1 FROM events WHERE code=? AND created_at=? AND who=? AND kind=? "
                "AND text IS NOT DISTINCT FROM ? AND lat IS NOT DISTINCT FROM ?",
                (r["code"], r["created_at"], r["who"], r["kind"], r["text"], r["lat"])).fetchone()
            if not exists:
                con.execute("INSERT INTO events(code, who, kind, text, lat, lon, created_at) VALUES(?,?,?,?,?,?,?)",
                            (r["code"], r["who"], r["kind"], r["text"], r["lat"], r["lon"], r["created_at"]))
                new += 1
        total_new += new
        print(f"{'events':10} — {len(rows)} qator manbada, {new} ta yangi qo'shildi, {len(rows) - new} ta allaqachon bor edi")

print(f"\nTayyor. Jami yangi qo'shildi: {total_new}")
