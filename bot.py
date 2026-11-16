"""Logistika kuzatuv boti (Telegram).

Matn / Excel / PDF: foydalanuvchi nima desa, bot shuni tashlaydi.
"""
import asyncio
import io
import logging
import os
import re
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandObject
from aiogram.types import BufferedInputFile, Message
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import database
from database import db, init_schema as init_db  # .env ham shu yerda o'qiladi
from aiohttp import web

STATUSES = ["Yangi", "Yo'lda", "Bojxonada", "Yetkazildi", "Bekor"]
COLUMNS = [
    ("code", "Kod"), ("client", "Mijoz"), ("origin", "Qayerdan"),
    ("destination", "Qayerga"), ("driver", "Haydovchi"),
    ("price", "Narx"), ("status", "Holat"), ("updated_at", "Yangilandi"),
]


# ---------- Baza: database.py da ----------


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def next_code(con):
    n = con.execute("SELECT COALESCE(MAX(id), 0) + 1 AS n FROM shipments").fetchone()["n"]
    return f"YK-{n:04d}"


def find_status(text):
    t = text.lower().replace("ʼ", "'").replace("‘", "'").replace("’", "'")
    for s in STATUSES:
        if s.lower() == t:
            return s
    return None


CODE_RE = re.compile(r"YK-\d{3,}", re.I)
EVENT_COLS = [("created_at", "Vaqt"), ("code", "Kod"), ("who", "Kim yubordi"),
              ("kind", "Turi"), ("text", "Izoh"), ("map", "Xarita")]


def fetch(status=None, code=None):
    where, args = [], []
    if status:
        where.append("status=?"); args.append(status)
    if code:
        where.append("code=?"); args.append(code)
    q = "SELECT * FROM shipments" + (" WHERE " + " AND ".join(where) if where else "")
    with db() as con:
        return [dict(r) for r in con.execute(q + " ORDER BY id DESC", args)]


def fetch_events(code=None):
    q, args = "SELECT * FROM events", ()
    if code:
        q, args = q + " WHERE code=?", (code,)
    with db() as con:
        rows = [dict(r) for r in con.execute(q + " ORDER BY id DESC", args)]
    for r in rows:
        r["map"] = f"https://maps.google.com/?q={r['lat']},{r['lon']}" if r["lat"] is not None else ""
    return rows


def add_event(code, who, kind, text="", lat=None, lon=None):
    with db() as con:
        con.execute("INSERT INTO events(code,who,kind,text,lat,lon,created_at) VALUES(?,?,?,?,?,?,?)",
                    (code, who, kind, text, lat, lon, now()))
        con.execute("UPDATE shipments SET updated_at=? WHERE code=?", (now(), code))


# ---------- Hisobotlar ----------
def fmt_one(r):
    return (f"📦 {r['code']}\n👤 {r['client']}\n🛣 {r['origin']} → {r['destination']}\n"
            f"🚚 {r['driver']}\n💰 {r['price']:,.0f}\n📍 Holat: {r['status']}\n🕒 {r['updated_at']}")


def fmt_list(rows):
    if not rows:
        return "Hozircha yuk yo'q."
    lines = [f"{r['code']} | {r['client']} | {r['origin']}→{r['destination']} | {r['status']}"
             for r in rows[:50]]
    extra = f"\n… yana {len(rows) - 50} ta (Excel/PDF so'rang)" if len(rows) > 50 else ""
    return f"Jami: {len(rows)} ta\n\n" + "\n".join(lines) + extra


def make_excel(rows, events=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Yuklar"
    ws.append([h for _, h in COLUMNS])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append([r[k] for k, _ in COLUMNS])
    for i, (k, h) in enumerate(COLUMNS, 1):
        w = max([len(h)] + [len(str(r[k])) for r in rows]) + 3
        ws.column_dimensions[get_column_letter(i)].width = min(w, 40)
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("Kuzatuv tarixi")
    ws2.append([h for _, h in EVENT_COLS])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for e in events:
        ws2.append([e[k] for k, _ in EVENT_COLS])
    for i, w in enumerate([18, 12, 18, 12, 50, 40], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _font():
    for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
              "C:/Windows/Fonts/arial.ttf",
              "/Library/Fonts/Arial.ttf",
              os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")):
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("UZ", p))
            return "UZ"
    return "Helvetica"


def make_pdf(rows, events=()):
    font = _font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24,
                            topMargin=24, bottomMargin=24)
    st = getSampleStyleSheet()["Title"]
    st.fontName = font
    data = [[h for _, h in COLUMNS]] + [[str(r[k]) for k, _ in COLUMNS] for r in rows]
    t = Table(data, repeatRows=1)
    style = TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font), ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3F8")]),
    ])
    t.setStyle(style)
    story = [Paragraph(f"Logistika hisoboti — {now()}", st), Spacer(1, 8), t]
    if events:
        cell = ParagraphStyle("c", fontName=font, fontSize=8, leading=10)
        ev = [[h for _, h in EVENT_COLS[:5]]] + [
            [Paragraph(str(e[k] or ""), cell) for k, _ in EVENT_COLS[:5]] for e in events]
        et = Table(ev, repeatRows=1, colWidths=[85, 55, 90, 60, 450])
        et.setStyle(style)
        story += [Spacer(1, 14), Paragraph("Kuzatuv tarixi", st), et]
    doc.build(story)
    return buf.getvalue()


# ---------- Bot ----------
ADMINS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
dp = Dispatcher()
import planner  # noqa: E402


def permitted(event) -> bool:
    """Umumiy to'siq: faqat adminlarning shaxsiy chati va ikki guruh. Qolganlarga bot JIM turadi.
    Hech narsa sozlanmagan bo'lsa ham hammaga yopiq (xavfsiz tomonga)."""
    user = getattr(event, "from_user", None)
    chat = getattr(event, "chat", None)
    if user is None or chat is None or user.is_bot:
        return False
    if chat.type == "private":
        return user.id in ADMINS
    if chat.id in (planner.COLLECT_CHAT, planner.LOAD_CHAT) and chat.id is not None:
        return True
    return user.id in ADMINS and (event.text or "").split("@")[0].strip() == "/id"  # guruh ID sini olish uchun


@dp.message.outer_middleware()
async def gate_message(handler, event, data):
    if permitted(event):
        return await handler(event, data)  # begona bo'lsa — hech narsa qaytarmaydi


@dp.edited_message.outer_middleware()
async def gate_edited(handler, event, data):
    if permitted(event):
        return await handler(event, data)

planner.setup(dp)  # umumiy matn ishlovchisidan oldin

HELP = (
    "🚛 Logistika kuzatuv boti\n\n"
    "Yozing: «ma'lumot tashla», «excel tashla» yoki «pdf tashla».\n\n"
    "/id — chat ID (guruhni sozlash uchun)\n"
    "/yuklangan — ortilganlar, /boshqa — alohida mavzu (2097)\n"
    "/royxat [holat] — yuklar ro'yxati\n"
    "/yuk KOD — bitta yuk\n"
    "/kuzat KOD — holati o'zgarsa xabar olish\n"
    "/boshla KOD — haydovchi: joylashuv va izohlarni shu yukka yozish\n"
    "«YK-0001 izoh matni» — kuzatuvga yozuv qo'shish\n"
    "«YK-0001 excel» yoki «YK-0001 pdf» — shu yuk tarixi\n"
    "/excel, /pdf — fayl\n\n"
    "📸 Admin: WeChat'dagi rasm va xabarlarni ketma-ket yuboring — bot ko'rib, kimning yuki / kg / kub so'raydi.\n"
    "/tovarlar, /tuzat ID soni=.. uzun=.. eni=.. bal=.. kg=.. sana=.., /qosh, /ochir ID\n"
    "/fura [uzun eni bal kg] — fura o'lchami, /reja [excel|pdf] — yuklash rejasi, /yuklandi N\n\n"
    "Admin:\n"
    "/yangi Mijoz | Qayerdan | Qayerga | Haydovchi | Narx\n"
    "/holat KOD Holat\n"
    f"Holatlar: {', '.join(STATUSES)}"
)


def is_admin(m: Message):
    return m.from_user.id in ADMINS


def parse_filter(text):
    code = CODE_RE.search(text or "")
    t = (text or "").replace("ʼ", "'").replace("’", "'").lower()
    status = next((s for s in STATUSES if s.lower() in t), None)
    return status, (code.group(0).upper() if code else None)


async def send_report(m: Message, kind: str, text=""):
    status, code = parse_filter(text)
    rows = fetch(status, code)
    events = fetch_events(code)
    name = code or "yuklar"
    if kind == "excel":
        await m.answer_document(BufferedInputFile(make_excel(rows, events), f"{name}.xlsx"))
    elif kind == "pdf":
        await m.answer_document(BufferedInputFile(make_pdf(rows, events), f"{name}.pdf"))
    elif code:
        await m.answer(fmt_history(rows, events))
    else:
        await m.answer(fmt_list(rows))


def fmt_history(rows, events):
    if not rows:
        return "Bunday kod topilmadi."
    out = [fmt_one(rows[0]), "", f"📋 Kuzatuv tarixi ({len(events)} ta):"]
    for e in events[:15]:
        pin = " 📍" if e["map"] else ""
        out.append(f"{e['created_at']} — {e['who']}: {e['text'] or e['kind']}{pin}")
    return "\n".join(out)


@dp.message(Command("start", "help"))
async def start(m: Message):
    await m.answer(HELP)


@dp.message(Command("yangi"))
async def yangi(m: Message, command: CommandObject):
    if not is_admin(m):
        return await m.answer("Bu buyruq faqat adminlar uchun.")
    parts = [p.strip() for p in (command.args or "").split("|")]
    if len(parts) != 5:
        return await m.answer("Format: /yangi Mijoz | Qayerdan | Qayerga | Haydovchi | Narx")
    try:
        price = float(parts[4].replace(" ", "").replace(",", ""))
    except ValueError:
        return await m.answer("Narx raqam bo'lishi kerak.")
    with db() as con:
        code = next_code(con)
        con.execute("INSERT INTO shipments(code,client,origin,destination,driver,price,status,updated_at)"
                    " VALUES(?,?,?,?,?,?,?,?)",
                    (code, parts[0], parts[1], parts[2], parts[3], price, "Yangi", now()))
    await m.answer(f"✅ Qo'shildi: {code}")


@dp.message(Command("holat"))
async def holat(m: Message, command: CommandObject, bot: Bot):
    if not is_admin(m):
        return await m.answer("Bu buyruq faqat adminlar uchun.")
    args = (command.args or "").split(maxsplit=1)
    status = find_status(args[1]) if len(args) == 2 else None
    if not status:
        return await m.answer(f"Format: /holat KOD Holat\nHolatlar: {', '.join(STATUSES)}")
    code = args[0].upper()
    with db() as con:
        cur = con.execute("UPDATE shipments SET status=?, updated_at=? WHERE code=?",
                          (status, now(), code))
        subs = [r["chat_id"] for r in con.execute("SELECT chat_id FROM subs WHERE code=?", (code,))]
    if not cur.rowcount:
        return await m.answer("Bunday kod topilmadi.")
    await m.answer(f"✅ {code} → {status}")
    for chat_id in subs:
        try:
            await bot.send_message(chat_id, f"🔔 {code} holati o'zgardi: {status}")
        except Exception:
            logging.exception("Xabar yuborilmadi: %s", chat_id)


@dp.message(Command("yuk"))
async def yuk(m: Message, command: CommandObject):
    code = (command.args or "").strip().upper()
    with db() as con:
        r = con.execute("SELECT * FROM shipments WHERE code=?", (code,)).fetchone()
    await m.answer(fmt_one(dict(r)) if r else "Bunday kod topilmadi. Misol: /yuk YK-0001")


@dp.message(Command("kuzat"))
async def kuzat(m: Message, command: CommandObject):
    code = (command.args or "").strip().upper()
    with db() as con:
        if not con.execute("SELECT 1 FROM shipments WHERE code=?", (code,)).fetchone():
            return await m.answer("Bunday kod topilmadi.")
        con.execute("INSERT INTO subs(code, chat_id) VALUES(?,?) ON CONFLICT DO NOTHING", (code, m.chat.id))
    await m.answer(f"👀 {code} kuzatuvga olindi. Holati o'zgarsa xabar beraman.")


@dp.message(Command("royxat"))
async def royxat(m: Message, command: CommandObject):
    await send_report(m, "text", command.args or "")


@dp.message(Command("excel"))
async def excel(m: Message, command: CommandObject):
    await send_report(m, "excel", command.args or "")


@dp.message(Command("pdf"))
async def pdf(m: Message, command: CommandObject):
    await send_report(m, "pdf", command.args or "")


def who(m: Message):
    u = m.from_user
    return u.full_name or str(u.id)


@dp.message(Command("boshla"))
async def boshla(m: Message, command: CommandObject):
    code = (command.args or "").strip().upper()
    with db() as con:
        if not con.execute("SELECT 1 FROM shipments WHERE code=?", (code,)).fetchone():
            return await m.answer("Bunday kod topilmadi.")
        con.execute("INSERT INTO active(chat_id, code) VALUES(?,?) "
                    "ON CONFLICT(chat_id) DO UPDATE SET code=excluded.code", (m.chat.id, code))
    await m.answer(f"✅ {code} faol. Endi joylashuv (📎 → Location) yuboring yoki "
                   f"«{code} izoh» deb yozing — hammasi saqlanadi.")


@dp.message(F.location)
async def got_location(m: Message, bot: Bot):
    with db() as con:
        r = con.execute("SELECT code FROM active WHERE chat_id=?", (m.chat.id,)).fetchone()
    if not r:
        return await m.answer("Avval /boshla KOD yuboring.")
    add_event(r["code"], who(m), "Joylashuv", "", m.location.latitude, m.location.longitude)
    await notify_subs(bot, r["code"], f"📍 {r['code']}: yangi joylashuv — "
                      f"https://maps.google.com/?q={m.location.latitude},{m.location.longitude}")
    await m.answer("📍 Joylashuv saqlandi.")


async def notify_subs(bot: Bot, code, text):
    with db() as con:
        subs = [x["chat_id"] for x in con.execute("SELECT chat_id FROM subs WHERE code=?", (code,))]
    for chat_id in subs:
        try:
            await bot.send_message(chat_id, text)
        except Exception:
            logging.exception("Xabar yuborilmadi: %s", chat_id)


@dp.message(F.text)
async def free_text(m: Message, bot: Bot):
    if m.chat.type != "private":  # guruhda ortiqcha javob bermaydi
        return
    t = m.text.lower()
    if "excel" in t or "eksel" in t:
        return await send_report(m, "excel", m.text)
    if "pdf" in t:
        return await send_report(m, "pdf", m.text)
    match = CODE_RE.match(m.text.strip())
    rest = m.text.strip()[match.end():].strip() if match else ""
    if match and rest:  # "YK-0001 Samarqandda yoqilg'i quydim" -> yozib qo'yiladi
        code = match.group(0).upper()
        with db() as con:
            if not con.execute("SELECT 1 FROM shipments WHERE code=?", (code,)).fetchone():
                return await m.answer("Bunday kod topilmadi.")
        add_event(code, who(m), "Izoh", rest)
        await notify_subs(bot, code, f"💬 {code}: {rest}")
        return await m.answer("✅ Saqlandi.")
    if any(w in t for w in ("ma'lumot", "malumot", "maʼlumot", "ro'yxat", "royxat", "tashla", "yuk", "tarix")):
        return await send_report(m, "text", m.text)
    await m.answer(HELP)


@dp.message(Command("zaxira"))
async def zaxira(m: Message):
    """Bazaning to'liq nusxasi (JSON). Supabase bepul rejasida avtomatik zaxira yo'q."""
    if not is_admin(m) or m.chat.type != "private":
        return
    data = await asyncio.to_thread(database.dump_all)
    await m.answer_document(BufferedInputFile(data, f"zaxira_{datetime.now():%Y-%m-%d}.json"))


async def health(request):
    """Render/ping uchun. Faqat "ok" qaytaradi; bazani ham "uyg'otib" turadi, lekin xato matnini chiqarmaydi."""
    try:
        await asyncio.to_thread(database.ping)
    except Exception as e:
        logging.error("health: baza javob bermadi (%s)", type(e).__name__)  # matni/manzili yozilmaydi
    return web.Response(text="ok")


async def start_health():
    app = web.Application()
    app.router.add_get("/health", health)
    app.router.add_get("/", health)
    runner = web.AppRunner(app)
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", int(os.getenv("PORT", "10000"))).start()
    return runner


async def main():
    logging.basicConfig(level=logging.INFO)
    init_db()
    runner = await start_health()
    try:
        await dp.start_polling(Bot(os.environ["BOT_TOKEN"]))
    finally:
        await runner.cleanup()


if __name__ == "__main__":
    asyncio.run(main())
