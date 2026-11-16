"""Rasm -> tovarlar -> tuzatish -> kelgan sana bo'yicha fura yuklash rejasi."""
import asyncio
import base64
import io
import json
import logging
import math
import os
import re
from datetime import date, datetime

from aiogram import Bot, F
from aiogram.filters import Command, CommandObject
from aiogram.types import BufferedInputFile, Message
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from database import db, init_schema as init_db  # noqa: E402
MODEL = os.getenv("VISION_MODEL", "claude-sonnet-5-5")
ADMINS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
def _int_env(name):
    try:
        return int(os.getenv(name, "").strip())
    except ValueError:
        return None


COLLECT_CHAT = _int_env("COLLECT_CHAT_ID")  # zbornik guruhi: faqat ma'lumot yig'iladi
LOAD_CHAT = _int_env("LOAD_CHAT_ID")        # yuklash guruhi: tayyor reja shu yerga tushadi
DEFAULT_TRUCK = {"l": 1360.0, "w": 245.0, "h": 270.0, "kg": 22000.0}  # sm, kg (standart tent fura)
FILL = 0.85  # hajmning amalda to'ldiriladigan ulushi (bo'shliq, yig'ish uchun)
LAST_PLAN: dict = {}

PROMPT = """You are reading a photo or document of cargo/goods (invoice, packing list, delivery note, label, or a photo of the products). Today is {today}.
User caption: {caption}
Extract every distinct product line. Reply with ONLY valid JSON, no markdown:
{{"items":[{{"name":"as written, keep original language","qty":number of units/boxes/pallets,"length_cm":number or null,"width_cm":number or null,"height_cm":number or null,"weight_kg":weight of ONE unit in kg or null,"arrival_date":"YYYY-MM-DD" or null}}]}}
Rules: use only values visible in the image or caption; never guess dimensions or weight (use null). Convert to cm and kg. If a line shows total weight and quantity, divide to get per-unit weight. Set arrival_date only if an arrival/delivery/invoice date is stated."""


# ---------- Baza: database.py da ----------


def norm(name):
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def num(x):
    try:
        v = float(str(x).replace(",", ".").replace(" ", ""))
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None


def parse_date(s, default=None):
    s = (s or "").strip().lower()
    if s in ("bugun", "today"):
        return date.today().isoformat()
    for f in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, f).date().isoformat()
        except ValueError:
            pass
    return default


def learn(name, l, w, h, kg):
    """Tovar o'lchamini eslab qoladi: keyingi safar rasmda bo'lmasa ham o'zi to'ldiradi."""
    with db() as con:
        old = con.execute("SELECT * FROM catalog WHERE key=?", (norm(name),)).fetchone()
        o = dict(old) if old else {}
        con.execute("INSERT INTO catalog(key, name, l, w, h, kg) VALUES(?,?,?,?,?,?) "
                    "ON CONFLICT(key) DO UPDATE SET name=excluded.name, l=excluded.l, w=excluded.w, "
                    "h=excluded.h, kg=excluded.kg",
                    (norm(name), name, l or o.get("l"), w or o.get("w"), h or o.get("h"), kg or o.get("kg")))


def add_item(name, qty, l, w, h, kg, arrival, owner=None, mark=None, vol=None):
    with db() as con:
        c = con.execute("SELECT * FROM catalog WHERE key=?", (norm(name),)).fetchone()
        if c:
            l, w, h, kg = l or c["l"], w or c["w"], h or c["h"], kg or c["kg"]
        cur = con.execute(
            "INSERT INTO items(name,qty,l,w,h,kg,arrival,owner,mark,vol,created_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?) RETURNING id",
            (name, qty or 1, l, w, h, kg, arrival, owner, mark, vol, datetime.now().strftime("%Y-%m-%d %H:%M")))
        return cur.fetchone()["id"]


def get_item(i):
    with db() as con:
        r = con.execute("SELECT * FROM items WHERE id=?", (i,)).fetchone()
    return dict(r) if r else None


def pending():
    with db() as con:
        return [dict(r) for r in con.execute(
            "SELECT * FROM items WHERE status='kutilmoqda' ORDER BY arrival, id")]


def get_truck(chat_id):
    with db() as con:
        r = con.execute("SELECT * FROM trucks WHERE chat_id=?", (chat_id,)).fetchone()
    return dict(r) if r else dict(DEFAULT_TRUCK)


# ---------- Rasmdan o'qish ----------
def parse_items(text):
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return []
    try:
        data = json.loads(m.group(0))
    except json.JSONDecodeError:
        return []
    items = data.get("items", []) if isinstance(data, dict) else []
    return [i for i in items if isinstance(i, dict) and i.get("name")]


async def extract(data: bytes, media: str, caption: str):
    import anthropic  # kechiktirilgan import: kalit bo'lmasa bot baribir ishga tushadi
    block_type = "document" if media == "application/pdf" else "image"
    block = {"type": block_type, "source": {
        "type": "base64", "media_type": media, "data": base64.b64encode(data).decode()}}
    client = anthropic.AsyncAnthropic()
    resp = await client.messages.create(
        model=MODEL, max_tokens=2500,
        messages=[{"role": "user", "content": [
            block, {"type": "text", "text": PROMPT.format(today=date.today().isoformat(),
                                                           caption=caption or "-")}]}])
    return parse_items("".join(b.text for b in resp.content if b.type == "text"))


# ---------- Formatlash ----------
def fmt_item(r):
    q = r["qty"] or 1

    def per(v, unit):
        return f"{v:.4g} {unit}" if q == 1 else f"{v:.4g} {unit}/dona (jami {v * q:.4g})"

    if all(r[k] for k in "lwh"):
        size = f"{r['l']:g}×{r['w']:g}×{r['h']:g} sm"
    elif r.get("vol"):
        size = per(r["vol"], "m³")
    else:
        size = "❓ kub yo'q"
    kg = per(r["kg"], "kg") if r["kg"] else "❓ kg yo'q"
    who = f" | 👤 {r['owner']}" if r.get("owner") else ""
    return f"#{r['id']} {r['name']} — {q:g} dona | {size} | {kg}{who} | 📅 {r['arrival']}"


def fmt_items(items):
    if not items:
        return "Kutilayotgan tovar yo'q."
    out, cur = [], None
    for r in items:
        if r["arrival"] != cur:
            cur = r["arrival"]
            out.append(f"\n📅 {cur}")
        out.append(fmt_item(r))
    return "\n".join(out).strip()


# ---------- Reja ----------
def build_plan(items, truck):
    """Kelgan sana bo'yicha (eski -> yangi) ketma-ket furalarga joylaydi. Dona-dona bo'lib ketishi mumkin."""
    tv = truck["l"] * truck["w"] * truck["h"] / 1e6
    cap_v, cap_kg = tv * FILL, truck["kg"]
    tdims = sorted([truck["l"], truck["w"], truck["h"]])
    trucks, skipped = [], []

    def new():
        return {"lines": [], "vol": 0.0, "kg": 0.0, "cap_v": cap_v, "cap_kg": cap_kg}

    cur = new()
    for it in sorted(items, key=lambda r: (r["arrival"], r.get("owner") or "", r["id"])):
        has_dims = all(it.get(k) for k in ("l", "w", "h"))
        uv = it["vol"] if it.get("vol") else (it["l"] * it["w"] * it["h"] / 1e6 if has_dims else None)
        if not uv or not it.get("kg"):
            skipped.append((it, "kub (hajm) yoki og'irlik kiritilmagan"))
            continue
        if has_dims:
            a, b, c = sorted([it["l"], it["w"], it["h"]])
            if a > tdims[0] or b > tdims[1] or c > tdims[2]:
                skipped.append((it, "bitta dona furaga sig'maydi"))
                continue
        ukg = it["kg"]
        if uv > cap_v or ukg > cap_kg:
            skipped.append((it, "bitta dona fura sig'imidan katta"))
            continue
        remaining = max(1, int(round(it["qty"])))
        while remaining > 0:
            n = int(min((cap_v - cur["vol"]) / uv, (cap_kg - cur["kg"]) / ukg) + 1e-9)
            n = min(n, remaining)
            if n <= 0:
                trucks.append(cur)
                cur = new()
                continue
            cur["lines"].append({"id": it["id"], "name": it["name"], "arrival": it["arrival"],
                                 "qty": n, "vol": n * uv, "kg": n * ukg, "owner": it.get("owner")})
            cur["vol"] += n * uv
            cur["kg"] += n * ukg
            remaining -= n
    if cur["lines"]:
        trucks.append(cur)
    return trucks, skipped


def fmt_plan(trucks, skipped, truck):
    head = (f"🚛 Yuklash rejasi\nFura: {truck['l']/100:g}×{truck['w']/100:g}×{truck['h']/100:g} m, "
            f"{truck['kg']/1000:g} t (hajmning {int(FILL*100)}% i hisobda)\n"
            f"Tartib: kelgan sana bo'yicha, eski → yangi\nJami: {len(trucks)} ta fura")
    out = [head]
    for n, t in enumerate(trucks, 1):
        out.append(f"\n━━ Fura #{n} — hajm {t['vol']/t['cap_v']*100:.0f}% | "
                   f"og'irlik {t['kg']/t['cap_kg']*100:.0f}% ({t['vol']:.1f} m³, {t['kg']:.0f} kg)")
        cur = None
        for ln in t["lines"]:
            if ln["arrival"] != cur:
                cur = ln["arrival"]
                out.append(f"📅 {cur}")
            who = f" 👤{ln['owner']}" if ln.get("owner") else ""
            out.append(f"  • #{ln['id']} {ln['name']}{who} × {ln['qty']} — {ln['vol']:.2f} m³, {ln['kg']:.0f} kg")
    if skipped:
        out.append("\n⚠️ Rejaga kirmadi:")
        out += [f"  • #{it['id']} {it['name']} — {why}" for it, why in skipped]
        out.append("Tuzatish: /tuzat ID jkg=2380 jkub=6.5 kim=Murod")
    out.append("\nFura yuklangach: /yuklandi 1  |  Excel/PDF: /reja excel, /reja pdf")
    return "\n".join(out)


PLAN_COLS = ["Fura", "Sana", "ID", "Tovar", "Egasi", "Soni", "Hajm m³", "Og'irlik kg"]


def plan_rows(trucks):
    rows = []
    for n, t in enumerate(trucks, 1):
        for ln in t["lines"]:
            rows.append([f"#{n}", ln["arrival"], ln["id"], ln["name"], ln.get("owner") or "", ln["qty"],
                         round(ln["vol"], 2), round(ln["kg"], 1)])
        rows.append([f"#{n} JAMI", "", "", f"hajm {t['vol']/t['cap_v']*100:.0f}% / og'irlik "
                     f"{t['kg']/t['cap_kg']*100:.0f}%", "", "", round(t["vol"], 2), round(t["kg"], 1)])
    return rows


def plan_excel(trucks, skipped):
    wb = Workbook()
    ws = wb.active
    ws.title = "Yuklash rejasi"
    ws.append(PLAN_COLS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for r in plan_rows(trucks):
        ws.append(r)
        if str(r[0]).endswith("JAMI"):
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
    for i, w in enumerate([10, 12, 6, 40, 18, 8, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if skipped:
        ws2 = wb.create_sheet("Kirmadi")
        ws2.append(["ID", "Tovar", "Sabab"])
        for it, why in skipped:
            ws2.append([it["id"], it["name"], why])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _font():
    for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/arial.ttf",
              "/Library/Fonts/Arial.ttf", os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")):
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("UZ", p))
            return "UZ"
    return "Helvetica"


def plan_pdf(trucks, skipped):
    font = _font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24,
                            topMargin=24, bottomMargin=24)
    st = getSampleStyleSheet()["Title"]
    st.fontName = font
    rows = plan_rows(trucks)
    t = Table([PLAN_COLS] + [[str(c) for c in r] for r in rows], repeatRows=1)
    style = [("FONTNAME", (0, 0), (-1, -1), font), ("FONTSIZE", (0, 0), (-1, -1), 8),
             ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
             ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.grey)]
    for i, r in enumerate(rows, 1):
        if str(r[0]).endswith("JAMI"):
            style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#DDE6F0")))
    t.setStyle(TableStyle(style))
    story = [Paragraph(f"Yuklash rejasi — {datetime.now():%Y-%m-%d %H:%M}", st), Spacer(1, 8), t]
    if skipped:
        story += [Spacer(1, 10), Paragraph("Rejaga kirmadi: " + "; ".join(
            f"#{it['id']} {it['name']} ({why})" for it, why in skipped), getSampleStyleSheet()["Normal"])]
    doc.build(story)
    return buf.getvalue()


# ---------- Telegram ----------
def is_admin(m: Message):
    return bool(m.from_user) and m.from_user.id in ADMINS


def in_group(m: Message, gid):
    return gid is not None and m.chat.id == gid


def allowed(m: Message):
    """Yig'ish va tahrirlash: zbornik guruhi (yoki guruh berilmasa — adminlar)."""
    if COLLECT_CHAT is None and LOAD_CHAT is None:
        return is_admin(m)
    return in_group(m, COLLECT_CHAT) or (m.chat.type == "private" and is_admin(m))


def can_plan(m: Message):
    """Reja, tovarlar ro'yxati, fura: yuklash guruhi ham qila oladi."""
    return allowed(m) or in_group(m, LOAD_CHAT)


async def send_long(bot: Bot, chat_id, text):
    while text:
        cut_at = text.rfind("\n", 0, 3800) if len(text) > 3800 else len(text)
        cut_at = cut_at if cut_at > 0 else 3800
        await bot.send_message(chat_id, text[:cut_at])
        text = text[cut_at:].lstrip("\n")


async def cmd_id(m: Message):
    await m.answer(f"Chat ID: {m.chat.id}\nTuri: {m.chat.type}")


COLLECT: dict = {}  # chat_id -> yig'ilayotgan rasm/xabarlar
DRAFT: dict = {}    # chat_id -> so'ralayotgan yuklar navbati
DEBOUNCE = 6        # oxirgi xabardan keyin shuncha soniya kutib, hammasini birga ko'radi

BATCH_PROMPT = """You help a logistics company in Uzbekistan. Below are photos and chat messages forwarded from Chinese warehouse agents (WeChat) about cargo that just arrived. Messages are often Uzbek/Russian/Chinese, e.g. "4433 6ta jumrak 2ta padon kaldi" = mark 4433, 6 valves and 2 pallets arrived. A number like 2380kg is a weight; a calculator screenshot like 6,960-4,580 = 2,380 means gross minus tare in kg. Numbers written on the cargo (e.g. 4433) are the shipping mark of the customer's order. The same photo may be forwarded several times: never duplicate lots.
Today is {today}. Forwarded messages:
{messages}
Look at ALL photos and messages together and group the cargo into lots (one kind of goods with one mark = one lot). Reply with ONLY valid JSON:
{"lots":[{"mark":"visible number or marking, or null","name":"what the goods are, short, in Uzbek","qty":number of pieces/pallets/bags or null,"package":"pallet|bag|box|other","weight_kg":total weight of this lot or null,"volume_m3":total volume of this lot or null,"length_cm":null,"width_cm":null,"height_cm":null,"owner":"customer name only if stated, else null","arrival_date":"YYYY-MM-DD or null","note":"short remark"}]}
Rules: never invent weight, volume or owner. If one total weight is given for several lots, put it in a lot only if it clearly belongs to that lot; otherwise null and say so in note. Convert to kg, m3, cm."""


def parse_lots(text):
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return []
    try:
        data = json.loads(m.group(0))
    except json.JSONDecodeError:
        return []
    lots = data.get("lots", []) if isinstance(data, dict) else []
    return [x for x in lots if isinstance(x, dict)]


async def extract_batch(files, texts):
    import anthropic  # kechiktirilgan import
    content = []
    for data, media in files[:8]:
        content.append({"type": "document" if media == "application/pdf" else "image", "source": {
            "type": "base64", "media_type": media, "data": base64.b64encode(data).decode()}})
    prompt = BATCH_PROMPT.replace("{today}", date.today().isoformat()).replace(
        "{messages}", "\n".join(f"- {t}" for t in texts) or "(no text)")
    content.append({"type": "text", "text": prompt})
    resp = await anthropic.AsyncAnthropic().messages.create(
        model=MODEL, max_tokens=3000, messages=[{"role": "user", "content": content}])
    return parse_lots("".join(b.text for b in resp.content if b.type == "text"))


def parse_reply(t):
    """'Murod 1200 kg 3.5 kub' -> {owner, kg, vol}. Kalit so'z bo'lmasa: 1-raqam kg, 2-raqam kub."""
    t, out = t.replace(",", "."), {}
    for key, pat in (("kg", r"(\d+(?:\.\d+)?)\s*(?:kg|кг)"),
                     ("vol", r"(\d+(?:\.\d+)?)\s*(?:kub|куб|m3|m³|m\^3)")):
        mm = re.search(pat, t, re.I)
        if mm:
            out[key] = float(mm.group(1))
            t = t.replace(mm.group(0), " ")
    nums = re.findall(r"\d+(?:\.\d+)?", t)
    for key in ("kg", "vol"):
        if key not in out and nums:
            out[key] = float(nums.pop(0))
    owner = re.sub(r"\d+(?:\.\d+)?", " ", t)
    owner = re.sub(r"\b(kim|egasi|yuki)\b\s*[:=]?", "", owner, flags=re.I)
    owner = re.sub(r"\s+", " ", owner).strip(" ,;:-")
    if owner:
        out["owner"] = owner
    return out


def lot_missing(lot):
    miss = []
    if not lot.get("owner"):
        miss.append("kimning yuki")
    if not num(lot.get("weight_kg")):
        miss.append("necha kg")
    if not (num(lot.get("volume_m3")) or all(num(lot.get(k)) for k in ("length_cm", "width_cm", "height_cm"))):
        miss.append("necha kub (m³)")
    return miss


def lot_question(lot, i, n, miss):
    known = []
    if lot.get("owner"):
        known.append(f"👤 {lot['owner']}")
    if num(lot.get("weight_kg")):
        known.append(f"{lot['weight_kg']:g} kg")
    if num(lot.get("volume_m3")):
        known.append(f"{lot['volume_m3']:g} m³")
    qty = f" ({lot['qty']:g} {lot.get('package') or 'dona'})" if num(lot.get("qty")) else ""
    lines = [f"📦 Yuk {i + 1}/{n}: {lot.get('mark') or '—'} · {lot.get('name') or '?'}{qty}"]
    if lot.get("note"):
        lines.append(f"📝 {lot['note']}")
    if known:
        lines.append("Bor: " + ", ".join(known))
    lines += [f"❓ Yozing: {', '.join(miss)}", "Masalan: Murod 1200 kg 3.5 kub", "O'tkazish: -   |   tugatish: /tayyor"]
    return "\n".join(lines)


def save_lots(lots):
    today, ids = date.today().isoformat(), []
    for lot in lots:
        qty = max(1, int(round(num(lot.get("qty")) or 1)))
        kg, vol = num(lot.get("weight_kg")), num(lot.get("volume_m3"))
        mark = (str(lot.get("mark") or "")).strip() or None
        name = " ".join(x for x in (mark, lot.get("name")) if x) or "Noma'lum yuk"
        ids.append(add_item(name, qty, num(lot.get("length_cm")), num(lot.get("width_cm")),
                            num(lot.get("height_cm")), kg / qty if kg else None,
                            parse_date(lot.get("arrival_date"), today), owner=lot.get("owner"),
                            mark=mark, vol=vol / qty if vol else None))
    return ids


async def finish(m: Message):
    d = DRAFT.pop(m.chat.id, None)
    if not d:
        return await m.answer("Saqlanadigan yuk yo'q.")
    rows = [get_item(i) for i in save_lots(d["lots"])]
    await m.answer("✅ Saqlandi:\n\n" + "\n".join(fmt_item(r) for r in rows) +
                   "\n\nTuzatish: /tuzat ID jkg=.. jkub=.. kim=..\nYuklash rejasi: /reja")
    if LOAD_CHAT and LOAD_CHAT != m.chat.id:
        await send_long(m.bot, LOAD_CHAT, "📥 Yangi yuk qo'shildi:\n\n" + "\n".join(fmt_item(r) for r in rows))


async def ask_next(m: Message):
    d = DRAFT.get(m.chat.id)
    while d and d["i"] < len(d["lots"]):
        lot = d["lots"][d["i"]]
        miss = lot_missing(lot)
        if miss:
            return await m.answer(lot_question(lot, d["i"], len(d["lots"]), miss))
        d["i"] += 1
    if d:
        await finish(m)


async def draft_answer(m: Message):
    d = DRAFT[m.chat.id]
    lot = d["lots"][d["i"]]
    if m.text.strip() not in ("-", "—"):
        p = parse_reply(m.text)
        if "owner" in p:
            lot["owner"] = p["owner"]
        if "kg" in p:
            lot["weight_kg"] = p["kg"]
        if "vol" in p:
            lot["volume_m3"] = p["vol"]
    d["i"] += 1
    await ask_next(m)


async def process_batch(m: Message, buf):
    if not os.getenv("ANTHROPIC_API_KEY"):
        return await m.answer("Rasmni ko'rish uchun ANTHROPIC_API_KEY kerak (.env ga yozing).")
    wait = await m.answer(f"🔍 {len(buf['files'])} ta rasm va {len(buf['texts'])} ta xabarni ko'ryapman…")
    try:
        lots = await extract_batch(buf["files"], buf["texts"])
    except Exception:
        logging.exception("Partiyani o'qishda xato")
        return await wait.edit_text("❌ O'qib bo'lmadi. Qayta yuboring.")
    if not lots:
        return await wait.edit_text("Yuk topilmadi. Aniqroq rasm yuboring yoki /qosh ishlating.")
    d = DRAFT.get(m.chat.id)
    if d:
        d["lots"].extend(lots)
        return await wait.edit_text(f"➕ {len(lots)} ta yuk navbatga qo'shildi.")
    DRAFT[m.chat.id] = {"lots": lots, "i": 0}
    await wait.edit_text(f"✅ {len(lots)} ta yuk topildi. Endi aniqlashtiramiz.")
    await ask_next(m)


def _buf(chat_id):
    return COLLECT.setdefault(chat_id, {"files": [], "texts": [], "task": None})


def _schedule(m: Message):
    b = _buf(m.chat.id)
    if b["task"]:
        b["task"].cancel()
    b["task"] = asyncio.create_task(_flush(m))


async def _flush(m: Message):
    try:
        await asyncio.sleep(DEBOUNCE)
    except asyncio.CancelledError:
        return
    buf = COLLECT.pop(m.chat.id, None)
    if buf and (buf["files"] or any(re.search(r"\d", t) for t in buf["texts"])):
        await process_batch(m, buf)


async def on_photo(m: Message, bot: Bot):
    if not allowed(m):
        return
    buf = io.BytesIO()
    await bot.download(m.photo[-1], destination=buf)
    b = _buf(m.chat.id)
    b["files"].append((buf.getvalue(), "image/jpeg"))
    if m.caption:
        b["texts"].append(m.caption)
    _schedule(m)


async def on_document(m: Message, bot: Bot):
    d = m.document
    if not allowed(m) or d.mime_type not in ("image/jpeg", "image/png", "image/webp", "application/pdf"):
        return
    if d.file_size and d.file_size > 15 * 1024 * 1024:
        return await m.answer("Fayl 15 MB dan katta.")
    buf = io.BytesIO()
    await bot.download(d, destination=buf)
    b = _buf(m.chat.id)
    b["files"].append((buf.getvalue(), d.mime_type))
    if m.caption:
        b["texts"].append(m.caption)
    _schedule(m)


def _forwarded(m: Message):
    return bool(getattr(m, "forward_origin", None) or getattr(m, "forward_date", None))


def batch_filter(m: Message):
    if not m.text or m.text.startswith("/") or not allowed(m):
        return False
    if not _forwarded(m) and m.text.lower().startswith("reja"):
        return False
    return (in_group(m, COLLECT_CHAT) or _forwarded(m)
            or m.chat.id in COLLECT or m.chat.id in DRAFT)


async def batch_text(m: Message):
    if m.chat.id in DRAFT and m.chat.id not in COLLECT and not _forwarded(m):
        return await draft_answer(m)
    _buf(m.chat.id)["texts"].append(m.text)
    _schedule(m)


async def cmd_tayyor(m: Message):
    if allowed(m):
        await finish(m)


async def cmd_bekor(m: Message):
    if not allowed(m):
        return
    b = COLLECT.pop(m.chat.id, None)
    if b and b["task"]:
        b["task"].cancel()
    DRAFT.pop(m.chat.id, None)
    await m.answer("Bekor qilindi.")


FIELDS = {"nom": "name", "soni": "qty", "uzun": "l", "eni": "w", "bal": "h", "kg": "kg", "sana": "arrival",
          "kim": "owner", "belgi": "mark", "kub": "vol", "jkg": "kg_total", "jkub": "vol_total"}


async def cmd_tuzat(m: Message, command: CommandObject):
    if not allowed(m):
        return
    args = command.args or ""
    parts = args.split(maxsplit=1)
    if not parts or not parts[0].lstrip("#").isdigit():
        return await m.answer('Format: /tuzat ID soni=20 jkg=2380 jkub=6.5 kim=Murod sana=20.11.2026 nom="Yangi nom"\n'
                              "(kg/kub — bitta dona uchun, jkg/jkub — jami)")
    item = get_item(int(parts[0].lstrip("#")))
    if not item:
        return await m.answer("Bunday ID yo'q.")
    upd = {}
    for k, v in re.findall(r'(\w+)=("[^"]*"|\S+)', parts[1] if len(parts) > 1 else ""):
        col, v = FIELDS.get(k.lower()), v.strip('"')
        if not col:
            return await m.answer(f"Noma'lum maydon: {k}. Mumkin: {', '.join(FIELDS)}")
        if col in ("name", "owner", "mark"):
            upd[col] = v
        elif col == "arrival":
            d = parse_date(v)
            if not d:
                return await m.answer("Sana formati: 2026-11-20 yoki 20.11.2026")
            upd[col] = d
        else:
            val = num(v)
            if val is None:
                return await m.answer(f"{k} musbat raqam bo'lishi kerak.")
            upd[col] = val
    if not upd:
        return await m.answer("Hech narsa o'zgarmadi.")
    qty = upd.get("qty", item["qty"]) or 1
    if "kg_total" in upd:
        upd["kg"] = upd.pop("kg_total") / qty
    if "vol_total" in upd:
        upd["vol"] = upd.pop("vol_total") / qty
    with db() as con:
        con.execute(f"UPDATE items SET {', '.join(c + '=?' for c in upd)} WHERE id=?", (*upd.values(), item["id"]))
    r = get_item(item["id"])
    if any(c in upd for c in ("l", "w", "h", "kg")):
        learn(r["name"], r["l"], r["w"], r["h"], r["kg"])
    await m.answer("✅ " + fmt_item(r) + ("\n🧠 O'lcham eslab qolindi." if any(c in upd for c in "lwh") else ""))


async def cmd_qosh(m: Message, command: CommandObject):
    if not allowed(m):
        return
    p = [x.strip() for x in (command.args or "").split("|")]
    if len(p) < 2:
        return await m.answer("Format: /qosh Nom | soni | uzun | eni | bal | kg | sana\n(o'lcham bo'lmasa — faqat nom | soni)")
    p += [""] * (7 - len(p))
    i = add_item(p[0], num(p[1]), num(p[2]), num(p[3]), num(p[4]), num(p[5]),
                 parse_date(p[6], date.today().isoformat()))
    await m.answer("✅ " + fmt_item(get_item(i)))


async def cmd_ochir(m: Message, command: CommandObject):
    if not allowed(m):
        return
    a = (command.args or "").lstrip("#")
    if not a.isdigit():
        return await m.answer("Format: /ochir ID")
    with db() as con:
        con.execute("DELETE FROM items WHERE id=?", (int(a),))
    await m.answer("🗑 O'chirildi.")


async def cmd_tovarlar(m: Message):
    if can_plan(m):
        await send_long(m.bot, m.chat.id, fmt_items(pending()))


async def cmd_fura(m: Message, command: CommandObject):
    if not can_plan(m):
        return
    key = LOAD_CHAT or m.chat.id
    a = (command.args or "").split()
    if a:
        v = [num(x) for x in a]
        if len(v) != 4 or None in v:
            return await m.answer("Format: /fura uzunlik eni balandlik kg  (sm va kg)\nMasalan: /fura 1360 245 270 22000")
        with db() as con:
            con.execute("INSERT INTO trucks(chat_id, l, w, h, kg) VALUES(?,?,?,?,?) ON CONFLICT(chat_id) DO UPDATE "
                        "SET l=excluded.l, w=excluded.w, h=excluded.h, kg=excluded.kg", (key, *v))
    t = get_truck(key)
    await m.answer(f"🚛 Fura: {t['l']:g}×{t['w']:g}×{t['h']:g} sm, {t['kg']:g} kg\n"
                   "O'zgartirish: /fura 1360 245 270 22000")


async def cmd_reja(m: Message, text: str = ""):
    """Reja har doim yuklash guruhiga tushadi (guruh berilmagan bo'lsa — shu chatga)."""
    if not can_plan(m):
        return
    items = pending()
    if not items:
        return await m.answer("Kutilayotgan tovar yo'q. Avval rasm yuboring.")
    target = LOAD_CHAT or m.chat.id
    truck = get_truck(target)
    trucks, skipped = build_plan(items, truck)
    LAST_PLAN[target] = trucks
    t = text.lower()
    if "excel" in t:
        await m.bot.send_document(target, BufferedInputFile(plan_excel(trucks, skipped), "yuklash_rejasi.xlsx"))
    elif "pdf" in t:
        await m.bot.send_document(target, BufferedInputFile(plan_pdf(trucks, skipped), "yuklash_rejasi.pdf"))
    else:
        await send_long(m.bot, target, fmt_plan(trucks, skipped, truck))
    if target != m.chat.id:
        await m.answer("✅ Reja yuklash guruhiga yuborildi.")


async def cmd_reja_cmd(m: Message, command: CommandObject):
    await cmd_reja(m, command.args or "")


async def reja_text(m: Message):
    await cmd_reja(m, m.text)


async def cmd_yuklandi(m: Message, command: CommandObject):
    if not can_plan(m):
        return
    key = LOAD_CHAT or m.chat.id
    a = (command.args or "").lstrip("#")
    plan = LAST_PLAN.get(key)
    if not a.isdigit() or not plan or not 1 <= int(a) <= len(plan):
        return await m.answer("Avval /reja qiling, keyin /yuklandi 1 (fura raqami).")
    with db() as con:
        for ln in plan[int(a) - 1]["lines"]:
            con.execute("UPDATE items SET qty=qty-? WHERE id=?", (ln["qty"], ln["id"]))
            con.execute("UPDATE items SET status='yuklandi', loaded_date=? WHERE id=? AND qty<=0.0001",
                        (date.today().isoformat(), ln["id"]))
    LAST_PLAN.pop(key, None)
    await m.answer(f"✅ Fura #{a} yuklandi deb belgilandi. Yangi reja uchun /reja.")


def by_status(status):
    with db() as con:
        return [dict(r) for r in con.execute("SELECT * FROM items WHERE status=? ORDER BY arrival, id", (status,))]


def total_kg(items):
    return sum((r["kg"] or 0) * (r["qty"] or 1) for r in items)


async def cmd_yuklangan(m: Message):
    if not can_plan(m):
        return
    items = by_status("yuklandi")
    head = f"✅ Ortilgan: {len(items)} qator, jami {total_kg(items):,.0f} kg\n\n"
    await send_long(m.bot, m.chat.id, head + fmt_items(items) if items else "Ortilgan yuk yo'q.")


async def cmd_boshqa(m: Message):
    if not can_plan(m):
        return
    items = by_status("boshqa")
    head = f"📂 Alohida mavzu (rejaga kirmaydi): {len(items)} qator, ma'lum og'irlik {total_kg(items):,.0f} kg\n\n"
    await send_long(m.bot, m.chat.id, head + fmt_items(items) if items else "Alohida yuk yo'q.")


def setup(dp):
    """bot.py dagi umumiy matn ishlovchisidan OLDIN chaqirilishi kerak."""
    r = dp.message.register
    r(cmd_id, Command("id"))
    r(on_photo, F.photo)
    r(on_document, F.document)
    r(cmd_tuzat, Command("tuzat"))
    r(cmd_qosh, Command("qosh"))
    r(cmd_ochir, Command("ochir"))
    r(cmd_tovarlar, Command("tovarlar"))
    r(cmd_fura, Command("fura"))
    r(cmd_reja_cmd, Command("reja"))
    r(cmd_yuklandi, Command("yuklandi"))
    r(cmd_tayyor, Command("tayyor"))
    r(cmd_bekor, Command("bekor"))
    r(cmd_yuklangan, Command("yuklangan"))
    r(cmd_boshqa, Command("boshqa"))
    r(batch_text, F.text, F.func(batch_filter))
    r(reja_text, F.text, F.func(lambda m: m.text.lower().startswith("reja") and can_plan(m)))
