"""Excel jadvalni botning bazasiga yuklaydi (boshlang'ich holat).

Ustunlar tartibi (boshlang'ich Excel fayllar ham shunday):
A №, B kelgan sana, C rasm, D markirovka, E tovar nomi, F joy soni,
G og'irlik kg (JAMI), H hajm CBM (JAMI), I kelgan rasm, J yuklangan sana, K izoh.

Qoida: "Yuklangan sana" to'ldirilgan qator = ortilgan. Bo'sh bo'lsa --status bo'yicha.
Qayta yuklasangiz ham takror qo'shilmaydi (manba + № bo'yicha tekshiriladi).

    python import_excel.py fayl.xlsx [--sheet 0] [--status kutilmoqda|yuklandi|boshqa]
                           [--loaded-date 2026-11-12] [--topic 2097] [--skip 17,18]
"""
import argparse
import os
import re
import sys
from datetime import date, datetime

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import planner  # noqa: E402

TR = {  # xitoycha nomlar -> o'zbekcha (asl nom izohda saqlanadi)
    "软密封闸阀": "Yumshoq zichlamali zadvijka (klapan)",
    "手柄对夹蝶阀": "Dastakli disk (kapalak) klapan",
    "冷却塔填料": "Sovutish minorasi to'ldirgichi",
    "板材": "List material (plita)",
    "装载机": "Yuk ortgich (frontal yuklagich)",
    "四个推土斗+1个铲斗": "4 ta buldozer cho'michi + 1 ta kovsh",
    "轮胎": "G'ildiraklar (shinalar)",
    "配件": "Ehtiyot qismlar",
    "油缸": "Gidravlik silindrlar",
    "油箱": "Moy baklari",
    "风机": "Ventilyatorlar",
    "装载机+4个轮胎": "Yuk ortgich + 4 ta g'ildirak",
    "装载机+1个斗子+黄色配件": "Yuk ortgich + 1 ta cho'mich + sariq ehtiyot qismlar",
}


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return v.date().isoformat() if isinstance(v, datetime) else v.isoformat()
    s = re.sub(r"\D", "", str(v).split(".")[0]) if re.fullmatch(r"\d{8}(\.0)?", str(v)) else str(v)
    for f in ("%Y%m%d", "%Y.%m.%d", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s.strip(), f).date().isoformat()
        except ValueError:
            pass
    return None


def to_num(v):
    """-> (son, izoh). '2765*2=5530' -> 5530 va izohda asl yozuv."""
    if v is None or v == "":
        return None, None
    if isinstance(v, (int, float)):
        return (float(v), None) if v > 0 else (None, None)
    s = str(v).strip()
    nums = re.findall(r"\d+(?:[.,]\d+)?", s)
    if "=" in s and nums:
        return float(nums[-1].replace(",", ".")), f"Jadvalda: {s}"
    if re.fullmatch(r"\d+(?:[.,]\d+)?", s):
        return float(s.replace(",", ".")), None
    return None, f"Raqamga aylanmadi: {s}"


def import_sheet(path, sheet=0, status="kutilmoqda", loaded_date=None, topic=None, skip=(), override=None):
    override = override or {}
    planner.init_db()
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[sheet] if isinstance(sheet, int) else wb[sheet]
    base, added, dup, last_date = os.path.basename(path), 0, 0, None
    with planner.db() as con:
        for row in ws.iter_rows(min_row=2, values_only=True):
            n, d, _, mark, name, pcs, kg, cbm, _, loaded, remark = (list(row) + [None] * 11)[:11]
            if n is None or (name is None and pcs is None and kg is None):
                continue  # "Jami" yoki bo'sh qator
            n = int(n)
            if n in skip:
                continue
            ov = override.get(n, {})
            arrival = to_date(d) or last_date or date.today().isoformat()
            last_date = arrival
            key = f"{base}|{ws.title}|{n}"
            if con.execute("SELECT 1 FROM items WHERE src=?", (key,)).fetchone():
                dup += 1
                continue
            qty = max(1, int(round(to_num(pcs)[0] or 1)))
            tot_kg, kg_note = to_num(kg)
            tot_cbm, _ = to_num(cbm)
            notes = [f"Manba: {base} №{n}"]
            mark = str(int(mark)) if isinstance(mark, (int, float)) else (str(mark).strip() if mark else None)
            if mark and re.search(r"kg", mark, re.I):
                notes.append(f"Markirovka ustunida: {mark}")
                mark = None
            uz = TR.get(name, name) or "Nomi yo'q"
            if name in TR:
                notes.append(f"Asl nom: {name}")
            full = f"{mark} {uz}" if mark else uz
            ld = to_date(loaded)
            st = "yuklandi" if ld else ov.get("status", status)
            if st == "yuklandi" and not ld:
                ld = loaded_date
            for x in (kg_note, remark, ov.get("note")):
                if x:
                    notes.append(str(x))
            con.execute(
                "INSERT INTO items(name,qty,l,w,h,kg,arrival,status,owner,mark,vol,created_at,loaded_date,note,topic,src)"
                " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (full, qty, None, None, None, tot_kg / qty if tot_kg else None, arrival, st, None, mark,
                 tot_cbm / qty if tot_cbm else None, datetime.now().strftime("%Y-%m-%d %H:%M"),
                 ld, " | ".join(notes), topic, key))
            added += 1
    print(f"{base} [{ws.title}]: {added} qo'shildi, {dup} takror (o'tkazildi)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--sheet", default="0")
    ap.add_argument("--status", default="kutilmoqda", choices=["kutilmoqda", "yuklandi", "boshqa"])
    ap.add_argument("--loaded-date")
    ap.add_argument("--topic")
    ap.add_argument("--skip", default="")
    a = ap.parse_args()
    import_sheet(a.file, int(a.sheet) if a.sheet.isdigit() else a.sheet, a.status, a.loaded_date, a.topic,
                 {int(x) for x in a.skip.split(",") if x.strip()})
