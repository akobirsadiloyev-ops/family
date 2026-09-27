"""
yakuniy_natija.xlsx yasaydi — 3 manbadan birlashtirib:
  * add_progress_run1.json + add_progress.json  -> har a'zoning yakuniy holati
  * base_old.xlsx (978 a'zo)                     -> a'zolar ro'yxati
  * olad_base/homes.json (670 uy)                -> barcha uy egalari

Varaqlar:
  1) "Oila a'zolari" — 978 a'zo, NATIJA rangli; uy egasi har xonadonda BIR MARTA
     (qolgan a'zolar qatorida egasi va № bo'sh), xonadonlar orasida ajratuvchi chiziq.
  2) "Xonadonlar (670)" — barcha 670 uy egasi + qo'shilgan a'zolar jamlanmasi.

Ishlatish:  python build_yakuniy.py
"""
import json
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "yakuniy_natija.xlsx"

# --- yakuniy holat (run-2 run-1 ustidan) ---
fin = {}
fin.update(json.load(open("add_progress_run1.json", encoding="utf-8")))
fin.update(json.load(open("add_progress.json", encoding="utf-8")))


def label(v):
    if v["kind"] == "done":
        return "done", "Qo'shildi"
    if v["kind"] == "dup":
        return "dup", "Dublikat (allaqachon shu uyda)"
    n = v["note"]
    if "GCP" in n:
        return "gcp", "GCP topilmadi (bazada yo'q/sana mos emas)"
    if "мавжуд" in n or "mavjud" in n.lower() or ":400" in n:
        return "band", "Band (boshqa xonadonga biriktirilgan)"
    return "other", n[:60]


# --- a'zolar (base_old.xlsx) ---
sw = openpyxl.load_workbook("base_old.xlsx", data_only=True, read_only=True)["Oila a'zolari"]
raw = []
for r in range(2, sw.max_row + 1):
    vals = [sw.cell(r, c).value for c in range(1, 9)]
    if vals[1] is None and vals[5] is None:
        continue
    hid = str(vals[4]).strip() if vals[4] is not None else ""
    mp = str(vals[5]).strip() if vals[5] is not None else ""
    kind, txt = label(fin[f"{hid}:{mp}"]) if f"{hid}:{mp}" in fin else ("none", "—")
    raw.append((vals, kind, txt, hid))

# --- uslublar ---
HF = PatternFill("solid", fgColor="1F4E78")
HFONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
GRP = Side(style="medium", color="9CB3D0")
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center")
ST = {
    "done": ("D5F0DD", Font(bold=True, color="1E7B3A")),
    "band": ("FDE9E7", Font(bold=True, color="C0392B")),
    "gcp": ("FFF2CC", Font(bold=True, color="8A6D00")),
    "dup": ("DDE7F5", Font(bold=True, color="1F4E78")),
    "other": ("EEEEEE", Font()),
    "none": ("FFFFFF", Font()),
}


def hdr_row(ws, headers):
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font, c.alignment = HF, HFONT, CTR
        c.border = Border(THIN, THIN, THIN, THIN)


wb = openpyxl.Workbook()

# ============ 1-varaq: Oila a'zolari ============
ws = wb.active
ws.title = "Oila a'zolari"
hdr_row(ws, ["№", "Uy egasi (F.I.O.)", "Uy egasi JSHSHIR", "Uy egasi sana",
             "Uy egasi ID", "A'zo JSHSHIR", "A'zo sana", "Qarindoshlik", "NATIJA"])
prev = None
for vals, kind, txt, hid in raw:
    first = hid != prev
    show = list(vals)
    if not first:
        show[0] = show[1] = show[2] = show[3] = show[4] = None
    ws.append(show + [txt])
    rr = ws.max_row
    for c in range(1, 10):
        top = GRP if (first and rr > 2) else THIN
        cell = ws.cell(rr, c)
        cell.border = Border(left=THIN, right=THIN, top=top, bottom=THIN)
        cell.alignment = LFT if c in (2, 9) else CTR
    fill, fnt = ST[kind]
    ws.cell(rr, 9).fill = PatternFill("solid", fgColor=fill)
    ws.cell(rr, 9).font = fnt
    prev = hid
for i, w in enumerate([5, 30, 16, 12, 10, 16, 12, 13, 42], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{ws.max_row}"
ws.sheet_view.showGridLines = False

# ============ jamlanma (uy bo'yicha) ============
agg = defaultdict(lambda: {"done": 0, "band": 0, "gcp": 0, "dup": 0, "other": 0, "jami": 0})
for vals, kind, txt, hid in raw:
    a = agg[hid]
    a["jami"] += 1
    if kind in a:
        a[kind] += 1

# ============ 2-varaq: Xonadonlar (670) ============
homes = json.load(open("olad_base/homes.json", encoding="utf-8"))
ss = wb.create_sheet("Xonadonlar (670)")
hdr_row(ss, ["Uy ID", "Uy raqami", "Uy egasi (F.I.O.)", "JSHSHIR", "Telefon",
             "So'rov sanasi", "Jami a'zo", "Qo'shildi", "Band", "GCP topilmadi",
             "Dublikat", "Izoh"])
GREEN = PatternFill("solid", fgColor="D5F0DD")


def sort_key(h):
    a = agg.get(str(h["id"]))
    return (-(a["done"] if a else -1), -(a["jami"] if a else 0))


for h in sorted(homes, key=sort_key):
    hid = str(h["id"])
    a = agg.get(hid)
    if a is None:
        izoh = "A'zo ro'yxatida yo'q"
        vals = [hid, h.get("home_num"), h.get("full_name"), h.get("pinfl"),
                h.get("mobile_phone"), h.get("survey_date"), 0, 0, 0, 0, 0, izoh]
    else:
        izoh = "" if a["done"] else "Hech biri qo'shilmadi"
        vals = [hid, h.get("home_num"), h.get("full_name"), h.get("pinfl"),
                h.get("mobile_phone"), h.get("survey_date"), a["jami"], a["done"],
                a["band"], a["gcp"], a["dup"], izoh]
    ss.append(vals)
    rr = ss.max_row
    for c in range(1, 13):
        cell = ss.cell(rr, c)
        cell.border = Border(THIN, THIN, THIN, THIN)
        cell.alignment = LFT if c in (3, 12) else CTR
        cell.number_format = "@" if c in (1, 4, 5) else cell.number_format
    if a and a["done"] > 0:
        ss.cell(rr, 8).fill = GREEN
        ss.cell(rr, 8).font = Font(bold=True, color="1E7B3A")
for i, w in enumerate([9, 8, 30, 16, 18, 12, 8, 9, 7, 13, 9, 22], 1):
    ss.column_dimensions[get_column_letter(i)].width = w
ss.freeze_panes = "A2"
ss.auto_filter.ref = f"A1:L{ss.max_row}"
ss.sheet_view.showGridLines = False

# ============ 3-varaq: Qolgan uylar (qarindosh topilmagan) ============
def decode_birth(pinfl):
    p = str(pinfl).strip()
    if len(p) != 14 or not p.isdigit():
        return ""
    c = int(p[0]); dd, mm, yy = int(p[1:3]), int(p[3:5]), int(p[5:7])
    base = {1: 1800, 2: 1800, 3: 1900, 4: 1900, 5: 2000, 6: 2000}.get(c)
    try:
        import datetime
        return datetime.date(base + yy, mm, dd).strftime("%d.%m.%Y")
    except Exception:
        return ""


qolgan = [h for h in homes if str(h["id"]) not in agg]
qs = wb.create_sheet("Qolgan uylar (qarindosh kerak)")
hdr_row(qs, ["№", "Uy ID", "Uy raqami", "Uy egasi (F.I.O.)", "JSHSHIR",
             "Egasi tug'ilgan sana", "Telefon", "So'rov sanasi", "Holat"])
for i, h in enumerate(qolgan, 1):
    qs.append([i, str(h["id"]), h.get("home_num"), h.get("full_name"),
               str(h.get("pinfl")), decode_birth(h.get("pinfl")),
               h.get("mobile_phone"), h.get("survey_date"), "Qarindoshi izlanmagan"])
    rr = qs.max_row
    for c in range(1, 10):
        cell = qs.cell(rr, c)
        cell.border = Border(THIN, THIN, THIN, THIN)
        cell.alignment = LFT if c in (4, 9) else CTR
        if c in (2, 5, 6):
            cell.number_format = "@"
for i, w in enumerate([5, 9, 8, 30, 16, 15, 18, 12, 22], 1):
    qs.column_dimensions[get_column_letter(i)].width = w
qs.freeze_panes = "A2"
qs.auto_filter.ref = f"A1:I{qs.max_row}"
qs.sheet_view.showGridLines = False

wb.save(OUT)
added = sum(1 for v in fin.values() if v["kind"] == "done")
print(f"{OUT} yasaldi:")
print(f"  1) Oila a'zolari: {len(raw)} a'zo (egasi har uyda bir marta)")
print(f"  2) Xonadonlar (670): {len(homes)} uy egasi")
print(f"  3) Qolgan uylar (qarindosh kerak): {len(qolgan)} uy")
print(f"  Jami qo'shilgan a'zo: {added}")
