"""
yakuniy_natija.xlsx ga 2 ta yangi varaq qo'shadi:
  1) "Oila holati (670)"       — 670 xonadon; oila topilgan / "Tarkibi topilmadi (bazada yo'q)"
  2) "Oila a'zolari (taxminiy)" — oila topilgan uylarning a'zolari + taxminiy qarindoshlik,
                                  uy egasi har uyda BIR MARTA.

Manbalar:
  ESKI (390 uy)  -> olad_base/base_old.xlsx  (allaqachon qarindoshligi bor)
  YANGI (280 uy) -> family_280_results.json  (guess_relations bilan qarindoshlik)
  A'zo ismlari   -> olad_base/family_members.xlsx (pinfl -> F.I.Sh.)
  Uy egalari     -> olad_base/homes.json (670)

Ishlatish:  & "C:\\Program Files\\Python311\\python.exe" build_oila.py
"""
import datetime
import json
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import guess_relations as gr

OUT = "yakuniy_natija.xlsx"
TODAY = datetime.date(2026, 9, 27)

homes = json.load(open("olad_base/homes.json", encoding="utf-8"))
pin2home = {str(h["pinfl"]): h for h in homes}

# --- a'zo ismlari xaritasi (pinfl -> F.I.Sh.) ---
name_map = {}
fm = openpyxl.load_workbook("olad_base/family_members.xlsx", data_only=True, read_only=True)["Oila a'zolari"]
for r in range(2, fm.max_row + 1):
    p = fm.cell(r, 5).value
    nm = fm.cell(r, 4).value
    if p and nm:
        name_map[str(p).strip()] = nm
for h in homes:
    name_map.setdefault(str(h["pinfl"]), h.get("full_name"))

household = OrderedDict()  # owner_pinfl -> {name, hid, manba, members:[...]}

# ===== ESKI: base_old.xlsx =====
bo = openpyxl.load_workbook("olad_base/base_old.xlsx", data_only=True, read_only=True)["Oila a'zolari"]
for r in range(2, bo.max_row + 1):
    opinfl = bo.cell(r, 3).value
    mpinfl = bo.cell(r, 6).value
    if opinfl is None or mpinfl is None:
        continue
    opinfl = str(opinfl).strip(); mpinfl = str(mpinfl).strip()
    h = pin2home.get(opinfl)
    hh = household.setdefault(opinfl, {"name": bo.cell(r, 2).value,
                                       "hid": str(h["id"]) if h else str(bo.cell(r, 5).value),
                                       "manba": "avval", "members": []})
    hh["members"].append({"fio": name_map.get(mpinfl, ""), "pinfl": mpinfl,
                          "sana": bo.cell(r, 7).value, "rel": bo.cell(r, 8).value, "ish": ""})

# ===== YANGI: 280 fetch + guess_relations =====
owners = json.load(open("owners_280.json", encoding="utf-8"))
new_people = [{"num": o["id"], "name": o["name"], "pinfl": o["pinfl"], "dob": o["dob"]} for o in owners]
new_results = json.load(open("family_280_results.json", encoding="utf-8"))
new_rows, _ = gr.build_rows(new_people, new_results, TODAY)
for line, conf in new_rows:
    opinfl = str(line[2]); mpinfl = line[5]; rel = line[10]
    h = pin2home.get(opinfl)
    if not h:
        continue
    hh = household.setdefault(opinfl, {"name": line[1], "hid": str(h["id"]),
                                       "manba": "yangi", "members": []})
    if not mpinfl or str(mpinfl) == opinfl or (rel and str(rel).startswith("O'zi")):
        continue
    hh["members"].append({"fio": line[4], "pinfl": str(mpinfl), "sana": line[6],
                          "rel": rel, "ish": line[11]})

with_fam = {k: v for k, v in household.items() if v["members"]}
print(f"Oila topilgan uylar: {len(with_fam)} | jami a'zo: {sum(len(v['members']) for v in with_fam.values())}")

# YANGI uylardagi "oila yo'q" sabablari: 'selfonly' (yolg'iz) yoki 'nodata' (bazada yo'q)
new_real_owners = set()
for line, conf in new_rows:
    op = str(line[2]); mp = line[5]; rel = line[10]
    if mp and str(mp) != op and not (rel and str(rel).startswith("O'zi")):
        new_real_owners.add(op)
new_status = {}
for o in owners:
    p = o["pinfl"]
    if p in new_real_owners or p in with_fam:
        continue
    r = new_results.get(p)
    new_status[p] = "nodata" if (not r or r.get("status") != "ok") else "selfonly"

# --- uslublar ---
HF = PatternFill("solid", fgColor="1F4E78"); HFONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9"); GRP = Side(style="medium", color="9CB3D0")
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center")
GREEN = PatternFill("solid", fgColor="D5F0DD"); RED = PatternFill("solid", fgColor="FDE9E7")


def hdr(ws, headers):
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font, c.alignment = HF, HFONT, CTR
        c.border = Border(THIN, THIN, THIN, THIN)


wb = openpyxl.load_workbook(OUT)
for old in ("Oila holati (670)", "Oila a'zolari (taxminiy)", "Qolgan uylar (qarindosh kerak)"):
    if old in wb.sheetnames:
        del wb[old]

# ===== 1) Oila holati (670) =====
ws = wb.create_sheet("Oila holati (670)")
hdr(ws, ["№", "Uy ID", "Uy egasi (F.I.O.)", "JSHSHIR", "Telefon", "So'rov sanasi",
         "Oila a'zosi soni", "Holat"])
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GRAY = PatternFill("solid", fgColor="EEEEEE")
CAT = {
    "fam": ("Oila topilgan", GREEN, "1E7B3A"),
    "selfonly": ("Yolg'iz (oila a'zosi yo'q)", YELLOW, "8A6D00"),
    "nodata": ("Tarkibi topilmadi (bazada yo'q)", RED, "C0392B"),
    "other": ("Izlanmagan", GRAY, "555555"),
}
ordered = sorted(homes, key=lambda h: -len(with_fam.get(str(h["pinfl"]), {"members": []})["members"]))
n = {"fam": 0, "selfonly": 0, "nodata": 0, "other": 0}
for i, h in enumerate(ordered, 1):
    p = str(h["pinfl"])
    v = with_fam.get(p)
    if v:
        kind = "fam"; cnt = len(v["members"])
    else:
        kind = new_status.get(p, "other"); cnt = 0
    n[kind] += 1
    holat, fill, color = CAT[kind]
    ws.append([i, str(h["id"]), h.get("full_name"), p,
               h.get("mobile_phone"), h.get("survey_date"), cnt, holat])
    rr = ws.max_row
    for c in range(1, 9):
        cell = ws.cell(rr, c)
        cell.border = Border(THIN, THIN, THIN, THIN)
        cell.alignment = LFT if c in (3, 8) else CTR
        if c in (2, 4):
            cell.number_format = "@"
    ws.cell(rr, 8).fill = fill
    ws.cell(rr, 8).font = Font(bold=True, color=color)
for i, w in enumerate([5, 9, 30, 16, 18, 12, 12, 30], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:H{ws.max_row}"; ws.sheet_view.showGridLines = False

# ===== 2) Oila a'zolari (taxminiy) — uy egasi bir marta =====
ms = wb.create_sheet("Oila a'zolari (taxminiy)")
hdr(ms, ["№", "Uy egasi (F.I.O.)", "Uy egasi JSHSHIR", "Uy ID", "A'zo (F.I.Sh.)",
         "A'zo JSHSHIR", "A'zo sana", "Taxminiy qarindoshlik", "Ishonch", "Manba"])
CONF = {"Yuqori": ("D5F0DD", "1E7B3A"), "O'rta": ("FFF2CC", "8A6D00"), "Past": ("FDE9E7", "C0392B")}
gi = 0
for opinfl, v in with_fam.items():
    gi += 1
    members = sorted(v["members"], key=lambda m: str(m["rel"] or ""))
    for j, m in enumerate(members):
        first = j == 0
        row = ([gi, v["name"], opinfl, v["hid"]] if first else [None, None, None, None])
        row += [m["fio"], m["pinfl"], m["sana"], m["rel"], m["ish"], v["manba"]]
        ms.append(row)
        rr = ms.max_row
        for c in range(1, 11):
            cell = ms.cell(rr, c)
            top = GRP if (first and rr > 2) else THIN
            cell.border = Border(left=THIN, right=THIN, top=top, bottom=THIN)
            cell.alignment = LFT if c in (2, 5, 8) else CTR
            if c in (3, 6):
                cell.number_format = "@"
        if m["ish"] in CONF:
            fill, color = CONF[m["ish"]]
            ms.cell(rr, 9).fill = PatternFill("solid", fgColor=fill)
            ms.cell(rr, 9).font = Font(bold=True, color=color)
for i, w in enumerate([5, 28, 16, 9, 28, 16, 12, 26, 8, 8], 1):
    ms.column_dimensions[get_column_letter(i)].width = w
ms.freeze_panes = "A2"; ms.auto_filter.ref = f"A1:J{ms.max_row}"; ms.sheet_view.showGridLines = False

wb.save(OUT)
print(f"{OUT} yangilandi:")
print(f"  'Oila holati (670)': {n['fam']} oila topilgan, {n['selfonly']} yolg'iz, "
      f"{n['nodata']} bazada yo'q, {n['other']} izlanmagan")
print(f"  'Oila a'zolari (taxminiy)': {len(with_fam)} uy, {sum(len(v['members']) for v in with_fam.values())} a'zo")
