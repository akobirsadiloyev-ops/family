"""
homes.json -> homes.xlsx (chiroyli, formatlangan Excel).

Ishlatish:
    python make_excel.py
"""

import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

IN_FILE = "homes.json"
OUT_FILE = "homes.xlsx"

# defect kalitlari -> o'zbekcha sarlavha
DEFECTS = {
    "study_level_defect": "Ta'lim darajasi",
    "employment_defect": "Bandlik",
    "income_defect": "Daromad",
    "month_income_defect": "Oylik daromad",
    "has_debt_defect": "Qarzdorlik",
    "wishes_defect": "Istaklar",
    "tadbirkorlik_defect": "Tadbirkorlik",
    "infratuzilma_defect": "Infratuzilma",
    "family_defect": "Oila",
}

# asosiy ustunlar: (json kaliti, sarlavha, matnmi?)
MAIN_COLS = [
    ("home_num", "Uy raqami", True),
    ("full_name", "Uy egasi (F.I.O.)", False),
    ("pinfl", "JSHSHIR", True),
    ("mobile_phone", "Telefon", True),
    ("cadaster_number", "Kadastr raqami", True),
    ("ownership_type", "Egalik turi", False),
    ("survey_date", "So'rov sanasi", True),
    ("id", "ID", True),
]


def parse_defects(rec):
    info = rec.get("defects_info") or {}
    raw = info.get("value") if isinstance(info, dict) else None
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return {}


def main():
    rows = json.load(open(IN_FILE, encoding="utf-8"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Uylar"

    # --- Sarlavhalar ---
    headers = ["№"] + [h for _, h, _ in MAIN_COLS] + list(DEFECTS.values())
    ws.append(headers)

    # --- Uslublar ---
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    defect_fill = PatternFill("solid", fgColor="C0392B")  # defect ustun sarlavhasi
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    yes_fill = PatternFill("solid", fgColor="FDE9E7")  # "Ha" katakcha fon

    n_main = 1 + len(MAIN_COLS)  # № + asosiy ustunlar soni
    for col_idx, cell in enumerate(ws[1], start=1):
        cell.fill = defect_fill if col_idx > n_main else header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    # --- Ma'lumot qatorlari ---
    for i, rec in enumerate(rows, start=1):
        defects = parse_defects(rec)
        line = [i]
        for key, _, _ in MAIN_COLS:
            line.append(rec.get(key, ""))
        for key in DEFECTS:
            line.append("Ha" if defects.get(key) else "Yo'q")
        ws.append(line)

        r = ws.max_row
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=col_idx)
            c.border = border
            c.alignment = left if col_idx == 2 else center  # F.I.O. chapga

    # --- JSHSHIR / telefon / kadastr / sana / id ni MATN qilish ---
    text_cols = {"№": False}
    # ustun harfini topamiz
    header_to_letter = {h: get_column_letter(idx + 1) for idx, h in enumerate(headers)}
    for key, title, is_text in MAIN_COLS:
        if is_text and title in header_to_letter:
            letter = header_to_letter[title]
            for row in range(2, ws.max_row + 1):
                ws[f"{letter}{row}"].number_format = "@"

    # --- "Ha" kataklarini bo'yash ---
    for title in DEFECTS.values():
        letter = header_to_letter[title]
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{letter}{row}"]
            if cell.value == "Ha":
                cell.fill = yes_fill
                cell.font = Font(bold=True, color="C0392B")

    # --- Ustun kengliklari ---
    widths = {
        "№": 6, "Uy raqami": 10, "Uy egasi (F.I.O.)": 32, "JSHSHIR": 17,
        "Telefon": 18, "Kadastr raqami": 22, "Egalik turi": 14,
        "So'rov sanasi": 14, "ID": 10,
    }
    for idx, h in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = widths.get(h, 13)

    # --- Sarlavhani muzlatish, filtr, balandlik ---
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.row_dimensions[1].height = 34
    ws.sheet_view.showGridLines = False

    wb.save(OUT_FILE)
    print(f"Tayyor: {len(rows)} qator -> {OUT_FILE}")
    print(f"Ustunlar: {len(headers)} ta")


if __name__ == "__main__":
    main()
