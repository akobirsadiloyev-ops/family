"""Excel fayldan shaxsiy ma'lumotlarni olib tashlaydi (sahifa uchun kerak emas).

Ishlatish:  python3 tozalash.py asl_fayl.xlsx data.xlsx

- Formulalar o'rniga ularning qiymatlari saqlanadi.
- Barcha varaqlarda F.I.Sh., JSHSHIR, telefon, ID, manzil, mahalla,
  diplom raqami, mas'ul xodim ustunlari tozalanadi.
"""
import sys
import warnings
import openpyxl

warnings.filterwarnings("ignore")

DROP = {
    "f.i.sh.", "fio", "jshshir", "telefon raqami", "id", "manzil", "mahalla",
    "diplom raqami", "tashkilot stir", "mas'ul xodim", "mas'ul kafedra mudiri",
}


def norm(v):
    return " ".join(str(v or "").replace("‘", "'").replace("’", "'")
                    .replace("`", "'").lower().split())


def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True)
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 12) + 1):
            cols = [c.column for c in ws[r] if norm(c.value) in DROP]
            if len(cols) >= 1 and any(norm(c.value) in ("f.i.sh.", "jshshir", "id") for c in ws[r]):
                for col in cols:
                    for rr in range(r + 1, ws.max_row + 1):
                        ws.cell(rr, col).value = None
                break
    wb.save(dst)
    print("Tayyor:", dst)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Ishlatish: python3 tozalash.py asl_fayl.xlsx data.xlsx")
    main(sys.argv[1], sys.argv[2])
