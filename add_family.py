"""
585 oilaning a'zolarini online-mahalla.uz ga (survey_homes_family) kiritadi.

Manba: yakuniy_natija.xlsx -> "Oila a'zolari (taxminiy)" varag'i (build_oila.py yasaydi).
Natija: yakuniy_natija.xlsx -> "Kiritish" varag'i (har a'zo: Holat + Izoh/sabab).

Har bir a'zo uchun:
    1) member-can-add/check  -> shu uyda allaqachon bor bo'lsa -> "Allaqachon bor" (qo'shilmaydi).
                                Tekshirib bo'lmasa ham qo'shilmaydi (ikki marta kiritmaslik uchun),
                                keyingi safar qayta tekshiriladi.
    2) ma'lumot to'liqmi     -> qarindoshlik va ma'lumoti (ta'lim) bo'lmasa -> "Ma'lumot yetishmaydi"
    3) gcp/pinfl             -> F.I.Sh., hujjat seriya/raqami, hujjat turi
    4) saqlash (POST)        -> "Qo'shildi" yoki "Xato" + sabab (masalan "boshqa xonadonga biriktirilgan")

Qo'lda to'ldiriladigan ustunlar ("Kiritish" varag'ida, sariq):
    * "Ma'lumoti"  — kattalar (18+) uchun MAJBURIY: Oliy / Tugallanmagan oliy / O'rta maxsus /
                     O'rta / Ma'lumotsiz. Dastur ta'limni o'ylab topmaydi. Bolalarga yoshidan
                     kelib chiqib avtomatik qo'yiladi (7 yoshgacha Ma'lumotsiz, 8-17 O'rta),
                     kerak bo'lsa o'zgartiring.
    * "Qarindoshlik (tasdiqlangan)" — taxmin ishonchi O'rta/Past bo'lsa MAJBURIY; to'ldirilsa
                     taxmin o'rniga shu ishlatiladi.
    Bu qiymatlar add_progress.json ga ham saqlanadi (varaq o'chib ketsa ham yo'qolmaydi).
    Dastur ishlayotganda Excel'ni o'zgartirmang — keyin to'ldirib, qayta ishga tushiring.

Ishlatish:
    python add_family.py tekshir    # HAMMASINI saytda tekshiradi: kim kiritilgan, kim yo'q (hech narsa qo'shmaydi)
    python add_family.py            # SINOV: dastlabki 3 ta kiritilmagan a'zo
    python add_family.py 20         # dastlabki 20 ta
    python add_family.py all        # HAMMASI
    python add_family.py xato       # "Xato" bo'lganlarni qayta urinadi
    python add_family.py member <JSHSHIR>   # faqat shu a'zo

Qayta ishga tushirsa bo'ladi — holat add_progress.json da saqlanadi.
Token eskirsa (401): python get_token.py  -> keyin add_family.py ni qayta ishga tushiring.
"""

import datetime
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from hisobot import TEKSHIR_ISHONCH, read_rows as read_sheet

# --- Sozlamalar ----------------------------------------------------------
FILE = "yakuniy_natija.xlsx"
SRC_SHEET = "Oila a'zolari (taxminiy)"   # build_oila.py yasaydi
SHEET = "Kiritish"                        # natija va qo'lda to'ldiriladigan ustunlar
PROGRESS = "add_progress.json"
PROGRESS_OLD = "add_progress_run1.json"   # 1-bosqich natijalari (build_yakuniy.py ham o'qiydi)
TOKEN_FILE = "token.txt"   # get_token.py yozib beradi; kodga token yozilmaydi
HEADERS_FILE = "headers.json"   # get_token.py yozadi: brauzer uy formasini ochgandagi sarlavhalar

API = "https://api.online-mahalla.uz"
PARAMS = {"obl_id": 6, "area_id": 608, "district_id": 608035, "street_id": 60800147}
DELAY = 0.3          # a'zolar orasidagi pauza (soniya) — saqlashning o'zi sekin
GCP_RETRY = 3        # tezlik cheklovi/tarmoq xatosida qayta urinish
EXCEL_EVERY = 25     # har shuncha a'zodan keyin "Kiritish" varag'i yangilanadi
# -------------------------------------------------------------------------

# Saytdagi "Qarindoshligi" ro'yxati (relationship id)
RELATIONS = {
    1: "Otasi", 2: "Onasi", 3: "Opasi", 4: "Akasi", 5: "Singlisi", 6: "Ukasi",
    7: "Turmush o'rtog'i", 8: "Qizi", 9: "O'g'li", 10: "Qaynotasi", 11: "Qaynonasi",
    12: "Kelini", 13: "Boshqa", 14: "Kuyovi", 15: "Nevarasi",
}
# Ro'yxatda alohida bandi yo'q qarindoshlar -> "Boshqa"
REL_BOSHQA = ("qaynisi", "qaynsinglisi", "jiyani", "bobosi", "buvisi", "pochchasi",
              "yangasi", "ogayogli", "ogayqizi")

# Saytdagi "Ma'lumoti" ro'yxati (study_level_id)
STUDY_LEVELS = {"1": "Oliy", "2": "Tugallanmagan oliy", "3": "O'rta maxsus", "4": "O'rta",
                "5": "Ma'lumotsiz"}

COL_REL_OK = "Qarindoshlik (tasdiqlangan)"
COL_STUDY = "Ma'lumoti"

TODAY = datetime.date.today()


def norm_rel(s):
    s = re.sub(r"\(.*?\)", "", str(s or "").lower())   # "Qaynisi (…izoh…)" -> "qaynisi"
    return re.sub(r"[^a-z]", "", s)


REL_MAP = {norm_rel(v): k for k, v in RELATIONS.items()}
REL_MAP.update({"eri": 7, "xotini": 7, "nabira": 15, "nabirasi": 15})
REL_MAP.update({r: 13 for r in REL_BOSHQA})
STUDY_MAP = {norm_rel(v): k for k, v in STUDY_LEVELS.items()}


def load_token():
    if os.path.exists(TOKEN_FILE):
        t = open(TOKEN_FILE, encoding="utf-8").read().strip()
        if t:
            return t
    sys.exit(f"{TOKEN_FILE} topilmadi yoki bo'sh. Avval: python get_token.py")


TOKEN = None
BASE_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://online-mahalla.uz",
    "Referer": "https://online-mahalla.uz/",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"),
}
HEADERS = dict(BASE_HEADERS)


def set_token():
    """token.txt + headers.json -> HEADERS. So'rovlar brauzernikidek ketadi
    (sayt token bilan birga brauzer sarlavhalarini ham tekshirishi mumkin)."""
    global TOKEN
    TOKEN = load_token()
    HEADERS.clear()
    HEADERS.update(BASE_HEADERS)
    if os.path.exists(HEADERS_FILE):
        HEADERS.update(json.load(open(HEADERS_FILE, encoding="utf-8")))   # brauzerniki ustun
    HEADERS["Authorization"] = f"Bearer {TOKEN}"


class AuthError(Exception):
    pass


def refresh_token():
    """401 bo'lganda get_token.py orqali tokenni yangilaydi."""
    print(">>> Token yangilanmoqda (get_token.py)...")
    try:
        subprocess.run([sys.executable, "get_token.py"], timeout=420)   # login uchun 5 daqiqagacha
    except Exception as e:
        print("   token yangilashda xato:", e)
    set_token()
    print("   yangi token:", TOKEN[:8] + "...")


def _req(method, url, body=None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=HEADERS, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        txt = e.read().decode("utf-8", "replace")
        if e.code in (401, 403):
            path = url.split("?")[0].replace(API, "")
            raise AuthError(f"HTTP {e.code} {method} {path}: {txt[:200]}")
        return e.code, txt
    except urllib.error.URLError as e:
        return 0, f"URLError: {e.reason}"


def decode_birth(pinfl):
    p = str(pinfl).strip()
    if len(p) != 14 or not p.isdigit():
        return None
    c = int(p[0]); dd, mm, yy = int(p[1:3]), int(p[3:5]), int(p[5:7])
    base = {1: 1800, 2: 1800, 3: 1900, 4: 1900, 5: 2000, 6: 2000}.get(c)
    if base is None:
        return None
    try:
        return datetime.date(base + yy, mm, dd)
    except ValueError:
        return None


def age_of(pinfl):
    d = decode_birth(pinfl)
    if not d:
        return None
    return TODAY.year - d.year - ((TODAY.month, TODAY.day) < (d.month, d.day))




def deep_message(txt):
    """Server javobidan eng ma'noli xabarni chiqaradi."""
    try:
        j = json.loads(txt)
    except Exception:
        return txt[:200]
    best = None

    def walk(o):
        nonlocal best
        if isinstance(o, dict):
            for k in ("message", "error", "result_message", "detail"):
                v = o.get(k)
                if isinstance(v, str) and v.strip() and v.strip().lower() not in ("null", "ok"):
                    best = v.strip()
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(j)
    return best or txt[:200]


# --- API chaqiruvlari -----------------------------------------------------

_home_cache = {}


def get_home(home_id):
    """Uy id -> (survey_uuid, owner_pinfl, owner_phone). Keshlanadi."""
    if home_id in _home_cache:
        return _home_cache[home_id]
    q = "&".join(f"{k}={v}" for k, v in PARAMS.items())
    url = f"{API}/web/v1/forms/survey_homes/{home_id}?{q}"
    code, txt = _req("GET", url)
    if code != 200:
        raise RuntimeError(f"uy formasi olinmadi (HTTP {code}): {deep_message(txt)}")
    fd = (json.loads(txt).get("formData") or {})
    info = (str(fd.get("survey_uuid") or ""), str(fd.get("pinfl") or ""),
            str(fd.get("mobile_phone") or ""))
    _home_cache[home_id] = info
    return info


def member_can_add(pinfl, survey_uuid):
    """True = shu uyda ALLAQACHON bor (qo'shmaymiz)."""
    url = f"{API}/api/v1/survey_homes/family/member-can-add/check"
    code, txt = _req("POST", url, {"pinfl": int(pinfl), "survey_uuid": survey_uuid})
    if code != 200:
        return None  # noaniq — davom etamiz
    try:
        return bool(json.loads(txt).get("data", {}).get("status"))
    except Exception:
        return None


def gcp_person(pinfl, dob):
    url = (f"{API}/api/v1/gcp/pinfl?pinfl={pinfl}"
           f"&birth_date={dob:%Y-%m-%d}&form_name=survey_homes_family")
    for attempt in range(GCP_RETRY):
        code, txt = _req("GET", url)
        if code == 200:
            try:
                pd = json.loads(txt).get("data", {}).get("data")
            except Exception:
                pd = None
            return pd
        if code in (429, 0, 500, 502, 503):   # tezlik/tarmoq — kutib qayta urinamiz
            time.sleep(3 * (attempt + 1))
            continue
        return None
    return None


def build_payload(owner_pinfl, member_pinfl, dob, rel_id, pd, owner_phone, study_id):
    cur = str(pd.get("current_document") or "").strip()
    m = re.match(r"^(.*?)(\d+)$", cur)
    if not m:
        return None, "hujjat raqami yo'q"
    doc_serial = m.group(1).upper()
    doc_number = m.group(2)

    # hujjat turi: guvohnoma (type 80) -> 2, aks holda ID/pasport -> 1
    doc_type = "1"
    docs = pd.get("documents") or []
    matched = next((d for d in docs if str(d.get("document")) == cur), None)
    if matched is not None:
        doc_type = "2" if matched.get("type_id") == 80 else "1"
    elif docs and all(d.get("type_id") == 80 for d in docs):
        doc_type = "2"

    # Telefon MAJBURIY EMAS. Uy egasining raqamini qayta-qayta ishlatib bo'lmaydi
    # (sayt "bu raqam N marta ishlatilgan" deb rad etadi), shuning uchun bo'sh qoldiramiz.
    phone = ""

    payload = {
        "pinfl": int(member_pinfl),
        "full_name": pd.get("full_name") or "",
        "address": None,
        "phone": phone,
        "birth_date": f"{dob:%Y-%m-%d}",
        "study_level_id": study_id,
        "doc_serial": doc_serial,
        "doc_number": doc_number,
        "relationship": str(rel_id),
        "document_type": doc_type,
    }
    return payload, None


def save_member(owner_pinfl, survey_uuid, payload):
    q = "&".join(f"{k}={v}" for k, v in PARAMS.items())
    url = (f"{API}/web/v1/forms/survey_homes_family?_target=modal"
           f"&pinfl={owner_pinfl}&{q}&survey_uuid={survey_uuid}&owner_pinfl={owner_pinfl}")
    code, txt = _req("POST", url, payload)
    if code in (200, 201):
        # 200 bo'lsa ham ichida xato bo'lishi mumkin
        low = txt.lower()
        if code == 201 or ('"error":null' in low.replace(" ", "") or '"id"' in low):
            if "мавжуд" in txt or "mavjud" in txt or "biriktir" in txt.lower():
                return False, deep_message(txt)
            return True, None
        return False, deep_message(txt)
    return False, deep_message(txt)




# --- A'zolarni o'qish -----------------------------------------------------

def key_of(row):
    return f"{row['home_id']}:{row['member']}"


def text(v):
    return "" if v is None else str(v).strip()


def read_members(progress):
    """SRC_SHEET dagi a'zolar + "Kiritish" varag'idagi (yoki progress'dagi) qo'lda kiritilganlar."""
    try:
        wb = openpyxl.load_workbook(FILE, data_only=True, read_only=True)
    except FileNotFoundError:
        sys.exit(f"{FILE} topilmadi. Skriptni loyiha papkasida ishga tushiring.")
    _, src = read_sheet(wb, SRC_SHEET)
    qolda = {}
    if SHEET in wb.sheetnames:
        _, old = read_sheet(wb, SHEET)
        for r in old:
            k = text(r.get("Uy ID")) + ":" + text(r.get("A'zo JSHSHIR"))
            qolda[k] = {"rel": text(r.get(COL_REL_OK)), "study": text(r.get(COL_STUDY))}
    wb.close()

    rows, seen, owner = [], set(), {}
    for r in src:
        if r.get("Uy egasi JSHSHIR") is not None:      # uy egasi faqat guruhning 1-qatorida
            owner = {"owner": text(r["Uy egasi JSHSHIR"]), "home_id": text(r.get("Uy ID")),
                     "owner_name": text(r.get("Uy egasi (F.I.O.)"))}
        row = dict(owner, member=text(r.get("A'zo JSHSHIR")), name=text(r.get("A'zo (F.I.Sh.)")),
                   sana=r.get("A'zo sana"), rel=text(r.get("Taxminiy qarindoshlik")),
                   ishonch=text(r.get("Ishonch")))
        if not row.get("home_id") or not row["member"] or key_of(row) in seen:
            continue                                     # bir a'zo bir uyga faqat bir marta
        seen.add(key_of(row))
        k = key_of(row)
        saved = progress.get(k, {}).get("qolda", {})
        q = qolda.get(k, saved)
        row["rel_ok"], row["study"] = q.get("rel", ""), q.get("study", "")
        if (row["rel_ok"], row["study"]) != (saved.get("rel", ""), saved.get("study", "")):
            progress.setdefault(k, {})["qolda"] = {"rel": row["rel_ok"], "study": row["study"]}
        rows.append(row)
    return rows


# 18+ uchun taqsimot: 50% O'rta maxsus(3), 30% Oliy(1), 20% O'rta(4) — JSHSHIR bo'yicha barqaror
STUDY_PATTERN = ["3", "3", "3", "3", "3", "1", "1", "1", "4", "4"]


def default_study(age, pinfl=None):
    """Yoshdan kelib chiqib avtomatik ta'lim (qo'lda kiritilgan bo'lsa, o'sha ustun turadi):
    0-7 Ma'lumotsiz, 8-17 O'rta, 18+ 50/30/20 (JSHSHIR bo'yicha barqaror)."""
    if age is None:
        return ""
    if age <= 7:
        return STUDY_LEVELS["5"]
    if age <= 17:
        return STUDY_LEVELS["4"]
    if pinfl is None or not str(pinfl).isdigit():
        return STUDY_LEVELS["3"]
    return STUDY_LEVELS[STUDY_PATTERN[int(pinfl) % 10]]


def resolve(row):
    """-> (rel_id, study_id, yetishmaydi). yetishmaydi bo'sh bo'lmasa, a'zo hali kiritilmaydi."""
    missing = []
    if row["rel_ok"]:
        rel_id = REL_MAP.get(norm_rel(row["rel_ok"]))
        if rel_id is None:
            missing.append(f"qarindoshlik noto'g'ri yozilgan: {row['rel_ok']}")
    elif row["ishonch"] in TEKSHIR_ISHONCH:
        rel_id = None
        missing.append(f"qarindoshlikni tasdiqlang (taxmin: {row['rel'] or '-'}, "
                       f"ishonch: {row['ishonch']})")
    else:
        rel_id = REL_MAP.get(norm_rel(row["rel"]))
        if rel_id is None:
            missing.append("qarindoshlik noma'lum: " + (row["rel"] or "bo'sh"))

    study = row["study"] or default_study(age_of(row["member"]), row["member"])
    study_id = study if study in STUDY_LEVELS else STUDY_MAP.get(norm_rel(study))
    if study_id is None:
        missing.append(f"ma'lumoti noto'g'ri yozilgan: {study}" if study
                       else "ma'lumotini (ta'lim) kiriting")
    return rel_id, study_id, "; ".join(missing)


# --- "Kiritish" varag'i ---------------------------------------------------

LABELS = {  # kind -> (Holat, fon, matn rangi)
    "done": ("Qo'shildi", "D5F0DD", "1E7B3A"),
    "dup": ("Allaqachon bor", "DDE7F5", "1F4E78"),
    "absent": ("Kiritilmagan", "FFF2CC", "8A6D00"),
    "wait": ("Ma'lumot yetishmaydi", "FCE4D6", "C55A11"),
    "error": ("Xato", "FDE9E7", "C0392B"),
    None: ("Tekshirilmagan", "EEEEEE", "555555"),
}
HEADER = ["№", "Uy ID", "Uy egasi (F.I.O.)", "Uy egasi JSHSHIR", "A'zo (F.I.Sh.)", "A'zo JSHSHIR",
          "Tug'ilgan sana", "Yoshi", "Qarindoshlik (taxminiy)", "Ishonch", COL_REL_OK, COL_STUDY,
          "Holat", "Izoh"]
WIDTHS = [6, 9, 28, 16, 28, 16, 12, 6, 22, 9, 22, 18, 20, 50]
EDIT_COLS = (HEADER.index(COL_REL_OK) + 1, HEADER.index(COL_STUDY) + 1)


def write_sheet(rows, progress):
    """Natijani yakuniy_natija.xlsx -> "Kiritish" ga yozadi. Excel ochiq bo'lsa, o'tkazib yuboradi."""
    try:
        wb = openpyxl.load_workbook(FILE)
    except PermissionError:
        print(f"   ! {FILE} ochiq — varaq keyinroq yangilanadi")
        return
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, 0)
    ws.append(HEADER)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    edit_fill = PatternFill("solid", fgColor="FFFBE6")
    for c in EDIT_COLS:
        ws.cell(1, c).fill = PatternFill("solid", fgColor="B7950B")

    for i, row in enumerate(rows, 1):
        st = progress.get(key_of(row), {})
        holat, fill, color = LABELS.get(st.get("kind"), LABELS[None])
        ws.append([i, row["home_id"], row["owner_name"], row["owner"], row["name"], row["member"],
                   row["sana"], age_of(row["member"]), row["rel"], row["ishonch"], row["rel_ok"],
                   row["study"] or default_study(age_of(row["member"]), row["member"]), holat, st.get("note", "")])
        r = ws.max_row
        for c in range(1, len(HEADER) + 1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if c in (3, 5, 14) else "center",
                                       vertical="center")
            if c in (2, 4, 6):
                cell.number_format = "@"
        for c in EDIT_COLS:
            ws.cell(r, c).fill = edit_fill
        ws.cell(r, len(HEADER) - 1).fill = PatternFill("solid", fgColor=fill)
        ws.cell(r, len(HEADER) - 1).font = Font(bold=True, color=color)

    last = ws.max_row
    for col, values in ((EDIT_COLS[0], RELATIONS.values()), (EDIT_COLS[1], STUDY_LEVELS.values())):
        dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{get_column_letter(col)}2:{get_column_letter(col)}{max(last, 2)}")
    for i, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}{last}"
    ws.sheet_view.showGridLines = False
    for s in wb.worksheets:           # bir nechta varaq birga tanlanib qolmasin
        s.sheet_view.tabSelected = False
    wb.active = 0
    ws.sheet_view.tabSelected = True
    try:
        wb.save(FILE)
    except PermissionError:
        print(f"   ! {FILE} ochiq — varaq keyinroq yangilanadi")


# --- Asosiy ---------------------------------------------------------------

CHECK_FAILED = "saytda tekshirib bo'lmadi — keyingi safar qayta tekshiriladi"


def process_member(row, check_only=False):
    """Bitta a'zo -> (kind, note). check_only: faqat tekshiradi, saqlamaydi.
    AuthError yuqoriga uzatiladi."""
    member = row["member"]
    dob = decode_birth(member)
    if not dob:
        return "error", "JSHSHIR noto'g'ri"

    survey_uuid, owner_pinfl, owner_phone = get_home(row["home_id"])
    if not survey_uuid or not owner_pinfl:
        return "error", "uy survey_uuid topilmadi"

    exists = member_can_add(member, survey_uuid)
    if exists is True:
        return "dup", "Allaqachon shu uyda bor"
    if exists is None:          # noaniq — ikki marta kiritib qo'ymaslik uchun saqlamaymiz
        return "wait", CHECK_FAILED

    rel_id, study_id, missing = resolve(row)
    if missing:
        return "wait", missing
    if check_only:
        return "absent", "Tayyor — qo'shish rejimida kiritiladi"

    pd = gcp_person(member, dob)
    if not pd or pd.get("birth_date") == "-":
        return "error", "GCP: ma'lumot topilmadi (sana/JSHSHIR)"

    payload, err = build_payload(owner_pinfl, member, dob, rel_id, pd, owner_phone, study_id)
    if err:
        return "error", err

    ok, err = save_member(owner_pinfl, survey_uuid, payload)
    if ok:
        return "done", "Qo'shildi"
    return "error", err or "saqlash xatosi"


def merge_check(prev, kind, note):
    """Tekshiruv natijasini avvalgi holat bilan birlashtiradi (tekshir rejimi)."""
    pk = prev.get("kind")
    if kind == "dup":
        return (pk, prev.get("note")) if pk == "done" else (kind, note)
    if kind == "absent" or (kind == "wait" and note != CHECK_FAILED):   # saytda yo'q ekani aniq
        if pk == "done":
            return kind, f"avval qo'shilgan edi, lekin saytda topilmadi; {note}"
        if pk == "error":
            return pk, prev.get("note")  # xato sababi saqlanadi ("xato" rejimida qayta uriniladi)
        return kind, note
    if pk:                               # tekshirib bo'lmadi / xato — ma'lum holat saqlanadi
        return pk, prev.get("note")
    return kind, note


def summary(rows, progress):
    cnt = {}
    for row in rows:
        k = progress.get(key_of(row), {}).get("kind")
        cnt[k] = cnt.get(k, 0) + 1
    g = lambda k: cnt.get(k, 0)
    print(f"\nJami a'zo: {len(rows)}")
    print(f"  Kiritilgan:            {g('done') + g('dup'):>5}  "
          f"(dastur qo'shgan {g('done')}, oldin bor edi {g('dup')})")
    print(f"  Kiritilmagan, tayyor:  {g('absent'):>5}")
    print(f"  Ma'lumot yetishmaydi:  {g('wait'):>5}")
    print(f"  Xato:                  {g('error'):>5}")
    print(f"  Tekshirilmagan:        {g(None):>5}")
    print(f"Batafsil: {FILE} -> '{SHEET}' varag'i")


def main():
    # argumentlar:
    #   tekshir -> hammasini tekshirish | (bo'sh) -> 3 ta sinov | son -> shuncha | all -> hammasi
    #   xato -> xatolarni qayta urinish | member <JSHSHIR> -> faqat shu a'zo
    arg = (sys.argv[1] if len(sys.argv) > 1 else "3").lower()
    check_only = arg == "tekshir"

    progress = {}
    if os.path.exists(PROGRESS):
        progress = json.load(open(PROGRESS, encoding="utf-8"))
    if os.path.exists(PROGRESS_OLD):   # 1-bosqichda qo'shilgan/xato bo'lganlar ham hisobga olinadi
        for k, v in json.load(open(PROGRESS_OLD, encoding="utf-8")).items():
            progress.setdefault(k, v)
    rows = read_members(progress)
    set_token()

    if arg == "member":
        target = sys.argv[2]
        todo = [r for r in rows if r["member"] == target]
        print(f"Aniq a'zo rejimi: {target} -> {len(todo)} ta qator")
    elif check_only:
        todo = rows
        print(f"TEKSHIRISH rejimi: {len(rows)} a'zo saytda tekshiriladi (hech narsa qo'shilmaydi).")
    elif arg == "xato":
        todo = [r for r in rows if progress.get(key_of(r), {}).get("kind") == "error"]
        print(f"Xato bo'lganlar qayta urinilmoqda: {len(todo)} ta.")
    else:
        limit = None if arg == "all" else int(arg)
        done_kinds = {"done", "dup", "error"}
        todo = [r for r in rows if progress.get(key_of(r), {}).get("kind") not in done_kinds]
        print(f"Jami {len(rows)} a'zo. Tayyor/skip: {len(rows)-len(todo)}. Qolgan: {len(todo)}.")
        if limit is not None:
            todo = todo[:limit]
            print(f"Bu safar {len(todo)} ta (sinov). Hammasi uchun: python add_family.py all")

    def save_progress():
        tmp = PROGRESS + ".tmp"
        json.dump(progress, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        os.replace(tmp, PROGRESS)

    def set_status(row, kind, note):
        progress.setdefault(key_of(row), {}).update(kind=kind, note=note)
        save_progress()

    save_progress()   # qo'lda kiritilganlar darhol zaxiraga
    counts = {}
    labels = {"done": "QO'SHILDI", "dup": "ALLAQACHON bor", "absent": "KIRITILMAGAN",
              "wait": "MA'LUMOT YETISHMAYDI", "error": "XATO"}
    try:
        for i, row in enumerate(todo, 1):
            tag = f"[{i}/{len(todo)}] uy {row['home_id']} a'zo {row['member']} ({row['rel']})"
            kind = note = None
            for attempt in (1, 2):
                try:
                    kind, note = process_member(row, check_only)
                    break
                except AuthError as e:
                    if attempt == 2:
                        print(f"\n!!! Sayt so'rovni rad etdi: {e}\n"
                              "    To'xtatildi.\n"
                              "    python get_token.py ni ishga tushirib, ochilgan brauzerda "
                              "online-mahalla.uz ga login qiling, keyin qayta ishga tushiring "
                              "— qolganidan davom etadi.")
                        return
                    refresh_token()
                except Exception as e:
                    kind, note = "error", f"kutilmagan xato: {str(e)[:150]}"
                    break
            if check_only:
                kind, note = merge_check(progress.get(key_of(row), {}), kind, note)
            set_status(row, kind, note)
            counts[kind] = counts.get(kind, 0) + 1
            print(f"{tag} -> {labels.get(kind, kind)}"
                  + (f": {note[:90]}" if kind in ("error", "wait") else ""))
            if i % EXCEL_EVERY == 0:
                write_sheet(rows, progress)
            time.sleep(DELAY)
    finally:
        write_sheet(rows, progress)
        summary(rows, progress)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    main()
