"""
Login qilingan brauzerdan (.browser_profile) online-mahalla.uz tokenini va brauzer
yuborgan sarlavhalarni oladi: token -> token.txt, sarlavhalar -> headers.json.
add_family.py 401 desa avtomatik chaqiriladi; qo'lda ham:  python get_token.py

Qanday aniqlaydi:
    Brauzer uy formasini (/web/v1/forms/survey_homes/...) MUVAFFAQIYATLI ochgan so'rov
    ushlanadi — aynan shu so'rovning tokeni va sarlavhalari olinadi. Eskirgan token
    bilan ketgan so'rov 401 oladi va hisobga olinmaydi; ochiq (login talab qilmaydigan)
    so'rovlar ham hisobga olinmaydi. add_family.py ham xuddi shu sarlavhalar bilan
    so'raydi, so'ng token shu so'rov bilan tekshiriladi.

Avval headless (ko'rinmas) urinadi; topolmasa ko'rinadigan brauzer ochadi —
online-mahalla.uz ga login qiling (LOGIN_WAIT soniya kutiladi). Login'dan keyin
sahifa boshqa joyga o'tib ketsa, dastur uy formasini o'zi qayta ochadi.
"""

import json
import sys
import urllib.error
import urllib.request

from playwright.sync_api import sync_playwright

PROFILE_DIR = ".browser_profile"
URL = ("https://online-mahalla.uz/forms/survey_homes/1445717"
       "?obl_id=6&area_id=608&district_id=608035&street_id=60800147")
# add_family.get_home() bilan bir xil so'rov
CHECK_URL = ("https://api.online-mahalla.uz/web/v1/forms/survey_homes/1445717"
             "?obl_id=6&area_id=608&district_id=608035&street_id=60800147")
FORM_API = "api.online-mahalla.uz/web/v1/forms/survey_homes/"   # login talab qiladi
TOKEN_FILE = "token.txt"
HEADERS_FILE = "headers.json"
HEADLESS_WAIT = 30     # soniya: ko'rinmas rejimda kutish
LOGIN_WAIT = 300       # soniya: ko'rinadigan brauzerda login qilish uchun
REOPEN_EVERY = 10      # soniya: login'dan keyin forma ochilmasa, qayta ochish oralig'i
# Saqlanmaydigan sarlavhalar: token alohida saqlanadi; qolganlarini urllib o'zi qo'yadi
SKIP_HEADERS = {"authorization", "host", "content-length", "connection", "accept-encoding",
                "content-type"}
DEFAULT_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://online-mahalla.uz",
    "Referer": "https://online-mahalla.uz/",
}


def browser_headers(all_headers):
    return {k: v for k, v in all_headers.items()
            if not k.startswith(":") and k.lower() not in SKIP_HEADERS}


def check(tok, headers):
    """add_family kabi so'rov (shu sarlavhalar bilan) -> (qabul qilindimi, izoh)."""
    h = dict(DEFAULT_HEADERS)
    h.update(headers)
    h["Authorization"] = f"Bearer {tok}"
    try:
        with urllib.request.urlopen(urllib.request.Request(CHECK_URL, headers=h), timeout=30) as r:
            return r.status == 200, f"HTTP {r.status}"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:200]}"
    except urllib.error.URLError as e:
        return False, f"ulanib bo'lmadi: {e.reason}"


def on_login_page(url):
    low = url.lower()
    return any(w in low for w in ("login", "auth", "sso", "oneid", "signin", "id.egov"))


def grab(headless):
    """Brauzer uy formasini muvaffaqiyatli ochgan so'rovdan -> (token, sarlavhalar) yoki (None, None)."""
    ok_resp = []
    rejected = set()

    def on_response(resp):
        try:
            if FORM_API in resp.url:
                if 200 <= resp.status < 300:
                    ok_resp.append(resp)
                else:
                    rejected.add(resp.status)
        except Exception:
            pass

    token, headers = None, None
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=headless, viewport=None,
            args=["--start-maximized"] if not headless else [])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("response", on_response)
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        if not headless:
            print(f">>> Brauzerda online-mahalla.uz ga login qiling ({LOGIN_WAIT // 60} daqiqa kutiladi).")
            print("    Login'dan keyin uy sahifasi o'zi ochiladi — hech narsa bosmang.")
        steps = (HEADLESS_WAIT if headless else LOGIN_WAIT) * 2
        for i in range(1, steps + 1):
            if ok_resp:
                try:
                    h = ok_resp[-1].request.all_headers()
                    a = h.get("authorization", "")
                    if a.lower().startswith("bearer "):
                        token, headers = a.split(" ", 1)[1].strip(), browser_headers(h)
                        break
                except Exception:
                    pass
                ok_resp.clear()
            # login'dan keyin sahifa bosh sahifaga o'tib ketsa — uy formasini qayta ochamiz
            if not headless and i % (REOPEN_EVERY * 2) == 0:
                try:
                    cur = page.url
                    if (cur.startswith("https://online-mahalla.uz") and not on_login_page(cur)
                            and "/forms/survey_homes/" not in cur):
                        page.goto(URL, wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
            try:
                page.wait_for_timeout(500)
            except Exception:       # brauzer yopib qo'yilsa
                break
        try:
            ctx.close()
        except Exception:
            pass
    if not token and rejected:
        print(f"   uy formasi so'rovi rad etildi (HTTP {', '.join(map(str, sorted(rejected)))})")
    return token, headers


def main():
    tok, headers = grab(headless=True)
    if not tok:
        print("Sessiya eskirgan, ko'rinadigan brauzer ochilyapti...")
        tok, headers = grab(headless=False)
    if not tok:
        print("Token olinmadi: brauzer uy formasini ocha olmadi. Login qilinganini tekshiring.")
        return 1
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        f.write(tok)
    with open(HEADERS_FILE, "w", encoding="utf-8") as f:
        json.dump(headers, f, ensure_ascii=False, indent=2)
    ok, info = check(tok, headers)
    if ok:
        print(f"token.txt yangilandi (sayt qabul qildi): {tok[:8]}...")   # to'liq token chiqmaydi
        return 0
    print(f"Brauzerda token ishladi, lekin dastur so'roviga sayt javobi: {info}")
    print("Sarlavhalar nomlari:", ", ".join(sorted(headers)))
    return 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
