"""
Brauzer sessiyasidan (login qilingan .browser_profile) joriy Bearer tokenni olib,
token.txt ga yozadi. add_family.py 401 (token eskirgan) desa avtomatik chaqiriladi;
qo'lda ham ishlatsa bo'ladi:  python get_token.py

Brauzer so'rovlaridagi har bir token add_family.py ishlatadigan so'rov bilan
(uy formasi) TEKSHIRILADI — faqat sayt qabul qilgani yoziladi. Ba'zi ochiq so'rovlar
eskirgan token bilan ham 200 qaytaradi, shuning uchun bunday tekshiruv kerak.

Avval headless (ko'rinmas) urinadi; topolmasa ko'rinadigan brauzer ochadi —
online-mahalla.uz ga login qiling (LOGIN_WAIT soniya kutiladi).
"""

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
HEADLESS_WAIT = 30     # soniya: ko'rinmas rejimda token kutish
LOGIN_WAIT = 300       # soniya: ko'rinadigan brauzerda login qilish uchun


def check(tok):
    """Token bilan uy formasini so'raydi -> (qabul qilindimi, izoh)."""
    req = urllib.request.Request(CHECK_URL, headers={
        "Authorization": f"Bearer {tok}",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://online-mahalla.uz",
        "Referer": "https://online-mahalla.uz/",
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"),
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status == 200, f"HTTP {r.status}"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:200]}"
    except urllib.error.URLError as e:
        return False, f"ulanib bo'lmadi: {e.reason}"


def grab(headless):
    """-> (token yoki None, oxirgi rad etilgan tokenning izohi)."""
    seen = []           # brauzer so'rovlarida uchragan tokenlar (tartib bilan)
    tried = {}          # token -> (ok, izoh)

    def on_request(req):
        try:
            if "api.online-mahalla.uz" in req.url:
                a = req.headers.get("authorization", "")
                if a.lower().startswith("bearer "):
                    t = a.split(" ", 1)[1].strip()
                    if t and t not in seen:
                        seen.append(t)
        except Exception:
            pass

    found, last = None, ""
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=headless, viewport=None,
            args=["--start-maximized"] if not headless else [])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("request", on_request)
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        if not headless:
            print(f">>> Brauzerda online-mahalla.uz ga login qiling ({LOGIN_WAIT // 60} daqiqa kutiladi)...")
        for _ in range((HEADLESS_WAIT if headless else LOGIN_WAIT) * 2):
            for t in list(seen):
                if t not in tried:
                    tried[t] = check(t)
                    if not tried[t][0]:
                        last = tried[t][1]
                        print(f"   token {t[:8]}... rad etildi ({last[:60]})")
                if tried[t][0]:
                    found = t
                    break
            if found:
                break
            try:
                page.wait_for_timeout(500)
            except Exception:       # brauzer yopib qo'yilsa
                break
        try:
            ctx.close()
        except Exception:
            pass
    return found, last


def main():
    tok, last = grab(headless=True)
    if not tok:
        print("Headless topolmadi, ko'rinadigan brauzer ochilyapti...")
        tok, last = grab(headless=False)
    if tok:
        with open("token.txt", "w", encoding="utf-8") as f:
            f.write(tok)
        print("token.txt yangilandi (sayt qabul qildi):", tok[:8] + "...")   # to'liq token chiqmaydi
        return 0
    print("Ishlaydigan token topilmadi. online-mahalla.uz ga login qilinganini tekshiring.")
    if last:
        print("Saytning oxirgi javobi:", last)
    return 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
