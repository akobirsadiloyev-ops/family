"""
Brauzer sessiyasidan (login qilingan .browser_profile) joriy Bearer tokenni olib,
token.txt ga yozadi. add_family.py 401 (token eskirgan) desa avtomatik chaqiriladi;
qo'lda ham ishlatsa bo'ladi:  python get_token.py

Faqat sayt QABUL QILGAN token olinadi (API javobi 2xx). Sessiya eskirgan bo'lsa,
ko'rinadigan brauzer ochiladi — online-mahalla.uz ga login qiling (LOGIN_WAIT soniya kutiladi).

Avval headless (ko'rinmas) urinadi; topolmasa ko'rinadigan brauzer ochadi.
"""

import sys
from playwright.sync_api import sync_playwright

PROFILE_DIR = ".browser_profile"
URL = ("https://online-mahalla.uz/forms/survey_homes/1445717"
       "?obl_id=6&area_id=608&district_id=608035&street_id=60800147")
HEADLESS_WAIT = 30     # soniya: ko'rinmas rejimda token kutish
LOGIN_WAIT = 300       # soniya: ko'rinadigan brauzerda login qilish uchun


def grab(headless):
    token = {}

    def on_response(resp):
        # Eskirgan token bilan ketgan so'rov 401 qaytaradi — faqat muvaffaqiyatli javobdagisi olinadi
        try:
            if "api.online-mahalla.uz" in resp.url and 200 <= resp.status < 300:
                a = resp.request.headers.get("authorization", "")
                if a.lower().startswith("bearer "):
                    token["v"] = a.split(" ", 1)[1].strip()
        except Exception:
            pass

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
            print(f">>> Brauzerda online-mahalla.uz ga login qiling ({LOGIN_WAIT // 60} daqiqa kutiladi)...")
        for _ in range((HEADLESS_WAIT if headless else LOGIN_WAIT) * 2):
            if "v" in token:
                break
            try:
                page.wait_for_timeout(500)
            except Exception:       # brauzer yopib qo'yilsa
                break
        try:
            ctx.close()
        except Exception:
            pass
    return token.get("v")


def main():
    tok = grab(headless=True)
    if not tok:
        print("Headless topolmadi, ko'rinadigan brauzer ochilyapti...")
        tok = grab(headless=False)
    if tok:
        with open("token.txt", "w", encoding="utf-8") as f:
            f.write(tok)
        print("token.txt yangilandi:", tok[:8] + "...")   # to'liq token ekranga chiqmaydi
        return 0
    print("Token topilmadi. online-mahalla.uz ga login qilinganini tekshiring.")
    return 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
