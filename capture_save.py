"""
online-mahalla.uz — "Oila a'zosi qo'shish" SAQLASH so'rovini ishonchli ushlash (v2).

Bu skript faqat SO'ROVLARni yozadi (javob tanasini O'QIMAYDI), shuning uchun
saqlash payload'i (POST) hech qachon tushib qolmaydi. Har bir POST darhol
faylga yoziladi.

Ishlatish:
    python capture_save.py
Keyin brauzerda BITTA yangi a'zoni qo'shing (Qidirish -> Saqlash), so'ng Ctrl+C.

Natija: explore_requests.json  (method, url, post_data, headers)
"""

import json
import sys

from playwright.sync_api import sync_playwright

PROFILE_DIR = ".browser_profile"
BASE = "https://online-mahalla.uz"
START_URL = (
    BASE + "/forms/survey_homes/1445717"
    "?obl_id=6&area_id=608&district_id=608035&street_id=60800147"
)
REQ_FILE = "explore_requests.json"
SEL_FILE = "explore_selects.json"

WANTED = {"xhr", "fetch", "document"}
KEEP = {"content-type", "accept", "x-requested-with", "authorization",
        "x-csrf-token", "x-xsrf-token"}

SELECT_JS = r"""
() => {
  const out = [];
  document.querySelectorAll('select').forEach(s => {
    const opts = [...s.options].map(o => ({value:o.value, text:(o.textContent||'').trim()}));
    if (opts.length) out.push({name:s.name||'', id:s.id||'', selected:s.value, options:opts});
  });
  return out;
}
"""

requests_log = []
best_selects = []


def save_reqs():
    with open(REQ_FILE, "w", encoding="utf-8") as f:
        json.dump(requests_log, f, ensure_ascii=False, indent=2)


def on_request(req):
    try:
        url = req.url
        if "online-mahalla.uz" not in url:
            return
        if req.resource_type not in WANTED:
            return
        pd = req.post_data  # sinxron, ishonchli
        entry = {
            "method": req.method,
            "url": url,
            "resource_type": req.resource_type,
            "headers": {k: v for k, v in req.headers.items() if k.lower() in KEEP},
            "post_data": pd,
        }
        requests_log.append(entry)
        if req.method in ("POST", "PUT", "PATCH"):
            n = len(pd or "")
            print(f"[{req.method}] {url[:115]}  <== payload {n} bayt")
            save_reqs()  # POSTni darhol saqlaymiz — yo'qolmasin
        else:
            print(f"[{req.method}] {url[:115]}")
    except Exception as e:
        print("  (req xato:", e, ")")


def main():
    global best_selects
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, viewport=None, args=["--start-maximized"]
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("request", on_request)

        print("\n>>> Brauzer ochildi. Kerak bo'lsa login qiling.")
        print(">>> Sahifa ochilyapti. Boshqa uyga o'tsangiz — URL id ni almashtiring.")
        try:
            page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print(f"    (ogohlantirish: {e})")

        print("\n>>> BITTA yangi a'zo qo'shing: '+qo'shish' -> to'ldiring -> Qidirish -> Saqlash")
        print(">>> Saqlash bosilganda '[POST] ... <== payload' qatorini ko'rasiz.")
        print(">>> TUGAGACH Ctrl+C bosing.\n")
        try:
            while True:
                page.wait_for_timeout(700)
                try:
                    sel = page.evaluate(SELECT_JS)
                    if sel:
                        best_selects = sel
                except Exception:
                    pass
        except KeyboardInterrupt:
            print("\n>>> To'xtatildi, saqlanmoqda...")
        finally:
            save_reqs()
            with open(SEL_FILE, "w", encoding="utf-8") as f:
                json.dump(best_selects, f, ensure_ascii=False, indent=2)
            print(f">>> {len(requests_log)} ta so'rov -> {REQ_FILE}")
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
