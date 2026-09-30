"""
demotest.uzedu.uz — saytning ichki API'sini aniqlash (razvedka).

Skript brauzerni ochadi va siz qo'lda test ishlaganingizda barcha XHR/fetch
so'rov va javoblarini faylga yozadi. Shundan keyin savollar qaysi URL'dan
kelayotgani aniq bo'ladi va to'liq yuklab oluvchi skript yozish mumkin.

Ishlatish:
    pip install playwright
    python -m playwright install chromium
    python demotest_capture.py

Keyin ochilgan brauzerda: (kerak bo'lsa login) -> fan/testni tanlang ->
1-2 ta savolni ko'ring -> terminalda Ctrl+C.

Natija:
    demotest_requests.json  — har bir so'rov: method, url, post_data, status,
                              content-type va javob tanasi (JSON bo'lsa to'liq,
                              boshqa bo'lsa boshi 2000 belgi).
"""

import json
import signal
import sys

from playwright.sync_api import sync_playwright

PROFILE_DIR = ".demotest_profile"
START_URL = "https://demotest.uzedu.uz/uz"
OUT_FILE = "demotest_requests.json"

HOST = "uzedu.uz"
WANTED = {"xhr", "fetch"}
KEEP_HEADERS = {"content-type", "accept", "authorization", "cookie",
                "x-csrf-token", "x-xsrf-token", "x-requested-with"}
BODY_LIMIT = 2000

log = []


def save():
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


def on_response(resp):
    try:
        req = resp.request
        if HOST not in req.url or req.resource_type not in WANTED:
            return

        entry = {
            "method": req.method,
            "url": req.url,
            "status": resp.status,
            "post_data": req.post_data,
            "req_headers": {k: v for k, v in req.all_headers().items()
                            if k.lower() in KEEP_HEADERS},
            "content_type": resp.headers.get("content-type", ""),
        }

        # Javob tanasi: JSON bo'lsa to'liq, aks holda qisqartirib.
        try:
            if "json" in entry["content_type"]:
                entry["json"] = resp.json()
            else:
                entry["text"] = resp.text()[:BODY_LIMIT]
        except Exception as e:          # javob o'qilmasa — o'tkazib yuboramiz
            entry["body_error"] = str(e)

        log.append(entry)
        save()
        print(f"[{len(log):3}] {entry['status']} {req.method} {req.url[:110]}")
    except Exception as e:
        print("  (xato, e'tiborsiz):", e, file=sys.stderr)


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, viewport={"width": 1400, "height": 900}
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        ctx.on("response", on_response)

        page.goto(START_URL, wait_until="domcontentloaded")
        print("\nBrauzer ochildi. Testni qo'lda boshlang (login -> fan -> savollar).")
        print(f"So'rovlar {OUT_FILE} ga yozilmoqda. Tugatish: Ctrl+C\n")

        try:
            while True:
                page.wait_for_timeout(1000)
        except KeyboardInterrupt:
            pass
        finally:
            save()
            print(f"\nTo'xtatildi. {len(log)} so'rov -> {OUT_FILE}")
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    signal.signal(signal.SIGINT, signal.default_int_handler)
    main()
