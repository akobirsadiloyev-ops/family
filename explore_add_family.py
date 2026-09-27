"""
online-mahalla.uz — "Xonadon a'zosi qo'shish" jarayonini O'RGANISH (capture) skripti.

Nima qiladi:
  1. Brauzer ochiladi (sessiya .browser_profile papkasida saqlanadi).
  2. Siz online-mahalla.uz ga LOGIN qilasiz.
  3. Misol uy ochiladi (ID 1445733) -> "Xonadon a'zolari" tab -> "+ qo'shish".
  4. Modalda BITTA a'zoni qo'lda qo'shasiz:
       - Hujjat turi: ID karta
       - "JSHSHIR bilan qidirish" ni yoqing
       - Tug'ilgan sana + JSHSHIR kiriting, Qarindoshligi tanlang
       - "Qidirish" -> ma'lumot chiqadi -> "Saqlash"
  5. Skript shu paytdagi BARCHA tarmoq so'rovlarini (qidirish + saqlash payload'lari)
     va modaldagi dropdown (Qarindoshligi / Ma'lumoti / Hujjat turi) qiymatlarini yozib oladi.
  6. Tugagach shu terminalda Ctrl+C bosing -> hammasi faylga saqlanadi.

Natija:
  explore_capture.json  — tarmoq so'rovlari (method, url, post_data, javob)
  explore_selects.json  — modaldagi <select> variantlari (value + matn)

Ishlatish:
    python explore_add_family.py
"""

import json
import sys

from playwright.sync_api import sync_playwright

PROFILE_DIR = ".browser_profile"
BASE = "https://online-mahalla.uz"
EXAMPLE_URL = (
    BASE + "/forms/survey_homes/1445733"
    "?obl_id=6&area_id=608&district_id=608035&street_id=60800147"
)
CAP_FILE = "explore_capture.json"
SEL_FILE = "explore_selects.json"

# Faqat shu turdagi so'rovlar qiziq (statik fayllar emas)
WANTED_TYPES = {"xhr", "fetch", "document"}
KEEP_REQ_HEADERS = {"content-type", "accept", "x-requested-with", "authorization"}

SELECT_JS = r"""
() => {
  const out = [];
  document.querySelectorAll('select').forEach(s => {
    const opts = [...s.options].map(o => ({value: o.value, text: (o.textContent||'').trim()}));
    if (opts.length) out.push({
      name: s.name || '', id: s.id || '',
      className: s.className || '',
      selected: s.value,
      options: opts,
    });
  });
  return out;
}
"""

captured = []
best_selects = []


def on_response(resp):
    try:
        req = resp.request
        url = resp.url
        if "online-mahalla.uz" not in url:
            return
        if req.resource_type not in WANTED_TYPES:
            return
        entry = {
            "method": req.method,
            "url": url,
            "resource_type": req.resource_type,
            "status": resp.status,
            "req_headers": {
                k: v for k, v in req.headers.items() if k.lower() in KEEP_REQ_HEADERS
            },
            "post_data": req.post_data,
        }
        ct = (resp.headers or {}).get("content-type", "")
        if any(t in ct for t in ("json", "text", "html", "xml")):
            try:
                entry["response_body"] = resp.text()[:30000]
            except Exception as e:  # ba'zi javoblarni o'qib bo'lmaydi — muhim emas
                entry["response_body"] = f"<o'qib bo'lmadi: {e}>"
        captured.append(entry)
        # jonli log — muhimlarini ajratib ko'rsatamiz
        mark = ""
        low = url.lower()
        if req.method == "POST":
            mark = "  <== POST"
        if any(w in low for w in ("family", "person", "search", "save", "store", "qidir")):
            mark += "  *"
        print(f"[{req.method}] {resp.status}  {url[:130]}{mark}")
    except Exception:
        pass


def dump(page):
    global best_selects
    try:
        sel = page.evaluate(SELECT_JS)
        if sel:
            best_selects = sel
    except Exception:
        pass


def main():
    global best_selects
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, viewport=None, args=["--start-maximized"]
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("response", on_response)

        print("\n>>> Brauzer ochildi. online-mahalla.uz ga LOGIN qiling.")
        print(">>> Misol uy sahifasi ochilyapti (ID 1445733)...")
        try:
            page.goto(EXAMPLE_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print(f"    (sahifa ochilishida ogohlantirish: {e})")

        print("\n>>> Endi qo'lda BITTA a'zo qo'shing:")
        print(">>>   'Xonadon a'zolari' tab -> '+ qo'shish' -> modalni to'ldiring -> Qidirish -> Saqlash")
        print(">>> Modal ochilganda dropdown variantlari avtomatik yozib olinadi.")
        print(">>> TUGAGACH shu oynada Ctrl+C bosing.\n")

        try:
            while True:
                page.wait_for_timeout(700)
                dump(page)  # modal ochiq bo'lsa, select variantlarini ushlaymiz
        except KeyboardInterrupt:
            print("\n>>> To'xtatildi, saqlanmoqda...")
        finally:
            dump(page)
            with open(CAP_FILE, "w", encoding="utf-8") as f:
                json.dump(captured, f, ensure_ascii=False, indent=2)
            with open(SEL_FILE, "w", encoding="utf-8") as f:
                json.dump(best_selects, f, ensure_ascii=False, indent=2)
            print(f">>> {len(captured)} ta so'rov -> {CAP_FILE}")
            print(f">>> {len(best_selects)} ta <select> -> {SEL_FILE}")
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
