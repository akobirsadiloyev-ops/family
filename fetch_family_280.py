"""
owners_280.json dagi 280 uy egasining oila a'zolarini ins.ihma.uz dan oladi.
Mavjud fetch_family.py mantig'ini qayta ishlatadi. Natija -> family_280_results.json (JSON).

Ishlatish:
    python fetch_family_280.py
Brauzer ochiladi -> ins.ihma.uz ga login qiling -> BITTA odamni qo'lda qidiring
(shablon ushlanadi) -> qolgani avtomatik. To'xtasa qayta ishga tushiring (resumable).
"""
import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

import fetch_family as ff

OWNERS = "owners_280.json"
RESULTS = "family_280_results.json"


def main():
    owners = json.load(open(OWNERS, encoding="utf-8"))
    results = json.load(open(RESULTS, encoding="utf-8")) if os.path.exists(RESULTS) else {}
    todo = [o for o in owners if results.get(o["pinfl"], {}).get("status") != "ok"]
    print(f"280 uy egasi. Tayyor: {len(owners)-len(todo)}, qolgan: {len(todo)}")
    if not todo:
        print("Hammasi tayyor.")
        return

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            ff.PROFILE_DIR, headless=False, viewport=None, args=["--start-maximized"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(ff.PAGE_URL)
        template, headers = ff.wait_for_template(page)
        print(f"Shablon ushlandi: {json.dumps(template, ensure_ascii=False)}")

        try:
            for i, o in enumerate(todo, 1):
                person = {"num": o["id"], "name": o["name"],
                          "pinfl": o["pinfl"], "dob": o["dob"]}
                if not o["dob"]:
                    results[o["pinfl"]] = {"status": "error", "error": "Tug'ilgan sana yo'q"}
                else:
                    try:
                        results[o["pinfl"]] = ff.fetch_family(page, template, headers, person)
                    except Exception as e:
                        results[o["pinfl"]] = {"status": "error", "error": str(e)[:200]}
                json.dump(results, open(RESULTS, "w", encoding="utf-8"),
                          ensure_ascii=False, indent=2)
                res = results[o["pinfl"]]
                info = (f"{len(res['members'])} a'zo" if res["status"] == "ok"
                        else f"XATO: {res['error']}")
                print(f"[{i}/{len(todo)}] {o['pinfl']} {o['name']} -> {info}")
                time.sleep(ff.DELAY)
        except KeyboardInterrupt:
            print("\nTo'xtatildi. Qayta ishga tushirsangiz qolganidan davom etadi.")
        finally:
            ctx.close()

    ok = sum(1 for v in results.values() if v.get("status") == "ok")
    members = sum(len(v.get("members", [])) for v in results.values() if v.get("status") == "ok")
    print(f"\nTayyor: {ok} uyda oila topildi, jami {members} a'zo -> {RESULTS}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
