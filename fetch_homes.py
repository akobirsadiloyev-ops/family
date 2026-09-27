"""
online-mahalla.uz — survey_homes_street ma'lumotini yig'ib, JSON qilib saqlaydi.

Ishlatish:
    python fetch_homes.py

Token eskirsa (401/403 chiqsa): brauzerda qayta Copy as cURL qilib,
pastdagi TOKEN ni yangilang.
"""

import json
import math
import sys
import urllib.request
import urllib.error

# --- Sozlamalar (kerak bo'lsa shularni o'zgartiring) ---------------------
TOKEN = "8c190733-0a86-41b0-a4a4-724de842db23"

BASE_URL = "https://api.online-mahalla.uz/api/v1/survey_homes_street/cache/data"
PARAMS = {
    "obl_id": 6,
    "area_id": 608,
    "district_id": 608035,
    "street_id": 60800147,
    "surveyed": 0,     # 0 = hali so'rovdan o'tmagan uylar
    "size": 100,       # bir sahifada nechta
}

OUT_FILE = "homes.json"
# -------------------------------------------------------------------------

HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "uz",
    "Authorization": f"Bearer {TOKEN}",
    "Origin": "https://online-mahalla.uz",
    "Referer": "https://online-mahalla.uz/",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
    ),
    "ngrok-skip-browser-warning": "1",
}


def fetch_page(page):
    query = "&".join(f"{k}={v}" for k, v in {**PARAMS, "page": page}.items())
    url = f"{BASE_URL}?{query}"
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        print(f"  XATO: HTTP {e.code} (page {page}). Javob: {detail}")
        if e.code in (401, 403):
            print("  -> Token eskirgan bo'lishi mumkin. TOKEN ni yangilang.")
        sys.exit(1)
    except urllib.error.URLError as e:
        print(f"  XATO: ulanib bo'lmadi (page {page}): {e.reason}")
        sys.exit(1)
    return json.loads(body)


def main():
    print("1-sahifa olinmoqda...")
    first = fetch_page(1)
    data = first.get("data") or {}
    total = data.get("total", 0)
    results = list(data.get("results") or [])
    size = PARAMS["size"]
    pages = max(1, math.ceil(total / size)) if total else 1
    print(f"Jami uy: {total}. Sahifalar soni: {pages}")

    for page in range(2, pages + 1):
        print(f"{page}-sahifa olinmoqda...")
        payload = fetch_page(page)
        results.extend((payload.get("data") or {}).get("results") or [])

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\nTayyor: {len(results)} ta yozuv -> {OUT_FILE}")
    if results:
        print("Ustunlar (birinchi yozuv kalitlari):")
        for k in results[0].keys():
            print(f"  - {k}")


if __name__ == "__main__":
    main()
