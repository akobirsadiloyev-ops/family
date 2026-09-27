# HOLAT — davom ettirish uchun eslatma

Yangi sessiyada (boshqa akaunt bo'lsa ham) shu faylni o'qib, ishni davom ettiring.

## Maqsad
online-mahalla.uz (survey_homes) dagi 670 xonadonning har biriga oila a'zolarini
(qarindoshlarini) topib qo'shish. Manba: ins.ihma.uz (oila a'zolari) + GCP (shaxs ma'lumoti).

## Hozirgi holat (2026-09-27)
- **619 oila a'zosi qo'shildi**, **289 / 670** xonadonga.
- Qo'shilmaganlar: 202 band (boshqa xonadonda), 110 GCP topilmadi, 16 dublikat.
- **281 xonadon** hali qarindoshi izlanmagan (670 − 389 ishlangan). Bular keyingi bosqich.

## Fayllar (FAQAT yakuniy_natija.xlsx bilan ishlash — boshqa Excel ochmaslik!)
- `yakuniy_natija.xlsx` — ASOSIY natija. Varaqlar: "Oila a'zolari" (978, egasi har uyda bir marta),
  "Xonadonlar (670)", va qo'shilishi kerak "Qolgan uylar (qarindosh kerak)" (281).
- `build_yakuniy.py` — yakuniy_natija.xlsx ni qayta yasaydi (progress + base_old + homes.json dan).
- `base_old.xlsx` — 978 a'zoning to'liq ro'yxati (dastlabki).
- `olad_base/homes.json` — 670 uy egasi (id, pinfl, telefon, ...).
- `add_family.py` — 585 oila a'zolarini ("Oila a'zolari (taxminiy)" varag'idan) online-mahalla.uz ga
  qo'shadi. Avval `python add_family.py tekshir` (kim kiritilgan / kim yo'q, hech narsa qo'shmaydi),
  keyin `python add_family.py all`. Natija va qo'lda to'ldiriladigan ustunlar (Ma'lumoti,
  Qarindoshlik (tasdiqlangan)) — `yakuniy_natija.xlsx` -> "Kiritish" varag'i. Kattalarning ta'limi
  o'ylab topilmaydi: kiritilmaguncha a'zo qo'shilmaydi. Resumable, `add_progress.json`.
- `hisobot.py` — sonli hisobot (shaxsiy ma'lumotsiz); `yakuniy_varaq.py` — "Yakuniy" va
  "Yakuniy xulosa" varaqlari.
- `add_progress_run1.json` + `add_progress.json` — qo'shish natijalari (run-1 + run-2).
- `fetch_family.py` — ins.ihma.uz dan oila a'zolarini oladi (Playwright, `.browser_profile`).
- `guess_relations.py` — qarindoshlikni taxmin qiladi.
- `get_token.py` / `token.txt` — online-mahalla Bearer token (401 bo'lsa yangilaydi).

## KUTILAYOTGAN ish (shu yerdan davom)
1. `yakuniy_natija.xlsx` yopilgach: `python build_yakuniy.py` — "Qolgan uylar (281)" varag'ini qo'shadi.
2. Keyin **281 uyning qarindoshini ins.ihma.uz dan topish**:
   - `fetch_family.py` ni shu 281 uy egasiga moslab ishga tushirish (Playwright, ins.ihma login + 1 ta qo'lda qidiruv shabloni).
   - So'ng `guess_relations.py` bilan qarindoshlik taxmini.
   - Topilgan a'zolarni `add_family.py` orqali online-mahalla.uz ga qo'shish.
   - Natijani yangi Excel ochmasdan `yakuniy_natija.xlsx` ga sheet qilib yozish.
3. Har bir a'zoga aniq izoh: ko'rildi / qo'shildi / bo'lmadi (sabab: band / GCP topilmadi / ...).

## Muhim eslatmalar
- O'zbekcha gaplashish.
- Xatolarni skip qilib, izohini yozish (dastur to'xtamasin).
- Saqlash so'rovi ~20-25s (server gov-registrni tekshiradi) — batch bir necha soat.
- add_family telefonni bo'sh qoldiradi (uy egasi raqami qayta ishlatilsa limitga uriladi; telefon majburiy emas).
