# Bitiruvchilar bandligi — dashboard

Fayllar:
- `index.html` — sahifaning o'zi
- `data.xlsx` — ma'lumotlar (sahifa shu faylni o'qiydi)
- `tozalash.py` — Excel'dan shaxsiy ma'lumotlarni olib tashlaydi

## Saytga joylash (bkm-buxdu.uz)
`index.html` va `data.xlsx` ni serverda **bitta papkaga** yuklang, masalan `/bandlik/`.
Sahifa manzili: `https://bkm-buxdu.uz/bandlik/`

Katta sayt sahifasi ichiga qo'yish uchun:
```html
<iframe src="/bandlik/index.html" style="width:100%;height:2600px;border:0"></iframe>
```

## Ma'lumotni yangilash
Serverdagi `data.xlsx` ni yangi Excel bilan almashtiring (nomi aynan `data.xlsx` bo'lsin).
Sahifa avtomatik yangi ma'lumotni ko'rsatadi.

Excel'da kamida `Kategoriya` ustuni bo'lishi kerak; qo'shimcha ravishda:
Ta'lim turi, Ta'lim shakli, To'lov turi, Jinsi, Mutaxassislik kodi, Mutaxassislik, Viloyat,
Diplom kategoriyasi, OTM nomi. `Qo'shimcha kategoriyalar` ustuni bor varaq — izohlar bloki uchun.

## MUHIM: shaxsiy ma'lumotlar
`data.xlsx` serverda ochiq turadi — uni har kim yuklab olishi mumkin.
Shuning uchun JSHSHIR, telefon, F.I.Sh., manzil, diplom raqamini **yuklashdan oldin olib tashlang**:
```
python3 tozalash.py asl_fayl.xlsx data.xlsx
```

## Bandlik hisobi
Bandlik darajasi = Band / (Jami − Uzrli).
- Band: yuridik shaxsda bandlar, o'zini band qilish, shartnoma asosida, YATT, davlat xizmatchisi
- Uzrli: 3 yoshgacha bola parvarishi, o'qishni davom ettirish, xorijga ketgan, nogiron, harbiy xizmat
- Band emas: kategoriyasi bo'sh
Guruhlarni `index.html` ichidagi `BAND_KEYS` / `UZR_KEYS` da o'zgartirish mumkin.
