# Bitiruvchilar bandligi — dashboard

Fayllar:
- `index.html` — sahifaning o'zi (qora fon)
- `data.xlsx` — ma'lumotlar (sahifa shu faylni o'qiydi)
- `tozalash.py` — Excel'dan shaxsiy ma'lumotlarni olib tashlaydi

## Saytga joylash (bkm-buxdu.uz)
`index.html` va `data.xlsx` ni serverda **bitta papkaga** yuklang, masalan `/bandlik/`.
Sahifa manzili: `https://bkm-buxdu.uz/bandlik/`

Katta sayt sahifasi ichiga qo'yish uchun:
```html
<iframe src="/bandlik/index.html" style="width:100%;height:7200px;border:0"></iframe>
```

## Ma'lumotni yangilash
1. Yangi "Bandlik shtabi indeks" Excel faylini tozalang:
   ```
   python3 tozalash.py Bandlik_shtabi_indeks_YANGI.xlsx data.xlsx
   ```
2. Serverdagi `data.xlsx` ni shu fayl bilan almashtiring. Sahifa o'zi yangilanadi.

Sahifa varaqlarni nomi va ustun sarlavhalari bo'yicha topadi, shuning uchun
varaq nomlari va sarlavhalarni o'zgartirmang:

| Varaq | Sahifadagi bo'lim |
|---|---|
| BANDLIK INDEKSI1 | Yuqoridagi raqamlar, indeks (30 balldan), yo'nalishlar jadvali, sana |
| Fakultetlar indeksi | Fakultetlar reytingi |
| Kafedralar indeksi | Kafedralar reytingi |
| KUNLIK, KUNLIK B | Kunlik monitoring (ishsiz / o'zini band / jami) |
| Viloyatlar (2) | Viloyatlar jadvali |
| Viloyat tumanlari | Buxoro tumanlari grafigi va jadvali |
| Nomma-nom | Faoliyat turlari, ishsizlar holati (ism-familiyasiz) |

Biror varaq bo'lmasa, sahifadagi o'sha bo'lim shunchaki ko'rinmaydi.

## MUHIM: shaxsiy ma'lumotlar
Serverdagi `data.xlsx` ni har kim yuklab olishi mumkin. Asl faylda (Nomma-nom varag'ida)
ism-familiya va ID bor. Shuning uchun har doim avval `tozalash.py` dan o'tkazing.
