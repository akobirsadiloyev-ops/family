# Bitiruvchilar bandligi — dashboard

Fayllar:
- `index.html` — sahifaning o'zi (qora fon, sidebar bilan)
- `data.xlsx` — ma'lumotlar (sahifa shu faylni o'qiydi)
- `tozalash.py` — Excel'dan shaxsiy ma'lumotlarni olib tashlaydi

## Saytga joylash (bkm-buxdu.uz)
`index.html` va `data.xlsx` ni serverda **bitta papkaga** yuklang, masalan `/bandlik/`.
Sahifa manzili: `https://bkm-buxdu.uz/bandlik/`

Katta sayt sahifasi ichiga qo'yish uchun:
```html
<iframe src="/bandlik/index.html" style="width:100%;height:100vh;border:0"></iframe>
```

## Ma'lumotni yangilash
1. Yangi "Bandlik shtabi indeks" Excel faylini tozalang:
   ```
   python3 tozalash.py Bandlik_shtabi_indeks_YANGI.xlsx data.xlsx
   ```
2. Serverdagi `data.xlsx` ni shu fayl bilan almashtiring. Sahifa o'zi yangilanadi.

## Sahifalar (sidebar)
Excel'dagi **har bir ko'rinadigan varaq — alohida sahifa**. Chap tomondagi sidebar'dan tanlanadi,
har bir sahifaning o'z manzili bor (masalan `index.html#KUNLIK`).

Har bir sahifada:
- tepada — shu varaq bo'yicha grafiklar va asosiy raqamlar;
- pastda — varaqning o'zi Excel'dagidek jadval ko'rinishida, **Excel ranglari bilan**:
  katak ranglari, qalin shrift, shartli formatlash (rang shkalasi, data bar, qoida bo'yicha
  bo'yash — masalan ▼/▲ o'zgarish, oylik chegaralari). Yashirin ustun/qatorlar ko'rsatilmaydi,
  birlashtirilgan kataklar va son formatlari saqlanadi, qidiruv bor.

"Kafedralar" va "Kafedralar indeksi" sahifalarida jadvaldan pastdagi yozuvlar (Tekshiruv,
Asos, Eslatma, ko'chirish jadvallari) ko'rsatilmaydi — faqat asosiy jadval va "Jami" qatori.
Boshqa varaqlarga ham qo'llash uchun `index.html` ichidagi `CUT_BELOW` ro'yxatiga varaq nomini qo'shing.

"BANDLIK INDEKSI1" Excel'da yashirin bo'lsa ham bosh sahifa sifatida chiqadi.
Boshqa yashirin varaqlar (JAMI, ISHSIZLAR va h.k.) ko'rsatilmaydi — kerak bo'lsa
`index.html` ichida `SHOW_HIDDEN_SHEETS = true` qiling.

"O'zini band (Buxoro)" va "Ishsiz+band yo'nalishlar" sahifalarida viloyat filtri ishlaydi
(Nomma-nom varag'idan hisoblanadi).

Excel'ga yangi varaq qo'shilsa, u ham avtomatik yangi sahifa bo'lib chiqadi.

## MUHIM: shaxsiy ma'lumotlar
Serverdagi `data.xlsx` ni har kim yuklab olishi mumkin. Asl faylda (Nomma-nom varag'ida)
ism-familiya va ID bor. Shuning uchun har doim avval `tozalash.py` dan o'tkazing.
