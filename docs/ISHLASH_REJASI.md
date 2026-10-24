# HISOB WEB — ish tartibi

## V1.0: faqat sayt interfeysi
- `index.html` — mustaqil, tashqi kutubxonasiz UI prototip.
- Dashboard: faqat kilogramm, Plan bir qatorda 4 loyiha, to'liq qabul qilingan loyihalar ro'yxati.
- Kunlik kirimlar: sana oralig'i, kg diagramma, metallik kalendar, tanlangan kun tafsilotlari, CSV namuna.
- Boshqa bo'limlar: keyingi bosqichda mavjud kod asosida to'ldiriladi.
- Ekrandagi ko'rsatkichlar TEST ma'lumotlari. Haqiqiy biznes ma'lumotlari emas.

## Xavfsizlik / ma'lumotlar
1. Hozirgi HISOB_BAZA, Apps Script va UMBRELLA botga tegilmaydi.
2. Repo ichiga BOT_TOKEN, parollar, Script Properties, sessiya tokenlari yoki haqiqiy mijoz ma'lumotlari kiritilmaydi.
3. GitHub kodni saqlaydi, Google Sheets esa keyingi bosqichda biznes ma'lumotlari va audit tarixini saqlaydi.
4. Google Sheetsga brauzerdan ochiq kalit bilan ulanmaslik kerak. Alohida autentifikatsiyalangan server/API orqali ulanish rejalashtiriladi.
5. Ko'chirish: avval statik sayt -> bo'limlar UI -> kirish/ruxsatlar -> read-only Google Sheets -> xavfsiz yozish/audit -> sinov -> ishga tushirish.

## Cheklov
Bu versiya ishlab chiqarish uchun tayyor emas: server yo'q, ruxsatlar yo'q, Google Sheets bilan real integratsiya yo'q. Faqat ko'rish va UI sinovi uchun.
