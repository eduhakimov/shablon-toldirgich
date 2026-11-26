# Ariza va shablon to'ldirgich: qo'llanma

## 1. Profil yarating
**Profillar** bo'limida **Qo'shish** ni bosing va ma'lumotlarni kiriting: F.I.Sh., tug'ilgan sana, manzil,
telefon, pasport, JSHSHIR, ish joyi, lavozim va boshqalar. Oila a'zolari uchun alohida profil yarating.
**Saqlash** ni bosing.

Pasport va JSHSHIR yashirin ko'rinadi. Ularni ko'rish uchun **Maxfiy ma'lumotni ko'rsatish** ni yoqing.

**Parol bilan himoya**: profillar shifrlanadi (AES). Dastur keyingi ochilishida Profillar bo'limida parol
so'raladi. Parolni unutsangiz, ma'lumotlarni tiklab bo'lmaydi.

## 2. Hujjat tayyorlash
1. **Hujjat tayyorlash** bo'limida chapdan shablonni tanlang (yoki **Boshqa .docx faylni ochish**).
2. **Kim uchun** ro'yxatidan profilni tanlang. Mos maydonlar avtomatik to'ladi.
3. Bo'sh qolgan qatorlarni (rangli) jadvalda to'ldiring, masalan ta'til sanalari.
4. **Ko'rinish** varag'ida natijani tekshiring.
5. Papka va fayl nomini belgilang, kerak bo'lsa **PDF ham yaratish** ni yoqing.
6. **Hujjatni yaratish** ni bosing.

Fayl nomida `{QISQA_FISH}`, `{FISH}`, `{SANA}` kabi nomlar qiymat bilan almashtiriladi.
Shablon fayli hech qachon o'zgartirilmaydi.

## 3. O'z shabloningiz
Word'da hujjat yozing va o'zgaradigan joylarga `{{NOM}}` qo'ying:

| O'rinbosar | Nima qo'yiladi |
|---|---|
| `{{FISH}}` | Familiya Ism Otasining ismi |
| `{{QISQA_FISH}}` | A.R. Karimov |
| `{{SANA}}` | bugungi sana, 27.11.2026 |
| `{{SANA_SOZ}}` | 2026-yil 27-noyabr |
| `{{MANZIL}}`, `{{TELEFON}}`, `{{PASPORT}}`, `{{JSHSHIR}}` | profildan |
| `{{TASHKILOT}}`, `{{BOLIM}}`, `{{LAVOZIM}}`, `{{GURUH}}` | profildan |
| `{{RAHBAR_LAVOZIM}}`, `{{RAHBAR_FISH}}` | ariza kimga yoziladi |

Ro'yxatda yo'q nom (masalan `{{KURS}}`) to'ldirishda qo'lda so'raladi va keyingi safar eslab qolinadi.
Formatlash (shrift, qalin, jadval, kolontitul) saqlanadi. Faylni **Shablonlar** bo'limida **Qo'shish** orqali
kutubxonaga qo'shing.

## 4. Excel'dan ommaviy
Bitta shablondan ko'p hujjat (masalan, guruh talabalari uchun):
1. Excel faylning birinchi qatoriga ustun nomlarini yozing: Familiya, Ism, Otasining ismi, Guruh...
2. **Excel'dan ommaviy** bo'limida shablon va faylni tanlang.
3. **Moslashtirish** jadvalida ustunlarni tekshiring. Hamma qatorlar uchun bir xil qiymatni
   **Doimiy qiymat** ustuniga yozing.
4. **Hujjatlarni yaratish** ni bosing. Har bir qator uchun alohida .docx (va kerak bo'lsa PDF) yaratiladi.

## PDF
PDF Microsoft Word orqali yaratiladi (Word o'rnatilgan bo'lishi kerak). Word bo'lmasa, LibreOffice ishlatiladi.
