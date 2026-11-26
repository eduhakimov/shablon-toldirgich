## Shablon yozish

Word hujjatida o'zgaradigan joylarga `{{NOM}}` qo'ying: `{{FISH}}`, `{{QISQA_FISH}}`, `{{SANA}}`,
`{{MANZIL}}`, `{{TELEFON}}`, `{{PASPORT}}`, `{{JSHSHIR}}`, `{{TASHKILOT}}`, `{{LAVOZIM}}`, `{{RAHBAR_FISH}}` va h.k.
To'liq ro'yxat dasturning **Shablonlar** bo'limida. Boshqa nomlar to'ldirishda qo'lda so'raladi.

O'rinbosar Word ichida bir nechta bo'lakka (run) bo'lingan bo'lsa ham topiladi; jadvallar, kolontitullar,
izohlar va matn bloklari ham to'ldiriladi. Qiymat o'rinbosar boshlangan joydagi formatni oladi.

## Ma'lumotlar qayerda saqlanadi

`%APPDATA%\BahtiyorSoft\shablon-toldirgich\` papkasida: `profiles.json` (parol qo'yilgan bo'lsa Fernet/AES
bilan shifrlangan, kalit PBKDF2-SHA256 orqali paroldan olinadi) va `templates\` (siz qo'shgan shablonlar nusxasi).
Internetga hech narsa yuborilmaydi.

## Manba koddan ishga tushirish

```
python -m pip install PySide6 python-docx openpyxl cryptography pywin32
python tools/make_templates.py
python main.py
```
