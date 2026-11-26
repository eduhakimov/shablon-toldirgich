# Template Filler: user guide

## 1. Create a profile
In **Profiles** press **Add** and enter the details: name, date of birth, address, phone, passport, PINFL,
workplace, position and so on. Create separate profiles for family members. Press **Save**.

Passport and PINFL are hidden. Turn on **Show hidden data** to see them.

**Password protection** encrypts profiles (AES). Next time the Profiles section asks for the password.
A forgotten password cannot be recovered.

## 2. New document
1. In **New document** pick a template on the left (or **Open another .docx**).
2. Choose a profile in **For**. Matching fields are filled in automatically.
3. Fill in the empty (highlighted) rows, e.g. leave dates.
4. Check the result on the **Preview** tab.
5. Set the folder and file name; turn on **Also create PDF** if needed.
6. Press **Create document**.

In the file name, `{QISQA_FISH}`, `{FISH}`, `{SANA}` are replaced with values. The template is never modified.

## 3. Your own template
Write a document in Word and put `{{NAME}}` where text changes:

| Placeholder | Filled with |
|---|---|
| `{{FISH}}` | Last First Patronymic |
| `{{QISQA_FISH}}` | A.R. Karimov |
| `{{SANA}}` | today's date, 27.11.2026 |
| `{{SANA_SOZ}}` | 2026-yil 27-noyabr |
| `{{MANZIL}}`, `{{TELEFON}}`, `{{PASPORT}}`, `{{JSHSHIR}}` | from the profile |
| `{{TASHKILOT}}`, `{{BOLIM}}`, `{{LAVOZIM}}`, `{{GURUH}}` | from the profile |
| `{{RAHBAR_LAVOZIM}}`, `{{RAHBAR_FISH}}` | addressee |

Names not in the list (e.g. `{{KURS}}`) are asked for when filling in and remembered. Formatting
(fonts, bold, tables, headers and footers) is kept. Add the file to the library in **Templates**.

## 4. Batch from Excel
1. Put column names in the first row: Familiya, Ism, Otasining ismi, Guruh...
2. In **Batch from Excel** choose the template and the file.
3. Check the **Mapping**. Put values shared by all rows into **Fixed value**.
4. Press **Create documents**: one .docx (and PDF) per row.

## PDF
PDF is made with Microsoft Word. Without Word, LibreOffice is used if installed.
