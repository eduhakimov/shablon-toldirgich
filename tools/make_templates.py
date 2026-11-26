"""Dastur bilan birga keladigan shablonlarni yaratadi: assets/templates/*.docx va index.json.

    python tools/make_templates.py
"""
import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT = Path(__file__).resolve().parents[1] / "assets" / "templates"
FONT = "Times New Roman"


def new_doc():
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(3), Cm(1.5)
    sec.top_margin, sec.bottom_margin = Cm(2), Cm(2)
    st = d.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = st.paragraph_format
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.15
    return d


def para(d, text="", align=None, bold=False, indent_cm=None, first_line=None, size=None, space_after=0,
         italic=False, caps=False):
    p = d.add_paragraph()
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "j":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent_cm is not None:
        p.paragraph_format.left_indent = Cm(indent_cm)
    if first_line is not None:
        p.paragraph_format.first_line_indent = Cm(first_line)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.all_caps = caps
        if size:
            r.font.size = Pt(size)
    return p


def addressee(d, sender_lines):
    """O'ng tomondagi "kimga / kimdan" bloki."""
    ind = 8.5
    para(d, "{{TASHKILOT}}", indent_cm=ind)
    para(d, "{{RAHBAR_LAVOZIM}}", indent_cm=ind)
    para(d, "{{RAHBAR_FISH}}ga", indent_cm=ind, bold=True, space_after=6)
    for line in sender_lines:
        para(d, line, indent_cm=ind)


def title(d, text):
    para(d, "")
    para(d, text, align="c", bold=True, space_after=12)


def body(d, text):
    return para(d, text, align="j", first_line=1.25, space_after=6)


def signature(d, with_date=True):
    para(d, "")
    para(d, "")
    p = d.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
    if with_date:
        p.add_run("{{SANA}}")
    p.add_run("\t_____________  ")
    r = p.add_run("{{QISQA_FISH}}")
    r.bold = True


def save(d, name):
    OUT.mkdir(parents=True, exist_ok=True)
    d.core_properties.title = name
    d.core_properties.author = ""
    d.save(OUT / name)


SENDER = ["{{BOLIM}} {{LAVOZIM}}", "{{FISH}}dan", "Manzil: {{MANZIL}}", "Tel.: {{TELEFON}}"]


def t_ishga_qabul():
    d = new_doc()
    addressee(d, ["{{FISH}}dan", "Manzil: {{MANZIL}}", "Tel.: {{TELEFON}}"])
    title(d, "A R I Z A")
    body(d, "Meni {{ISHGA_KIRISH_SANASI}}dan e'tiboran {{BOLIM}}ga {{LAVOZIM}} lavozimiga ishga qabul "
            "qilishingizni so'rayman.")
    body(d, "Ilova: shaxsiy varaqa (obyektivka), pasport nusxasi, ma'lumoti to'g'risidagi hujjat nusxasi, "
            "mehnat daftarchasi (elektron mehnat daftarchasidan ko'chirma).")
    signature(d)
    save(d, "ishga_qabul.docx")


def t_tatil():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "A R I Z A")
    body(d, "Menga {{ISH_YILI}} ish yili uchun {{TATIL_BOSHLANISHI}}dan {{TATIL_TUGASHI}}gacha "
            "{{TATIL_KUNLARI}} kalendar kunlik navbatdagi mehnat ta'tili berishingizni so'rayman.")
    body(d, "Ta'til davrida mening vazifalarimni vaqtincha {{VAZIFANI_BAJARUVCHI}} bajarib turadi.")
    signature(d)
    save(d, "tatil.docx")


def t_ishdan_boshash():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "A R I Z A")
    body(d, "Meni {{BOSHASH_SANASI}}dan e'tiboran o'z xohishimga ko'ra egallab turgan {{LAVOZIM}} lavozimimdan "
            "ozod qilishingizni so'rayman.")
    body(d, "Barcha hujjatlar va moddiy boyliklarni belgilangan tartibda topshirishga tayyorman.")
    signature(d)
    save(d, "ishdan_boshash.docx")


def t_malumotnoma():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "A R I Z A")
    body(d, "{{MALUMOTNOMA_QAYERGA}}ga taqdim etish uchun menga {{TASHKILOT}}da {{LAVOZIM}} lavozimida "
            "ishlayotganligim to'g'risida ma'lumotnoma berishingizni so'rayman.")
    signature(d)
    save(d, "malumotnoma_sorash.docx")


def t_tushuntirish():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "TUSHUNTIRISH XATI")
    body(d, "Men, {{FISH}}, {{VOQEA_SANASI}} kuni {{VOQEA_MAZMUNI}} yuzasidan quyidagilarni ma'lum qilaman.")
    body(d, "{{TUSHUNTIRISH}}")
    body(d, "Kelgusida bunday holatga yo'l qo'ymaslikka harakat qilaman.")
    signature(d)
    save(d, "tushuntirish_xati.docx")


def t_bildirishnoma():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "BILDIRISHNOMA")
    body(d, "Sizga shuni ma'lum qilamanki, {{BILDIRISHNOMA_MATNI}}")
    body(d, "Yuqoridagilarni inobatga olib, {{SORALADIGAN_CHORA}} bo'yicha tegishli ko'rsatma berishingizni "
            "so'rayman.")
    signature(d)
    save(d, "bildirishnoma.docx")


def t_ishonchnoma():
    d = new_doc()
    para(d, "I S H O N C H N O M A", align="c", bold=True, space_after=6)
    p = d.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run("{{SHAHAR}} shahri\t{{SANA_SOZ}}")
    p.paragraph_format.space_after = Pt(12)
    body(d, "Men, {{FISH}}, {{TUGILGAN_SANA}} yilda tug'ilganman, pasport (ID-karta) {{PASPORT}}, "
            "{{PASPORT_BERILGAN}}, JSHSHIR {{JSHSHIR}}, {{MANZIL}} manzilida yashayman.")
    body(d, "Ushbu ishonchnoma bilan {{ISHONCHLI_SHAXS_FISH}}ga (pasport (ID-karta) {{ISHONCHLI_SHAXS_PASPORT}}) "
            "{{ISHONCHNOMA_MAQSADI}} bo'yicha mening nomimdan ish yuritish, zarur hujjatlarni topshirish va "
            "olish, ularga imzo qo'yish vakolatini beraman.")
    body(d, "Ishonchnoma {{AMAL_QILISH_MUDDATI}}gacha amal qiladi. Vakolatlar boshqa shaxsga o'tkazilmaydi.")
    signature(d, with_date=False)
    save(d, "ishonchnoma.docx")


def _table(d, rows, widths=(6.0, 10.5), bold_left=True):
    t = d.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for a, b in rows:
        cells = t.add_row().cells
        cells[0].width, cells[1].width = Cm(widths[0]), Cm(widths[1])
        r = cells[0].paragraphs[0].add_run(a)
        r.bold = bold_left
        cells[1].paragraphs[0].add_run(b)
    return t


def t_rezyume():
    d = new_doc()
    para(d, "R E Z Y U M E", align="c", bold=True, space_after=4)
    para(d, "{{FISH}}", align="c", bold=True, size=16, space_after=12)
    _table(d, [
        ("Tug'ilgan sana", "{{TUGILGAN_SANA}}"),
        ("Manzil", "{{MANZIL}}"),
        ("Telefon", "{{TELEFON}}"),
        ("E-pochta", "{{EMAIL}}"),
        ("Ma'lumoti", "{{MALUMOTI}}, {{TAMOMLAGAN}}"),
        ("Mutaxassisligi", "{{MUTAXASSISLIGI}}"),
        ("Ilmiy darajasi", "{{ILMIY_DARAJA}}"),
        ("Hozirgi ish joyi", "{{TASHKILOT}}, {{BOLIM}}"),
        ("Lavozimi", "{{LAVOZIM}}"),
        ("Chet tillari", "{{TILLAR}}"),
    ])
    para(d, "")
    para(d, "Ish tajribasi", bold=True, space_after=4)
    body(d, "{{ISH_TAJRIBASI}}")
    para(d, "Kasbiy ko'nikmalar", bold=True, space_after=4)
    body(d, "{{KONIKMALAR}}")
    para(d, "Qo'shimcha ma'lumot", bold=True, space_after=4)
    body(d, "{{QOSHIMCHA}}")
    save(d, "rezyume.docx")


def t_obyektivka():
    d = new_doc()
    para(d, "MA'LUMOTNOMA", align="c", bold=True, space_after=6)
    para(d, "{{FISH}}", align="c", bold=True, space_after=10)
    p = para(d, "{{ISHGA_KIRGAN_YIL}} yildan: {{TASHKILOT}} {{BOLIM}} {{LAVOZIM}}", align="j", space_after=10)
    t = d.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    pairs = [
        (("Tug'ilgan yili:", "{{TUGILGAN_SANA}}"), ("Tug'ilgan joyi:", "{{TUGILGAN_JOY}}")),
        (("Millati:", "{{MILLATI}}"), ("Partiyaviyligi:", "{{PARTIYAVIYLIGI}}")),
        (("Ma'lumoti:", "{{MALUMOTI}}"), ("Tamomlagan:", "{{TAMOMLAGAN}}")),
        (("Ma'lumoti bo'yicha mutaxassisligi:", "{{MUTAXASSISLIGI}}"), None),
        (("Ilmiy darajasi:", "{{ILMIY_DARAJA}}"), ("Ilmiy unvoni:", "{{ILMIY_UNVONI}}")),
        (("Qaysi chet tillarini biladi:", "{{TILLAR}}"), None),
        (("Davlat mukofotlari bilan taqdirlanganmi (qanaqa):", "{{MUKOFOTLAR}}"), None),
        (("Xalq deputatlari kengashi deputatimi yoki boshqa saylanadigan organlar a'zosimi:",
          "{{DEPUTATLIK}}"), None),
    ]
    for left, right in pairs:
        cells = t.add_row().cells
        for cell, item in ((cells[0], left), (cells[1], right)):
            if not item:
                continue
            r = cell.paragraphs[0].add_run(item[0])
            r.bold = True
            cell.add_paragraph(item[1])
        if right is None:
            cells[0].merge(cells[1])
    para(d, "")
    para(d, "MEHNAT FAOLIYATI", align="c", bold=True, space_after=6)
    body(d, "{{MEHNAT_FAOLIYATI}}")
    save(d, "obyektivka.docx")


def t_umumiy():
    d = new_doc()
    addressee(d, SENDER)
    title(d, "A R I Z A")
    body(d, "{{ARIZA_MATNI}}")
    signature(d)
    save(d, "ariza_umumiy.docx")


def t_talaba():
    d = new_doc()
    ind = 8.5
    para(d, "{{TASHKILOT}}", indent_cm=ind)
    para(d, "{{RAHBAR_LAVOZIM}}", indent_cm=ind)
    para(d, "{{RAHBAR_FISH}}ga", indent_cm=ind, bold=True, space_after=6)
    para(d, "{{GURUH}} guruh talabasi", indent_cm=ind)
    para(d, "{{FISH}}dan", indent_cm=ind)
    para(d, "Tel.: {{TELEFON}}", indent_cm=ind)
    title(d, "A R I Z A")
    body(d, "{{ARIZA_MATNI}}")
    signature(d)
    save(d, "talaba_arizasi.docx")


INDEX = [
    ("ishga_qabul.docx", "ariza", ("Ishga qabul qilish to'g'risida ariza", "Заявление о приёме на работу",
                                   "Application for employment")),
    ("tatil.docx", "ariza", ("Mehnat ta'tili berish to'g'risida ariza", "Заявление на отпуск",
                             "Leave application")),
    ("ishdan_boshash.docx", "ariza", ("Ishdan bo'shash to'g'risida ariza", "Заявление об увольнении",
                                      "Resignation letter")),
    ("malumotnoma_sorash.docx", "ariza", ("Ish joyidan ma'lumotnoma so'rash", "Запрос справки с места работы",
                                          "Request for an employment certificate")),
    ("ariza_umumiy.docx", "ariza", ("Ariza (umumiy shakl)", "Заявление (общая форма)", "Application (general)")),
    ("talaba_arizasi.docx", "ariza", ("Talaba arizasi", "Заявление студента", "Student application")),
    ("tushuntirish_xati.docx", "xat", ("Tushuntirish xati", "Объяснительная записка", "Explanatory note")),
    ("bildirishnoma.docx", "xat", ("Bildirishnoma", "Докладная записка", "Memo to management")),
    ("ishonchnoma.docx", "boshqa", ("Ishonchnoma", "Доверенность", "Power of attorney")),
    ("rezyume.docx", "shaxsiy", ("Rezyume", "Резюме", "CV / resume")),
    ("obyektivka.docx", "shaxsiy", ("Ma'lumotnoma (obyektivka)", "Объективка (личная справка)",
                                    "Personal record (obyektivka)")),
]


def main():
    for fn in (t_ishga_qabul, t_tatil, t_ishdan_boshash, t_malumotnoma, t_tushuntirish, t_bildirishnoma,
               t_ishonchnoma, t_rezyume, t_obyektivka, t_umumiy, t_talaba):
        fn()
    idx = [{"file": f, "category": c, "title": {"uz": t[0], "ru": t[1], "en": t[2]}} for f, c, t in INDEX]
    (OUT / "index.json").write_text(json.dumps(idx, ensure_ascii=False, indent=1), encoding="utf-8")
    print("OK", len(idx))


if __name__ == "__main__":
    main()
