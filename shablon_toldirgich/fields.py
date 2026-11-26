"""Profil maydonlari, hisoblanadigan qiymatlar va o'rinbosar nomlarini moslashtirish."""
import re
from datetime import date
from typing import Dict, List, Optional


def _t(uz, ru, en):
    return {"uz": uz, "ru": ru, "en": en}


# (kalit, yorliq, maxfiy, qo'shimcha nomlar)
FIELDS = [
    ("FAMILIYA", _t("Familiya", "Фамилия", "Last name"), False, ["SURNAME", "LASTNAME", "ФАМИЛИЯ", "FAMILIYASI"]),
    ("ISM", _t("Ism", "Имя", "First name"), False, ["NAME", "FIRSTNAME", "ИМЯ", "ISMI"]),
    ("OTASINING_ISMI", _t("Otasining ismi", "Отчество", "Patronymic"), False,
     ["SHARIF", "OTCHESTVO", "ОТЧЕСТВО", "PATRONYMIC", "OTASININGISMI"]),
    ("TUGILGAN_SANA", _t("Tug'ilgan sana", "Дата рождения", "Date of birth"), False,
     ["TSANA", "BIRTHDATE", "DOB", "ДАТАРОЖДЕНИЯ", "TUGILGANSANASI", "TUGILGANYILI"]),
    ("TUGILGAN_JOY", _t("Tug'ilgan joy", "Место рождения", "Place of birth"), False,
     ["BIRTHPLACE", "МЕСТОРОЖДЕНИЯ", "TUGILGANJOYI"]),
    ("MILLATI", _t("Millati", "Национальность", "Nationality"), False, ["NATIONALITY", "НАЦИОНАЛЬНОСТЬ", "MILLAT"]),
    ("MANZIL", _t("Manzil", "Адрес", "Address"), False, ["ADDRESS", "АДРЕС", "MANZILI", "YASHASHJOYI"]),
    ("TELEFON", _t("Telefon", "Телефон", "Phone"), False, ["TEL", "PHONE", "ТЕЛЕФОН", "TELEFONRAQAMI"]),
    ("EMAIL", _t("E-pochta", "Эл. почта", "Email"), False, ["EPOCHTA", "MAIL", "ПОЧТА", "ELPOCHTA"]),
    ("PASPORT", _t("Pasport seriyasi va raqami", "Серия и номер паспорта", "Passport series and number"), True,
     ["PASSPORT", "ПАСПОРТ", "PASPORTSERIYA", "PASPORTRAQAMI", "IDKARTA"]),
    ("PASPORT_BERILGAN", _t("Pasport kim tomonidan va qachon berilgan", "Кем и когда выдан паспорт",
                            "Passport issued by / on"), False, ["PASPORTBERILGAN", "КЕМВЫДАН", "ISSUEDBY"]),
    ("JSHSHIR", _t("JSHSHIR", "ПИНФЛ", "Personal ID (PINFL)"), True, ["PINFL", "ПИНФЛ", "ЖШШИР", "JSHIR"]),
    ("TASHKILOT", _t("Ish joyi (tashkilot)", "Место работы (организация)", "Workplace (organisation)"), False,
     ["ISHJOYI", "TASHKILOTNOMI", "ORGANIZATION", "WORKPLACE", "МЕСТОРАБОТЫ", "ОРГАНИЗАЦИЯ", "MUASSASA"]),
    ("BOLIM", _t("Bo'lim / kafedra", "Отдел / кафедра", "Department"), False,
     ["KAFEDRA", "DEPARTMENT", "ОТДЕЛ", "КАФЕДРА", "BOLIMI", "FAKULTET"]),
    ("LAVOZIM", _t("Lavozim", "Должность", "Position"), False, ["POSITION", "ДОЛЖНОСТЬ", "LAVOZIMI"]),
    ("GURUH", _t("Guruh", "Группа", "Group"), False, ["GROUP", "ГРУППА", "GURUHI"]),
    ("MALUMOTI", _t("Ma'lumoti", "Образование", "Education"), False, ["EDUCATION", "ОБРАЗОВАНИЕ", "MALUMOT"]),
    ("TAMOMLAGAN", _t("Tamomlagan (OTM, yil)", "Окончил (вуз, год)", "Graduated from (year)"), False,
     ["GRADUATED", "ОКОНЧИЛ", "OTM"]),
    ("MUTAXASSISLIGI", _t("Mutaxassisligi", "Специальность", "Speciality"), False,
     ["SPECIALITY", "СПЕЦИАЛЬНОСТЬ", "MUTAXASSISLIK"]),
    ("ILMIY_DARAJA", _t("Ilmiy darajasi / unvoni", "Учёная степень / звание", "Academic degree / title"), False,
     ["ILMIYDARAJASI", "DEGREE", "СТЕПЕНЬ", "ILMIYUNVON"]),
    ("TILLAR", _t("Chet tillari", "Иностранные языки", "Languages"), False, ["LANGUAGES", "ЯЗЫКИ", "CHETTILLARI"]),
    ("RAHBAR_LAVOZIM", _t("Rahbar lavozimi", "Должность руководителя", "Manager's position"), False,
     ["RAHBARLAVOZIMI", "BOSHLIQLAVOZIMI", "ДОЛЖНОСТЬРУКОВОДИТЕЛЯ"]),
    ("RAHBAR_FISH", _t("Rahbar F.I.Sh.", "ФИО руководителя", "Manager's name"), False,
     ["RAHBAR", "BOSHLIQ", "РУКОВОДИТЕЛЬ", "DIREKTOR", "REKTOR", "RAHBARFISH"]),
]

FIELD_KEYS = [f[0] for f in FIELDS]
MASKED = {f[0] for f in FIELDS if f[2]}

# Hisoblanadigan (profil va sanadan avtomatik) qiymatlar
COMPUTED = {
    "FISH": _t("F.I.Sh. (to'liq)", "ФИО полностью", "Full name"),
    "QISQA_FISH": _t("F.I.Sh. (qisqa: B.M. Hakimov)", "ФИО кратко", "Short name"),
    "SANA": _t("Bugungi sana (27.11.2026)", "Сегодняшняя дата", "Today's date"),
    "SANA_SOZ": _t("Bugungi sana (2026-yil 27-noyabr)", "Дата прописью", "Today's date in words"),
    "YIL": _t("Joriy yil", "Текущий год", "Current year"),
}
COMPUTED_ALIASES = {
    "FISH": ["FIO", "FIOSH", "FULLNAME", "ФИО", "ФИШ", "FAMILIYAISMI", "FAMILIYAISMSHARIF", "ISMFAMILIYA"],
    "QISQA_FISH": ["QISQAFISH", "FISHQISQA", "SHORTNAME", "ИНИЦИАЛЫ", "IMZOFISH"],
    "SANA": ["DATE", "ДАТА", "BUGUN", "TODAY", "SANASI"],
    "SANA_SOZ": ["SANASOZ", "SANAMATN", "ДАТАПРОПИСЬЮ"],
    "YIL": ["YEAR", "ГОД"],
}

MONTHS_UZ = ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avgust", "sentabr", "oktabr",
             "noyabr", "dekabr"]

_CYR = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i",
        "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
        "у": "u", "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "", "ы": "i", "ь": "",
        "э": "e", "ю": "yu", "я": "ya", "ў": "o", "қ": "q", "ғ": "g", "ҳ": "h"}


def norm(name: str) -> str:
    """O'rinbosar nomini solishtirish uchun: katta harf, faqat harf va raqam."""
    s = (name or "").strip().upper()
    return re.sub(r"[^0-9A-ZА-ЯЁЎҚҒҲ]", "", s.replace("'", "").replace("`", "").replace("’", "").replace("ʻ", ""))


def _translit(s: str) -> str:
    return "".join(_CYR.get(ch, ch) for ch in s.lower()).upper()


def _index() -> Dict[str, str]:
    idx = {}
    for key, _, _, aliases in FIELDS:
        for a in [key] + aliases:
            idx[norm(a)] = key
            idx[norm(_translit(a))] = key
    for key, aliases in COMPUTED_ALIASES.items():
        for a in [key] + aliases:
            idx[norm(a)] = key
            idx[norm(_translit(a))] = key
    return idx


_INDEX = _index()


def match_field(placeholder: str) -> Optional[str]:
    """O'rinbosar nomiga mos profil maydoni (yoki hisoblanadigan qiymat) kaliti, topilmasa None."""
    n = norm(placeholder)
    if n in _INDEX:
        return _INDEX[n]
    t = norm(_translit(placeholder))
    return _INDEX.get(t)


def full_name(p: Dict[str, str]) -> str:
    return " ".join(x for x in (p.get("FAMILIYA", ""), p.get("ISM", ""), p.get("OTASINING_ISMI", "")) if x).strip()


def short_name(p: Dict[str, str]) -> str:
    """O'zbek rasmiy uslubi: "B.M. Hakimov"."""
    fam, ism, ota = (p.get(k, "").strip() for k in ("FAMILIYA", "ISM", "OTASINING_ISMI"))
    ini = ""
    if ism:
        ini += ism[0].upper() + "."
    if ota:
        ini += ota[0].upper() + "."
    return f"{ini} {fam}".strip()


def date_words(d: date) -> str:
    return f"{d.year}-yil {d.day}-{MONTHS_UZ[d.month - 1]}"


def values_for(profile: Dict[str, str], today: date = None) -> Dict[str, str]:
    """Profil + hisoblanadigan qiymatlar."""
    today = today or date.today()
    v = {k: (profile.get(k) or "").strip() for k in FIELD_KEYS}
    v["FISH"] = full_name(profile)
    v["QISQA_FISH"] = short_name(profile)
    v["SANA"] = today.strftime("%d.%m.%Y")
    v["SANA_SOZ"] = date_words(today)
    v["YIL"] = str(today.year)
    return v


def mask(value: str, keep_start: int = 2, keep_end: int = 2) -> str:
    """Maxfiy qiymatni ko'rsatish uchun: "AB1234567" -> "AB*****67"."""
    s = (value or "").replace(" ", "")
    if len(s) <= keep_start + keep_end:
        return "*" * len(s)
    return s[:keep_start] + "*" * (len(s) - keep_start - keep_end) + s[-keep_end:]


def label(key: str, lang: str) -> str:
    for k, lab, _, _ in FIELDS:
        if k == key:
            return lab.get(lang) or lab["uz"]
    if key in COMPUTED:
        return COMPUTED[key].get(lang) or COMPUTED[key]["uz"]
    return key


def all_keys() -> List[str]:
    return FIELD_KEYS + list(COMPUTED)


def validate_jshshir(v: str) -> bool:
    s = (v or "").strip()
    return not s or (s.isdigit() and len(s) == 14)


def validate_passport(v: str) -> bool:
    s = (v or "").replace(" ", "").upper()
    return not s or bool(re.fullmatch(r"[A-Z]{2}\d{7}", s))
