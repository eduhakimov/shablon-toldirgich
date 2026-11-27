"""Word (.docx) shablonlaridagi {{O'RINBOSAR}} larni topish va to'ldirish.

Word ko'pincha bitta o'rinbosarni bir nechta "run" ga bo'lib yuboradi ({{FI, SH}}, bold va h.k.).
Shu sabab matn har bir paragrafning barcha <w:t> elementlari bo'yicha birlashtirilib qidiriladi,
almashtirilgan qiymat esa o'rinbosar boshlangan run formatini oladi. Asosiy matn, jadvallar
(ichma-ich ham), sarlavha/pastki kolontitullar, izohlar va matn bloklari (text box) qayta ishlanadi.
"""
import copy
import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from docx import Document
from docx.oxml.ns import qn

PH_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
W_P, W_T, W_R, W_BR = qn("w:p"), qn("w:t"), qn("w:r"), qn("w:br")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_PART_RE = re.compile(r"/word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")


class FillError(Exception):
    def __init__(self, key, **kw):
        super().__init__(key)
        self.key, self.kw = key, kw


def _parts(doc) -> Iterable:
    """Hujjatning matnli XML qismlari (asosiy, kolontitullar, izohlar)."""
    seen = set()
    for part in doc.part.package.iter_parts():
        name = str(part.partname)
        if _PART_RE.search(name) and hasattr(part, "element") and name not in seen:
            seen.add(name)
            yield part.element


def _paragraphs(root) -> Iterable:
    return root.iter(W_P)


def _texts(p) -> List:
    """Paragrafga tegishli <w:t> lar (ichki paragraflar — matn bloklari — bundan mustasno)."""
    out = []
    for t in p.iter(W_T):
        anc = t.getparent()
        while anc is not None and anc.tag != W_P:
            anc = anc.getparent()
        if anc is p:
            out.append(t)
    return out


def open_doc(path):
    try:
        return Document(str(path))
    except Exception as e:
        raise FillError("err_open_docx", name=Path(path).name, detail=str(e)[:200])


def scan(path) -> List[str]:
    """Shablondagi o'rinbosarlar (birinchi uchragan tartibda, takrorsiz)."""
    doc = open_doc(path)
    found: List[str] = []
    for root in _parts(doc):
        for p in _paragraphs(root):
            ts = _texts(p)
            if not ts:
                continue
            text = "".join(t.text or "" for t in ts)
            for m in PH_RE.finditer(text):
                name = m.group(1).strip()
                if name not in found:
                    found.append(name)
    return found


_CTRL_RE = re.compile(r"[\x00-\x08\x0c\x0e-\x1f\x7f\ufffe\uffff]")


def clean_value(v) -> str:
    """Qiymatni Word XML ga yaroqli qiladi: Excel yoki CSV dan kelgan CR LF va VT (Alt+Enter)
    yangi qatorga aylanadi, boshqa boshqaruv belgilari olib tashlanadi (aks holda lxml xato beradi)."""
    s = str(v).replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n")
    return _CTRL_RE.sub("", s)


def _set_text(t, text):
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set(XML_SPACE, "preserve")


def _insert_multiline(t, value: str):
    """Qiymatda yangi qator bo'lsa: birinchi qator shu <w:t> da, keyingilari <w:br/> + <w:t> bilan."""
    lines = value.split("\n")
    _set_text(t, lines[0])
    prev = t
    for line in lines[1:]:
        br = t.makeelement(W_BR, {})
        nt = t.makeelement(W_T, {})
        nt.set(XML_SPACE, "preserve")
        nt.text = line
        prev.addnext(br)
        br.addnext(nt)
        prev = nt


def _fill_paragraph(p, values: Dict[str, str], keep_missing: bool) -> Tuple[int, List[str]]:
    ts = _texts(p)
    if not ts:
        return 0, []
    full = "".join(t.text or "" for t in ts)
    if "{{" not in full:
        return 0, []
    matches = list(PH_RE.finditer(full))
    if not matches:
        return 0, []
    # har bir <w:t> ning umumiy matndagi boshlanish nuqtasi
    starts, pos = [], 0
    for t in ts:
        starts.append(pos)
        pos += len(t.text or "")

    def locate(off):
        for i in range(len(ts) - 1, -1, -1):
            if starts[i] <= off:
                return i, off - starts[i]
        return 0, 0

    count, missing = 0, []
    for m in reversed(matches):
        name = m.group(1).strip()
        if name in values and values[name] is not None:
            val = clean_value(values[name])
        else:
            missing.append(name)
            if keep_missing:
                continue
            val = ""
        si, so = locate(m.start())
        ei, eo = locate(m.end() - 1)
        eo += 1
        if si == ei:
            t = ts[si]
            txt = t.text or ""
            new = txt[:so] + val + txt[eo:]
            if "\n" in val:
                _insert_multiline(t, new)
            else:
                _set_text(t, new)
        else:
            first, last = ts[si], ts[ei]
            tail = (last.text or "")[eo:]
            _set_text(last, tail)
            for k in range(si + 1, ei):
                _set_text(ts[k], "")
            head = (first.text or "")[:so]
            if "\n" in val:
                _insert_multiline(first, head + val)
            else:
                _set_text(first, head + val)
        count += 1
    return count, missing


def fill(src, dst, values: Dict[str, str], keep_missing: bool = False) -> Dict[str, object]:
    """Shablonni to'ldirib, yangi faylga saqlaydi. Kirish fayli o'zgarmaydi.

    values — {o'rinbosar_nomi: qiymat}. keep_missing=True bo'lsa topilmaganlari {{...}} holida qoladi,
    aks holda bo'sh qatorga almashtiriladi.
    Natija: {"replaced": soni, "missing": [nomlar]}.
    """
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise FillError("err_same_file")
    doc = open_doc(src)
    total, missing = 0, []
    for root in _parts(doc):
        for p in list(_paragraphs(root)):
            c, miss = _fill_paragraph(p, values, keep_missing)
            total += c
            for x in miss:
                if x not in missing:
                    missing.append(x)
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(dst.name + ".tmp")
    doc.save(str(tmp))
    tmp.replace(dst)
    return {"replaced": total, "missing": missing}


def preview_paragraphs(path, values: Dict[str, str] = None) -> List[Tuple[str, str]]:
    """Oldindan ko'rish uchun asosiy matn paragraflari: [(tur, matn)], tur: 'p' | 'h' | 'row'.
    O'rinbosarlar o'rnida \x01nom\x02 belgilari qoladi (interfeys ularni ajratib ko'rsatadi)."""
    doc = open_doc(path)
    out = []
    body = doc.element.body

    def para_text(p):
        return "".join(t.text or "" for t in _texts(p))

    for child in body.iterchildren():
        if child.tag == W_P:
            txt = para_text(child)
            style = child.find(qn("w:pPr"))
            align = ""
            if style is not None:
                jc = style.find(qn("w:jc"))
                if jc is not None:
                    align = jc.get(qn("w:val")) or ""
            out.append((align or "p", txt))
        elif child.tag == qn("w:tbl"):
            for tr in child.iter(qn("w:tr")):
                cells = []
                for tc in tr.iter(qn("w:tc")):
                    cells.append(" ".join(para_text(p) for p in tc.iter(W_P)))
                out.append(("row", "\t".join(cells)))
    return out


def unique_path(path) -> Path:
    p = Path(path)
    if not p.exists():
        return p
    i = 2
    while True:
        c = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if not c.exists():
            return c
        i += 1


def safe_filename(text: str, maxlen: int = 120) -> str:
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", text or "")
    t = re.sub(r"\s+", " ", t).strip(" .")
    return t[:maxlen].strip() or "hujjat"


def render_name(pattern: str, values: Dict[str, str]) -> str:
    """Fayl nomi namunasi: "{FISH} - ariza" -> "Hakimov B. - ariza"."""
    def rep(m):
        return str(values.get(m.group(1).strip(), ""))
    return safe_filename(re.sub(r"\{([^{}]+)\}", rep, pattern or ""))
