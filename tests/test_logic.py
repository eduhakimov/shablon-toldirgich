"""Shablon to'ldirish mantiqining sinovlari (vaqtinchalik papkada)."""
import os
import sys
from datetime import date
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shablon_toldirgich import batch, docx_fill, fields
from shablon_toldirgich.store import ProfileStore, StoreLocked, WrongPassword


def make_split_doc(path: Path):
    """O'rinbosarlar runlarga bo'lingan, jadval, kolontitul va matn blokisiz hujjat."""
    d = Document()
    p = d.add_paragraph()
    r1 = p.add_run("Hurmatli {{RAH")
    r1.bold = True
    r2 = p.add_run("BAR_FISH}}, men ")
    r2.italic = True
    p.add_run("{{")
    p.add_run("FISH")
    r = p.add_run("}}dan ariza.")
    r.font.size = Pt(18)
    d.add_paragraph("Sana: {{SANA}}; ikki marta: {{FISH}} / {{ FISH }}")
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "Telefon"
    t.cell(0, 1).text = "{{TELEFON}}"
    inner = t.cell(0, 1).add_table(rows=1, cols=1)
    inner.cell(0, 0).text = "Ichki: {{JSHSHIR}}"
    d.sections[0].header.paragraphs[0].text = "Kolontitul {{TASHKILOT}}"
    d.sections[0].footer.paragraphs[0].text = "Pastki {{YIL}}"
    d.add_paragraph("Noma'lum: {{BOSHQA_NARSA}}")
    d.add_paragraph("Ko'p qator: {{IZOH}} tugadi")
    d.save(path)
    return path


def all_text(path):
    d = Document(str(path))
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                parts.append(c.text)
                for it in c.tables:
                    parts.append(it.cell(0, 0).text)
    for s in d.sections:
        parts += [p.text for p in s.header.paragraphs] + [p.text for p in s.footer.paragraphs]
    return "\n".join(parts)


PROFILE = {"FAMILIYA": "Karimov", "ISM": "Anvar", "OTASINING_ISMI": "Rustamovich", "TELEFON": "+998 90 123 45 67",
           "JSHSHIR": "31234567890123", "TASHKILOT": "Toshkent davlat texnika universiteti",
           "RAHBAR_FISH": "A.B. Yusupov"}


def test_scan_finds_split_placeholders(tmp_path):
    p = make_split_doc(tmp_path / "t.docx")
    names = docx_fill.scan(p)
    assert names == ["RAHBAR_FISH", "FISH", "SANA", "TELEFON", "JSHSHIR", "BOSHQA_NARSA", "IZOH", "TASHKILOT",
                     "YIL"] or set(names) == {"RAHBAR_FISH", "FISH", "SANA", "TELEFON", "JSHSHIR", "TASHKILOT",
                                              "YIL", "BOSHQA_NARSA", "IZOH"}


def test_fill_everywhere_and_keep_format(tmp_path):
    src = make_split_doc(tmp_path / "t.docx")
    dst = tmp_path / "out.docx"
    vals = fields.values_for(PROFILE, date(2026, 11, 27))
    vals["IZOH"] = "birinchi\nikkinchi"
    res = docx_fill.fill(src, dst, vals)
    assert res["missing"] == ["BOSHQA_NARSA"]
    txt = all_text(dst)
    assert "Hurmatli A.B. Yusupov, men Karimov Anvar Rustamovichdan ariza." in txt
    assert "Sana: 27.11.2026; ikki marta: Karimov Anvar Rustamovich / Karimov Anvar Rustamovich" in txt
    assert "+998 90 123 45 67" in txt and "Ichki: 31234567890123" in txt
    assert "Kolontitul Toshkent davlat texnika universiteti" in txt and "Pastki 2026" in txt
    assert "{{" not in txt and "Noma'lum: " in txt
    assert "birinchi\nikkinchi tugadi" in txt
    # format: o'rinbosar boshlangan run (bold) qiymatni oladi
    d = Document(str(dst))
    runs = d.paragraphs[0].runs
    assert runs[0].bold and "A.B. Yusupov" in runs[0].text
    # asl fayl o'zgarmagan
    assert "{{RAH" in Document(str(src)).paragraphs[0].runs[0].text


def test_keep_missing(tmp_path):
    src = make_split_doc(tmp_path / "t.docx")
    dst = tmp_path / "o2.docx"
    docx_fill.fill(src, dst, {"FISH": "X"}, keep_missing=True)
    txt = all_text(dst)
    assert "{{BOSHQA_NARSA}}" in txt and "{{SANA}}" in txt and "X" in txt


def test_same_file_guard(tmp_path):
    src = make_split_doc(tmp_path / "t.docx")
    with pytest.raises(docx_fill.FillError):
        docx_fill.fill(src, src, {})


def test_field_matching():
    assert fields.match_field("FISH") == "FISH"
    assert fields.match_field("F.I.Sh.") == "FISH"
    assert fields.match_field("ФИО") == "FISH"
    assert fields.match_field("tel") == "TELEFON"
    assert fields.match_field("Pasport") == "PASPORT"
    assert fields.match_field("ПИНФЛ") == "JSHSHIR"
    assert fields.match_field("ish joyi") == "TASHKILOT"
    assert fields.match_field("Tug'ilgan sana") == "TUGILGAN_SANA"
    assert fields.match_field("date") == "SANA"
    assert fields.match_field("ISHONCHLI_SHAXS_FISH") is None


def test_computed_values():
    v = fields.values_for(PROFILE, date(2026, 3, 8))
    assert v["FISH"] == "Karimov Anvar Rustamovich"
    assert v["QISQA_FISH"] == "A.R. Karimov"
    assert v["SANA"] == "08.03.2026" and v["SANA_SOZ"] == "2026-yil 8-mart" and v["YIL"] == "2026"


def test_mask_and_validation():
    assert fields.mask("AB1234567") == "AB*****67"
    assert fields.validate_passport("AB 1234567") and not fields.validate_passport("A1234567")
    assert fields.validate_jshshir("31234567890123") and not fields.validate_jshshir("123")


def test_store_plain_and_encrypted(tmp_path):
    s = ProfileStore(tmp_path / "profiles.json")
    s.load()
    p = s.add("Men")
    p["fields"].update(PROFILE)
    s.save()
    s2 = ProfileStore(tmp_path / "profiles.json")
    assert s2.load()[0]["fields"]["FAMILIYA"] == "Karimov"
    s2.set_password("maxfiy")
    raw = (tmp_path / "profiles.json").read_text(encoding="utf-8")
    assert "Karimov" not in raw and '"encrypted": true' in raw
    s3 = ProfileStore(tmp_path / "profiles.json")
    with pytest.raises(StoreLocked):
        s3.load()
    with pytest.raises(WrongPassword):
        s3.load("xato")
    assert s3.load("maxfiy")[0]["title"] == "Men"
    s3.set_password(None)
    assert "Karimov" in (tmp_path / "profiles.json").read_text(encoding="utf-8")


def test_builtin_templates_scan_and_fill(tmp_path):
    tdir = Path(__file__).resolve().parents[1] / "assets" / "templates"
    files = sorted(tdir.glob("*.docx"))
    assert len(files) >= 10
    vals = fields.values_for(PROFILE)
    for f in files:
        names = docx_fill.scan(f)
        assert names, f.name
        full = {n: vals.get(fields.match_field(n) or n, "") or "X" for n in names}
        res = docx_fill.fill(f, tmp_path / f.name, full)
        assert res["missing"] == []
        assert "{{" not in all_text(tmp_path / f.name)


def test_batch_from_excel(tmp_path):
    import openpyxl
    from datetime import datetime
    tpl = tmp_path / "t.docx"
    d = Document()
    d.add_paragraph("{{FISH}} — {{GURUH}} — {{TUGILGAN_SANA}} — {{SANA}} — {{KURS}}")
    d.save(tpl)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Familiya", "Ism", "Otasining ismi", "Guruh", "Tug'ilgan sana"])
    ws.append(["Aliyev", "Vali", "Karimovich", "911-21", datetime(2003, 5, 1)])
    ws.append(["Soliyeva", "Gulnora", "", "912-21", datetime(2004, 1, 9)])
    ws.append([None, None, None, None, None])
    x = tmp_path / "r.xlsx"
    wb.save(x)
    headers, rows, sheets = batch.read_table(x)
    assert len(rows) == 2 and headers[0] == "Familiya"
    m = batch.map_columns(docx_fill.scan(tpl), headers)
    assert m["GURUH"] == "Guruh" and m["TUGILGAN_SANA"] == "Tug'ilgan sana"
    out = batch.run(tpl, rows, tmp_path / "out", "{FISH}", defaults={"KURS": "3"})
    assert [p.name for p in out] == ["Aliyev Vali Karimovich.docx", "Soliyeva Gulnora.docx"]
    t = Document(str(out[0])).paragraphs[0].text
    assert t.startswith("Aliyev Vali Karimovich — 911-21 — 01.05.2003 — ") and t.endswith(" — 3")


def test_batch_csv(tmp_path):
    c = tmp_path / "a.csv"
    c.write_text("FISH;Telefon\nA B;1\nC D;2\n", encoding="utf-8")
    headers, rows, _ = batch.read_table(c)
    assert headers == ["FISH", "Telefon"] and rows[1]["Telefon"] == "2"


def test_render_name():
    assert docx_fill.render_name("{FISH} - ariza", {"FISH": "A/B"}) == "A B - ariza"
    assert docx_fill.render_name("", {}) == "hujjat"


def test_library(tmp_path):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from shablon_toldirgich.library import Library
    lib = Library(tmp_path / "lib")
    n = len(lib.items)
    assert n >= 10 and all(i["builtin"] for i in lib.items)
    src = tmp_path / "my.docx"
    Document().save(src)
    it = lib.add(src, "Mening shablonim", "xat")
    assert len(lib.items) == n + 1 and not it["builtin"] and src.exists()
    assert lib.remove(it["id"]) and len(lib.items) == n
    assert lib.remove(lib.items[0]["id"]) is False
