"""Chuqur ssenariylar: Excel/CSV dagi "iflos" qiymatlar, bitta ustunli CSV, ommaviy to'ldirish."""
import sys
from pathlib import Path

import pytest
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shablon_toldirgich import batch, docx_fill  # noqa: E402
from test_logic import all_text, make_split_doc  # noqa: E402


@pytest.mark.parametrize("value, shown", [
    ("Ali\r\nValiyev", "Ali\nValiyev"),
    ("Ali\x0bValiyev", "Ali\nValiyev"),           # Excel Alt+Enter
    ("nol\x00belgi\x07", "nolbelgi"),
    ("O'tkir Ҳасанов «Рустамович»", "O'tkir Ҳасанов «Рустамович»"),
])
def test_dirty_values_do_not_break_docx(tmp_path, value, shown):
    tpl = make_split_doc(tmp_path / "t.docx")
    out = tmp_path / "o.docx"
    docx_fill.fill(tpl, out, {"FISH": value, "IZOH": value})
    text = all_text(out)
    assert "{{FISH}}" not in text
    assert shown.split("\n")[0] in text and shown.split("\n")[-1] in text
    Document(str(out))                     # qayta ochiladi


def test_single_column_csv_and_cp1251(tmp_path):
    c = tmp_path / "bir.csv"
    c.write_text("FISH\nAli Valiyev\nO'tkir Hasanov\n", encoding="utf-8")
    headers, rows, _ = batch.read_table(c)
    assert headers == ["FISH"] and [r["FISH"] for r in rows] == ["Ali Valiyev", "O'tkir Hasanov"]
    c2 = tmp_path / "win.csv"
    c2.write_bytes("ФИО;Телефон\nАли Валиев;1\n".encode("cp1251") + b"\x98;2\n")
    headers, rows, _ = batch.read_table(c2)
    assert headers == ["ФИО", "Телефон"] and rows[0]["ФИО"] == "Али Валиев" and len(rows) == 2


def test_batch_from_excel_with_line_breaks_and_duplicates(tmp_path):
    import openpyxl
    tpl = make_split_doc(tmp_path / "t.docx")
    x = tmp_path / "royxat.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["FISH", "Izoh", "Telefon"])
    ws.append(["Aliyev Vali", "birinchi\nikkinchi", "+998 90 000 00 01"])
    ws.append(["Aliyev Vali", "x\r\ny", None])
    ws.append([None, None, None])
    ws.append(["Soliyeva / Gulnora", "", 998901234567])
    wb.save(x)
    headers, rows, _ = batch.read_table(x)
    assert len(rows) == 3
    out = batch.run(tpl, rows, tmp_path / "chiqdi", "{FISH} - ariza")
    assert [p.name for p in out] == ["Aliyev Vali - ariza.docx", "Aliyev Vali - ariza (2).docx",
                                     "Soliyeva Gulnora - ariza.docx"]
    assert "998901234567" in all_text(out[2])
    for p in out:
        Document(str(p))


def test_template_unchanged_and_missing_kept(tmp_path):
    tpl = make_split_doc(tmp_path / "t.docx")
    before = tpl.read_bytes()
    res = docx_fill.fill(tpl, tmp_path / "o.docx", {"FISH": "X"}, keep_missing=True)
    assert tpl.read_bytes() == before
    assert "BOSHQA_NARSA" in res["missing"]
    assert "{{BOSHQA_NARSA}}" in all_text(tmp_path / "o.docx")


def test_broken_docx_reports_error(tmp_path):
    bad = tmp_path / "buzuq.docx"
    bad.write_bytes(b"PK\x03\x04 bu docx emas")
    with pytest.raises(docx_fill.FillError) as e:
        docx_fill.scan(bad)
    assert e.value.key == "err_open_docx"
