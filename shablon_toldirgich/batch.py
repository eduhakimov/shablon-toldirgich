"""Ommaviy to'ldirish: bitta shablon × Excel jadvalidagi har bir qator."""
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from . import docx_fill, fields


class BatchError(Exception):
    def __init__(self, key, **kw):
        super().__init__(key)
        self.key, self.kw = key, kw


def _cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_table(path, sheet: str = None, limit: int = 100000) -> Tuple[List[str], List[Dict[str, str]], List[str]]:
    """Excel (.xlsx) yoki CSV: (sarlavhalar, qatorlar, varaqlar_nomlari). Birinchi qator — sarlavha."""
    p = Path(path)
    if p.suffix.lower() == ".csv":
        import csv
        raw = p.read_bytes()
        text = None
        for enc in ("utf-8-sig", "cp1251"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        dialect = csv.Sniffer().sniff(text[:2000], delimiters=",;\t") if text.strip() else csv.excel
        rows = list(csv.reader(text.splitlines(), dialect))
        sheets = []
    else:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        except Exception as e:
            raise BatchError("err_excel", detail=str(e)[:200])
        sheets = wb.sheetnames
        ws = wb[sheet] if sheet and sheet in sheets else wb[sheets[0]]
        rows = []
        for r in ws.iter_rows(values_only=True):
            rows.append([_cell(v) for v in r])
            if len(rows) > limit:
                break
        wb.close()
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise BatchError("err_excel_empty")
    headers = [str(h).strip() or f"C{i + 1}" for i, h in enumerate(rows[0])]
    data = []
    for r in rows[1:]:
        d = {}
        for i, h in enumerate(headers):
            d[h] = _cell(r[i]) if i < len(r) else ""
        data.append(d)
    return headers, data, sheets


def map_columns(placeholders: List[str], headers: List[str]) -> Dict[str, Optional[str]]:
    """O'rinbosar -> ustun nomi. Avval aynan nom, keyin normallashtirilgan nom, keyin maydon kaliti bo'yicha."""
    by_norm = {fields.norm(h): h for h in headers}
    by_field = {}
    for h in headers:
        k = fields.match_field(h)
        if k and k not in by_field:
            by_field[k] = h
    out = {}
    for ph in placeholders:
        if ph in headers:
            out[ph] = ph
        elif fields.norm(ph) in by_norm:
            out[ph] = by_norm[fields.norm(ph)]
        else:
            k = fields.match_field(ph)
            out[ph] = by_field.get(k) if k else None
    return out


def row_values(placeholders, mapping, row: Dict[str, str], defaults: Dict[str, str] = None,
               today: date = None) -> Dict[str, str]:
    """Bitta qator uchun qiymatlar: ustundan, bo'lmasa hisoblanadigan/standart qiymatdan."""
    # qatordan profil ko'rinishini tiklaymiz (FAMILIYA, ISM ... ustunlari bo'lsa FISH hisoblanadi)
    prof = {}
    for h, v in row.items():
        k = fields.match_field(h)
        if k in fields.FIELD_KEYS and v:
            prof[k] = v
    computed = fields.values_for(prof, today)
    defaults = defaults or {}
    out = {}
    for ph in placeholders:
        col = mapping.get(ph)
        v = row.get(col, "") if col else ""
        if not v:
            k = fields.match_field(ph)
            v = computed.get(k, "") if k else ""
        if not v:
            v = defaults.get(ph, "")
        out[ph] = v
    return out


def run(template, rows: List[Dict[str, str]], outdir, name_pattern: str, defaults: Dict[str, str] = None,
        progress=None) -> List[Path]:
    """Har bir qator uchun alohida hujjat. Natija: yaratilgan fayllar."""
    placeholders = docx_fill.scan(template)
    if not rows:
        raise BatchError("err_excel_empty")
    headers = list(rows[0].keys())
    mapping = map_columns(placeholders, headers)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    res = []
    for k, row in enumerate(rows):
        if progress:
            if getattr(progress, "cancelled", None) and progress.cancelled():
                break
            progress(int(100 * k / len(rows)), f"{k + 1}/{len(rows)}")
        vals = row_values(placeholders, mapping, row, defaults)
        name_src = dict(row)
        name_src.update(vals)
        name = docx_fill.render_name(name_pattern, name_src) or f"{k + 1:03d}"
        dst = docx_fill.unique_path(outdir / f"{name}.docx")
        docx_fill.fill(template, dst, vals)
        res.append(dst)
    return res
