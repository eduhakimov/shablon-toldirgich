"""DOCX -> PDF: Microsoft Word (COM) orqali, bo'lmasa LibreOffice (agar o'rnatilgan bo'lsa)."""
import os
import shutil
import subprocess
from pathlib import Path
from typing import List, Optional

WD_FORMAT_PDF = 17


class ConvertError(Exception):
    def __init__(self, key, **kw):
        super().__init__(key)
        self.key, self.kw = key, kw


def word_available() -> bool:
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
        return True
    except Exception:
        return False


def _soffice() -> Optional[str]:
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        p = Path(base) / "LibreOffice" / "program" / "soffice.exe"
        if p.exists():
            return str(p)
    return shutil.which("soffice")


def pdf_available() -> bool:
    return word_available() or bool(_soffice())


def to_pdf_many(docs: List[str], progress=None) -> List[Path]:
    """Bir nechta .docx ni PDF ga o'tkazadi (bitta Word nusxasi bilan). Natija: PDF yo'llari."""
    docs = [str(Path(d).resolve()) for d in docs]
    out: List[Path] = []
    if word_available():
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for k, d in enumerate(docs):
                pdf = str(Path(d).with_suffix(".pdf"))
                doc = word.Documents.Open(d, False, True, False)   # ConfirmConversions, ReadOnly, AddToRecent
                try:
                    doc.SaveAs2(pdf, FileFormat=WD_FORMAT_PDF)
                finally:
                    doc.Close(False)
                out.append(Path(pdf))
                if progress:
                    progress(k + 1, len(docs))
            return out
        except Exception as e:
            if not _soffice():
                raise ConvertError("err_pdf_word", detail=str(e)[:200])
        finally:
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
    so = _soffice()
    if not so:
        raise ConvertError("err_pdf_none")
    for k, d in enumerate(docs):
        r = subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", str(Path(d).parent), d],
                           capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                           timeout=180)
        pdf = Path(d).with_suffix(".pdf")
        if r.returncode != 0 or not pdf.exists():
            raise ConvertError("err_pdf_lo")
        out.append(pdf)
        if progress:
            progress(k + 1, len(docs))
    return out
