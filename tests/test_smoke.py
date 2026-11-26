"""Tutun testi: oyna xatosiz yaratiladimi."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def test_window_creates():
    from PySide6.QtWidgets import QApplication
    from uzkit import app as appmod, i18n
    from uzkit.settings import Settings
    appmod.META = appmod.load_meta()
    appmod.SETTINGS = Settings(appmod.META["slug"] + "_test")
    qapp = QApplication.instance() or QApplication([])
    from shablon_toldirgich.window import MainWindow
    for lang in ("uz", "ru", "en"):
        i18n.set_language(lang)
        w = MainWindow()
        w.close()


def test_translations_complete():
    from uzkit import i18n
    import shablon_toldirgich.window  # noqa: registers strings
    assert i18n.missing_keys() == {}


def test_fill_page_flow(tmp_path):
    from PySide6.QtWidgets import QApplication
    from uzkit import app as appmod, i18n
    from uzkit.settings import Settings
    appmod.META = appmod.load_meta()
    appmod.SETTINGS = Settings(appmod.META["slug"] + "_test")
    QApplication.instance() or QApplication([])
    i18n.set_language("uz")
    from shablon_toldirgich.window import MainWindow
    w = MainWindow()
    tpl = next(it for it in w.state.library.items if "tatil" in it["path"])
    w.use_template(tpl)
    assert w.fill.template == tpl["path"]
    assert w.fill.table.rowCount() == len(w.fill.placeholders) > 5
    w.fill.update_preview()
    assert "A R I Z A" in w.fill.preview.toPlainText()
    w.close()
