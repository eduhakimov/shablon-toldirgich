"""Asosiy oyna: yuqorida bo'limlar paneli; to'ldirish, profillar, shablonlar, Excel'dan ommaviy."""
import html
import os
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Optional

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMenu, QProgressBar, QPushButton, QScrollArea, QSplitter,
                               QStackedWidget, QTableWidget, QTableWidgetItem, QTabWidget, QTextBrowser,
                               QVBoxLayout, QWidget)

from uzkit import app as appmod
from uzkit.i18n import language, register, tr
from uzkit.paths import data_dir, documents_dir
from uzkit.window import BaseMainWindow
from uzkit.widgets import Card, PathPicker, Worker, ask, header, info, open_path, primary_button, run_worker, warn

from . import batch, convert, docx_fill, fields
from .library import CATEGORIES, Library
from .store import ProfileStore, StoreLocked, WrongPassword
from .strings import STRINGS

register(STRINGS)

SLUG = "shablon-toldirgich"
DOCX_FILTER = "Word (*.docx)"


def is_demo() -> bool:
    return "--screenshot" in sys.argv


def err_text(e: Exception) -> str:
    key = getattr(e, "key", None)
    if key:
        return tr(key, **getattr(e, "kw", {}))
    if isinstance(e, PermissionError):
        return tr("err_permission")
    return tr("err_generic", detail=str(e)[:300])


def title_of(item: Dict) -> str:
    t = item.get("title", {})
    return t.get(language()) or t.get("uz") or Path(item["path"]).stem


def make_table(cols, stretch_col=None) -> QTableWidget:
    t = QTableWidget(0, len(cols))
    t.setHorizontalHeaderLabels(cols)
    t.verticalHeader().setVisible(False)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    h = t.horizontalHeader()
    for i in range(len(cols)):
        h.setSectionResizeMode(i, QHeaderView.Stretch if i == stretch_col else QHeaderView.ResizeToContents)
    return t


def muted(text) -> QLabel:
    l = QLabel(text)
    l.setObjectName("Muted")
    l.setWordWrap(True)
    return l


# =========================================================================== Holat (umumiy ma'lumotlar)
class AppState:
    def __init__(self, win):
        self.win = win
        base = data_dir(SLUG)
        if is_demo():
            base = Path(tempfile.mkdtemp(prefix="shablon_demo_"))
        self.base = base
        self.store = ProfileStore(base / "profiles.json")
        self.library = Library(base / "templates")
        self.locked = False
        try:
            self.store.load()
        except StoreLocked:
            self.locked = True
        except Exception:
            self.store.profiles = []
        if is_demo():
            self._demo()

    def _demo(self):
        demo = [
            ("Men", {"FAMILIYA": "Karimov", "ISM": "Anvar", "OTASINING_ISMI": "Rustamovich",
                     "TUGILGAN_SANA": "14.06.1988", "TUGILGAN_JOY": "Namangan viloyati", "MILLATI": "o'zbek",
                     "MANZIL": "Toshkent sh., Chilonzor tumani, Bunyodkor ko'chasi, 12-uy",
                     "TELEFON": "+998 90 123-45-67", "EMAIL": "a.karimov@mail.uz", "PASPORT": "AD1234567",
                     "PASPORT_BERILGAN": "Chilonzor tumani IIB, 12.03.2021", "JSHSHIR": "31406880120034",
                     "TASHKILOT": "Toshkent davlat texnika universiteti", "BOLIM": "Informatika kafedrasi",
                     "LAVOZIM": "katta o'qituvchi", "MALUMOTI": "oliy", "TAMOMLAGAN": "TDTU, 2010",
                     "MUTAXASSISLIGI": "Axborot texnologiyalari", "ILMIY_DARAJA": "PhD",
                     "TILLAR": "rus, ingliz", "RAHBAR_LAVOZIM": "rektori", "RAHBAR_FISH": "B.A. Yusupov"}),
            ("Turmush o'rtog'im", {"FAMILIYA": "Karimova", "ISM": "Dilnoza", "OTASINING_ISMI": "Akramovna",
                                   "TELEFON": "+998 91 765-43-21", "TASHKILOT": "45-maktab", "LAVOZIM": "o'qituvchi"}),
            ("O'g'lim", {"FAMILIYA": "Karimov", "ISM": "Jasur", "OTASINING_ISMI": "Anvarovich", "GURUH": "911-23"}),
        ]
        self.store.profiles = []
        for title, f in demo:
            p = self.store.add(title)
            p["fields"].update(f)

    def save_profiles(self):
        if not is_demo():
            self.store.save()


# =========================================================================== 1. To'ldirish
class FillPage(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win, self.st = win, win.state
        self.template: Optional[str] = None
        self.placeholders: List[str] = []
        self.values: Dict[str, str] = {}
        self._last_docs: List[Path] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)
        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)

        # chap: shablonlar
        left = Card(tr("templates"))
        self.search = QLineEdit()
        self.search.setPlaceholderText(tr("search_templates"))
        self.search.setClearButtonEnabled(True)
        left.add(self.search)
        self.tlist = QListWidget()
        self.tlist.setObjectName("TemplateList")
        self.tlist.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tlist.setTextElideMode(Qt.ElideRight)
        left.add(self.tlist)
        self.open_btn = QPushButton(tr("open_docx"))
        left.add(self.open_btn)
        left.setMinimumWidth(230)
        split.addWidget(left)

        # o'ng: ma'lumotlar
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        top = Card()
        row = QHBoxLayout()
        self.tpl_lbl = QLabel(tr("no_template"))
        self.tpl_lbl.setObjectName("SectionTitle")
        self.tpl_lbl.setWordWrap(True)
        row.addWidget(self.tpl_lbl, 1)
        row.addWidget(QLabel(tr("person")))
        self.profile = QComboBox()
        self.profile.setMinimumWidth(190)
        row.addWidget(self.profile)
        top.add(row)
        self.tabs = QTabWidget()
        self.table = make_table([tr("col_placeholder"), tr("col_field"), tr("col_value")], stretch_col=2)
        self.table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.tabs.addTab(self.table, tr("tab_fields"))
        self.preview = QTextBrowser()
        self.preview.setObjectName("Paper")
        self.tabs.addTab(self.preview, tr("tab_preview"))
        top.add(self.tabs)
        srow = QHBoxLayout()
        self.missing_lbl = QLabel("")
        self.missing_lbl.setObjectName("Muted")
        self.show_secret = QCheckBox(tr("show_secret"))
        srow.addWidget(self.missing_lbl, 1)
        srow.addWidget(self.show_secret)
        top.add(srow)
        rl.addWidget(top, 1)
        split.addWidget(right)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([260, 700])
        lay.addWidget(split, 1)

        # pastki qator: chiqish sozlamalari
        out = Card()
        g = QGridLayout()
        g.setHorizontalSpacing(10)
        self.outdir = PathPicker("dir")
        self.outdir.set_path(self.win.settings.get("out_dir") or str(documents_dir() / "Arizalar"))
        self.fname = QLineEdit()
        self.pdf = QCheckBox(tr("also_pdf"))
        self.pdf.setChecked(bool(self.win.settings.get("also_pdf", False)))
        if not convert.pdf_available():
            self.pdf.setChecked(False)
            self.pdf.setEnabled(False)
            self.pdf.setToolTip(tr("pdf_unavailable"))
        g.addWidget(QLabel(tr("save_folder")), 0, 0)
        g.addWidget(self.outdir, 0, 1)
        g.addWidget(QLabel(tr("file_name")), 1, 0)
        g.addWidget(self.fname, 1, 1)
        g.setColumnStretch(1, 1)
        out.add(g)
        brow = QHBoxLayout()
        self.result = QLabel("")
        self.result.setWordWrap(True)
        self.open_doc = QPushButton(tr("open_document"))
        self.open_dir = QPushButton(tr("open_folder"))
        self.open_doc.hide()
        self.open_dir.hide()
        self.make_btn = primary_button(tr("make_document"))
        self.make_btn.setMinimumWidth(170)
        brow.addWidget(self.pdf)
        brow.addWidget(self.result, 1)
        brow.addWidget(self.open_doc)
        brow.addWidget(self.open_dir)
        brow.addWidget(self.make_btn)
        out.add(brow)
        lay.addWidget(out)

        self.search.textChanged.connect(self.fill_templates)
        self.tlist.currentItemChanged.connect(self._tpl_selected)
        self.open_btn.clicked.connect(self._open_file)
        self.profile.currentIndexChanged.connect(self._profile_changed)
        self.table.itemChanged.connect(self._value_edited)
        self.show_secret.toggled.connect(self.refresh_table)
        self.tabs.currentChanged.connect(lambda i: self.update_preview() if i == 1 else None)
        self.make_btn.clicked.connect(self.make)
        self.open_doc.clicked.connect(lambda: self._last_docs and open_path(str(self._last_docs[0])))
        self.open_dir.clicked.connect(lambda: self._last_docs and open_path(str(self._last_docs[0].parent)))
        self.fill_templates()
        self.fill_profiles()
        self.make_btn.setEnabled(False)

    # -- shablonlar ro'yxati
    def fill_templates(self):
        q = self.search.text().strip().lower()
        cur = self.tlist.currentItem().data(Qt.UserRole) if self.tlist.currentItem() else None
        self.tlist.blockSignals(True)
        self.tlist.clear()
        for cat in CATEGORIES:
            items = [it for it in self.st.library.items if it["category"] == cat and q in title_of(it).lower()]
            if not items:
                continue
            h = QListWidgetItem(tr("cat_" + cat).upper())
            h.setFlags(Qt.NoItemFlags)
            f = h.font()
            f.setPointSizeF(8)
            f.setBold(True)
            h.setFont(f)
            self.tlist.addItem(h)
            for it in items:
                li = QListWidgetItem(title_of(it))
                li.setData(Qt.UserRole, it["path"])
                li.setToolTip(it["path"])
                self.tlist.addItem(li)
                if it["path"] == cur:
                    self.tlist.setCurrentItem(li)
        self.tlist.blockSignals(False)

    def _tpl_selected(self, item, _prev=None):
        if item is not None and item.data(Qt.UserRole):
            self.set_template(item.data(Qt.UserRole), item.text())

    def _open_file(self):
        p, _ = QFileDialog.getOpenFileName(self, tr("select_file"), "", DOCX_FILTER)
        if p:
            self.tlist.clearSelection()
            self.set_template(p, Path(p).stem)

    def set_template(self, path, title):
        try:
            names = docx_fill.scan(path)
        except Exception as e:
            self.result.setText(f"<span style='color:#b42318'>{err_text(e)}</span>")
            return
        self.template, self.placeholders = str(path), names
        self.tpl_title = title
        self.tpl_lbl.setText(title)
        self.fname.setText(f"{docx_fill.safe_filename(title)} - {{QISQA_FISH}}")
        self.recompute()
        self.make_btn.setEnabled(True)
        self.result.setText("")
        self.open_doc.hide()
        self.open_dir.hide()

    # -- profillar
    def fill_profiles(self):
        cur = self.profile.currentData()
        self.profile.blockSignals(True)
        self.profile.clear()
        if self.st.locked:
            self.profile.addItem(tr("profiles_locked"), None)
        else:
            self.profile.addItem(tr("no_profile"), None)
            for p in self.st.store.profiles:
                self.profile.addItem(p["title"] or fields.full_name(p["fields"]) or "?", p["id"])
            i = self.profile.findData(cur if cur else self.win.settings.get("last_profile"))
            if i > 0:
                self.profile.setCurrentIndex(i)
            elif self.profile.count() > 1:
                self.profile.setCurrentIndex(1)
        self.profile.blockSignals(False)
        self.recompute()

    def _profile_changed(self):
        pid = self.profile.currentData()
        if pid and not is_demo():
            self.win.settings.set("last_profile", pid)
        self.recompute()

    def recompute(self):
        """Qiymatlarni profil, hisoblanadigan va eslab qolingan qiymatlardan yig'adi."""
        pid = self.profile.currentData()
        prof = self.st.store.get(pid) if pid else None
        base = fields.values_for(prof["fields"] if prof else {})
        remembered = self.win.settings.get("extra_values", {}) if not is_demo() else {}
        vals = {}
        for ph in self.placeholders:
            k = fields.match_field(ph)
            v = base.get(k, "") if k else ""
            if not v:
                v = remembered.get(ph, "")
            vals[ph] = v
        self.values = vals
        self.refresh_table()

    def refresh_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.placeholders))
        miss = 0
        warn_bg = QColor(self.win.palette_color("hover"))
        for r, ph in enumerate(self.placeholders):
            k = fields.match_field(ph)
            a = QTableWidgetItem("{{" + ph + "}}")
            a.setFlags(a.flags() & ~Qt.ItemIsEditable)
            b = QTableWidgetItem(fields.label(k, language()) if k else tr("custom_field"))
            b.setFlags(b.flags() & ~Qt.ItemIsEditable)
            if not k:
                b.setForeground(QColor(self.win.palette_color("muted")))
            v = self.values.get(ph, "")
            secret = k in fields.MASKED and not self.show_secret.isChecked()
            c = QTableWidgetItem(fields.mask(v) if secret and v else v)
            c.setData(Qt.UserRole, ph)
            if secret and v:
                c.setFlags(c.flags() & ~Qt.ItemIsEditable)
                c.setToolTip(tr("secret_hint"))
            if not v:
                miss += 1
                for it in (a, b, c):
                    it.setBackground(warn_bg)
                c.setToolTip(tr("enter_value"))
            for col, it in enumerate((a, b, c)):
                self.table.setItem(r, col, it)
        self.table.blockSignals(False)
        if self.placeholders:
            self.missing_lbl.setText(tr("n_fields", n=len(self.placeholders)) +
                                     (("  ·  " + tr("n_missing", n=miss)) if miss else ""))
        else:
            self.missing_lbl.setText(tr("no_placeholders") if self.template else "")
        if self.tabs.currentIndex() == 1:
            self.update_preview()

    def _value_edited(self, item):
        ph = item.data(Qt.UserRole)
        if ph:
            self.values[ph] = item.text()
            self.refresh_table()

    def update_preview(self):
        if not self.template:
            self.preview.setHtml("")
            return
        try:
            paras = docx_fill.preview_paragraphs(self.template)
        except Exception as e:
            self.preview.setPlainText(err_text(e))
            return
        acc = appmod.META.get("accent", "#2c5f8d")

        def render(text):
            out, pos = [], 0
            for m in docx_fill.PH_RE.finditer(text):
                out.append(html.escape(text[pos:m.start()]))
                name = m.group(1).strip()
                v = self.values.get(name, "")
                if v:
                    out.append(f"<span style='color:{acc}'>{html.escape(v).replace(chr(10), '<br>')}</span>")
                else:
                    out.append(f"<span style='background:#f6d7c3;color:#7a2e0e'>&nbsp;{html.escape(name)}&nbsp;</span>")
                pos = m.end()
            out.append(html.escape(text[pos:]))
            return "".join(out)

        parts = ["<div style='font-family:Cambria, \"Times New Roman\"; font-size:12pt; color:#222;'>"]
        for kind, text in paras:
            if kind == "row":
                cells = text.split("\t")
                parts.append("<table width='100%' cellpadding='3' style='border-collapse:collapse'><tr>" +
                             "".join(f"<td style='border:1px solid #bbb'>{render(c)}</td>" for c in cells) +
                             "</tr></table>")
                continue
            align = {"center": "center", "right": "right", "both": "justify"}.get(kind, "left")
            parts.append(f"<p style='margin:0 0 4px 0; text-align:{align}'>{render(text) or '&nbsp;'}</p>")
        parts.append("</div>")
        self.preview.setHtml("".join(parts))

    # -- yaratish
    def make(self):
        if not self.template:
            return
        empty = [ph for ph in self.placeholders if not self.values.get(ph)]
        if empty and not ask(self, tr("ask_empty", n=len(empty), names=", ".join(empty[:6]))):
            return
        outdir = Path(self.outdir.path() or documents_dir() / "Arizalar")
        name_vals = dict(self.values)
        pid = self.profile.currentData()
        prof = self.st.store.get(pid) if pid else None
        name_vals.update({k: v for k, v in fields.values_for(prof["fields"] if prof else {}).items() if v})
        name = docx_fill.render_name(self.fname.text() or self.tpl_title, name_vals)
        try:
            outdir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.result.setText(f"<span style='color:#b42318'>{err_text(e)}</span>")
            return
        dst = docx_fill.unique_path(outdir / f"{name}.docx")
        want_pdf = self.pdf.isChecked() and self.pdf.isEnabled()
        if not is_demo():
            self.win.settings.set("out_dir", str(outdir))
            self.win.settings.set("also_pdf", want_pdf)
            rem = self.win.settings.get("extra_values", {})
            for ph, v in self.values.items():
                if v and not fields.match_field(ph):
                    rem[ph] = v
            self.win.settings.set("extra_values", rem)

        def job(tpl, dst, vals, want_pdf):
            docx_fill.fill(tpl, dst, vals)
            files = [dst]
            pdf_err = None
            if want_pdf:
                try:
                    files += convert.to_pdf_many([str(dst)])
                except Exception as e:
                    pdf_err = err_text(e)
            return files, pdf_err

        self.make_btn.setEnabled(False)
        self.result.setText(tr("working"))
        w = Worker(job, self.template, dst, dict(self.values), want_pdf)
        w.signals.result.connect(self._made)
        w.signals.error.connect(lambda t: self._failed(t))
        run_worker(w)

    def _made(self, res):
        files, pdf_err = res
        self.make_btn.setEnabled(True)
        self._last_docs = [Path(f) for f in files]
        names = ", ".join(Path(f).name for f in files)
        msg = tr("created", names=names)
        if pdf_err:
            msg += f"<br><span style='color:#b42318'>{pdf_err}</span>"
        self.result.setText(msg)
        self.open_doc.show()
        self.open_dir.show()
        self.win.status(tr("done"))

    def _failed(self, text):
        self.make_btn.setEnabled(True)
        self.result.setText(f"<span style='color:#b42318'>{tr('err_generic', detail=text)}</span>")


# =========================================================================== 2. Profillar
GROUPS = [
    ("grp_personal", ["FAMILIYA", "ISM", "OTASINING_ISMI", "TUGILGAN_SANA", "TUGILGAN_JOY", "MILLATI"]),
    ("grp_contact", ["MANZIL", "TELEFON", "EMAIL"]),
    ("grp_docs", ["PASPORT", "PASPORT_BERILGAN", "JSHSHIR"]),
    ("grp_work", ["TASHKILOT", "BOLIM", "LAVOZIM", "GURUH", "MALUMOTI", "TAMOMLAGAN", "MUTAXASSISLIGI",
                  "ILMIY_DARAJA", "TILLAR"]),
    ("grp_boss", ["RAHBAR_LAVOZIM", "RAHBAR_FISH"]),
]


class ProfilesPage(QWidget):
    changed = Signal()

    def __init__(self, win):
        super().__init__()
        self.win, self.st = win, win.state
        self.current = None
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(12)

        left = Card(tr("people"))
        left.setFixedWidth(250)
        self.list = QListWidget()
        self.list.setObjectName("PeopleList")
        left.add(self.list)
        self.add_btn = QPushButton(tr("add"))
        left.add(self.add_btn)
        r = QHBoxLayout()
        self.dup_btn = QPushButton(tr("duplicate"))
        self.del_btn = QPushButton(tr("delete"))
        for b in (self.dup_btn, self.del_btn):
            r.addWidget(b)
        left.add(r)
        self.lock_lbl = muted("")
        left.add(self.lock_lbl)
        self.pw_btn = QPushButton(tr("protect_store"))
        left.add(self.pw_btn)
        lay.addWidget(left)

        self.stack = QStackedWidget()
        # qulflangan holat
        lockw = Card(tr("store_locked_title"))
        lockw.add(muted(tr("store_locked_text")))
        lr = QHBoxLayout()
        self.unlock_pw = QLineEdit()
        self.unlock_pw.setEchoMode(QLineEdit.Password)
        self.unlock_pw.setPlaceholderText(tr("password"))
        self.unlock_btn = primary_button(tr("unlock"))
        lr.addWidget(self.unlock_pw, 1)
        lr.addWidget(self.unlock_btn)
        lockw.add(lr)
        self.unlock_msg = QLabel("")
        lockw.add(self.unlock_msg)
        lockw.lay.addStretch(1)
        self.stack.addWidget(lockw)
        # bo'sh holat
        empty = Card()
        empty.add(muted(tr("profiles_empty")))
        empty.lay.addStretch(1)
        self.stack.addWidget(empty)
        # forma
        formw = QWidget()
        fl = QVBoxLayout(formw)
        fl.setContentsMargins(0, 0, 0, 0)
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        inner = QWidget()
        il = QVBoxLayout(inner)
        il.setContentsMargins(0, 0, 8, 0)
        il.setSpacing(10)
        tcard = Card()
        tr_row = QHBoxLayout()
        tr_row.addWidget(QLabel(tr("profile_title")))
        self.ptitle = QLineEdit()
        self.ptitle.setPlaceholderText(tr("profile_title_ph"))
        tr_row.addWidget(self.ptitle, 1)
        tcard.add(tr_row)
        il.addWidget(tcard)
        self.edits: Dict[str, QLineEdit] = {}
        for gkey, keys in GROUPS:
            c = Card(tr(gkey))
            g = QGridLayout()
            g.setHorizontalSpacing(12)
            g.setVerticalSpacing(6)
            for i, k in enumerate(keys):
                e = QLineEdit()
                self.edits[k] = e
                if k in fields.MASKED:
                    e.setEchoMode(QLineEdit.Password)
                cell = QVBoxLayout()
                cell.setSpacing(2)
                lab = QLabel(fields.label(k, language()))
                lab.setObjectName("FieldLabel")
                cell.addWidget(lab)
                cell.addWidget(e)
                g.addLayout(cell, i // 2, i % 2)
            g.setColumnStretch(0, 1)
            g.setColumnStretch(1, 1)
            c.add(g)
            if gkey == "grp_docs":
                self.reveal = QCheckBox(tr("show_secret"))
                c.add(self.reveal)
                self.reveal.toggled.connect(self._reveal)
                self.doc_warn = QLabel("")
                self.doc_warn.setObjectName("Warn")
                self.doc_warn.setWordWrap(True)
                c.add(self.doc_warn)
            il.addWidget(c)
        il.addStretch(1)
        sc.setWidget(inner)
        fl.addWidget(sc, 1)
        srow = QHBoxLayout()
        self.saved_lbl = muted("")
        self.save_btn = primary_button(tr("save"))
        self.save_btn.setMinimumWidth(150)
        srow.addWidget(self.saved_lbl, 1)
        srow.addWidget(self.save_btn)
        fl.addLayout(srow)
        self.stack.addWidget(formw)
        lay.addWidget(self.stack, 1)

        self.list.currentItemChanged.connect(self._select)
        self.add_btn.clicked.connect(self._add)
        self.dup_btn.clicked.connect(self._dup)
        self.del_btn.clicked.connect(self._del)
        self.save_btn.clicked.connect(self._save)
        self.unlock_btn.clicked.connect(self._unlock)
        self.unlock_pw.returnPressed.connect(self._unlock)
        self.pw_btn.clicked.connect(self._password_menu)
        for k in ("PASPORT", "JSHSHIR"):
            self.edits[k].textChanged.connect(self._validate)
        self.reload()

    def reload(self):
        locked = self.st.locked
        for b in (self.add_btn, self.dup_btn, self.del_btn):
            b.setEnabled(not locked)
        self.lock_lbl.setText(tr("store_encrypted") if (self.st.store.password or locked) else tr("store_plain"))
        self.pw_btn.setEnabled(not locked)
        self.list.blockSignals(True)
        self.list.clear()
        if not locked:
            for p in self.st.store.profiles:
                it = QListWidgetItem(p["title"] or fields.full_name(p["fields"]) or "?")
                it.setData(Qt.UserRole, p["id"])
                sub = fields.full_name(p["fields"])
                if sub:
                    it.setToolTip(sub)
                self.list.addItem(it)
        self.list.blockSignals(False)
        if locked:
            self.stack.setCurrentIndex(0)
        elif not self.st.store.profiles:
            self.stack.setCurrentIndex(1)
        else:
            self.stack.setCurrentIndex(2)
            row = 0
            if self.current:
                for i in range(self.list.count()):
                    if self.list.item(i).data(Qt.UserRole) == self.current:
                        row = i
            self.list.setCurrentRow(row)
            self._select(self.list.currentItem())

    def _select(self, item, _prev=None):
        if item is None:
            return
        self.current = item.data(Qt.UserRole)
        p = self.st.store.get(self.current)
        if not p:
            return
        self.ptitle.setText(p["title"])
        for k, e in self.edits.items():
            e.setText(p["fields"].get(k, ""))
        self.saved_lbl.setText("")
        self.stack.setCurrentIndex(2)

    def _reveal(self, on):
        for k in fields.MASKED:
            self.edits[k].setEchoMode(QLineEdit.Normal if on else QLineEdit.Password)

    def _validate(self):
        msgs = []
        if not fields.validate_passport(self.edits["PASPORT"].text()):
            msgs.append(tr("warn_passport"))
        if not fields.validate_jshshir(self.edits["JSHSHIR"].text()):
            msgs.append(tr("warn_jshshir"))
        self.doc_warn.setText("  ".join(msgs))

    def _add(self):
        p = self.st.store.add(tr("new_profile"))
        self.current = p["id"]
        self.st.save_profiles()
        self.reload()
        self.ptitle.setFocus()
        self.ptitle.selectAll()
        self.changed.emit()

    def _dup(self):
        if not self.current:
            return
        p = self.st.store.duplicate(self.current, " (2)")
        if p:
            self.current = p["id"]
            self.st.save_profiles()
            self.reload()
            self.changed.emit()

    def _del(self):
        p = self.st.store.get(self.current) if self.current else None
        if not p or not ask(self, tr("ask_delete_profile", name=p["title"])):
            return
        self.st.store.remove(self.current)
        self.current = None
        self.st.save_profiles()
        self.reload()
        self.changed.emit()

    def _save(self):
        p = self.st.store.get(self.current) if self.current else None
        if not p:
            return
        p["title"] = self.ptitle.text().strip() or fields.full_name({k: e.text() for k, e in self.edits.items()})
        p["fields"] = {k: e.text().strip() for k, e in self.edits.items() if e.text().strip()}
        try:
            self.st.save_profiles()
        except Exception as e:
            warn(self, err_text(e))
            return
        self.saved_lbl.setText(tr("profile_saved"))
        cur = self.list.currentItem()
        if cur:
            cur.setText(p["title"] or "?")
        self.changed.emit()

    def _unlock(self):
        try:
            self.st.store.load(self.unlock_pw.text())
        except WrongPassword:
            self.unlock_msg.setText(f"<span style='color:#b42318'>{tr('wrong_password')}</span>")
            return
        except Exception as e:
            self.unlock_msg.setText(err_text(e))
            return
        self.st.locked = False
        self.unlock_pw.clear()
        self.unlock_msg.setText("")
        self.reload()
        self.changed.emit()

    def _password_menu(self):
        m = QMenu(self)
        a_set = m.addAction(tr("pw_change") if self.st.store.password else tr("pw_set"))
        a_rem = m.addAction(tr("pw_remove"))
        a_rem.setEnabled(bool(self.st.store.password))
        act = m.exec(self.pw_btn.mapToGlobal(self.pw_btn.rect().bottomLeft()))
        if act is a_set:
            p1, ok = QInputDialog.getText(self, tr("protect_store"), tr("pw_new"), QLineEdit.Password)
            if not ok or not p1:
                return
            p2, ok = QInputDialog.getText(self, tr("protect_store"), tr("pw_repeat"), QLineEdit.Password)
            if not ok:
                return
            if p1 != p2:
                warn(self, tr("pw_mismatch"))
                return
            if not is_demo():
                self.st.store.set_password(p1)
            else:
                self.st.store.password = p1
            info(self, tr("pw_set_done"))
        elif act is a_rem:
            if ask(self, tr("pw_remove_ask")):
                if not is_demo():
                    self.st.store.set_password(None)
                else:
                    self.st.store.password = None
        self.reload()


# =========================================================================== 3. Shablonlar
class TemplatesPage(QWidget):
    changed = Signal()

    def __init__(self, win):
        super().__init__()
        self.win, self.st = win, win.state
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(12)
        left = QVBoxLayout()
        self.table = make_table([tr("col_name"), tr("col_category"), tr("col_source"), tr("col_fields")], 0)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        left.addWidget(self.table, 1)
        r = QHBoxLayout()
        self.add_btn = QPushButton(tr("add_template"))
        self.word_btn = QPushButton(tr("open_in_word"))
        self.copy_btn = QPushButton(tr("make_copy"))
        self.ren_btn = QPushButton(tr("rename"))
        self.del_btn = QPushButton(tr("remove"))
        for b in (self.add_btn, self.word_btn, self.copy_btn, self.ren_btn, self.del_btn):
            r.addWidget(b)
        r.addStretch(1)
        self.use_btn = primary_button(tr("use_template"))
        r.addWidget(self.use_btn)
        left.addLayout(r)
        lay.addLayout(left, 1)

        help_card = Card(tr("how_to_make"))
        help_card.setFixedWidth(330)
        help_card.add(muted(tr("how_to_text")))
        self.ref = make_table([tr("col_placeholder"), tr("col_meaning")], 1)
        self.ref.setEditTriggers(QAbstractItemView.NoEditTriggers)
        keys = fields.all_keys()
        self.ref.setRowCount(len(keys))
        for i, k in enumerate(keys):
            self.ref.setItem(i, 0, QTableWidgetItem("{{" + k + "}}"))
            self.ref.setItem(i, 1, QTableWidgetItem(fields.label(k, language())))
        help_card.add(self.ref)
        help_card.add(muted(tr("copy_hint")))
        lay.addWidget(help_card)

        self.add_btn.clicked.connect(self._add)
        self.word_btn.clicked.connect(self._word)
        self.copy_btn.clicked.connect(self._copy)
        self.ren_btn.clicked.connect(self._rename)
        self.del_btn.clicked.connect(self._remove)
        self.use_btn.clicked.connect(self._use)
        self.table.itemSelectionChanged.connect(self._sel)
        self.table.doubleClicked.connect(lambda _: self._use())
        self.ref.doubleClicked.connect(self._copy_ph)
        self.reload()

    def _count(self, path):
        try:
            return str(len(docx_fill.scan(path)))
        except Exception:
            return "?"

    def reload(self):
        items = self.st.library.items
        self.table.setRowCount(len(items))
        for r, it in enumerate(items):
            vals = [title_of(it), tr("cat_" + it["category"]), tr("src_builtin") if it["builtin"] else tr("src_user"),
                    self._count(it["path"])]
            for c, v in enumerate(vals):
                cell = QTableWidgetItem(v)
                cell.setData(Qt.UserRole, it["id"])
                self.table.setItem(r, c, cell)
        self._sel()

    def _cur(self):
        r = self.table.currentRow()
        if r < 0 or not self.table.item(r, 0):
            return None
        return self.st.library.get(self.table.item(r, 0).data(Qt.UserRole))

    def _sel(self):
        it = self._cur()
        for b in (self.word_btn, self.copy_btn, self.use_btn):
            b.setEnabled(it is not None)
        user = bool(it) and not it["builtin"]
        self.ren_btn.setEnabled(user)
        self.del_btn.setEnabled(user)
        self.copy_btn.setEnabled(bool(it) and it["builtin"])

    def _add(self):
        ps, _ = QFileDialog.getOpenFileNames(self, tr("select_file"), "", DOCX_FILTER)
        for p in ps:
            try:
                docx_fill.scan(p)
            except Exception as e:
                warn(self, err_text(e))
                continue
            self.st.library.add(p, Path(p).stem, "boshqa")
        if ps:
            self.reload()
            self.changed.emit()

    def _word(self):
        it = self._cur()
        if not it:
            return
        path = it["path"]
        if it["builtin"]:
            tmp = Path(tempfile.gettempdir()) / "shablon_korish"
            tmp.mkdir(exist_ok=True)
            dst = docx_fill.unique_path(tmp / Path(path).name)
            import shutil
            shutil.copyfile(path, dst)
            path = str(dst)
        open_path(path)

    def _copy(self):
        it = self._cur()
        if it and it["builtin"]:
            self.st.library.copy_builtin(it["id"], " (" + tr("copy_suffix") + ")")
            self.reload()
            self.changed.emit()

    def _rename(self):
        it = self._cur()
        if not it or it["builtin"]:
            return
        name, ok = QInputDialog.getText(self, tr("rename"), tr("col_name"), text=title_of(it))
        if ok and name.strip():
            cats = [tr("cat_" + c) for c in CATEGORIES]
            cat, ok2 = QInputDialog.getItem(self, tr("rename"), tr("col_category"), cats,
                                            CATEGORIES.index(it["category"]) if it["category"] in CATEGORIES else 3,
                                            False)
            self.st.library.rename(it["id"], name.strip(), CATEGORIES[cats.index(cat)] if ok2 else None)
            self.reload()
            self.changed.emit()

    def _remove(self):
        it = self._cur()
        if it and not it["builtin"] and ask(self, tr("ask_remove_template", name=title_of(it))):
            self.st.library.remove(it["id"])
            self.reload()
            self.changed.emit()

    def _use(self):
        it = self._cur()
        if it:
            self.win.use_template(it)

    def _copy_ph(self, index):
        from PySide6.QtGui import QGuiApplication
        txt = self.ref.item(index.row(), 0).text()
        QGuiApplication.clipboard().setText(txt)
        self.win.status(tr("copied") + ": " + txt)


# =========================================================================== 4. Excel'dan ommaviy
class BatchPage(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win, self.st = win, win.state
        self.headers: List[str] = []
        self.rows: List[Dict[str, str]] = []
        self.placeholders: List[str] = []
        self.worker = None
        self.out_files: List[Path] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)
        top = Card()
        g = QGridLayout()
        g.setHorizontalSpacing(10)
        self.tpl = QComboBox()
        self.xls = PathPicker("open", "Excel / CSV (*.xlsx *.xlsm *.csv)", tr("excel_ph"))
        self.sheet = QComboBox()
        self.sheet.setMinimumWidth(120)
        g.addWidget(QLabel(tr("template")), 0, 0)
        g.addWidget(self.tpl, 0, 1, 1, 3)
        g.addWidget(QLabel(tr("excel_file")), 1, 0)
        g.addWidget(self.xls, 1, 1)
        g.addWidget(QLabel(tr("sheet")), 1, 2)
        g.addWidget(self.sheet, 1, 3)
        g.setColumnStretch(1, 1)
        top.add(g)
        top.add(muted(tr("batch_help")))
        lay.addWidget(top)

        mid = QHBoxLayout()
        dc = Card(tr("data_rows"))
        self.data = QTableWidget(0, 0)
        self.data.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.data.setAlternatingRowColors(True)
        dc.add(self.data)
        mid.addWidget(dc, 3)
        mc = Card(tr("mapping"))
        self.map = make_table([tr("col_placeholder"), tr("col_column"), tr("col_default")], 2)
        mc.add(self.map)
        mid.addWidget(mc, 2)
        lay.addLayout(mid, 1)

        bc = Card()
        g2 = QGridLayout()
        self.pattern = QLineEdit("{FISH}")
        self.outdir = PathPicker("dir")
        self.outdir.set_path(str(documents_dir() / "Arizalar" / "Ommaviy"))
        self.pdf = QCheckBox(tr("also_pdf"))
        if not convert.pdf_available():
            self.pdf.setEnabled(False)
            self.pdf.setToolTip(tr("pdf_unavailable"))
        g2.addWidget(QLabel(tr("name_pattern")), 0, 0)
        g2.addWidget(self.pattern, 0, 1)
        g2.addWidget(QLabel(tr("save_folder")), 1, 0)
        g2.addWidget(self.outdir, 1, 1)
        g2.setColumnStretch(1, 1)
        bc.add(g2)
        r = QHBoxLayout()
        self.prog = QProgressBar()
        self.prog.setFixedWidth(180)
        self.prog.setTextVisible(False)
        self.prog.hide()
        self.msg = QLabel("")
        self.msg.setWordWrap(True)
        self.open_dir = QPushButton(tr("open_folder"))
        self.open_dir.hide()
        self.cancel_btn = QPushButton(tr("cancel"))
        self.cancel_btn.hide()
        self.run_btn = primary_button(tr("make_all"))
        self.run_btn.setMinimumWidth(170)
        r.addWidget(self.pdf)
        r.addWidget(self.msg, 1)
        r.addWidget(self.prog)
        r.addWidget(self.open_dir)
        r.addWidget(self.cancel_btn)
        r.addWidget(self.run_btn)
        bc.add(r)
        lay.addWidget(bc)

        self.tpl.currentIndexChanged.connect(self._tpl_changed)
        self.xls.changed.connect(self._load_excel)
        self.sheet.currentIndexChanged.connect(lambda _: self._load_excel(self.xls.path(), keep_sheet=True))
        self.run_btn.clicked.connect(self._run)
        self.cancel_btn.clicked.connect(lambda: self.worker and self.worker.cancel())
        self.open_dir.clicked.connect(lambda: self.out_files and open_path(str(self.out_files[0].parent)))
        self.fill_templates()

    def fill_templates(self):
        cur = self.tpl.currentData()
        self.tpl.blockSignals(True)
        self.tpl.clear()
        for it in self.st.library.items:
            self.tpl.addItem(title_of(it), it["path"])
        i = self.tpl.findData(cur)
        self.tpl.setCurrentIndex(max(i, 0))
        self.tpl.blockSignals(False)
        self._tpl_changed()

    def _tpl_changed(self):
        p = self.tpl.currentData()
        try:
            self.placeholders = docx_fill.scan(p) if p else []
        except Exception:
            self.placeholders = []
        self._fill_mapping()

    def _load_excel(self, path, keep_sheet=False):
        if not path or not Path(path).is_file():
            return
        sheet = self.sheet.currentText() if keep_sheet else None
        try:
            self.headers, self.rows, sheets = batch.read_table(path, sheet)
        except Exception as e:
            self.msg.setText(f"<span style='color:#b42318'>{err_text(e)}</span>")
            self.headers, self.rows = [], []
            self._fill_data()
            return
        if not keep_sheet:
            self.sheet.blockSignals(True)
            self.sheet.clear()
            self.sheet.addItems(sheets)
            self.sheet.blockSignals(False)
        self.msg.setText(tr("n_rows", n=len(self.rows)))
        self._fill_data()
        self._fill_mapping()

    def _fill_data(self):
        self.data.setColumnCount(len(self.headers))
        self.data.setHorizontalHeaderLabels(self.headers)
        shown = self.rows[:200]
        self.data.setRowCount(len(shown))
        for r, row in enumerate(shown):
            for c, h in enumerate(self.headers):
                self.data.setItem(r, c, QTableWidgetItem(row.get(h, "")))
        self.data.resizeColumnsToContents()

    def _fill_mapping(self):
        mapping = batch.map_columns(self.placeholders, self.headers) if self.headers else {}
        self.map.clearContents()
        self.map.setRowCount(0)
        self.map.setRowCount(len(self.placeholders))
        for r, ph in enumerate(self.placeholders):
            a = QTableWidgetItem("{{" + ph + "}}")
            a.setFlags(a.flags() & ~Qt.ItemIsEditable)
            self.map.setItem(r, 0, a)
            combo = QComboBox()
            k = fields.match_field(ph)
            combo.addItem(tr("col_auto") if k in fields.COMPUTED else tr("col_none"), None)
            for h in self.headers:
                combo.addItem(h, h)
            col = mapping.get(ph)
            if col:
                combo.setCurrentIndex(combo.findData(col))
            self.map.setCellWidget(r, 1, combo)
            d = QTableWidgetItem("")
            if k in fields.COMPUTED:
                d.setToolTip(tr("auto_value"))
            self.map.setItem(r, 2, d)

    def _run(self):
        tpl = self.tpl.currentData()
        if not tpl or not self.rows:
            self.msg.setText(f"<span style='color:#b42318'>{tr('need_excel')}</span>")
            return
        mapping, defaults = {}, {}
        for r, ph in enumerate(self.placeholders):
            w = self.map.cellWidget(r, 1)
            mapping[ph] = w.currentData() if w else None
            it = self.map.item(r, 2)
            if it and it.text().strip():
                defaults[ph] = it.text().strip()
        # Ustun tanlovi foydalanuvchiniki: qatorlarni o'rinbosar nomlari bilan qayta nomlaymiz
        rows = []
        for row in self.rows:
            nr = dict(row)
            for ph, col in mapping.items():
                nr[ph] = row.get(col, "") if col else ""
            rows.append(nr)
        outdir = self.outdir.path() or str(documents_dir() / "Arizalar" / "Ommaviy")
        pattern = self.pattern.text().strip() or "{FISH}"
        want_pdf = self.pdf.isChecked() and self.pdf.isEnabled()

        def job(progress, tpl, rows, outdir, pattern, defaults, want_pdf):
            files = batch.run(tpl, rows, outdir, pattern, defaults, progress)
            pdf_err = None
            if want_pdf and files and not progress.cancelled():
                try:
                    convert.to_pdf_many([str(f) for f in files])
                except Exception as e:
                    pdf_err = err_text(e)
            return files, pdf_err

        self.prog.show()
        self.prog.setValue(0)
        self.cancel_btn.show()
        self.run_btn.setEnabled(False)
        self.open_dir.hide()
        w = Worker(job, tpl, rows, outdir, pattern, defaults, want_pdf)
        self.worker = w
        w.signals.progress.connect(lambda p, m: (self.prog.setValue(p), self.msg.setText(f"{tr('working')} {m}")))
        w.signals.result.connect(self._done)
        w.signals.error.connect(self._fail)
        run_worker(w)

    def _end(self):
        self.worker = None
        self.prog.hide()
        self.cancel_btn.hide()
        self.run_btn.setEnabled(True)

    def _done(self, res):
        self._end()
        files, pdf_err = res
        self.out_files = [Path(f) for f in files]
        msg = tr("batch_done", n=len(files))
        if pdf_err:
            msg += f"  <span style='color:#b42318'>{pdf_err}</span>"
        self.msg.setText(msg)
        self.open_dir.setVisible(bool(files))

    def _fail(self, text):
        self._end()
        self.msg.setText(f"<span style='color:#b42318'>{tr('err_generic', detail=text)}</span>")

    def demo(self):
        import openpyxl
        p = Path(tempfile.gettempdir()) / "shablon_demo_talabalar.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Talabalar"
        ws.append(["Familiya", "Ism", "Otasining ismi", "Guruh", "Telefon"])
        for row in [("Aliyev", "Sardor", "Karimovich", "911-23", "+998 90 111-22-33"),
                    ("Rahimova", "Madina", "Olimovna", "911-23", "+998 93 444-55-66"),
                    ("Tursunov", "Bekzod", "Anvarovich", "912-23", "+998 97 777-88-99"),
                    ("Qodirova", "Nilufar", "Shavkatovna", "912-23", "+998 99 123-00-11")]:
            ws.append(row)
        wb.save(p)
        i = self.tpl.findText([title_of(t) for t in self.st.library.items if "talaba" in t["path"]][0]
                              if any("talaba" in t["path"] for t in self.st.library.items) else "")
        if i >= 0:
            self.tpl.setCurrentIndex(i)
        self.xls.set_path(str(p))


# =========================================================================== Asosiy oyna
class MainWindow(BaseMainWindow):
    def __init__(self):
        super().__init__()
        self.state = AppState(self)
        root = QWidget()
        v = QVBoxLayout(root)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        bar = QFrame()
        bar.setObjectName("TopBar")
        h = QHBoxLayout(bar)
        h.setContentsMargins(20, 10, 20, 0)
        h.setSpacing(18)
        brand = QLabel(appmod.app_name())
        brand.setObjectName("Brand")
        h.addWidget(brand)
        self._nav = QListWidget()
        self._nav.setObjectName("TopNav")
        self._nav.setFlow(QListWidget.LeftToRight)
        self._nav.setFixedHeight(44)
        self._nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._nav.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        h.addWidget(self._nav, 1)
        v.addWidget(bar)
        self._stack = QStackedWidget()
        holder = QWidget()
        hl = QVBoxLayout(holder)
        hl.setContentsMargins(20, 14, 20, 14)
        hl.addWidget(self._stack)
        v.addWidget(holder, 1)
        self.setCentralWidget(root)

        self.fill = FillPage(self)
        self.profiles = ProfilesPage(self)
        self.templates = TemplatesPage(self)
        self.batch = BatchPage(self)
        for key, w in (("nav_fill", self.fill), ("nav_profiles", self.profiles), ("nav_templates", self.templates),
                       ("nav_batch", self.batch)):
            self._stack.addWidget(w)
            it = QListWidgetItem(tr(key))
            it.setTextAlignment(Qt.AlignCenter)
            f = QFont(self._nav.font())
            f.setBold(True)
            f.setPointSizeF(max(f.pointSizeF(), 10) * 1.1)
            from PySide6.QtGui import QFontMetrics
            it.setSizeHint(QSize(QFontMetrics(f).horizontalAdvance(tr(key)) + 44, 40))
            self._nav.addItem(it)
        self._nav.currentRowChanged.connect(self._stack.setCurrentIndex)
        self._nav.setCurrentRow(0)
        self.profiles.changed.connect(self.fill.fill_profiles)
        self.templates.changed.connect(self._templates_changed)
        if is_demo():
            self._demo()

    def palette_color(self, key):
        from uzkit.theme import PALETTES
        dark = self.palette().window().color().lightness() < 128
        mode = "dark" if dark else "light"
        c = dict(PALETTES[mode])
        c.update((appmod.META.get("palette", {}) or {}).get(mode, {}))
        return c[key]

    def on_theme_changed(self):
        self.fill.refresh_table()

    def _templates_changed(self):
        self.state.library.reload()
        self.fill.fill_templates()
        self.batch.fill_templates()

    def use_template(self, item):
        self.set_page(0)
        for i in range(self.fill.tlist.count()):
            li = self.fill.tlist.item(i)
            if li.data(Qt.UserRole) == item["path"]:
                self.fill.tlist.setCurrentItem(li)
                return
        self.fill.set_template(item["path"], title_of(item))

    def _demo(self):
        self.fill.fill_profiles()
        self.profiles.reload()
        for i in range(self.fill.tlist.count()):
            li = self.fill.tlist.item(i)
            if li.data(Qt.UserRole) and "tatil" in li.data(Qt.UserRole):
                self.fill.tlist.setCurrentItem(li)
                break
        demo_vals = {"ISH_YILI": "2025-2026", "TATIL_BOSHLANISHI": "01.07.2026", "TATIL_TUGASHI": "18.08.2026",
                     "TATIL_KUNLARI": "48"}
        self.fill.values.update(demo_vals)
        self.fill.refresh_table()
        try:
            self.batch.demo()
        except Exception:
            pass
