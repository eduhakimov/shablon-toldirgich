"""Shablonlar kutubxonasi: dastur bilan kelgan (o'zgarmas) va foydalanuvchi qo'shgan shablonlar."""
import json
import shutil
import uuid
from pathlib import Path
from typing import Dict, List, Optional

from uzkit.paths import resource

CATEGORIES = ["ariza", "xat", "shaxsiy", "boshqa"]


class Library:
    def __init__(self, user_dir):
        self.user_dir = Path(user_dir)
        self.index_path = self.user_dir / "index.json"
        self.items: List[Dict] = []
        self.reload()

    def reload(self):
        items = []
        bdir = resource("assets", "templates")
        try:
            idx = json.loads((bdir / "index.json").read_text(encoding="utf-8"))
        except Exception:
            idx = []
        for e in idx:
            p = bdir / e["file"]
            if p.exists():
                items.append({"id": "b:" + e["file"], "path": str(p), "title": e["title"],
                              "category": e.get("category", "boshqa"), "builtin": True})
        for e in self._user_index():
            p = self.user_dir / e["file"]
            if p.exists():
                items.append({"id": "u:" + e["id"], "path": str(p), "title": e["title"],
                              "category": e.get("category", "boshqa"), "builtin": False})
        self.items = items

    def _user_index(self) -> List[Dict]:
        try:
            return json.loads(self.index_path.read_text(encoding="utf-8"))
        except Exception:
            return []

    def _save_user_index(self, idx):
        self.user_dir.mkdir(parents=True, exist_ok=True)
        self.index_path.write_text(json.dumps(idx, ensure_ascii=False, indent=1), encoding="utf-8")

    def get(self, tid) -> Optional[Dict]:
        for it in self.items:
            if it["id"] == tid:
                return it
        return None

    def add(self, src, title: str = None, category: str = "boshqa") -> Dict:
        """Faylni kutubxonaga nusxalaydi (asl fayl joyida qoladi)."""
        src = Path(src)
        self.user_dir.mkdir(parents=True, exist_ok=True)
        uid = uuid.uuid4().hex[:10]
        fname = f"{uid}_{src.name}"
        shutil.copyfile(src, self.user_dir / fname)
        t = title or src.stem
        idx = self._user_index()
        idx.append({"id": uid, "file": fname, "title": {"uz": t, "ru": t, "en": t}, "category": category})
        self._save_user_index(idx)
        self.reload()
        return self.get("u:" + uid)

    def rename(self, tid, title: str, category: str = None):
        if not tid.startswith("u:"):
            return
        idx = self._user_index()
        for e in idx:
            if e["id"] == tid[2:]:
                e["title"] = {"uz": title, "ru": title, "en": title}
                if category:
                    e["category"] = category
        self._save_user_index(idx)
        self.reload()

    def remove(self, tid):
        """Faqat foydalanuvchi shablonini kutubxonadan olib tashlaydi (nusxani o'chiradi)."""
        if not tid.startswith("u:"):
            return False
        idx = self._user_index()
        keep = []
        for e in idx:
            if e["id"] == tid[2:]:
                try:
                    (self.user_dir / e["file"]).unlink()
                except Exception:
                    pass
            else:
                keep.append(e)
        self._save_user_index(keep)
        self.reload()
        return True

    def copy_builtin(self, tid, suffix: str = " (2)") -> Optional[Dict]:
        """Tayyor shablondan tahrirlanadigan nusxa yaratadi."""
        it = self.get(tid)
        if not it:
            return None
        return self.add(it["path"], it["title"].get("uz", "") + suffix, it["category"])
