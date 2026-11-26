"""Profillar ombori: data_dir/profiles.json. Ixtiyoriy parol bilan shifrlash (Fernet, PBKDF2-SHA256)."""
import base64
import json
import os
import tempfile
import uuid
from pathlib import Path
from typing import Dict, List, Optional

ITERATIONS = 390_000


class StoreLocked(Exception):
    """Ombor shifrlangan, parol kerak."""


class WrongPassword(Exception):
    pass


def _key(password: str, salt: bytes) -> bytes:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS)
    return base64.urlsafe_b64encode(kdf.derive(password.encode("utf-8")))


def new_profile(title: str = "") -> Dict:
    return {"id": uuid.uuid4().hex[:12], "title": title, "fields": {}}


class ProfileStore:
    def __init__(self, path):
        self.path = Path(path)
        self.profiles: List[Dict] = []
        self.password: Optional[str] = None
        self.loaded = False

    # -- holat
    def exists(self) -> bool:
        return self.path.exists()

    def is_encrypted(self) -> bool:
        try:
            d = json.loads(self.path.read_text(encoding="utf-8"))
            return bool(d.get("encrypted"))
        except Exception:
            return False

    # -- o'qish / yozish
    def load(self, password: str = None) -> List[Dict]:
        if not self.path.exists():
            self.profiles, self.password, self.loaded = [], password, True
            return self.profiles
        d = json.loads(self.path.read_text(encoding="utf-8"))
        if d.get("encrypted"):
            if not password:
                raise StoreLocked()
            from cryptography.fernet import Fernet, InvalidToken
            salt = base64.b64decode(d["salt"])
            try:
                raw = Fernet(_key(password, salt)).decrypt(d["data"].encode("ascii"))
            except InvalidToken:
                raise WrongPassword()
            self.profiles = json.loads(raw.decode("utf-8"))
            self.password = password
        else:
            self.profiles = d.get("profiles", [])
            self.password = None
        self.loaded = True
        return self.profiles

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.password:
            from cryptography.fernet import Fernet
            salt = os.urandom(16)
            token = Fernet(_key(self.password, salt)).encrypt(
                json.dumps(self.profiles, ensure_ascii=False).encode("utf-8"))
            d = {"version": 1, "encrypted": True, "salt": base64.b64encode(salt).decode("ascii"),
                 "data": token.decode("ascii")}
        else:
            d = {"version": 1, "encrypted": False, "profiles": self.profiles}
        fd, tmp = tempfile.mkstemp(dir=str(self.path.parent), suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.path)

    def set_password(self, password: Optional[str]):
        """Parol qo'yish / o'zgartirish (None yoki "" — olib tashlash). Darhol saqlaydi."""
        self.password = password or None
        self.save()

    # -- profillar
    def get(self, pid) -> Optional[Dict]:
        for p in self.profiles:
            if p["id"] == pid:
                return p
        return None

    def add(self, title: str) -> Dict:
        p = new_profile(title)
        self.profiles.append(p)
        return p

    def duplicate(self, pid, suffix=" (2)") -> Optional[Dict]:
        src = self.get(pid)
        if not src:
            return None
        p = new_profile(src["title"] + suffix)
        p["fields"] = dict(src["fields"])
        self.profiles.append(p)
        return p

    def remove(self, pid):
        self.profiles = [p for p in self.profiles if p["id"] != pid]
