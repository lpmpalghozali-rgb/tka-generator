"""Penyimpanan lokal paket soal sebagai berkas JSON (tanpa dependensi Streamlit agar mudah diuji).

Folder bawaan: `saved_packages/` di samping app.py (dapat diganti lewat env `TKA_SAVE_DIR`).
Berkas yang sama dapat diunduh/diimpor lewat dumps_package / loads_package.
"""
import json
import os
import re
from datetime import datetime
from pathlib import Path

from core import QUESTION_TYPES

FORMAT_NAME = "tka-generator-package"
FORMAT_VERSION = 1
MAX_IMPORT_BYTES = 10 * 1024 * 1024
DEFAULT_DIR = Path(os.getenv("TKA_SAVE_DIR") or Path(__file__).resolve().parent / "saved_packages")
_NAME_RE = re.compile(r"^[A-Za-z0-9_\-]{1,120}\.json$")


class StorageError(ValueError):
    """Berkas paket tidak ada, rusak, atau tidak sesuai format."""


def _base(base_dir):
    return Path(base_dir) if base_dir else DEFAULT_DIR


def _check_name(name):
    if not isinstance(name, str) or not _NAME_RE.match(name):
        raise StorageError("Nama berkas tidak valid.")
    return name


def _dicts(items, must_have=()):
    if not isinstance(items, list):
        return []
    return [x for x in items if isinstance(x, dict) and all(k in x for k in must_have)]


def validate_package(data, meta):
    """Periksa dan lengkapi data/meta hasil muat agar aman dipakai build_docx & pratinjau."""
    if not isinstance(data, dict) or not isinstance(data.get("questions"), list) or not data["questions"]:
        raise StorageError("Berkas tidak berisi daftar soal.")
    questions = []
    for i, q in enumerate(data["questions"], 1):
        if not isinstance(q, dict) or not all(k in q for k in ("type", "question", "answer")):
            raise StorageError(f"Soal {i} tidak lengkap (butuh type, question, answer).")
        q = dict(q)
        q["number"] = i
        q.setdefault("options", {})
        if not isinstance(q["options"], dict):
            raise StorageError(f"Opsi soal {i} tidak valid.")
        for key, default in (("difficulty", "Sedang"), ("explanation", ""), ("scoring_rubric", ""),
                             ("passage_id", ""), ("stimulus", ""), ("indicator", "")):
            q.setdefault(key, default)
        questions.append(q)

    n = len(questions)
    clean = dict(data)
    clean["questions"] = questions
    clean["passages"] = _dicts(data.get("passages"), ("id", "text"))
    for key in ("blueprint", "qualitative_analysis"):
        items = data.get(key)
        ok = isinstance(items, list) and len(items) == n and all(isinstance(x, dict) for x in items)
        clean[key] = [dict(x, number=i) for i, x in enumerate(items, 1)] if ok else []
    for key in ("title", "instructions", "material_summary"):
        clean[key] = str(data.get(key) or "")

    meta = dict(meta) if isinstance(meta, dict) else {}
    types = list(dict.fromkeys(q["type"] for q in questions))
    types.sort(key=lambda t: QUESTION_TYPES.index(t) if t in QUESTION_TYPES else len(QUESTION_TYPES))
    defaults = {
        "subject": "Soal tersimpan", "grade": "10", "semester": "Semester 1", "topic": "",
        "output_language": "Indonesia", "difficulty": "Campuran", "qtype": ", ".join(types),
        "qtypes": types, "option_format": "ABCD", "model": "-",
        "include_key": True, "include_explanations": True, "include_blueprint": True,
        "include_analysis": True, "include_rubric": True, "include_summary": True,
    }
    for key, value in defaults.items():
        meta.setdefault(key, value)
    return clean, meta


def dumps_package(data, meta):
    payload = {"format": FORMAT_NAME, "version": FORMAT_VERSION,
               "saved_at": datetime.now().isoformat(timespec="seconds"), "meta": meta, "data": data}
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def loads_package(raw):
    """Baca isi berkas (bytes/str) -> (data, meta) yang sudah divalidasi."""
    if isinstance(raw, str):
        raw = raw.encode("utf-8")
    if len(raw) > MAX_IMPORT_BYTES:
        raise StorageError("Berkas terlalu besar (maks. 10 MB).")
    try:
        payload = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError) as exc:
        raise StorageError("Berkas bukan JSON yang valid.") from exc
    if not isinstance(payload, dict) or payload.get("format") != FORMAT_NAME:
        raise StorageError("Berkas ini bukan paket soal TKA Generator.")
    return validate_package(payload.get("data"), payload.get("meta"))


def _new_name(meta, directory):
    subject = re.sub(r"[^A-Za-z0-9_-]+", "_", str(meta.get("subject") or "Soal")).strip("_")[:50] or "Soal"
    grade = re.sub(r"[^A-Za-z0-9]+", "", str(meta.get("grade") or ""))[:4]
    stem = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{subject}" + (f"_Kelas_{grade}" if grade else "")
    name, k = f"{stem}.json", 2
    while (directory / name).exists():
        name, k = f"{stem}_{k}.json", k + 1
    return name


def save_package(data, meta, base_dir=None, name=None):
    """Simpan paket. Jika `name` diberikan, berkas itu ditimpa. Mengembalikan nama berkas."""
    directory = _base(base_dir)
    directory.mkdir(parents=True, exist_ok=True)
    name = _check_name(name) if name else _new_name(meta, directory)
    target = directory / name
    tmp = target.with_name(target.name + ".tmp")
    tmp.write_bytes(dumps_package(data, meta))
    os.replace(tmp, target)
    return name


def load_package(name, base_dir=None):
    path = _base(base_dir) / _check_name(name)
    if not path.is_file():
        raise StorageError("Paket tidak ditemukan.")
    return loads_package(path.read_bytes())


def delete_package(name, base_dir=None):
    path = _base(base_dir) / _check_name(name)
    if not path.is_file():
        raise StorageError("Paket tidak ditemukan.")
    path.unlink()


def list_packages(base_dir=None):
    """Daftar paket tersimpan, terbaru di atas. Berkas rusak dilewati."""
    directory = _base(base_dir)
    if not directory.is_dir():
        return []
    items = []
    for path in directory.glob("*.json"):
        if not _NAME_RE.match(path.name):
            continue
        try:
            data, meta = loads_package(path.read_bytes())
        except (StorageError, OSError):
            continue
        items.append({
            "name": path.name, "title": data.get("title") or meta["subject"],
            "subject": meta["subject"], "grade": meta["grade"], "semester": meta["semester"],
            "count": len(data["questions"]), "qtype": meta["qtype"],
            "saved_at": datetime.fromtimestamp(path.stat().st_mtime).strftime("%d %b %Y %H:%M"),
            "mtime": path.stat().st_mtime,
        })
    items.sort(key=lambda x: x["mtime"], reverse=True)
    return items
