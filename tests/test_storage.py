"""Tes penyimpanan lokal paket soal. Jalankan: python tests/test_storage.py"""
import json
import os
import sys
import tempfile
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core  # noqa: E402
import storage  # noqa: E402
import test_core as helpers  # noqa: E402


def make_package(count=6, qtypes=("Pilihan Ganda", "Essay")):
    cfg = helpers.base_cfg(count=count, qtypes=list(qtypes), subject="Biologi")
    data = core.generate_package(helpers.make_llm(cfg, []), cfg)
    meta = {"subject": "Biologi", "grade": "10", "semester": "Semester 1", "topic": "t",
            "output_language": "Indonesia", "difficulty": "Campuran", "qtype": ", ".join(qtypes),
            "qtypes": list(qtypes), "option_format": "ABCDE", "model": "tiruan"}
    return data, meta


def raises(fn, *args, **kw):
    try:
        fn(*args, **kw)
    except storage.StorageError:
        return True
    return False


def test_save_load_list_delete_roundtrip():
    data, meta = make_package()
    with tempfile.TemporaryDirectory() as tmp:
        assert storage.list_packages(tmp) == []
        name = storage.save_package(data, meta, tmp)
        assert name.endswith(".json") and "Biologi_Kelas_10" in name
        listed = storage.list_packages(tmp)
        assert [p["name"] for p in listed] == [name]
        assert listed[0]["count"] == 6 and listed[0]["subject"] == "Biologi"
        data2, meta2 = storage.load_package(name, tmp)
        assert [q["question"] for q in data2["questions"]] == [q["question"] for q in data["questions"]]
        assert [q["type"] for q in data2["questions"]] == ["Pilihan Ganda"] * 3 + ["Essay"] * 3
        assert meta2["subject"] == "Biologi" and meta2["model"] == "tiruan"
        # hasil muat bisa langsung diekspor ke Word dan pratinjau
        assert core.build_docx(data2, meta2)[:2] == b"PK"
        assert "Pilihan Ganda" in core.build_preview_html(data2)
        storage.delete_package(name, tmp)
        assert storage.list_packages(tmp) == []
        assert raises(storage.load_package, name, tmp)


def test_overwrite_same_name_and_unique_new_names():
    data, meta = make_package()
    with tempfile.TemporaryDirectory() as tmp:
        a = storage.save_package(data, meta, tmp)
        again = storage.save_package(data, dict(meta, topic="diubah"), tmp, name=a)
        assert again == a and len(storage.list_packages(tmp)) == 1
        assert storage.load_package(a, tmp)[1]["topic"] == "diubah"
        b = storage.save_package(data, meta, tmp)  # nama baru, tidak menimpa
        assert b != a and len(storage.list_packages(tmp)) == 2
        assert not [f for f in os.listdir(tmp) if f.endswith(".tmp")]


def test_path_traversal_and_bad_names_rejected():
    data, meta = make_package()
    with tempfile.TemporaryDirectory() as tmp:
        for bad in ("../x.json", "a/b.json", "..\\x.json", "x.txt", "", ".json", "x.json/..", None):
            assert raises(storage.load_package, bad, tmp), bad
            assert raises(storage.delete_package, bad, tmp), bad
            if bad:  # name kosong/None berarti "buat nama baru", bukan error
                assert raises(storage.save_package, data, meta, tmp, name=bad), bad


def test_corrupt_and_foreign_files():
    data, meta = make_package()
    with tempfile.TemporaryDirectory() as tmp:
        good = storage.save_package(data, meta, tmp)
        with open(os.path.join(tmp, "rusak.json"), "w") as f:
            f.write("{bukan json")
        with open(os.path.join(tmp, "asing.json"), "w") as f:
            json.dump({"hello": "world"}, f)
        assert [p["name"] for p in storage.list_packages(tmp)] == [good]  # yang rusak dilewati
        assert raises(storage.load_package, "rusak.json", tmp)
        assert raises(storage.load_package, "asing.json", tmp)


def test_loads_package_validation_and_defaults():
    assert raises(storage.loads_package, b"")
    assert raises(storage.loads_package, b"[]")
    assert raises(storage.loads_package, b"x" * (storage.MAX_IMPORT_BYTES + 1))
    base = {"format": storage.FORMAT_NAME, "version": 1}
    assert raises(storage.loads_package, json.dumps(dict(base, data={"questions": []})))
    assert raises(storage.loads_package, json.dumps(dict(base, data={"questions": [{"type": "Essay"}]})))
    # paket minimal: field opsional dilengkapi, nomor dibuat ulang, aman untuk Word & pratinjau
    minimal = dict(base, data={"questions": [
        {"type": "Essay", "question": "Jelaskan <b>X</b>", "answer": "jawab", "number": 99},
        {"type": "Pilihan Ganda", "question": "Q2", "answer": "A", "options": {"A": "a", "B": "b"}},
    ], "blueprint": ["salah"], "passages": ["salah"]})
    data, meta = storage.loads_package(json.dumps(minimal, ensure_ascii=False))
    assert [q["number"] for q in data["questions"]] == [1, 2]
    assert data["blueprint"] == [] and data["passages"] == []
    assert meta["qtypes"] == ["Pilihan Ganda", "Essay"] and meta["include_key"] is True
    assert core.build_docx(data, meta)[:2] == b"PK"
    html = core.build_preview_html(data)
    assert "<b>X</b>" not in html and "&lt;b&gt;X&lt;/b&gt;" in html


def test_dumps_loads_roundtrip_with_arabic_and_bom():
    data, meta = make_package(count=3, qtypes=("Pilihan Ganda",))
    data["questions"][0]["question"] = "ما هو الجواب؟ ✓ é"
    raw = storage.dumps_package(data, meta)
    assert "ما هو الجواب" in raw.decode("utf-8")  # tidak di-escape \uXXXX
    assert storage.loads_package(raw)[0]["questions"][0]["question"] == data["questions"][0]["question"]
    assert storage.loads_package(b"\xef\xbb\xbf" + raw)[0]["title"] == data["title"]  # BOM Windows


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK ", name)
