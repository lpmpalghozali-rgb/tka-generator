import os
import re
import time
from datetime import datetime

import streamlit as st
from groq import Groq

from core import (
    SCHOOL_NAME, QUESTION_TYPES, OPTION_FORMATS, DIFFICULTIES,
    BATCH_SIZE, GenerationError, generate_package, build_docx, build_preview_html,
)

st.set_page_config(page_title="TKA Generator AI | SMA ISLAM AL-GHOZALI", page_icon="📘", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root {
  --brand-navy: #10233f;
  --brand-blue: #2457c5;
  --brand-teal: #16a6a1;
  --surface: #ffffff;
  --muted: #65758b;
}
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: linear-gradient(180deg, #f4f7fc 0%, #f8fafc 35%, #f4f7fc 100%); }
.block-container { max-width: 1320px; padding-top: 1.5rem; padding-bottom: 3rem; }
h1, h2, h3 { font-family: 'Manrope', sans-serif !important; letter-spacing: -0.035em; color: #122641; }
h1 { font-size: clamp(2rem, 4vw, 3.05rem) !important; font-weight: 800 !important; line-height: 1.08 !important; }
h2 { font-size: 1.55rem !important; font-weight: 750 !important; }
h3 { font-size: 1.12rem !important; font-weight: 700 !important; }
p, label, .stCaption { color: #334155; }
[data-testid="stSidebar"] { background: #101f36; border-right: 1px solid #203651; }
[data-testid="stSidebar"] * { color: #eef4ff; }
[data-testid="stSidebar"] input { color: #10233f !important; }
[data-testid="stSidebar"] [data-baseweb="input"] { background: #fff; border-radius: 10px; }
[data-testid="stSidebar"] [data-testid="stAlert"] { background: rgba(255,255,255,.08); }
div[data-testid="stVerticalBlockBorderWrapper"] { border-color: #dbe4f0 !important; border-radius: 16px !important; }
div[data-testid="stForm"] { background: #ffffff; border: 1px solid #dbe4f0; border-radius: 18px; padding: 1.1rem 1.2rem 1rem; box-shadow: 0 8px 28px rgba(16,35,63,.045); }
div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea,
div[data-testid="stSelectbox"] [data-baseweb="select"], div[data-testid="stNumberInput"] input {
  border-radius: 10px !important;
}
.stButton button, .stDownloadButton button, div[data-testid="stFormSubmitButton"] button {
  border-radius: 11px !important; border: 0 !important; font-weight: 700 !important;
  min-height: 2.75rem; transition: transform .15s ease, box-shadow .15s ease;
}
.stButton button[kind="primary"], div[data-testid="stFormSubmitButton"] button[kind="primary"],
.stDownloadButton button[kind="primary"] {
  background: linear-gradient(135deg, #2457c5, #167caa) !important;
  color: #fff !important; box-shadow: 0 8px 18px rgba(36,87,197,.22);
}
.stButton button:hover, .stDownloadButton button:hover { transform: translateY(-1px); box-shadow: 0 10px 22px rgba(16,35,63,.12); }
div[data-testid="stTabs"] button { font-weight: 650; }
div[data-testid="stTabs"] [aria-selected="true"] { color: #2457c5 !important; border-bottom-color: #2457c5 !important; }
div[data-testid="stAlert"] { border-radius: 12px; }
[data-testid="stMetric"] { background: #fff; border: 1px solid #dbe4f0; border-radius: 14px; padding: 1rem; }
.hero-panel {
  background: linear-gradient(120deg, #10233f 0%, #17416d 58%, #167caa 100%);
  color: #f7fbff; border-radius: 22px; padding: 1.65rem 1.8rem;
  margin: .25rem 0 1.25rem; box-shadow: 0 16px 36px rgba(16,35,63,.14);
  position: relative; overflow: hidden;
}
.hero-panel:after { content: ''; position: absolute; width: 240px; height: 240px; border-radius: 50%;
  right: -50px; top: -110px; border: 1px solid rgba(255,255,255,.18);
  box-shadow: 0 0 0 28px rgba(255,255,255,.035), 0 0 0 58px rgba(255,255,255,.025); }
.hero-kicker { text-transform: uppercase; letter-spacing: .16em; font-size: .72rem; font-weight: 800; opacity: .78; }
.hero-title { color: #fff !important; font: 800 clamp(1.8rem, 4vw, 2.8rem)/1.1 'Manrope', sans-serif; letter-spacing: -.04em; margin: .5rem 0 .65rem; }
.hero-sub { color: #e0ebfa !important; max-width: 780px; font-size: 1rem; line-height: 1.6; margin: 0; }
.hero-pills { display:flex; gap:.5rem; flex-wrap:wrap; margin-top:1.1rem; }
.hero-pill { border:1px solid rgba(255,255,255,.22); background:rgba(255,255,255,.1); border-radius:999px; padding:.35rem .7rem; color:#f7fbff; font-size:.78rem; font-weight:650; }
.section-label { color:#2457c5; font-size:.76rem; letter-spacing:.12em; text-transform:uppercase; font-weight:800; margin-bottom:.35rem; }
.helper-copy { color:#64748b; font-size:.9rem; line-height:1.55; }
.st-key-topic_ textarea { line-height: 1.55; }
@media (max-width: 700px) {
  .block-container { padding-left: 1rem; padding-right: 1rem; }
  .hero-panel { padding: 1.25rem; border-radius: 16px; }
}
</style>
""", unsafe_allow_html=True)

SUBJECTS = [
    # Mata pelajaran umum
    "Pendidikan Agama dan Budi Pekerti",
    "Pendidikan Pancasila dan Kewarganegaraan (PPKn)",
    "Bahasa Indonesia",
    "Bahasa Inggris",
    "Bahasa Arab",
    "Bahasa Sunda",
    "Matematika (Wajib)",
    "Sejarah Indonesia",
    "Pendidikan Jasmani, Olahraga, dan Kesehatan (PJOK)",
    "Seni Budaya (Rupa/Musik/Tari/Teater)",
    "Prakarya dan Kewirausahaan (PKWU)",
    "Informatika / Teknologi Informasi",
    # Peminatan IPA
    "Matematika Peminatan",
    "Fisika",
    "Kimia",
    "Biologi",
    # Peminatan IPS
    "Ekonomi",
    "Sosiologi",
    "Geografi",
    "Sejarah Peminatan",
    # Latihan lintas bidang
    "Literasi",
    "Numerasi",
    "Campuran",
]

OUTPUT_LANGUAGES = ["Indonesia", "Arab", "Inggris", "Sunda"]

TOPIC_SUGGESTIONS = {
    "Bahasa Arab": {
        "10": "Mufradat sehari-hari, ta'aruf, dhamir, isim isyarah, jumlah ismiyah dan fi'liyah, qira'ah sederhana",
        "11": "Hiwar, qira'ah, tarkib nahwu-sharaf dasar, fi'il madhi/mudhari/amr, jumlah dan pemahaman teks",
        "12": "Fahmul maqru', insya' sederhana, tarkib lanjutan, balaghah dasar, pemahaman teks tematik",
    },
    "Bahasa Sunda": {
        "10": "Undak-usuk basa, paguneman, aksara Sunda, carita pondok, budaya Sunda",
        "11": "Biantara, sisindiran, sajak, artikel basa Sunda, wawacan jeung pamahaman eusi",
        "12": "Carpon, novel Sunda, drama, aksara Sunda, tradisi jeung budaya lokal, nulis jeung nganalisis téks",
    },
    "Matematika Peminatan": {
        "10": "Eksponen, logaritma, fungsi, vektor dasar",
        "11": "Trigonometri analitika, lingkaran, suku banyak (polinomial)",
        "12": "Limit fungsi aljabar/trigonometri, turunan, integral, statistika inferensial",
    },
    "Fisika": {
        "10": "Besaran dan satuan, gerak lurus, hukum Newton, optik geometri",
        "11": "Dinamika rotasi, kesetimbangan benda tegar, fluida, suhu, kalor, termodinamika",
        "12": "Gelombang bunyi/cahaya, listrik statis/dinamis, medan magnet, fisika modern",
    },
    "Kimia": {
        "10": "Struktur atom, tabel periodik, ikatan kimia, hukum dasar kimia",
        "11": "Hidrokarbon, termokimia, laju reaksi, kesetimbangan kimia, asam-basa, kelarutan",
        "12": "Sifat koligatif larutan, redoks, elektrokimia, kimia unsur, makromolekul",
    },
    "Biologi": {
        "10": "Keanekaragaman hayati, virus, bakteri, ekosistem",
        "11": "Struktur dan fungsi sel, jaringan tumbuhan/hewan, sistem organ manusia",
        "12": "Metabolisme sel, genetika, pembelahan sel, evolusi, bioteknologi",
    },
    "Ekonomi": {
        "10": "Konsep dasar ekonomi, kelangkaan, pasar, lembaga jasa keuangan",
        "11": "Pendapatan nasional, pertumbuhan ekonomi, ketenagakerjaan, kebijakan fiskal/moneter, APBN/APBD",
        "12": "Akuntansi perusahaan jasa, akuntansi perusahaan dagang",
    },
    "Sosiologi": {
        "10": "Fungsi sosiologi, interaksi sosial, gejala sosial di masyarakat",
        "11": "Kelompok sosial, permasalahan sosial, kesetaraan, konflik, kekerasan, perdamaian",
        "12": "Perubahan sosial, globalisasi, ketimpangan sosial, kearifan lokal, evaluasi pemberdayaan masyarakat",
    },
    "Geografi": {
        "10": "Pengetahuan dasar geografi, pemetaan, penginderaan jauh, jagat raya",
        "11": "Posisi strategis Indonesia, flora dan fauna, sumber daya alam, ketahanan pangan, dinamika kependudukan",
        "12": "Konsep wilayah dan tata ruang, interaksi desa-kota, pemanfaatan SIG, negara maju dan berkembang",
    },
    "Sejarah Peminatan": {
        "10": "Ruang lingkup ilmu sejarah, sumber sejarah, penelitian sejarah, manusia purba",
        "11": "Kerajaan Hindu-Buddha/Islam, reformasi gereja, revolusi besar dunia, Perang Dunia I dan II",
        "12": "Perang Dingin, organisasi global/regional, sejarah kontemporer dunia, konflik di berbagai belahan dunia",
    },
}

SILABUS = {
    "Matematika Peminatan": {
        "10": {
            "Semester 1": "Fungsi eksponen: sifat, grafik, persamaan dan pertidaksamaan eksponen. Fungsi logaritma: sifat, grafik, persamaan dan pertidaksamaan logaritma.",
            "Semester 2": "Vektor dua dimensi R2: operasi vektor, panjang vektor, dot product. Vektor tiga dimensi R3: proyeksi ortogonal dan sudut antarvektor.",
        },
        "11": {
            "Semester 1": "Trigonometri analitika: jumlah/selisih sudut, sudut ganda, perkalian dan penjumlahan sinus/kosinus. Persamaan trigonometri sin x=a, cos x=a, tan x=a.",
            "Semester 2": "Lingkaran berpusat di (0,0) dan (a,b), garis singgung lingkaran. Polinomial: operasi polinomial, Teorema Sisa, Teorema Faktor, akar rasional.",
        },
        "12": {
            "Semester 1": "Limit fungsi trigonometri, limit di ketakhinggaan aljabar/trigonometri, turunan fungsi trigonometri dan aplikasi maksimum/minimum serta garis singgung.",
            "Semester 2": "Statistika inferensial, distribusi binomial, variabel acak, uji hipotesis, distribusi normal, kurva normal, nilai Z dan penerapannya.",
        },
    },
    "Fisika": {
        "10": {
            "Semester 1": "Hakikat fisika, metode ilmiah, keselamatan laboratorium, pengukuran, ketidakpastian, angka penting, jangka sorong, mikrometer, vektor, GLB dan GLBB.",
            "Semester 2": "Gerak parabola dan melingkar, Hukum Newton I/II/III, bidang miring dan katrol, gravitasi dan hukum Kepler, usaha, energi, kekekalan energi mekanik.",
        },
        "11": {
            "Semester 1": "Dinamika rotasi, kesetimbangan benda tegar, torsi, momen inersia, titik berat, elastisitas, Hukum Hooke, fluida statis dan dinamis, kontinuitas, Bernoulli.",
            "Semester 2": "Suhu, kalor, pemuaian, asas Black, perpindahan kalor, teori kinetik gas, gas ideal, termodinamika, siklus Carnot, mesin kalor, gelombang mekanik, gelombang berjalan/stasioner.",
        },
        "12": {
            "Semester 1": "Rangkaian DC, Hukum Ohm dan Kirchhoff, elektrostatika, gaya Coulomb, medan/potensial listrik, kapasitor, medan magnet, gaya Lorentz, induksi Faraday-Lenz, generator, transformator, rangkaian AC RLC dan resonansi.",
            "Semester 2": "Gelombang elektromagnetik, spektrum dan dampak, relativitas khusus, efek fotolistrik, efek Compton, radiasi benda hitam, teknologi digital, transmisi data, radioaktivitas, reaksi inti dan energi nuklir.",
        },
    },
    "Kimia": {
        "10": {
            "Semester 1": "Hakikat kimia, metode ilmiah, laboratorium, struktur atom, teori atom, partikel atom, nomor atom/massa, isotop, konfigurasi elektron Bohr/kuantum, bilangan kuantum, sistem periodik dan sifat periodik unsur.",
            "Semester 2": "Ikatan ion/kovalen/logam, kepolaran, gaya antarmolekul, gaya Van der Waals, ikatan hidrogen, bentuk molekul VSEPR dan hibridisasi, tata nama dan penyetaraan reaksi, hukum dasar kimia, konsep mol, stoikiometri, kadar zat, rumus empiris/molekul.",
        },
        "11": {
            "Semester 1": "Hidrokarbon: alkana, alkena, alkuna, struktur, tata nama dan sifat; minyak bumi dan angka oktan; termokimia, entalpi, Hess, energi ikatan; laju reaksi, teori tumbukan, faktor, orde dan persamaan laju.",
            "Semester 2": "Kesetimbangan dinamis, Kc/Kp, Le Chatelier, teori asam-basa, pH kuat/lemah, hidrolisis garam, larutan penyangga, titrasi asam-basa, kurva titrasi, sistem koloid, efek Tyndall, gerak Brown, elektroforesis dan koagulasi.",
        },
        "12": {
            "Semester 1": "Sifat koligatif larutan elektrolit/nonelektrolit, tekanan uap, titik didih/beku, osmosis, redoks, penyetaraan reaksi, sel Volta, potensial sel, korosi, elektrolisis dan Hukum Faraday.",
            "Semester 2": "Kimia unsur: gas mulia, halogen, alkali, alkali tanah, periode 3 dan transisi periode 4; turunan alkana; benzena dan turunannya; polimerisasi adisi/kondensasi, karbohidrat, protein dan lipida.",
        },
    },
    "Biologi": {
        "10": {
            "Semester 1": "Ruang lingkup biologi, organisasi kehidupan, keselamatan dan metode ilmiah, keanekaragaman hayati tingkat gen/jenis/ekosistem, klasifikasi 5/6 kingdom, kunci dikotomi, virus dan replikasi litik/lisogenik, Monera dan bakteri.",
            "Semester 2": "Protista, fungi dan simbiosis liken/mikoriza, Plantae (lumut, paku, tumbuhan berbiji), Animalia (invertebrata dan vertebrata), ekosistem, rantai/jaring makanan, piramida ekologi, siklus biogeokimia, pencemaran dan perubahan lingkungan.",
        },
        "11": {
            "Semester 1": "Struktur dan komponen sel, sel hewan/tumbuhan, difusi, osmosis, transpor aktif, jaringan tumbuhan, kultur jaringan, jaringan hewan, sistem gerak manusia, tulang, sendi, otot, sistem sirkulasi, darah, jantung dan golongan darah.",
            "Semester 2": "Sistem pencernaan dan uji makanan, sistem pernapasan, sistem ekskresi dan pembentukan urine, sistem koordinasi (saraf, indra, hormon), pengaruh NAPZA, sistem reproduksi, gametogenesis, siklus menstruasi, fertilisasi, ASI, sistem imun dan imunisasi.",
        },
        "12": {
            "Semester 1": "Materi yang tersedia: metabolisme sel, genetika, pembelahan sel, evolusi dan bioteknologi. Sesuaikan rincian dengan silabus sekolah.",
            "Semester 2": "Silakan sesuaikan dengan silabus semester 2 yang digunakan sekolah; topik dapat diisi melalui kolom Topik/Materi.",
        },
    },
}


SUBJECT_DEFAULT_LANGUAGE = {
    "Bahasa Arab": "Arab", "Bahasa Sunda": "Sunda", "Bahasa Inggris": "Inggris",
}


def get_api_key():
    """Read API key from Streamlit Cloud Secrets or local environment."""
    try:
        key = st.secrets.get("GROQ_API_KEY", "")
        if key:
            return str(key).strip()
    except Exception:
        pass
    return os.getenv("GROQ_API_KEY", "").strip()


def choose_model(client):
    """Pick a usable text-generation model from models currently returned by Groq."""
    models = client.models.list()
    model_ids = []
    excluded = ("whisper", "guard", "safeguard", "speech", "tts", "audio", "playai",
                "orpheus", "compound", "embed", "vision-preview")
    for model in getattr(models, "data", []):
        model_id = getattr(model, "id", "") or ""
        if not model_id or getattr(model, "active", True) is False:
            continue
        if any(token in model_id.lower() for token in excluded):
            continue
        model_ids.append(model_id)
    for preferred in ("openai/gpt-oss-120b", "llama-3.3-70b-versatile",
                      "openai/gpt-oss-20b", "llama-3.1-8b-instant"):
        if preferred in model_ids:
            return preferred
    if model_ids:
        return model_ids[0]
    raise RuntimeError("Tidak ditemukan model teks yang tersedia pada akun Groq ini.")


def _status_code(exc):
    code = getattr(exc, "status_code", None)
    if isinstance(code, int):
        return code
    m = re.search(r"\b(4\d\d|5\d\d)\b", str(exc))
    return int(m.group(1)) if m else None


def _retry_wait_seconds(exc):
    m = re.search(r"try again in\s+(?:(\d+)m)?\s*(\d+(?:\.\d+)?)(ms|s)", str(exc), flags=re.I)
    if not m:
        return 10.0
    seconds = int(m.group(1) or 0) * 60 + float(m.group(2)) / (1000 if m.group(3).lower() == "ms" else 1)
    return min(seconds + 1.0, 45.0)


def make_llm(client, model):
    """Bungkus pemanggilan Groq: mode JSON, reasoning rendah, dan retry saat kena rate limit."""
    reasoning = "gpt-oss" in model.lower()

    def llm(messages, max_tokens, temperature=0.2):
        kwargs = dict(model=model, messages=messages, temperature=temperature, max_tokens=max_tokens)
        if reasoning:
            # Token reasoning ikut dihitung dalam max_tokens; tekan agar JSON tidak terpotong.
            kwargs["extra_body"] = {"reasoning_effort": "low"}
        use_json_mode = True
        for attempt in range(4):
            try:
                call_kwargs = dict(kwargs)
                if use_json_mode:
                    call_kwargs["response_format"] = {"type": "json_object"}
                resp = client.chat.completions.create(**call_kwargs)
                choice = resp.choices[0]
                return choice.message.content or "", getattr(choice, "finish_reason", None)
            except Exception as exc:  # noqa: BLE001
                code = _status_code(exc)
                if code == 429 and attempt < 3:
                    time.sleep(_retry_wait_seconds(exc))
                    continue
                if code == 400 and use_json_mode:
                    use_json_mode = False  # model/tier tidak mendukung atau gagal validasi JSON mode
                    continue
                raise
        raise RuntimeError("Gagal memanggil Groq setelah beberapa percobaan.")

    return llm


def describe_error(exc, key_source):
    text = str(exc)
    code = _status_code(exc)
    if code == 401 or "invalid_api_key" in text.lower():
        return ("Autentikasi Groq gagal (401). Key yang aktif berasal dari "
                f"{key_source}. Periksa key tersebut; kosongkan kolom sidebar jika ingin memakai key dari Secrets.")
    if code == 403:
        return "Akses ditolak (403). Periksa izin akun/organisasi Groq atau pembatasan wilayah pada API key."
    if code == 413 or "request too large" in text.lower():
        return "Permintaan terlalu besar untuk batas token Groq (413). Persingkat instruksi tambahan atau kurangi topik."
    if code == 429:
        return "Batas kuota/rate limit Groq tercapai (429). Tunggu beberapa saat lalu coba lagi, atau pakai API key lain."
    return f"Gagal membuat soal: {text}"


# ------------------------------------------------------------------ hero & sidebar
st.markdown("""
<div class="hero-panel">
  <div class="hero-kicker">Platform asesmen digital • SMA Islam Al-Ghozali</div>
  <div class="hero-title">TKA Generator AI</div>
  <p class="hero-sub">Susun paket latihan yang terstruktur, berbasis materi, dan siap diekspor ke Microsoft Word. Didukung Groq AI dengan pengaturan API yang fleksibel.</p>
  <div class="hero-pills">
    <span class="hero-pill">✦ Kelas 10–12</span>
    <span class="hero-pill">✦ 4 bahasa output</span>
    <span class="hero-pill">✦ Multi jenis soal</span>
    <span class="hero-pill">✦ Ekspor .DOCX</span>
  </div>
</div>
""", unsafe_allow_html=True)
st.info("Untuk Bahasa Indonesia, Inggris, Sunda, Arab, dan Literasi, generator menyertakan wacana panjang sebagai stimulus soal.")

default_api_key = get_api_key()

with st.sidebar:
    st.header("⚙️ Pengaturan API & Soal")
    user_api_key = st.text_input(
        "Groq API Key milik Anda (opsional)",
        type="password",
        key="user_groq_api_key",
        placeholder="Tempel API key Groq pribadi di sini",
        help="Jika diisi, aplikasi memakai API key Anda. Jika kosong, aplikasi memakai API key default pengelola.",
    ).strip()
    api_key = user_api_key if user_api_key else default_api_key
    key_source = "kolom sidebar (key pribadi)" if user_api_key else "Streamlit Secrets / environment"
    if user_api_key:
        st.success("Mode API: menggunakan API key pribadi Anda.")
        st.caption("Permintaan AI menggunakan kuota akun Groq Anda.")
    elif default_api_key:
        st.info("Mode API: menggunakan API key default aplikasi.")
        st.caption("Masukkan API key pribadi di atas untuk memakai kuota Groq Anda.")
    else:
        st.warning("Belum ada API key. Masukkan API key Groq pribadi di atas.")
    if st.button("🔌 Uji API Key aktif", use_container_width=True, disabled=not api_key):
        try:
            Groq(api_key=api_key, timeout=20).models.list()
            st.success(f"API key valid (sumber: {key_source}).")
        except Exception as exc:  # noqa: BLE001
            st.error(describe_error(exc, key_source))
    st.caption("API key hanya disimpan dalam sesi aplikasi dan tidak ditampilkan sebagai teks biasa. Jangan bagikan key kepada orang lain.")
    st.divider()
    st.caption("Model Groq dipilih otomatis dari model teks yang tersedia.")

if not api_key:
    st.error("Masukkan Groq API Key pribadi pada panel Pengaturan, atau minta pengelola mengonfigurasi API key default melalui Streamlit Secrets.")
    st.stop()

# ------------------------------------------------------------------------ form
st.markdown('<div class="section-label">01 / Rancang paket soal</div>', unsafe_allow_html=True)
st.markdown("## Atur materi dan format asesmen")
st.markdown('<p class="helper-copy">Pilih mata pelajaran, kelas, semester, bahasa, dan format soal. Materi akan menyesuaikan pilihan Anda secara otomatis.</p>', unsafe_allow_html=True)

# Selektor di luar form agar materi langsung menyesuaikan saat pilihan berubah.
col1, col2 = st.columns(2)
with col1:
    subject = st.selectbox("Mata pelajaran", SUBJECTS, key="subject_selector")
    grade = st.selectbox("Kelas", ["10", "11", "12"], key="grade_selector")
    semester = st.selectbox("Semester", ["Semester 1", "Semester 2"], key="semester_selector")

with col2:
    # Bahasa output otomatis mengikuti mata pelajaran bahasa; tetap bisa diubah manual.
    if st.session_state.get("_lang_subject") != subject:
        st.session_state["language_selector"] = SUBJECT_DEFAULT_LANGUAGE.get(subject, "Indonesia")
        st.session_state["_lang_subject"] = subject
    output_language = st.selectbox(
        "Bahasa output soal", OUTPUT_LANGUAGES, key="language_selector",
        help="Bahasa untuk soal, opsi, kunci, pembahasan, kisi-kisi, rubrik, dan rangkuman. Otomatis mengikuti mata pelajaran bahasa, dapat diubah.",
    )
    curriculum = st.text_input("Kurikulum / acuan (opsional)", placeholder="Contoh: Kurikulum Merdeka")

suggested_topic = SILABUS.get(subject, {}).get(grade, {}).get(
    semester, TOPIC_SUGGESTIONS.get(subject, {}).get(grade, ""))
topic_key = f"topic_{subject}_{grade}_{semester}"
if st.session_state.get("_topic_selection_key") != topic_key:
    st.session_state[topic_key] = suggested_topic
    st.session_state["_topic_selection_key"] = topic_key
topic = st.text_area("Topik / materi (terisi otomatis, bisa diedit)", key=topic_key, height=110,
                     placeholder="Materi mengikuti silabus kelas dan semester; dapat diedit.")

with st.form("generator_form"):
    f1, f2 = st.columns(2)
    with f1:
        count = st.number_input("Jumlah soal", min_value=1, max_value=100, value=10, step=1)
        qtypes = st.multiselect(
            "Jenis soal (bisa pilih lebih dari satu)", QUESTION_TYPES, default=["Pilihan Ganda"],
            help="Jika lebih dari satu jenis, soal divariasikan secara bergantian dalam satu paket.",
        )
    with f2:
        difficulty = st.selectbox("Tingkat kesulitan", DIFFICULTIES, index=3)
        option_format = st.selectbox("Opsi jawaban Pilihan Ganda", OPTION_FORMATS,
                                     help="Hanya berlaku untuk soal Pilihan Ganda.")
    st.markdown("**Bagian dokumen Word**")
    a, b, c = st.columns(3)
    with a:
        include_key = st.checkbox("Kunci jawaban", value=True)
        include_explanations = st.checkbox("Pembahasan", value=True)
    with b:
        include_blueprint = st.checkbox("Kisi-kisi", value=True)
        include_analysis = st.checkbox("Analisis kualitatif", value=True)
    with c:
        include_rubric = st.checkbox("Rubrik penilaian", value=True)
        include_summary = st.checkbox("Rangkuman materi", value=True)
    extra_instructions = st.text_area(
        "Instruksi tambahan (opsional)",
        placeholder="Contoh: gunakan stimulus kontekstual dan ukur penalaran tingkat tinggi.")
    submitted = st.form_submit_button("✨ Buat Soal", type="primary", use_container_width=True)

for state_key in ("generated_data", "generated_meta", "generated_docx"):
    st.session_state.setdefault(state_key, None)

# ------------------------------------------------------------------- generation
if submitted:
    if not qtypes:
        st.error("Pilih minimal satu jenis soal.")
    else:
        try:
            client = Groq(api_key=api_key, timeout=90.0, max_retries=1)
            with st.spinner("Memeriksa model Groq yang tersedia..."):
                model = choose_model(client)
            cfg = {
                "subject": subject, "grade": grade, "semester": semester, "topic": topic.strip(),
                "count": int(count), "difficulty": difficulty, "qtypes": list(qtypes),
                "option_format": option_format, "output_language": output_language,
                "curriculum": curriculum.strip(), "include_explanations": include_explanations,
                "extra_instructions": extra_instructions.strip(),
            }
            bar = st.progress(0.0, text=f"Memulai dengan model {model}...")

            def on_progress(done, total, message):
                bar.progress(min(done / total, 1.0), text=f"{message} (model {model})")

            data = generate_package(make_llm(client, model), cfg, progress=on_progress)
            bar.progress(1.0, text="Selesai.")
            meta = {
                "subject": subject, "grade": grade, "semester": semester, "topic": cfg["topic"],
                "output_language": output_language, "difficulty": difficulty,
                "qtype": ", ".join(qtypes), "qtypes": list(qtypes), "option_format": option_format,
                "model": model, "include_key": include_key,
                "include_explanations": include_explanations,
                "include_blueprint": include_blueprint, "include_analysis": include_analysis,
                "include_rubric": include_rubric, "include_summary": include_summary,
            }
            st.session_state.generated_data = data
            st.session_state.generated_meta = meta
            st.session_state.generated_docx = build_docx(
                data, meta, include_key=include_key, include_explanations=include_explanations,
                include_blueprint=include_blueprint, include_analysis=include_analysis,
                include_rubric=include_rubric, include_summary=include_summary)
            st.success(f"Paket soal berhasil dibuat dengan model `{model}`.")
        except GenerationError as exc:
            st.error(f"Gagal memproses jawaban AI: {exc}")
            st.warning(f"Soal dibuat per batch ({BATCH_SIZE} soal per panggilan) dengan beberapa percobaan otomatis. "
                       "Coba ulangi, persingkat instruksi tambahan, atau kurangi jumlah soal.")
        except Exception as exc:  # noqa: BLE001
            st.error(describe_error(exc, key_source))

# ---------------------------------------------------------------------- results
data = st.session_state.generated_data
meta = st.session_state.generated_meta
if data and meta:
    st.divider()
    st.markdown('<div class="section-label">02 / Hasil generator</div>', unsafe_allow_html=True)
    st.subheader(data.get("title", "Pratinjau Paket Soal"))
    st.caption(f'{SCHOOL_NAME} • {meta["subject"]} • Kelas {meta["grade"]} • {meta["semester"]} • '
               f'Bahasa: {meta["output_language"]} • {meta["qtype"]} • Model: {meta["model"]}')
    m1, m2, m3 = st.columns(3)
    m1.metric("Total soal", len(data["questions"]))
    m2.metric("Bahasa output", meta["output_language"])
    m3.metric("Jenis soal", len(meta["qtypes"]))
    tab_questions, tab_key, tab_json = st.tabs(["📄 Pratinjau soal", "🔑 Kunci & pembahasan", "🧩 Data JSON"])
    with tab_questions:
        st.write(data.get("instructions", "Baca setiap soal dengan cermat."))
        st.markdown(build_preview_html(data), unsafe_allow_html=True)
    with tab_key:
        for i, q in enumerate(data["questions"], 1):
            st.markdown(f"**{i}. Jawaban:** {q['answer']}")
            if q.get("explanation"):
                st.text(q["explanation"])
    with tab_json:
        st.json(data)

    safe_subject = re.sub(r"[^A-Za-z0-9_-]+", "_", meta["subject"]).strip("_")
    filename = f"TKA_SMA_ISLAM_AL_GHOZALI_{safe_subject}_Kelas_{meta['grade']}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    st.download_button(
        "⬇️ Ekspor ke Microsoft Word (.docx)",
        data=st.session_state.generated_docx,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

st.divider()
st.caption("Hasil AI perlu ditinjau guru sebelum digunakan. Aplikasi ini menghasilkan latihan orisinal dan bukan soal resmi TKA.")
