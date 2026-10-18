# Audit v16 → v17

## Bug yang diperbaiki

| # | Temuan di v16 | Dampak | Perbaikan |
|---|---|---|---|
| 1 | `"\\n\\n"` (backslash literal) pada prompt, prompt perbaikan, dan penggabungan `material_summary` | Teks `\n\n` tampil apa adanya di prompt dan rangkuman Word | Diganti baris baru sungguhan |
| 2 | `json.loads` mode ketat | Wacana panjang berisi baris baru mentah ditolak (`Invalid control character`) — kemungkinan besar penyebab error JSON berulang sejak v14 | `strict=False`, `raw_decode`, hapus koma sisa, buang blok `<think>`, mode JSON Groq |
| 3 | Model `gpt-oss-120b` (prioritas utama) memakai token reasoning dari `max_tokens` | Respons kosong/terpotong | `reasoning_effort=low`, cek `finish_reason == "length"`, retry dengan batas token lebih besar |
| 4 | Prompt menyuruh menyalin wacana penuh ke `stimulus` setiap soal, dan minimal 1–2 wacana per batch | Output membengkak ×3 → terpotong | Wacana ditulis sekali per batch di `passages`, soal merujuk `passage_id`; tampil sekali di Word/pratinjau |
| 5 | Batch saling tidak tahu | Soal bisa berulang; campuran jenis soal tidak seimbang; judul bisa menyebut batch | Rencana jenis soal pasti per nomor, daftar soal sebelumnya dikirim, judul untuk seluruh paket |
| 6 | Tidak ada validasi isi | Kunci PG tak cocok opsi, jumlah opsi salah, jenis soal tak dipilih lolos | `normalize_batch`: validasi opsi, kunci, Benar/Salah, jenis, wacana; regenerasi otomatis (maks. 3 percobaan) |
| 7 | Pratinjau memakai `unsafe_allow_html` dengan teks AI mentah | Injeksi HTML; baris kosong memecah blok HTML; `$` dibaca sebagai rumus | Semua teks di-escape, newline → `<br>`, `$` di-escape, `dir="auto"` untuk Arab |
| 8 | Word: `w:shd` tanpa `w:val` dan urutan elemen salah; tabel soal 7,07" melebihi area cetak; kertas Letter | Risiko peringatan/kerusakan di Word, tabel meluber | Perbaikan XML, A4, lebar tabel 6,7" |
| 9 | Teks Arab di Word tidak RTL | Bacaan Arab terbalik/tidak rata kanan | `w:bidi`, `w:rtl`, font complex-script otomatis |
| 10 | Widget kolom kanan ditaruh di `with col2` di dalam `st.form` | Sebenarnya di luar form → rerun tiap perubahan | Kolom dibuat di dalam form |
| 11 | README v13 menyebut tombol "Uji API Key aktif" yang tidak ada | Fitur hilang | Ditambahkan; pesan 401 menyebut sumber key |
| 12 | Bahasa output default Indonesia untuk Bahasa Arab/Sunda/Inggris | Salah pilih bahasa tanpa sadar | Default otomatis mengikuti mata pelajaran, tetap bisa diubah |
| 13 | Deteksi error lewat substring (`"401"`, `"batch"`) | Salah klasifikasi | Memakai `status_code`; retry otomatis untuk 429 |
| 14 | Rumus ditulis LaTeX oleh AI | Tampil mentah di Word | Prompt melarang LaTeX/`$`, memakai Unicode |
| 15 | Kisi-kisi memakai header bahasa Inggris mentah (`cognitive_level`, dst.); semester tidak tampil di Word | Tampilan kurang rapi | Label Indonesia; semester ditambahkan |
| 16 | `choose_model` bisa jatuh ke model acak (mis. compound); contoh key `gsk-` | Model tak cocok; contoh salah | Daftar pengecualian diperluas; contoh `gsk_` |

## Belum diuji / perlu Anda ketahui

- **Belum diuji langsung ke Groq dan UI Streamlit** (sandbox tanpa jaringan dan tanpa Streamlit). Logika inti diuji dengan LLM tiruan (`tests/test_core.py`), dokumen Word dirender dan divalidasi.
- Perilaku `reasoning_effort` dan mode JSON bergantung pada model/akun; jika Groq menolak mode JSON (HTTP 400), aplikasi otomatis mengulang tanpa mode itu.
- Paket 100 soal = 34 panggilan berurutan; pada tier gratis bisa lambat karena batas TPM (ada retry otomatis).
- Mata pelajaran bahasa kini menghasilkan 1 wacana per 3 soal (lebih banyak dari minimum lama).
- Silabus Biologi kelas 12 tetap belum rinci.
- Peringatan skema `settings.xml/zoom` pada validator berasal dari template bawaan python-docx, bukan kode ini; Word tetap membukanya.
