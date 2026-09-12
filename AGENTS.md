# AGENTS.md — Konvensi Kerja

## Project
- **MyKopinka** — Aplikasi Google Apps Script (GAS) untuk Portal Koperasi Kopinka.
- Struktur file:
  - `Code.gs` — server-side GAS (Google Sheets sebagai database).
  - `Index.html` — frontend (HTML + CSS + JS inline, satu halaman).
  - `workers.js` — file pekerja (worker script eksternal).
  - `appsscript.json` — konfigurasi manifest GAS.

## Aturan Commit & Push (WAJIB)
- **Setiap selesai melakukan perubahan/menyelesaikan fitur, HARUS langsung:**
  1. `git add <file yang relevan>` (hanya file terkait).
  2. `git commit -m "<pesan>"`.
  3. `git push origin main`.
- **Gaya pesan commit:** lowercase, tilde snake_case tanpa spasi, deskriptif, misal:
  - `add_fitur_rekap_kredit_per_anggota_export_excel`
  - `fix_parse_date_iso_string_cache_rekap_kredit_toko`
  - `update_animasi_loading_filter_rekap_kredit_toko`
- Jangan commit data sensitif (kunci/secret).

## Verifikasi Sebelum Commit
- Cek sintaks `Code.gs` via node: `node -e "new (require('vm').Script)(require('fs').readFileSync('Code.gs','utf8')); console.log('OK')"`.
- Cek sintaks blok `<script>` utama di `Index.html` dengan node (parse per blok script).

## Catatan
- Platform: Windows / PowerShell 5.1. `&&` tidak didukung — gunakan `;` atau `if ($?) { ... }`.
- Peringatan `LF will be replaced by CRLF` dari git bersifat normal, bisa diabaikan.
- Remote: `origin` → `https://github.com/isyaau/Mykopinka-GAS.git`, branch `main`.