# MyKopinka v2.0 — Portal Koperasi Kopinka

Aplikasi web Portal Koperasi Kopinka yang dibangun dengan **Google Apps Script (GAS)**, memakai **Google Sheets** sebagai database, dan diakses melalui **Cloudflare Worker** untuk mendukung fitur kamera & scan barcode lintas-origin.

## Fitur Utama

### Autentikasi & Peran (Role)
- Login berbasis username & password dengan deteksi peran otomatis:
  - **Kasir** — transaksi redeem voucher & pencatatan kredit toko.
  - **Admin Potongan** — kelola data potongan gaji.
  - **Admin** — akses penuh (pengguna, voucher, piutang, notifikasi, unit, rekap, pengaturan).
  - **Anggota** — profil mandiri, cek voucher, histori piutang, rekap bulanan.

### Voucher
- Cek status voucher (aktif, expired, used, diblokir).
- Redeem tunggal & **redeem massal**, dengan dukungan nominal piutang.
- Blokir / aktifkan voucher, perpanjang masa aktif (satu, per anggota, atau semua expired).
- Cari voucher multi-anggota dan histori pemakaian bulanan.

### Kredit Toko / Piutang
- Pencatatan piutang: manual, dari redeem voucher, redeem massal, dan import tagihan Excel.
- **Batas kredit global & per anggota**, termasuk fitur blokir anggota/toko.
- Pembayaran piutang (manual, import `PembayaranPiutang`) dengan simulasi **biaya jasa 1,5%**.
- Rekap kredit per anggota & rekap pembayaran, filter bulan/tahun/tanggal/toko, export Excel.
- Sinkronisasi tagihan otomatis ke pembayaran piutang.
- Notifikasi & pengingat hutang melalui **Email** dan **WhatsApp**.

### Potongan (Admin Potongan)
- Upload potongan gaji via Excel, input manual, edit & hapus per baris atau massal.
- Generator slip potongan per anggota per bulan/tahun.

### Notifikasi & Pengumuman
- Kirim notifikasi (file/ lampiran), import massal via Excel, arsip/hapus, pagination.
- Panel notifikasi khusus admin.
- Pengaturan nomor WhatsApp toko & template pesan WA (struk, rekap admin/kasir, group).

### Manajemen Data
- **Users**: CRUD, blokir piutang, reset profil mandiri (foto, NIP, KTP, email, bank, dll).
- **Unit/Kelompok**: CRUD unit & kelompok.
- **Toko**: daftar toko & nomor WA.
- **Pengaturan global**: limit piutang, limit kredit aktif, tampilan limit untuk anggota, template WA.

### Kamera & Scan Barcode (via Cloudflare Worker)
- Verifikasi foto anggota langsung dari kamera (*snapshot*).
- Scan barcode voucher (Code 39 / Code 128 / QR) dengan `html5-qrcode`.
- Overlay kamera lintas-origin yang menembus batasan iframe Apps Script.
- Foto disimpan ke Google Drive melalui endpoint `doPost`.

### Rekap & Laporan / Dashboard
- Dashboard admin dengan statistik & grafik (Chart.js).
- Histori mutasi kasir dengan filter toko/tanggal, pagination, dan export Excel (SheetJS/XLSX).
- Rekap kredit kasir per toko.

## Struktur File

| File | Deskripsi |
|------|-----------|
| `Code.gs` | Backend server-side GAS (~3.000 baris): logika bisnis, akses Google Sheets, Drive, MailApp, CacheService. |
| `Index.html` | Frontend satu halaman (HTML + CSS + JS): UI, `google.script.run`, Chart.js, SweetAlert2, SheetJS, html5-qrcode. |
| `workers.js` | Cloudflare Worker: membungkus halaman Apps Script, overlay kamera & scanner barcode lintas-origin. |
| `appsscript.json` | Manifest GAS (timezone Asia/Jakarta, runtime V8, OAuth scopes). |

## Arsitektur

```
Browser
  └─ Cloudflare Worker (workers.js)
       └─ iframe → Google Apps Script (platform /u/.../exec)
            ├─ Index.html (frontend)
            └─ Code.gs (backend)
                 ├─ Google Sheets  (database: Users, Voucher, Mutasi, Potongan, Notifikasi, Piutang, PembayaranPiutang, Unit, Settings)
                 └─ Google Drive   (foto verifikasi, lampiran notifikasi, bukti piutang)
```

PHP script Apps Script berkomunikasi dengan frontend melalui `google.script.run` (untuk fungsi internal) dan `ContentService`/`doPost` (untuk unggah foto dari worker).

## Sheet Database (Google Sheets)

| Sheet | Kegunaan |
|-------|----------|
| `Users` | Data pengguna: anggota, kasir, admin; termasuk blokir piutang & limit anggota. |
| `Voucher` | Kode voucher, status, nilai, masa aktif, pemilik. |
| `Mutasi` | Riwayat redeem / transaksi. |
| `Potongan` | Data potongan gaji. |
| `Notifikasi` | Daftar notifikasi & pengumuman. |
| `Piutang` | Tagihan kredit toko. |
| `PembayaranPiutang` | Riwayat pembayaran piutang. |
| `Unit` | Daftar unit / kelompok. |
| `Settings` | Pengaturan key-value (limit piutang, template WA, flag fitur). |

## Folder Google Drive

- `FOLDER_ID` — penyimpanan foto verifikasi kamera.
- `FOLDER_NOTIF_ID` — lampiran file notifikasi.
- `FOLDER_KREDIT_TOKO_ID` — bukti foto pembayaran / tagihan piutang.

## Pengembangan & Deployment

1. Buka [script.google.com](https://script.google.com), buat project baru, tempel isi `Code.gs`, `Index.html`, dan `appsscript.json`.
2. Atur ID folder Google Drive di bagian atas `Code.gs`.
3. Buat sheet-sheet database yang dibutuhkan di spreadsheet yang sama.
4. Deploy sebagai **Web App** (`Execute as: User accessing the web app`, `Who has access: Anyone`).
5. Salin URL `/exec` ke konstanta `APPS_SCRIPT_EXEC_URL` di `workers.js`.
6. Deploy `workers.js` ke **Cloudflare Workers** sebagai pintu masuk utama.

## Catatan Teknis

- **Cache**: `CacheService` berputar berbasis potongan (*chunked*, ~90KB ± entry) untuk data piutang, pengguna, dan pembayaran guna mengurangi hit Google Sheets.
- **Masa aktif / expired**: voucher otomatis mendapat status `Expired` sesuai `tglExp`, plus fungsi perpanjangan massal.
- **Biaya jasa piutang**: tarif 1,5% (lihat `_getBiayaJasaRate`).
- Tidak menyimpan file sensitif/secret di repository; konfigurasi sensitif via Properties/Settings sheet.

## Kontribusi

Ikuti konvensi pada `AGENTS.md`:

- Pesan commit: lowercase `snake_case` deskriptif (contoh: `fix_parse_date_iso_string_cache_rekap_kredit_toko`).
- Setiap selesai perubahan: `git add` (file terkait) → `git commit` → `git push origin main`.
- Platform Windows/PowerShell 5.1: gunakan `;` — `&&` tidak didukung.