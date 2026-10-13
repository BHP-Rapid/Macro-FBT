# FBT Reconciliation Tool

Aplikasi rekonsiliasi Fringe Benefit Tax (FBT) yang membandingkan data **PeopleSoft** dengan data **SAP**
untuk modul **MBA**, **C&B**, dan **FSA**. Seluruh proses berjalan di browser Anda; tidak ada server dan
tidak ada data yang dikirim keluar.

## Cara Menjalankan

1. Buka folder proyek.
2. **Klik dua kali `index.html`** (otomatis terbuka di browser, disarankan Microsoft Edge atau Chrome).
3. Aplikasi siap dipakai. Tidak perlu instalasi, Node.js, atau server.

> **Perlu internet.** `index.html` memuat Tailwind CSS, SheetJS, ExcelJS, dan font Inter dari CDN.
> Pastikan komputer terhubung ke internet saat pertama kali membuka aplikasi.

## Cara Memakai

1. **Upload file.** Unggah satu file PeopleSoft dan satu file SAP (`.xlsx` atau `.csv`) dari HDC.
   Nama file bebas; jenis file dikenali dari header kolom.
   - Jika file terpasang di slot yang salah, tekan **Swap Files**.
   - Jika modul kedua file berbeda (misalnya PeopleSoft MBA dengan SAP C&B), tombol mulai nonaktif
     dan muncul pesan *Module Type Mismatch*. Unggah pasangan file dari modul yang sama.
2. **Start Reconciliation.** Sistem mencocokkan data per Matching Key:

   | Modul | Kunci PeopleSoft | Kunci SAP |
   |---|---|---|
   | MBA | `HMS_INVOICE_NBR` | `REFERENCE` |
   | C&B | `ECS_RMB_SAP_REF` | `REFERENCE` |
   | FSA | `HMS_FSA_TRANS_ID` | 10 karakter terakhir `TEXT254` |

3. **Tinjau hasil di dashboard.**
   - Kartu KPI mengikuti filter yang aktif.
   - Cari, urutkan, filter status (All, Match, Mismatch, Not Found), dan atur jumlah baris per halaman.
   - Tampilkan atau sembunyikan kolom lewat menu **Filters**.
   - Klik sebuah baris untuk melihat detail invoice dan menulis catatan.
4. **Download.** Unduh dua laporan Excel: versi PeopleSoft dan versi SAP. Keduanya memuat data asli
   ditambah kolom `REKON_STATUS`, total sisi lawan, `DIFFERENCE_SAP_vs_PS`, dan `REMARKS`.

### Arti Status

| Status | Arti |
|---|---|
| MATCH | Ada di kedua sistem, nominal sama |
| SAP > PS | Nominal SAP lebih besar |
| PS > SAP | Nominal PeopleSoft lebih besar |
| PeopleSoft Only | Hanya ada di PeopleSoft |
| SAP Only | Hanya ada di SAP |

## Struktur Folder

```
.
├── index.html        Halaman aplikasi (buka file ini)
├── css/
│   └── style.css     Seluruh gaya tampilan khusus aplikasi
├── js/
│   ├── app.js        Logika aplikasi (dimuat oleh index.html)
│   ├── inject.js     Skrip bantu pengembangan
│   ├── patch_logic.js  Skrip bantu pengembangan
│   └── split.js      Skrip bantu pengembangan
└── README.md
```

Hanya `index.html`, `css/style.css`, dan `js/app.js` yang dipakai saat aplikasi berjalan.
`inject.js`, `patch_logic.js`, dan `split.js` tidak dimuat oleh `index.html`; fungsinya adalah skrip bantu
saat pengembangan dan tidak perlu dijalankan oleh pengguna.

## Privasi dan Batasan

- Data diproses sepenuhnya di browser dan tidak diunggah ke mana pun. Menutup tab menghapus data hasil.
- Aplikasi tidak mengubah file sumber dan tidak menyediakan fitur koreksi data.
- Data Concur dan KD KR hanya berasal dari SAP, sehingga tidak direkonsiliasi.
- Ukuran data yang dituju: 8.000 sampai 15.000 baris per file.

## Pemecahan Masalah

| Gejala | Penyebab dan solusi |
|---|---|
| Tampilan berantakan atau tombol tidak berfungsi | Tidak ada koneksi internet (CDN gagal dimuat). Sambungkan internet lalu muat ulang halaman. |
| Pesan *Module Type Mismatch* | Pasangan file berbeda modul. Unggah file dari modul yang sama. |
| File tidak terdeteksi | Header kolom tidak sesuai format HDC. Pastikan file belum diubah strukturnya. |
| Sisi file terbalik | Tekan **Swap Files**. |
| Gaya tidak termuat | Pastikan folder `css/` berada di sebelah `index.html`. |
