# INFOR Excel Parser & Label Automation 🚀

Aplikasi desktop berbasis Python untuk mengotomatisasi ekstraksi dan *parsing* data *batch* Excel (INFOR) menjadi ratusan file `.txt` terstandardisasi. Sistem ini dirancang khusus untuk mempercepat proses pembuatan dan pencetakan label Haspel/Drum di area produksi.

**Business Impact:**
Implementasi aplikasi ini berhasil mengeliminasi pekerjaan pemindahan data manual oleh operator dan **memangkas waktu pemrosesan hingga 83% (dari 30 menit menjadi hanya 5 menit per siklus)**.

---

## 📸 Tampilan Aplikasi
<img width="1194" height="733" alt="image" src="https://github.com/user-attachments/assets/2274ea4c-6145-47f5-9ee4-f0e0d5816f7b" />
<img width="1049" height="903" alt="image" src="https://github.com/user-attachments/assets/26ede5bd-3de2-4f95-b627-eaec41f8a92d" />
<img width="1048" height="623" alt="image" src="https://github.com/user-attachments/assets/9377f2f0-034f-4791-9340-7c10a0dabbe1" />
<img width="1205" height="469" alt="image" src="https://github.com/user-attachments/assets/23316fb8-6608-4957-ae5c-238a5f9ab22b" />


## ✨ Fitur Utama
- **Automated Data Parsing:** Mengonversi data tabular dari Excel (.xlsx) menjadi struktur direktori dan spesifik file teks (9 file per *Lot*) secara instan.
- **Real-time API Integration:** Terintegrasi langsung dengan API internal untuk mengambil data berat aktual Haspel secara dinamis guna menghitung nilai *Netto* dan *Gross*.
- **Interactive GUI:** Dibangun menggunakan PyQt5 dengan mode *Fullscreen* untuk memudahkan operator lapangan dalam memilih, mencari, dan memproses data *Lot* tertentu.
- **Smart Calculation:** Melakukan kalkulasi otomatis untuk berat produk berdasarkan *quantity* dan standar deskripsi produk.
- **Built-in Logging System:** Dilengkapi dengan layar *Log* interaktif untuk memantau status keberhasilan proses ekstraksi secara *real-time*.

---

## 🛠️ Teknologi yang Digunakan
- **Bahasa Pemrograman:** Python 3.x
- **GUI Framework:** PyQt5
- **Data Manipulation:** Pandas
- **Networking/API:** Requests, JSON, urllib3
- **File System:** OS, Shutil

---

## ⚙️ Cara Kerja Sistem
1. **Load Data:** Operator memasukkan file Excel dari sistem INFOR ke dalam aplikasi.
2. **Fetch Data:** Aplikasi secara otomatis menarik data berat Haspel terbaru dari API perusahaan.
3. **Select & Process:** Operator dapat menggunakan fitur *search* atau *select all* untuk memilih nomor *Lot* yang ingin diproses.
4. **Output Generation:** Aplikasi membuat folder untuk setiap *Lot* dan secara otomatis memecah atribut data ke dalam 9 file `.txt` berbeda (misal: `No.Drum.txt`, `Netto.txt`, `Gross.txt`, `Customer.txt`) yang siap dibaca oleh mesin cetak label.

---

## ⚠️ Catatan Penting
Aplikasi ini melakukan penarikan data (*fetch*) dari API internal perusahaan. Untuk menjalankan aplikasi ini secara fungsional dengan data berat Haspel yang akurat, sistem **wajib terhubung dengan jaringan internal (Intranet/WiFi)** dari lokasi pabrik. Jika tidak terhubung, aplikasi akan tetap berjalan namun kalkulasi berat tambahan Haspel akan diatur ke *default* (0).

---
*Developed by Annisa Ashifa - Data Engineer*
