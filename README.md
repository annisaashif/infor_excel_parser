# INFOR Excel Parser & Label Automation 🚀

Aplikasi desktop berbasis Python untuk mengotomatisasi ekstraksi dan *parsing* data *batch* Excel (INFOR) menjadi ratusan file `.txt` terstandardisasi. Sistem ini dirancang khusus untuk mempercepat proses pembuatan dan pencetakan label Haspel/Drum di area produksi.

**Business Impact:**
Implementasi aplikasi ini berhasil mengeliminasi pekerjaan pemindahan data manual oleh operator dan **memangkas waktu pemrosesan hingga 83% (dari 30 menit menjadi hanya 5 menit per siklus)**.

---

## 📸 Tampilan Aplikasi
<img width="1193" height="751" alt="image" src="https://github.com/user-attachments/assets/fb905f5b-0fab-4922-bce0-3b8691735d10" />
<img width="884" height="896" alt="image" src="https://github.com/user-attachments/assets/d74aa081-0f28-4614-af6d-cdd12335cd0e" />
<img width="1105" height="581" alt="image" src="https://github.com/user-attachments/assets/ead9fc2a-71f9-46bb-87e2-ab996da13a32" />

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
