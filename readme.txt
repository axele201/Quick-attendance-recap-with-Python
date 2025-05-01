Sorry, friends, I deleted my latest post. Apparently, there are some people who can't appreciate something that was made for free and instead blamed the program. 
Just to clarify, the program was already noted as not being directly connected to the database. It manually takes a copy of the data from the database and creates a new Excel document with the same structure and data. The attendance times were only used for recap purposes.

README - PROGRAM REKAP ABSENSI HARIAN

======================================
⚠️  PERHATIAN SEBELUM MENJALANKAN PROGRAM
======================================

1. FILE absensi.xlsx ADALAH FILE ABSENSI UNTUK 1 HARI SAJA.
2. JIKA MENGGUNAKAN SPREADSHEET PERHATIKAN DAN SESUAIKAN SUSUNAN, URUTAN, ATAU STRUKTUR NAMA PADA FILE rekap_absensi.xlsx SESUAI DENGAN DAFTAR NAMA YANG ADA DI SPREADSHEET.
3. PERTAHANKAN URUTAN NAMA SESUAI DENGAN DATABASE.
   NAMA TERAKHIR PADA DATA ASLI ADALAH "..." (SESUAIKAN PADA DATA ASLI)
4. PENAMBAHAN NAMA BARU AKAN DILAKUKAN SETELAH NAMA "..." .
5. JANGAN TAMBAH/HAPUS NAMA SECARA MANUAL DI FILE rekap_absensi.xlsx SAJA. JIKA ADA PENGHAPUSAN ATAU PERUBAHAN HARUS DISESUAIKAN SUSUNAN DAN URUTAN NAMA YANG ADA DI FILE rekap_absensi.xlsx SAMA YANG ADA DI SPREADSHEET KARENA NANTI YANG DIAMBIL HANYA JAM NYA SAJA

===========================
✅ PERSIAPAN SEBELUM MENJALANKAN
===========================

- PASTIKAN FILE rekap_absensi.xlsx SUDAH BERSIH DARI DATA JAM SEBELUMNYA
- (JIKA MENGGUNAKAN SPREADSHEET) PASTIKAN FILE DATA SPREADSHEET DAN rekap_absensi.xlsx SUDAH SESUAI SECARA SUSUNAN DAN URUTAN NAMA.
- TUTUP KEDUA FILE rekap_absensi.xlsx DAN absensi.xlsx DI MICROSOFT EXCEL SEBELUM MENJALANKAN PROGRAM KARENA JIKA MASIH TERBUKA MAKA PROGRAM TIDAK DAPAT AKSES BUAT MERUBAH FILE TERSEBUT.

================
📄 CARA KERJA PROGRAM
================

- MEMBACA DATA absensi.xlsx
- MENYALIN KOLOM: Nama, Clock In, Clock Out, Lembur
- JIKA NAMA SUDAH ADA: DATA Clock In, Clock Out, DAN Lembur AKAN DIPERBARUI
- JIKA NAMA BELUM ADA ATAU TIDAK SESUAI: DATA BARU AKAN DITAMBAHKAN SETELAH NAMA TERAKHIR PADA FILE rekap_absensi.xlsx "..."
- HASIL DISIMPAN KEMBALI KE rekap_absensi.xlsx

1. **Program rekap.py**:
   - Mencakup kolom: **Nama**, **Clock In**, dan **Clock Out**.
   - Program ini hanya memperbarui atau menambahkan data berdasarkan nama, waktu masuk, dan waktu keluar.
   
2. **Program rekapv2.py**:
   - Mencakup kolom: **Nama**, **Clock In**, **Clock Out**, dan **Lembur**.
   - Program ini memperbarui data sesuai dengan **Clock In**, **Clock Out**, dan **Lembur**.  
   - Program ini cocok untuk kasus di mana data lembur perlu dicatat.

==========================
📊 RUMUS EXCEL YANG DIGUNAKAN
==========================

1. **Mengubah angka di `Clock In` menjadi format waktu**:
   Rumus berikut digunakan untuk mengubah nilai angka di kolom **Clock In** menjadi format waktu (jam:menit):
   Keterangan:  
      - Fungsi ini akan mengubah angka di kolom **Clock In** menjadi format **hh:mm** dengan menit tetap 00.  
      - Jika kolom **Clock In** kosong, hasilnya akan kosong juga.

   --------------->     =IF(B2:B313<>"",TEXT(B2:B313,"00")&":00","")    <------------------

2. **Mengubah angka di `Clock Out` dan menambahkan jam lembur**:
   Rumus berikut digunakan untuk mengubah nilai di **Clock Out** dengan menambahkan jam lembur yang tercatat:

   Keterangan:  
      - Fungsi ini pertama-tama memeriksa apakah **Clock Out** (C2) dan **Lembur** (D2) kosong.  
      - Jika keduanya kosong, hasilnya juga kosong.
      - Jika ada data, rumus ini akan menjumlahkan **Clock Out** dengan **Lembur** (dalam satuan jam) dan mengubahnya menjadi format waktu **hh:mm**.

   --------------->     =IF(AND(C2="", D2=""), "", TEXT((C2 + D2)/24, "hh:mm"))  <-------------


=========================
📌 FORMAT KOLOM YANG DIGUNAKAN
=========================

| Kolom      | Keterangan                    |
|------------|-------------------------------|
| Nama       | Nama lengkap karyawan         |
| Clock In   | Angka masuk (absen masuk)     |
| Clock Out  | Angka pulang (absen keluar)   | 
| Lembur     | angka lembur                  |

===============================
JIKA SEMUA INSTRUKSI DIIKUTI,
PROGRAM AKAN BERJALAN DENGAN AMAN
===============================
