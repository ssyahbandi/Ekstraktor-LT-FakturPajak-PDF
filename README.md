# Ekstraksi NPWP dari PDF

## Deskripsi  
Proyek ini adalah skrip Python yang digunakan untuk mengekstrak data NPWP (Nomor Pokok Wajib Pajak) dari faktur pajak dalam format PDF. Skrip ini dirancang untuk menangkap NPWP16 dan NITKU jika tersedia, serta mengabaikan format NPWP lainnya.  

## Fitur  
- Menangkap NPWP16 dalam dua format:  
  - NPWP dengan 15 digit diikuti oleh 16 digit.  
  - NPWP hanya dengan 16 digit.  
- Menyimpan data yang diekstrak ke dalam file Excel.  
- Mengabaikan NPWP15 dan hanya fokus pada NPWP16.  
- Mengambil NITKU jika tersedia.  
  
## Prasyarat  
- Python 3.x  
- Library yang diperlukan:  
  - `pdfplumber`  
  - `pandas`  
  
## Instalasi  
1. Clone repositori ini:
   
   ```git clone https://github.com/ssyahbandi/Ekstraktor-LT-FakturPajak-PDF```
   
   ```cd Ekstraktor-LT-FakturPajak-PDF```
   
3. Install library yang diperlukan:
   
   ```pip install pdfplumber pandas```
## Format yang didukung

  ### Format 1
  
  | Pembeli Barang Kena Pajak / Penerima Jasa Kena Pajak          |
  ----------------------------------------------------------------|
  | Nama : PT ABC                                                 |
  | Alamat : Kav. 40-42, Jl. Gatot Subroto No.7 7, RT.7/RW.1, Senayan, Kec. Kby. Baru, Kota Jakarta Selatan, Daerah Khusus Ibukota Jakarta 12190                     |
  | NPWP : 123456789012345 / 1234567890123456                     |
  | NITKU : 1234567890123456000000                                |
  
  ### Format 2

  | Pembeli Barang Kena Pajak / Penerima Jasa Kena Pajak          |
  ----------------------------------------------------------------|
  | Nama : PT DCE                                                 |
  | Alamat : Kav. 40-42, Jl. Gatot Subroto No.7 7, RT.7/RW.1, Senayan, Kec. Kby. Baru, Kota Jakarta Selatan, Daerah Khusus Ibukota Jakarta 12190                    |
  | NPWP : 1234567890123456                                       |
  | NITKU : -                                                     |
  
  
## Penggunaan  
1. Jalankan skrip:
   
   ```python run.py```

2. Ikuti instruksi di terminal untuk memasukkan nama file output dan path folder yang berisi file PDF.

   Contoh Path          : ```F:\Extrak PDF Faktur Pajak\Perusahaan\2024\11 November 2024```
   
   Contoh Path Lainnya : ```F:\Extrak PDF Faktur Pajak\Perusahaan\2024\12 Desember 2024```

   Dan Pastikan di dalam Folder Path tersebut ada file PDF

   ![Desain tanpa judul (1)](https://github.com/user-attachments/assets/1b5b198e-f73c-4094-be13-d97827503c69)

## Contoh  
Setelah menjalankan skrip, data yang diekstrak akan disimpan dalam file Excel dengan format berikut:  
  
  | Nama                   | Alamat                                           | NPWP16              | NITKU                | File                     |  
  |------------------------|--------------------------------------------------|---------------------|----------------------|--------------------------|  
  | [NAMA LT]              | [ALAMAT_LT]                                      | [NPWP16]            | -                    | NAMA_PDF.pdf  |  
  | [NAMA LT]              | [ALAMAT_LT]                                      | [NPWP16]            | [NITKU/IDTKU]        | NAMA_PDF.pdf      |  
  
## Kontribusi  
Jika Anda ingin berkontribusi pada proyek ini, silakan buat pull request atau buka isu untuk diskusi.  
  
## Lisensi  
Proyek ini dilisensikan di bawah MIT License. Lihat file LICENSE untuk detail lebih lanjut.  
