#Ekstraksi NPWP dari PDF

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
   ```git clone https://github.com/username/repo-name.git```
   ```cd repo-name```
   
3. Install library yang diperlukan:  
  ```pip install pdfplumber pandas```

## Penggunaan  
1. Jalankan skrip:
  ```python run.py```

2. Ikuti instruksi di terminal untuk memasukkan nama file output dan path folder yang berisi file PDF.  
  
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
