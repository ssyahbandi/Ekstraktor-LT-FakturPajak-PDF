import os
import re
import time
import pdfplumber
import pandas as pd

def welcome_message():
    # Menampilkan welcome message
    print("Script by Syahbandi")
    print("Tidak diperkenankan untuk diperjual belikan")
    print("GRATIS")
    time.sleep(3)  # Tunggu selama 3 detik

def extract_tax_data(pdf_path, folder_path):
    # Buka file PDF
    with pdfplumber.open(pdf_path) as pdf:
        # Gabungkan semua teks dari semua halaman PDF
        text = " ".join(page.extract_text() for page in pdf.pages)

    # Ekspresi reguler untuk mencari semua faktur pajak
    pattern = r"(Pembeli Barang Kena Pajak\s*/\s*Penerima Jasa Kena Pajak.*?)(?=\s*(Harga\s*Jual|$))"
    matches = re.findall(pattern, text, re.DOTALL)

    # Jika tidak ada faktur pajak, kembalikan None
    if not matches:
        return None

    all_data = []
    for match in matches:
        extracted_text = match[0].strip()

        # Ekspresi reguler untuk data yang akan diambil
        name_pattern = r"Nama\s*:\s*(.*)"
        address_pattern = r"Alamat\s*:\s*(.*)"
        npwp15_pattern = r"NPWP\s*:\s*(\d{15})"
        npwp16_pattern = r"NPWP\s*:\s*\d{15}\s*/\s*(\d{16})"
        nitku_pattern = r"NITKU\s*:\s*(\d{22})"
        
        # Pencarian data menggunakan regex
        name = re.search(name_pattern, extracted_text)
        address = re.search(address_pattern, extracted_text)
        npwp15 = re.search(npwp15_pattern, extracted_text)
        npwp16 = re.search(npwp16_pattern, extracted_text)
        nitku = re.search(nitku_pattern, extracted_text)

        # Ambil hasilnya, jika tidak ditemukan beri nilai None
        all_data.append({
            "Nama": name.group(1).strip() if name else None,
            "Alamat": address.group(1).strip() if address else None,
            "NPWP15": npwp15.group(1) if npwp15 else None,
            "NPWP16": npwp16.group(1) if npwp16 else None,
            "NITKU": nitku.group(1) if nitku else None,
            "File": os.path.basename(pdf_path),  # Nama file PDF
            "Path": folder_path  # Path folder yang memuat file PDF
        })

    return all_data

def process_pdfs_in_folder(folder_path, processed_data_set):
    all_data = []  # List untuk menyimpan semua data

    # Iterasi semua file PDF di folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):  # Pastikan hanya memproses file PDF
            pdf_path = os.path.join(folder_path, filename)
            print(f"Memproses file: {filename}")
            data_list = extract_tax_data(pdf_path, folder_path)  # Mendapatkan daftar data untuk setiap faktur pajak
            if data_list:
                # Filter data yang sudah ada sebelumnya untuk menghindari duplikat
                for data in data_list:
                    # Buat tuple hanya berdasarkan kolom-kolom yang relevan untuk menghindari duplikasi
                    data_tuple = (
                        data["Nama"], data["Alamat"], data["NPWP15"], 
                        data["NPWP16"], data["NITKU"]
                    )  # Menggunakan tuple kolom penting saja
                    if data_tuple not in processed_data_set:
                        all_data.append(data)
                        processed_data_set.add(data_tuple)  # Tambahkan data yang baru ke set
    return all_data

def save_data_to_excel(all_data, output_name):
    # Pastikan output_name memiliki ekstensi .xlsx
    if not output_name.endswith('.xlsx'):
        output_name += '.xlsx'

    # Cek apakah file output sudah ada
    if os.path.exists(output_name):
        # Jika file ada, baca data lama dan gabungkan dengan data baru
        existing_data = pd.read_excel(output_name)

        # Gabungkan data lama dan baru, dan pastikan duplikasi dihapus berdasarkan kolom-kolom penting
        df_existing = existing_data.drop_duplicates(subset=["Nama", "Alamat", "NPWP15", "NPWP16", "NITKU"], keep="first")
        df_new = pd.DataFrame(all_data)

        # Gabungkan dan hapus duplikasi lagi setelah digabung
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        df_combined = df_combined.drop_duplicates(subset=["Nama", "Alamat", "NPWP15", "NPWP16", "NITKU"], keep="first")
    else:
        # Jika file tidak ada, buat baru
        df_combined = pd.DataFrame(all_data)
    
    # Simpan data gabungan ke file Excel
    df_combined.to_excel(output_name, index=False)
    print(f"Data berhasil disimpan ke {output_name}")

def main():
    welcome_message()  # Menampilkan welcome message

    # Meminta nama output sebelum memproses file PDF
    output_name = input("Masukkan nama file output (tanpa ekstensi .xlsx): ")

    all_data = []  # Menyimpan data dari semua folder
    processed_data_set = set()  # Set untuk melacak data yang sudah diproses

    # Loop untuk meminta input dan memproses data
    while True:
        folder_path = input("Masukkan path folder yang berisi file PDF: ")
        if os.path.exists(folder_path):
            data_from_folder = process_pdfs_in_folder(folder_path, processed_data_set)
            all_data.extend(data_from_folder)  # Gabungkan data dari folder ini ke dalam all_data

            # Menanyakan apakah ingin memproses PDF lainnya
            lagi = input("Mau ekstrak PDF lainnya? (y/n): ").strip().lower()
            if lagi != "y":
                break
        else:
            print(f"Folder path {folder_path} tidak ditemukan. Coba lagi.")

    # Simpan hasil ke Excel setelah semua folder diproses
    save_data_to_excel(all_data, output_name)

if __name__ == "__main__":
    main()
