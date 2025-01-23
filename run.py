import os  
import re  
import pdfplumber  
import pandas as pd  
  
def welcome_message():  
    print("Script by Syahbandi")  
    print("Tidak diperkenankan untuk diperjual belikan")  
    print("GRATIS")  
  
def extract_tax_data(pdf_path, folder_path):  
    with pdfplumber.open(pdf_path) as pdf:  
        text = " ".join(page.extract_text() for page in pdf.pages)  
  
    pattern = r"(Pembeli Barang Kena Pajak\s*/\s*Penerima Jasa Kena Pajak.*?)(?=\s*(Harga\s*Jual|$))"  
    matches = re.findall(pattern, text, re.DOTALL)  
  
    if not matches:  
        return None  
  
    all_data = []  
    for match in matches:  
        extracted_text = match[0].strip()  
  
        name_pattern = r"Nama\s*:\s*(.*)"  
        address_pattern = r"Alamat\s*:\s*(.*)"  
        npwp16_pattern = r"NPWP\s*:\s*(\d{16})|NPWP\s*:\s*\d{15}\s*/\s*(\d{16})"  
        nitku_pattern = r"NITKU\s*:\s*(\d{22})"  
          
        name = re.search(name_pattern, extracted_text)  
        address = re.search(address_pattern, extracted_text)  
        npwp16 = re.search(npwp16_pattern, extracted_text)  
        nitku = re.search(nitku_pattern, extracted_text)  
  
        npwp16_value = npwp16.group(1) if npwp16 and npwp16.group(1) else (npwp16.group(2) if npwp16 else None)  
        nitku_value = nitku.group(1) if nitku else None  
  
        all_data.append({  
            "Nama": name.group(1).strip() if name else None,  
            "Alamat": address.group(1).strip() if address else None,  
            "NPWP16": npwp16_value,  
            "NITKU": nitku_value,  
            "File": os.path.basename(pdf_path),  
            "Path": folder_path  
        })  
  
    return all_data  
  
def process_pdfs_in_folder(folder_path, processed_data_set):  
    all_data = []  
  
    for filename in os.listdir(folder_path):  
        if filename.endswith(".pdf"):  
            pdf_path = os.path.join(folder_path, filename)  
            print(f"Memproses file: {filename}")  
            data_list = extract_tax_data(pdf_path, folder_path)  
            if data_list:  
                for data in data_list:  
                    data_tuple = (  
                        data["Nama"], data["Alamat"], data["NPWP16"],   
                        data["NITKU"]  
                    )  
                    if data_tuple not in processed_data_set:  
                        all_data.append(data)  
                        processed_data_set.add(data_tuple)  
    return all_data  
  
def save_data_to_excel(all_data, output_name):  
    if not output_name.endswith('.xlsx'):  
        output_name += '.xlsx'  
  
    if os.path.exists(output_name):  
        existing_data = pd.read_excel(output_name)  
        df_existing = existing_data.drop_duplicates(subset=["Nama", "Alamat", "NPWP16", "NITKU"], keep="first")  
        df_new = pd.DataFrame(all_data)  
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)  
        df_combined = df_combined.drop_duplicates(subset=["Nama", "Alamat", "NPWP16", "NITKU"], keep="first")  
    else:  
        df_combined = pd.DataFrame(all_data)  
      
    df_combined.to_excel(output_name, index=False)  
    print(f"Data berhasil disimpan ke {output_name}")  
  
def main():  
    welcome_message()  
  
    output_name = input("Masukkan nama file output (tanpa ekstensi .xlsx): ")  
  
    all_data = []  
    processed_data_set = set()  
  
    while True:  
        folder_path = input("Masukkan path folder yang berisi file PDF: ")  
        if os.path.exists(folder_path):  
            data_from_folder = process_pdfs_in_folder(folder_path, processed_data_set)  
            all_data.extend(data_from_folder)  
  
            lagi = input("Mau ekstrak PDF lainnya? (y/n): ").strip().lower()  
            if lagi != "y":  
                break  
        else:  
            print(f"Folder path {folder_path} tidak ditemukan. Coba lagi.")  
  
    save_data_to_excel(all_data, output_name)  
  
if __name__ == "__main__":  
    main()  
