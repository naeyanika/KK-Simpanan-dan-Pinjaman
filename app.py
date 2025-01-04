import pandas as pd

# File paths for input files
file_pinjaman = 'DbPinjaman.xlsx'
file_simpanan = 'DbSimpanan.xlsx'

# Read the Excel files
pinjaman = pd.read_excel(file_pinjaman, skiprows=1)
simpanan = pd.read_excel(file_simpanan, skiprows=1)

# Filter DbSimpanan based on 'Sts. Anggota' = 'AKTIF'
simpanan_filtered = simpanan[simpanan['Sts. Anggota'].str.upper() == 'AKTIF']

# Filter DbPinjaman based on 'Tanggal Keluar' is null
pinjaman_filtered = pinjaman[pinjaman['Tgl. Keluar'].str.upper() == 'AKTIF']

# Function to prepare KK Pinjaman
def create_kk_pinjaman(data):
    kk_pinjaman = pd.DataFrame({
        'No.': range(1, len(data) + 1),
        'Client ID': data['Client ID'],
        'Loan No.': data['Loan No.'],
        'Client Name': data['Client Name'],
        'Meeting Day': data['Meeting Day'],
        'Product Name': data['Product Name'],
        'Loan Amount': data['Loan Amount'],
        'Outstanding': data['Outstanding'],
        'Purpose Name': data['Purpose Name'],
        'Officer Name': data['Officer Name'],
        'Disb. Date': data['Disb. Date'],
        'Saldo Buku': '',  # Empty
        'Saldo Sistem': data['Outstanding'],
        'SLS': '',  # Empty
        'Identitas': '',  # Empty
        'Form UK': '',  # Empty
        'Form Keanggotaan': '',  # Empty
        'Form P3': '',  # Empty
        'Akad': '',  # Empty
        'Monitoring Pembiayaan': '',  # Empty
        'KPA': '',  # Empty
        'Sesuai/ Tidak sesuai': '',  # Empty
        'KETERANGAN (Kelemahan)': ''  # Empty
    })
    return kk_pinjaman

# Function to prepare KK Simpanan
def create_kk_simpanan(data):
    kk_simpanan = pd.DataFrame({
        'No.': range(1, len(data) + 1),
        'Client ID': data['Client ID'],
        'Account No.': data['Account No'],
        'Client Name': data['Client Name'],
        'Product Name': data['Product Name'],
        'Officer Name': data['Officer Name'],
        'Saldo': data['Saldo'],
        'Saldo Buku': data['Saldo'],
        'Saldo Selisih': '',  # Empty
        'Buku (SPO, SWA, SSU, SPE)': '',  # Empty
        'Kartu SIHARA': '',  # Empty
        'Kartu Kurban': '',  # Empty
        'Kartu Sipadan': '',  # Empty
        'Informasi Buku Simpanan (Sesuai/Tidak Sesuai)': '',  # Empty
        'KETERANGAN (Kelemahan)': ''  # Empty
    })
    return kk_simpanan

# Create the KK sheets
kk_pinjaman = create_kk_pinjaman(pinjaman_filtered)
kk_simpanan = create_kk_simpanan(simpanan_filtered)

# Save to Excel output
output_file = 'Kertas_Kerja_Output.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    kk_pinjaman.to_excel(writer, sheet_name='KK Pinjaman', index=False)
    kk_simpanan.to_excel(writer, sheet_name='KK Simpanan', index=False)

print(f"Output saved to {output_file}")
