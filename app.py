import streamlit as st
import pandas as pd

# Judul aplikasi
st.title("Kertas Kerja Simpanan dan Pinjaman")
st.write("""File ini berisikan DbSimpanan.xlsx dan DbPinjaman.xlsx sesuai dengan periode audit, jangan hapus header. Upload Sesuai dengan kolom upload yang tersedia""")

# Upload file untuk DbPinjaman.xlsx
file_pinjaman = st.file_uploader("Unggah file DbPinjaman.xlsx", type=["xlsx"])
# Upload file untuk DbSimpanan.xlsx
file_simpanan = st.file_uploader("Unggah file DbSimpanan.xlsx", type=["xlsx"])

# Validasi bahwa kedua file harus diunggah
if file_pinjaman and file_simpanan:
    try:
        # Membaca file Excel
        pinjaman = pd.read_excel(file_pinjaman, skiprows=1)
        simpanan = pd.read_excel(file_simpanan, skiprows=1)

        # Filter data sesuai kriteria
        pinjaman = pinjaman[pinjaman["Status Pinjaman"] == "AKTIF"]  # Hanya pinjaman aktif
        simpanan = simpanan[simpanan["Sts. Simpanan"] == "AKTIF"]  # Hanya anggota aktif

        # Konversi format tanggal sebelum membuat KK Pinjaman
        pinjaman["Disb. Date"] = pd.to_datetime(pinjaman["Disb. Date"]).dt.strftime("%d/%m/%Y")

        # Input filter Center ID
        center_ids = st.multiselect("Pilih Center ID", options=pinjaman["Center ID"].unique())
        if center_ids:
            pinjaman = pinjaman[pinjaman["Center ID"].isin(center_ids)]
            simpanan = simpanan[simpanan["Center ID"].isin(center_ids)]

        # Membuat dataframe untuk KK Pinjaman
        kk_pinjaman = pinjaman[[
            "Loan No.", "Client ID", "Client Name", "Meeting Day", "Product Name",
            "Loan Amount", "Outstanding", "Purpose Name", "Officer Name", "Disb. Date"
        ]].copy()
        kk_pinjaman["Saldo Buku"] = None  # Kosongkan
        kk_pinjaman["Saldo Sistem"] = kk_pinjaman["Outstanding"]  # Ambil dari Outstanding
        kk_pinjaman["SLS"] = None
        kk_pinjaman["Identitas"] = None
        kk_pinjaman["Form UK"] = None
        kk_pinjaman["Form Keanggotaan"] = None
        kk_pinjaman["Form P3"] = None
        kk_pinjaman["Akad"] = None
        kk_pinjaman["Monitoring Pembiayaan"] = None
        kk_pinjaman["KPA"] = None
        kk_pinjaman["Sesuai/ Tidak sesuai"] = None
        kk_pinjaman["KETERANGAN (Kelemahan)"] = None

        # Membuat dataframe untuk KK Simpanan
        kk_simpanan = simpanan[[
            "Account No", "Client ID", "Client Name", "Product Name", "Officer Name", "Saldo"
        ]].copy()
        kk_simpanan["Saldo Buku"] = None
        kk_simpanan["Saldo Selisih"] = None
        kk_simpanan["Buku (SPO, SWA, SSU, SPE)"] = None
        kk_simpanan["Kartu SIHARA"] = None
        kk_simpanan["Kartu Kurban"] = None
        kk_simpanan["Kartu Sipadan"] = None
        kk_simpanan["Informasi Buku Simpanan (Sesuai/Tidak Sesuai)"] = None
        kk_simpanan["KETERANGAN (Kelemahan)"] = None

        # Tampilkan dataframe
        st.subheader("KK Pinjaman")
        st.dataframe(kk_pinjaman)

        st.subheader("KK Simpanan")
        st.dataframe(kk_simpanan)

        # Tombol untuk mengunduh file Excel KK Pinjaman
        if st.button("Simpan Kertas Kerja Pinjaman"):
            try:
                # Menulis KK Pinjaman ke file
                with pd.ExcelWriter("KK Pinjaman.xlsx") as writer:
                    kk_pinjaman.to_excel(writer, sheet_name="KK Pinjaman", index=False)
                
                # Membaca file sebagai binary untuk unduhan
                with open("KK Pinjaman.xlsx", "rb") as f:
                    data = f.read()
                
                # Tombol unduhan
                st.download_button(
                    label="Unduh KK Pinjaman",
                    data=data,
                    file_name="KK Pinjaman.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat mengunduh KK Pinjaman: {e}")
        
        # Tombol untuk mengunduh file Excel KK Simpanan
        if st.button("Simpan Kertas Kerja Simpanan"):
            try:
                # Menulis KK Simpanan ke file
                with pd.ExcelWriter("KK Simpanan.xlsx") as writer:
                    kk_simpanan.to_excel(writer, sheet_name="KK Simpanan", index=False)
                
                # Membaca file sebagai binary untuk unduhan
                with open("KK Simpanan.xlsx", "rb") as f:
                    data = f.read()
                
                # Tombol unduhan
                st.download_button(
                    label="Unduh KK Simpanan",
                    data=data,
                    file_name="KK Simpanan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat mengunduh KK Simpanan: {e}")
    except Exception as e:
        st.error(f"Terjadi kesalahan dalam pemrosesan file: {e}")
