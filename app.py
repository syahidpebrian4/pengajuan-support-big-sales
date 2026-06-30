import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import gspread

# --- LOAD DATABASE PRODUK ---
@st.cache_data
def load_product_db():
    df = pd.read_excel("prodname.xlsx", dtype={'Barcode': str})
    df['Barcode'] = df['Barcode'].str.strip()
    return df

product_db = load_product_db()

# --- SETUP ---
st.set_page_config(page_title="Pengajuan Voucher", layout="wide")

# --- DATA REGION & TOKO ---
data_toko = {
    "Region 1": ["Pasar Rebo", "Kelapa Gading", "Ciputat", "Alam Sutera", "Medan", "Palembang", "Pekannbaru", "Jatake", "Serang", "Batam", "Lampung", "Serpong"],
    "Region 2": ["Meruya", "Bandung", "Cibitung", "Bekasi", "Cikarang", "Cirebon", "Bogor", "Tasikmalaya", "Pakansari", "Kerawang", "Cimahi", "Tegal"],
    "Region 3": ["Sidoarjo", "Denpasar", "Semarang", "Makasar", "Yogyakarta", "Banjarmasin", "Solo", "Balikpapan", "Mastrip", "Samarinda", "Manado", "Mataram"]
}

if 'daftar_pengajuan' not in st.session_state: st.session_state.daftar_pengajuan = []

# --- FUNGSI GOOGLE SHEETS ---
def simpan_ke_googlesheets(data_list):
    try:
        creds_dict = dict(st.secrets["gcp"])
        client = gspread.service_account_from_dict(creds_dict)
        sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1qcixEUIKDJHebYxqqEHnZNDqhX_kEdKcbSfJx4guvwU/edit").sheet1
        
        # Cari baris kosong pertama
        data_exist = sheet.get_all_values()
        next_row = len(data_exist) + 1 
        
        for item in data_list:
            # Urutan harus pas dengan kolom A sampai O di Spreadsheet
            row_values = [
                next_row - 5,              # A: No
                item.get('No Cust'),       # B: No Customer
                item.get('Nama'),          # C: Nama Customer
                item.get('Barcode'),       # D: Barcode
                item.get('Prodname'),      # E: Prodname
                item.get('QTY'),           # F: QTY
                item.get('HK'),            # G: HK Buyer
                item.get('No Pengajuan'),  # H: No Pengajuan HK Buyer
                item.get('Harga'),         # I: Harga Customer
                item.get('gap'),           # J: Gap Value
                item.get('persen'),        # K: Gap %
                item.get('potensi'),       # L: Potensi Sales
                item.get('support'),       # M: Support Total
                item.get('rasio'),         # N: Support %
                item.get('Keterangan')     # O: Keterangan
            ]
            sheet.append_row(row_values)
            next_row += 1
    except Exception as e:
        st.error(f"Gagal simpan ke Sheets: {e}")

# --- UI INPUT ---
st.title("Form Pengajuan Voucher Big Sales")
# ... (UI Input sama seperti sebelumnya)

# --- TOMBOL TAMBAH KE DAFTAR ---
if 'temp' in st.session_state:
    if st.button("Tambah ke Daftar"):
        # Gabungkan data
        data_baru = {
            "No Cust": no_cust, "Nama": nama_cust, "Barcode": input_barcode, 
            "Prodname": auto_prodname, "QTY": qty, "HK": hk_buyer, 
            "Harga": harga_cust, "No Pengajuan": no_pengajuan, "Keterangan": ket
        }
        data_baru.update(st.session_state.temp)
        st.session_state.daftar_pengajuan.append(data_baru)
        del st.session_state.temp
        st.rerun()

# --- GENERATE EXCEL ---
# ... (Logika generate excel tetap sama, pastikan urutan data_row sama dengan row_values di Sheets)
