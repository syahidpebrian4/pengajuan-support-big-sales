import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import gspread

# --- 1. LOAD DATABASE PRODUK ---
@st.cache_data
def load_product_db():
    df = pd.read_excel("prodname.xlsx", dtype={'Barcode': str})
    df['Barcode'] = df['Barcode'].str.strip()
    return df

product_db = load_product_db()

# --- 2. SETUP ---
st.set_page_config(page_title="Pengajuan Voucher", layout="wide")

data_toko = {
    "Region 1": ["Pasar Rebo", "Kelapa Gading", "Ciputat", "Alam Sutera", "Medan", "Palembang", "Pekannbaru", "Jatake", "Serang", "Batam", "Lampung", "Serpong"],
    "Region 2": ["Meruya", "Bandung", "Cibitung", "Bekasi", "Cikarang", "Cirebon", "Bogor", "Tasikmalaya", "Pakansari", "Kerawang", "Cimahi", "Tegal"],
    "Region 3": ["Sidoarjo", "Denpasar", "Semarang", "Makasar", "Yogyakarta", "Banjarmasin", "Solo", "Balikpapan", "Mastrip", "Samarinda", "Manado", "Mataram"]
}

if 'daftar_pengajuan' not in st.session_state: st.session_state.daftar_pengajuan = []

# --- 3. FUNGSI GOOGLE SHEETS ---
def simpan_ke_googlesheets(data_list):
    try:
        creds_dict = dict(st.secrets["gcp"])
        client = gspread.service_account_from_dict(creds_dict)
        sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1qcixEUIKDJHebYxqqEHnZNDqhX_kEdKcbSfJx4guvwU/edit").sheet1
        
        data_exist = sheet.get_all_values()
        next_row = len(data_exist) + 1 
        
        for item in data_list:
            row_values = [
                next_row - 5,           # A: No
                item.get('No Cust'),    # B: No Customer
                item.get('Nama'),       # C: Nama Customer
                item.get('Barcode'),    # D: Barcode
                item.get('Prodname'),   # E: Prodname
                item.get('QTY'),        # F: QTY
                item.get('HK'),         # G: HK Buyer
                item.get('No Pengajuan'),# H: No Pengajuan HK Buyer
                item.get('Harga'),      # I: Harga Customer
                item.get('gap'),        # J: Gap Value
                item.get('persen'),     # K: Gap %
                item.get('potensi'),    # L: Potensi Sales
                item.get('support'),    # M: Support Total
                item.get('rasio'),      # N: Support %
                item.get('Keterangan')  # O: Keterangan
            ]
            sheet.append_row(row_values)
            next_row += 1
    except Exception as e:
        st.error(f"Gagal simpan ke Sheets: {e}")

# --- 4. UI INPUT ---
st.title("Form Pengajuan Voucher Big Sales")
with st.container(border=True):
    col1, col2 = st.columns(2)
    region = col1.selectbox("Region", list(data_toko.keys()))
    store = col2.selectbox("Store", data_toko[region])
with st.container(border=True):
    col3, col4 = st.columns(2)
    no_cust = col3.text_input("No. Customer")
    nama_cust = col4.text_input("Nama Customer")

with st.container(border=True):
    st.write("##### 3. Input Produk")
    input_barcode = st.text_input("Masukkan Barcode")
    auto_prodname = ""
    if input_barcode:
        match = product_db[product_db['Barcode'] == input_barcode]
        if not match.empty:
            auto_prodname = match.iloc[0]['Prodname']
            st.success(f"Produk ditemukan: {auto_prodname}")
        else:
            st.warning("Barcode tidak ditemukan.")

    with st.form("input_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        barcode = c1.text_input("Barcode", value=input_barcode, disabled=True)
        prodname = c2.text_input("Prodname", value=auto_prodname, disabled=True)
        
        c3, c4, c5 = st.columns(3)
        qty = c3.number_input("QTY", min_value=0, step=1, format="%d")
        hk_buyer = c4.number_input("HK Buyer", min_value=0, step=1, format="%d")
        harga_cust = c5.number_input("Harga Customer", min_value=0, step=1, format="%d")
        no_pengajuan = st.text_input("No Pengajuan HK Buyer")
        ket = st.text_input("Keterangan")
        submitted = st.form_submit_button("Cek Perhitungan")

    if submitted:
        st.session_state.temp = {
            "barcode": input_barcode, "prodname": auto_prodname,
            "gap": hk_buyer - harga_cust,
            "persen": ((hk_buyer - harga_cust) / hk_buyer * 100) if hk_buyer != 0 else 0,
            "potensi": qty * hk_buyer,
            "support": qty * (hk_buyer - harga_cust),
            "rasio": ((qty * (hk_buyer - harga_cust)) / (qty * hk_buyer) * 100) if (qty * hk_buyer) != 0 else 0
        }

    if 'temp' in st.session_state:
        if st.button("Tambah ke Daftar"):
            data_baru = {
                "No Cust": no_cust, "Nama": nama_cust, "Barcode": input_barcode, 
                "Prodname": auto_prodname, "QTY": qty, "HK": hk_buyer, 
                "Harga": harga_cust, "No Pengajuan": no_pengajuan, "Keterangan": ket
            }
            data_baru.update(st.session_state.temp)
            st.session_state.daftar_pengajuan.append(data_baru)
            del st.session_state.temp
            st.rerun()

# --- 5. GENERATE & DOWNLOAD ---
if st.session_state.daftar_pengajuan:
    st.write("### Daftar Pengajuan")
    st.table(pd.DataFrame(st.session_state.daftar_pengajuan))
    
    if st.button("Generate & Download Excel"):
        simpan_ke_googlesheets(st.session_state.daftar_pengajuan)
        # (Lanjutkan logika Openpyxl untuk Excel di sini seperti sebelumnya)
        st.success("Data tersimpan ke Sheets!")
