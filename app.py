import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO

st.set_page_config(page_title="Pengajuan Voucher", layout="wide")

# --- DATA REGION & TOKO ---
data_toko = {
    "Region 1": ["Pasar Rebo", "Kelapa Gading", "Ciputat", "Alam Sutera", "Medan", "Palembang", "Pekannbaru", "Jatake", "Serang", "Batam", "Lampung", "Serpong"],
    "Region 2": ["Meruya", "Bandung", "Cibitung", "Bekasi", "Cikarang", "Cirebon", "Bogor", "Tasikmalaya", "Pakansari", "Kerawang", "Cimahi", "Tegal"],
    "Region 3": ["Sidoarjo", "Denpasar", "Semarang", "Makasar", "Yogyakarta", "Banjarmasin", "Solo", "Balikpapan", "Mastrip", "Samarinda", "Manado", "Mataram"]
}

if 'daftar_pengajuan' not in st.session_state: st.session_state.daftar_pengajuan = []

st.title("Form Pengajuan Voucher Big Sales")

# --- KOTAK 1 & 2 ---
with st.container(border=True):
    col1, col2 = st.columns(2)
    region = col1.selectbox("Region", list(data_toko.keys()))
    store = col2.selectbox("Store", data_toko[region])
with st.container(border=True):
    col3, col4 = st.columns(2)
    no_cust = col3.text_input("No. Customer")
    nama_cust = col4.text_input("Nama Customer")

# --- KOTAK 3: INPUT PRODUK ---
with st.container(border=True):
    st.write("##### 3. Input Produk")
    with st.form("input_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        barcode = c1.text_input("Barcode")
        prodname = c2.text_input("Prodname")
        c3, c4, c5 = st.columns(3)
        qty = c3.number_input("QTY", min_value=0, step=1, format="%d")
        hk_buyer = c4.number_input("HK Buyer", min_value=0, step=1, format="%d")
        harga_cust = c5.number_input("Harga Customer", min_value=0, step=1, format="%d")
        no_pengajuan = st.text_input("No Pengajuan HK Buyer")
        submitted = st.form_submit_button("Cek Perhitungan")
    
    if submitted:
        st.session_state.temp = {
            "gap": hk_buyer - harga_cust,
            "persen": ((hk_buyer - harga_cust) / hk_buyer * 100) if hk_buyer != 0 else 0,
            "potensi": qty * hk_buyer,
            "support": qty * (hk_buyer - harga_cust),
            "rasio": ((qty * (hk_buyer - harga_cust)) / (qty * hk_buyer) * 100) if (qty * hk_buyer) != 0 else 0
        }

    if 'temp' in st.session_state:
        m1, m2, m3 = st.columns(3)
        m1.metric("Gap (Value)", f"{st.session_state.temp['gap']:,.0f}")
        m2.metric("Gap (%)", f"{st.session_state.temp['persen']:,.2f}%")
        m3.metric("Potensi Sales", f"{st.session_state.temp['potensi']:,.0f}")
        m4, m5 = st.columns(2)
        m4.metric("Support Voucher", f"{st.session_state.temp['support']:,.0f}")
        m5.metric("Rasio Voucher", f"{st.session_state.temp['rasio']:,.2f}%")
        
        if st.button("Tambah ke Daftar"):
            st.session_state.daftar_pengajuan.append({
                "No Cust": no_cust, "Nama": nama_cust, "Barcode": barcode, "Prodname": prodname,
                "QTY": qty, "HK": hk_buyer, "Harga": harga_cust, "No Pengajuan": no_pengajuan,
                **st.session_state.temp
            })
            del st.session_state.temp
            st.rerun()

# --- DAFTAR & GENERATE ---
if st.session_state.daftar_pengajuan:
    st.write("### Daftar Pengajuan")
    st.table(pd.DataFrame(st.session_state.daftar_pengajuan))
    
    if st.button("Generate & Download Excel"):
        try:
            wb = openpyxl.load_workbook("VCR TEMPLATE.xlsx")
            ws = wb.active
            center_align = Alignment(horizontal='center', vertical='center')

            # 1. Isi Nama TTD di M9 berdasarkan Region
            nama_ttd = {
                "Region 1": "AZKAHADI PUTRA",
                "Region 2": "DESKY BAYU AJI",
                "Region 3": "ADE CHANDRA"
            }
            # Menulis ke M9 (kolom 13, baris 9)
            ws.cell(row=9, column=13, value=nama_ttd.get(region, ""))
            # Pastikan rata tengah
            ws.cell(row=9, column=13).alignment = center_align
            
            # 2. Bersihkan Merge Cells di area data (baris 14 ke bawah) agar aman
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row >= 14:
                    ws.unmerge_cells(str(merged_range))

            # 3. Tulis Data Produk mulai dari baris 14
            formats = {2: '@', 3: '@', 4: '@', 5: '@', 6: '0', 7: '#,##0', 8: '@', 9: '#,##0', 10: '#,##0', 12: '#,##0', 13: '#,##0'}
            
            for i, item in enumerate(st.session_state.daftar_pengajuan):
                row = 14 + i
                data_row = [
                    i + 1, str(item.get('No Cust', '')), str(item.get('Nama', '')), 
                    str(item.get('Barcode', '')), str(item.get('Prodname', '')),
                    int(item.get('QTY', 0)), int(item.get('HK', 0)), 
                    str(item.get('No Pengajuan', '')), int(item.get('Harga', 0)), 
                    int(item.get('gap', 0)), item.get('persen', 0) / 100,
                    int(item.get('potensi', 0)), int(item.get('support', 0)), 
                    item.get('rasio', 0) / 100
                ]
                
                for col_idx, value in enumerate(data_row, start=1):
                    cell = ws.cell(row=row, column=col_idx, value=value)
                    cell.alignment = center_align 
                    
                    if col_idx in [11, 14]: cell.number_format = '0.00%'
                    elif col_idx in formats: cell.number_format = formats[col_idx]

            buffer = BytesIO()
            wb.save(buffer)
            st.download_button("Download Excel Rapi", buffer.getvalue(), f"Pengajuan_{store}.xlsx")
            st.success("File berhasil disiapkan dengan TTD yang sesuai!")
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

    if st.button("Reset Daftar"):
        st.session_state.daftar_pengajuan = []
        st.rerun()
