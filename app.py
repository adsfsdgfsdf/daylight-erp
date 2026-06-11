import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
import time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- CẤU HÌNH ---
COMPANY_NAME = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
COMPANY_MST = "2301380133"
COMPANY_ADDR = "Đông Lâu, Đại Đồng, Tiên Du, Bắc Ninh"

# Cấu hình UI
st.set_page_config(page_title="DAYLIGHT ERP FULL", layout="centered")
st.markdown("""<style>.stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }</style>""", unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM ERP")

# --- KẾT NỐI (TỰ ĐỘNG LÀM MỚI SAU 5 PHÚT = 300 GIÂY) ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl=300) 
except:
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state: st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB 1: KHO ---
with tab1:
    st.dataframe(df_kho, use_container_width=True)
    if st.expander("➕ Nhập hàng"):
        name = st.text_input("Tên thiết bị")
        q = st.number_input("Số lượng", min_value=1)
        p = st.number_input("Giá gốc", min_value=0)
        if st.button("Ghi dữ liệu"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.rerun()

# --- TAB 2: BÁO GIÁ ---
with tab2:
    col1, col2 = st.columns(2)
    with col1: c_name = st.text_input("Tên khách")
    with col2: c_phone = st.text_input("SĐT")
    
    sp_selected = st.selectbox("Chọn vật tư", df_kho["Tên sản phẩm"].tolist() if not df_kho.empty else [])
    sl = st.number_input("SL bán", min_value=1)
    vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])

    if st.button("➕ Thêm"):
        row = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
        st.session_state.quote.append({"Sản phẩm": sp_selected, "SL": sl, "Đơn giá": float(row["Giá"]), "VAT": vat, "Thành tiền": sl * float(row["Giá"]) * (1 + int(vat.replace("%",""))/100)})
        st.rerun()

    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        
        # HÀM XUẤT FILE CHUẨN A4
        def generate_excel():
            wb = Workbook(); ws = wb.active
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1
            ws.append(["STT", "Sản phẩm", "ĐVT", "SL", "Đơn giá", "VAT", "Thành tiền"])
            for i, r in enumerate(st.session_state.quote, 1):
                ws.append([i, r.get("Sản phẩm", ""), "Cái", r.get("SL", 0), r.get("Đơn giá", 0), r.get("VAT", "0%"), r.get("Thành tiền", 0)])
            out = io.BytesIO(); wb.save(out); return out.getvalue()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("📥 Excel A4", generate_excel(), f"BaoGia_{c_name}.xlsx", use_container_width=True)
        with col2:
            if st.button("📸 Bill Zalo"): st.session_state.show_zalo = True
        with col3:
            if st.button("🗑️ Xóa hàng"):
                if st.session_state.quote: st.session_state.quote.pop()
                st.rerun()
        
        if st.session_state.get("show_zalo"):
            bill = f"<div style='border:1px solid #ccc; padding:15px; border-radius:10px;'><h3>{COMPANY_NAME}</h3><p>KH: {c_name}</p><hr>"+ "".join([f"<p>{r.get('Sản phẩm')}: {r.get('Thành tiền',0):,.0f} đ</p>" for r in st.session_state.quote]) + "</div>"
            st.markdown(bill, unsafe_allow_html=True)
            if st.button("Đóng Bill"): st.session_state.show_zalo = False; st.rerun()

# --- TAB 3: HỢP ĐỒNG ---
with tab3:
    if st.button("✨ AI SOẠN HỢP ĐỒNG"):
        st.text_area("Nội dung", value=f"HỢP ĐỒNG KINH TẾ\nBên A: {c_name}", height=300)
