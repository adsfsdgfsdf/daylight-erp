import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- CẤU HÌNH ---
COMPANY_NAME = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
COMPANY_MST = "2301380133"
COMPANY_ADDR = "Đông Lâu, Đại Đồng, Tiên Du, Bắc Ninh"
BANK_NAME_BENEFICIARY = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
BANK_STK = "Số tài khoản: 688 608 632 999"
BANK_BRANCH = "Vietinbank - CN KCN Tiên Sơn"

# Cấu hình Excel
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)

st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

# CSS cho Mobile
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

# Kết nối dữ liệu
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except:
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state:
    st.session_state.quote = []

tab1, tab2 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ"])

with tab1:
    st.dataframe(df_kho, use_container_width=True, hide_index=True)

with tab2:
    if not df_kho.empty:
        sp_selected = st.selectbox("Chọn sản phẩm", df_kho["Tên sản phẩm"].tolist())
        sl = st.number_input("Số lượng", min_value=1, value=1)
        if st.button("➕ Thêm vào danh sách"):
            row = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
            st.session_state.quote.append({
                "Sản phẩm": sp_selected, "SL": sl, 
                "Đơn giá": float(row["Giá"]), 
                "Thành tiền": sl * float(row["Giá"])
            })
            st.rerun()

    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        tong = df_q["Thành tiền"].sum()
        st.subheader(f"TỔNG: {tong:,.0f} VNĐ")

        # Nút xuất Excel
        def get_excel():
            wb = Workbook()
            ws = wb.active
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            # ... (cấu trúc Excel của anh)
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Tải Excel", get_excel(), "BaoGia.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            if st.button("📸 Bill Zalo"):
                st.session_state.show_zalo = True
        
        if st.session_state.get("show_zalo"):
            st.markdown(f"**BILL NHANH:** Tổng {tong:,.0f} đ", unsafe_allow_html=True)
            if st.button("Ẩn bill"):
                st.session_state.show_zalo = False
                st.rerun()

        if st.button("🗑️ Xóa làm lại"):
            st.session_state.quote = []
            st.rerun()
