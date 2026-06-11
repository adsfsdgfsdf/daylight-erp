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
BANK_BRANCH = "Ngân hàng: TMCP Công thương Việt Nam (Vietinbank) - CN KCN Tiên Sơn"

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)

st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

st.markdown("""<style>.main { background-color: #F1F5F9; } .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; font-weight: bold; } .stTabs [data-baseweb="tab"] { font-weight: bold; font-size: 16px; }</style>""", unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except Exception as e:
    st.error(f"Lỗi kết nối Google Drive: {e}")
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state:
    st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

with tab1:
    st.subheader("Trạng thái tồn kho")
    st.dataframe(df_kho, use_container_width=True, hide_index=True)
    with st.expander("➕ Nhập thêm thiết bị"):
        name = st.text_input("Tên thiết bị")
        u = st.text_input("Đơn vị", value="Cái")
        q = st.number_input("Số lượng", min_value=1, value=1)
        p = st.number_input("Giá gốc", min_value=0, step=1000)
        if st.button("Xác nhận"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.cache_data.clear()
            st.rerun()

with tab2:
    if not df_kho.empty:
        sp_selected = st.selectbox("Chọn sản phẩm", df_kho["Tên sản phẩm"].tolist())
        sl = st.number_input("SL bán", min_value=1)
        vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])
        if st.button("➕ Thêm vào báo giá"):
            row = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
            thanh_tien = sl * float(row["Giá"]) * (1 + int(vat.replace("%",""))/100)
            st.session_state.quote.append({"Sản phẩm": sp_selected, "SL": sl, "Đơn giá": float(row["Giá"]), "VAT": vat, "Thành tiền": thanh_tien})
            st.rerun()
    
    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG THANH TOÁN: {tong:,.0f} VNĐ")
        if st.button("🗑️ XÓA LÀM LẠI"):
            st.session_state.quote = []
            st.rerun()

with tab3:
    st.subheader("Soạn hợp đồng")
    if st.button("✨ Soạn thảo"):
        st.write("Đang phát triển...")
