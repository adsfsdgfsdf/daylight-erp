import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io

# Thư viện xử lý Excel chuẩn form
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- CẤU HÌNH CỐ ĐỊNH ---
COMPANY_NAME = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
COMPANY_MST = "2301380133"
COMPANY_ADDR = "Đông Lâu, Đại Đồng, Tiên Du, Bắc Ninh"
BANK_NAME_BENEFICIARY = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
BANK_STK = "Số tài khoản: 688 608 632 999"
BANK_BRANCH = "Ngân hàng: TMCP Công thương Việt Nam (Vietinbank) - CN KCN Tiên Sơn"

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #F1F5F9; }
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; font-weight: bold; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; font-size: 16px; }
    </style>
    """, unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

# ================= KẾT NỐI GOOGLE SHEETS =================
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except Exception as e:
    st.error(f"Chưa kết nối được Google Drive: {e}")
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state:
    st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# ================= TAB 1: QUẢN LÝ KHO =================
with tab1:
    st.subheader("Trạng thái tồn kho thực tế")
    if not df_kho.empty:
        st.dataframe(df_kho, use_container_width=True, hide_index=True)
    else:
        st.info("Kho hàng trống hoặc đang chờ kết nối dữ liệu...")

    with st.expander("➕ Nhập thêm vật tư / Thiết bị"):
        name = st.text_input("Tên thiết bị")
        u = st.text_input("Đơn vị tính (ĐVT)", value="Cái")
        q = st.number_input("Số lượng nhập", min_value=1, value=1)
        p = st.number_input("Giá gốc (VNĐ)", min_value=0, step=1000)

        if st.button("Xác nhận ghi vào Google Drive"):
            if name:
                new_row = pd.DataFrame([{"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p}])
                updated_df = pd.concat([df_kho, new_row], ignore_index=True)
                conn.update(worksheet="KHO", data=updated_df)
                st.cache_data.clear()
                st.success(f"Đã đồng bộ sản phẩm '{name}' lên kho!")
                st.rerun()
            else:
                st.warning("Vui lòng nhập tên sản phẩm!")

# ================= TAB 2: LẬP BÁO GIÁ =================
with tab2:
    st.subheader("Thông tin khách hàng")
    c_name = st.text_input("Tên khách hàng / Đơn vị")
    c_phone = st.text_input("Số điện thoại")
    c_addr = st.text_input("Địa chỉ khách hàng")

    st.subheader("Chọn vật tư xuất bán")
    if not df_kho.empty:
        list_sp = df_kho["Tên sản phẩm"].tolist()
        sp_selected = st.selectbox("Sản phẩm trong kho", list_sp)

        col1, col2 = st.columns(2)
        with col1:
            sl = st.number_input("SL bán", min_value=1, value=1)
        with col2:
            vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])

        if st.button("➕ Thêm vào danh sách"):
            row_sp = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
            gia_goc = float(row_sp["Giá"])
            thanh_tien = sl * gia_goc * (1 + int(vat.replace("%",""))/100)

            st.session_state.quote.append({
                "Sản phẩm": sp_selected, "SL": sl, "Đơn giá": gia_goc, "VAT": vat, "Thành tiền": thanh_tien
            })
            st.toast("Đã thêm vào giỏ hàng!")
    else:
        st.warning("Không có sản phẩm nào trong kho.")

    if st.session_state.quote:
        st.markdown("---")
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True, hide_index=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.
