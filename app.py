import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
import math
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
st.markdown("""<style>.main { background-color: #F1F5F9; } .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; font-weight: bold; } .stTabs [data-baseweb="tab"] { font-weight: bold; font-size: 16px; }</style>""", unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

# ================= KẾT NỐI GOOGLE SHEETS =================
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except Exception as e:
    st.error(f"Chưa kết nối được Google Drive: {e}")
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state: st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB 1 ---
with tab1:
    st.dataframe(df_kho, use_container_width=True, hide_index=True)
    with st.expander("➕ Nhập thêm vật tư"):
        name = st.text_input("Tên thiết bị")
        u = st.text_input("ĐVT", value="Cái")
        q = st.number_input("Số lượng nhập", min_value=1, value=1)
        p = st.number_input("Giá gốc (VNĐ)", min_value=0, step=1000)
        if st.button("Xác nhận ghi vào Google Drive"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.cache_data.clear(); st.rerun()

# --- TAB 2 ---
with tab2:
    c_name = st.text_input("Tên khách hàng / Đơn vị")
    c_phone = st.text_input("Số điện thoại")
    c_addr = st.text_input("Địa chỉ khách hàng")
    sp_selected = st.selectbox("Sản phẩm trong kho", df_kho["Tên sản phẩm"].tolist() if not df_kho.empty else [])
    sl = st.number_input("SL bán", min_value=1, value=1)
    vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])

    if st.button("➕ Thêm vào danh sách"):
        gia_goc = float(df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]["Giá"])
        thanh_tien = sl * gia_goc * (1 + int(vat.replace("%",""))/100)
        st.session_state.quote.append({"Sản phẩm": sp_selected, "SL": sl, "Đơn giá": gia_goc, "VAT": vat, "Thành tiền": thanh_tien})
        st.rerun()

    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG: {tong:,.0f} VNĐ")

        def generate_excel():
            wb = Workbook(); ws = wb.active
            # CẤU HÌNH A4
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.append(["STT", "Sản phẩm", "ĐVT", "SL", "Đơn giá", "VAT", "Thành tiền"])
            for i, r in enumerate(st.session_state.quote, 1):
                ws.append([i, r["Sản phẩm"], "Cái", r["SL"], r["Đơn giá"], r["VAT"], r["Thành tiền"]])
            out = io.BytesIO(); wb.save(out); return out.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 TẢI EXCEL A4", generate_excel(), f"BaoGia_{c_name}.xlsx", use_container_width=True)
        with col2:
            if st.button("📸 BILL ZALO", type="primary"): st.session_state.show_zalo = True

        if st.session_state.get("show_zalo"):
            # BILL ZALO TỐI ƯU MOBILE
            html_rows = "".join([f"<tr><td>{r['Sản phẩm']}</td><td style='text-align:center'>{r['SL']}</td><td style='text-align:right'>{r['Thành tiền']:,.0f}</td></tr>" for r in st.session_state.quote])
            zalo_bill = f"<div style='border:1px solid #ddd; padding:15px; border-radius:10px; font-family:Arial; max-width:400px; margin:auto;'><h3>{COMPANY_NAME}</h3><hr><table style='width:100%'><tr><th>SP</th><th>SL</th><th>Tổng</th></tr>{html_rows}</table><hr><b>TỔNG: {tong:,.0f} đ</b></div>"
            st.markdown(zalo_bill, unsafe_allow_html=True)
            if st.button("Đóng Bill"): st.session_state.show_zalo = False; st.rerun()

# --- TAB 3 ---
with tab3:
    if st.button("✨ AI TỰ ĐỘNG SOẠN THẢO"):
        st.text_area("Hợp đồng", value=f"HỢP ĐỒNG KINH TẾ\nKhách hàng: {c_name}", height=300)
