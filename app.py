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

st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

# --- KẾT NỐI ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except:
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state: st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB 1 ---
with tab1:
    st.dataframe(df_kho, use_container_width=True, hide_index=True)
    with st.expander("➕ Nhập thêm vật tư"):
        name = st.text_input("Tên thiết bị")
        u = st.text_input("ĐVT", value="Cái")
        q = st.number_input("Số lượng", min_value=1, value=1)
        p = st.number_input("Giá gốc", min_value=0, step=1000)
        if st.button("Ghi dữ liệu"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.cache_data.clear(); st.rerun()

# --- TAB 2 ---
with tab2:
    c_name = st.text_input("Tên khách hàng")
    c_phone = st.text_input("Số ĐT")
    c_addr = st.text_input("Địa chỉ")
    
    sp_selected = st.selectbox("Chọn vật tư", df_kho["Tên sản phẩm"].tolist() if not df_kho.empty else [])
    sl = st.number_input("Số lượng", min_value=1, value=1)
    vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])

    if st.button("➕ Thêm vào danh sách"):
        row = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
        gia = float(row["Giá"])
        thanh_tien = sl * gia * (1 + int(vat.replace("%",""))/100)
        st.session_state.quote.append({"Sản phẩm": sp_selected, "SL": sl, "Đơn giá": gia, "VAT": vat, "Thành tiền": thanh_tien})
        st.rerun()

    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG: {tong:,.0f} VNĐ")

        def generate_excel():
            wb = Workbook(); ws = wb.active
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1 # Ép vào trang A4
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
            if st.button("🗑️ Xóa hàng"): st.session_state.quote = []; st.rerun()
        
        if st.session_state.get("show_zalo"):
            bill_html = f"<div style='border:1px solid #ccc; padding:15px; border-radius:10px;'><h3>{COMPANY_NAME}</h3><p>KH: {c_name}</p><hr><b>TỔNG: {tong:,.0f} VNĐ</b></div>"
            st.markdown(bill_html, unsafe_allow_html=True)
            if st.button("Đóng Bill"): st.session_state.show_zalo = False; st.rerun()

# --- TAB 3 ---
with tab3:
    if st.button("✨ Soạn Hợp Đồng AI"):
        st.text_area("Bản thảo", value=f"Hợp đồng kinh tế\nKhách hàng: {c_name}\nTổng tiền: {sum(item['Thành tiền'] for item in st.session_state.quote):,.0f} VNĐ", height=300)
