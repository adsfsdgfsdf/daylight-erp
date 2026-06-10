import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Cấu hình giao diện di động
st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

# --- PHONG CÁCH THƯƠNG MẠI ---
st.markdown("""
    <style>
    .main { background-color: #F1F5F9; }
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; font-size: 16px; }
    </style>
    """, unsafe_allow_html=True)

# --- KHỞI TẠO DỮ LIỆU TẠM (Sẽ kết nối Google Sheets ở bước sau) ---
if 'db' not in st.session_state:
    st.session_state.db = [
        {"Tên sản phẩm": "Thiết bị ATS 100A", "ĐVT": "Bộ", "Tồn": 50, "Giá": 1200000},
        {"Tên sản phẩm": "MCB 2P 32A Feeo", "ĐVT": "Cái", "Tồn": 100, "Giá": 150000}
    ]
if 'quote' not in st.session_state:
    st.session_state.quote = []

st.title("☀️ DAYLIGHT VIETNAM")

tab1, tab2, tab3 = st.tabs(["📦 KHO", "📄 BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB QUẢN LÝ KHO ---
with tab1:
    st.subheader("Trạng thái tồn kho")
    df_inventory = pd.DataFrame(st.session_state.db)
    st.dataframe(df_inventory, use_container_width=True, hide_index=True)
    
    with st.expander("➕ Nhập hàng mới"):
        name = st.text_input("Tên thiết bị")
        q = st.number_input("Số lượng", min_value=1)
        p = st.number_input("Giá gốc", min_value=0)
        u = st.text_input("ĐVT", value="Cái")
        if st.button("Xác nhận nhập"):
            st.session_state.db.append({"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p})
            st.success("Đã nhập hàng!")
            st.rerun()

# --- TAB LẬP BÁO GIÁ ---
with tab2:
    st.subheader("Thông tin khách hàng")
    c_name = st.text_input("Tên khách hàng")
    c_phone = st.text_input("Số điện thoại")
    
    st.subheader("Chọn vật tư")
    list_sp = [x["Tên sản phẩm"] for x in st.session_state.db]
    sp_selected = st.selectbox("Sản phẩm", list_sp)
    
    col1, col2 = st.columns(2)
    with col1:
        sl = st.number_input("SL bán", min_value=1)
    with col2:
        vat = st.selectbox("VAT", ["0%", "5%", "8%", "10%"])
        
    if st.button("➕ Thêm vào báo giá"):
        gia_goc = next(item["Giá"] for item in st.session_state.db if item["Tên sản phẩm"] == sp_selected)
        thanh_tien = sl * gia_goc * (1 + int(vat.replace("%",""))/100)
        st.session_state.quote.append({"Sản phẩm": sp_selected, "SL": sl, "Đơn giá": gia_goc, "VAT": vat, "Thành tiền": thanh_tien})
        st.toast("Đã thêm!")

    if st.session_state.quote:
        st.markdown("---")
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True, hide_index=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG: {tong:,.0f} VNĐ")
        
        if st.button("🟢 XUẤT FILE BÁO GIÁ"):
            st.success("Tính năng xuất file đang được khởi tạo...")

# --- TAB HỢP ĐỒNG ---
with tab3:
    st.subheader("Soạn hợp đồng nhanh")
    if st.button("✨ AI SOẠN NỘI DUNG"):
        content = f"CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nKhách hàng: {c_name.upper()}\nSĐT: {c_phone}\n..."
        st.text_area("Bản xem trước", value=content, height=300)
