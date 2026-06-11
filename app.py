import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Cấu hình trang
st.set_page_config(page_title="Daylight ERP", layout="centered")

st.title("☀️ DAYLIGHT VIETNAM ERP")

# Kết nối Google Sheets
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(worksheet="KHO")
    
    # Hiển thị Tab
    tab1, tab2 = st.tabs(["📦 KHO HÀNG", "📄 BÁO GIÁ"])
    
    with tab1:
        st.subheader("Tồn kho thực tế")
        if not df.empty:
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.warning("Kho hàng trống.")
            
    with tab2:
        st.subheader("Lập báo giá nhanh")
        if not df.empty:
            san_pham = st.selectbox("Chọn sản phẩm", df["Tên sản phẩm"].tolist())
            sl = st.number_input("Số lượng", min_value=1)
            if st.button("Tính tiền"):
                don_gia = df.loc[df["Tên sản phẩm"] == san_pham, "Giá"].values[0]
                thanh_tien = sl * don_gia
                st.success(f"Tổng tiền: {thanh_tien:,.0f} VNĐ")
        else:
            st.error("Không có dữ liệu để báo giá.")

except Exception as e:
    st.error("Lỗi kết nối dữ liệu. Vui lòng kiểm tra lại cấu hình Secrets.")
    with st.expander("Chi tiết lỗi"):
        st.write(e)
