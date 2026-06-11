import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Cấu hình giao diện trang web
st.set_page_config(page_title="Daylight ERP - Kho Hàng", layout="wide")

st.title("☀️ QUẢN LÝ KHO - DAYLIGHT VIETNAM")

try:
    # Kết nối tới Google Sheets thông qua Secrets
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    # Đọc dữ liệu từ tab 'KHO'
    # usecols=None để đọc tất cả các cột, hoặc anh có thể sửa thành [0, 1, 2, 3]
    df = conn.read(worksheet="KHO")
    
    # Kiểm tra xem dữ liệu có trống không
    if df.empty:
        st.warning("Dữ liệu trong tab KHO đang trống!")
    else:
        st.success("Đã kết nối thành công kho hàng!")
        st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error(f"Lỗi kết nối tới Google Sheets: {e}")
    st.info("Kiểm tra lại xem file Secrets đã lưu chưa hoặc tab KHO có tồn tại không.")

# Nút làm mới dữ liệu
if st.button("Làm mới dữ liệu"):
    st.rerun()
