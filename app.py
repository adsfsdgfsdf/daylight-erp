import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

# Cấu hình giao diện di động
st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #F1F5F9; }
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; font-size: 16px; }
    </style>
    """, unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

# ================= KẾT NỐI GOOGLE SHEETS =================
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    # Tự động làm mới bảng hàng hóa từ Google Sheets sau mỗi 1 phút (ttl="1m")
    df_kho = conn.read(worksheet="KHO", ttl="1m").dropna(how="all")
except Exception as e:
    st.error(f"Chưa kết nối được Google Drive: {e}")
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state:
    st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB 1: QUẢN LÝ KHO ---
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
                # Logic thêm dòng mới
                new_row = pd.DataFrame([{"Tên sản phẩm": name, "ĐVT": u, "Tồn": q, "Giá": p}])
                updated_df = pd.concat([df_kho, new_row], ignore_index=True)

                # Lệnh ghi đè trực tiếp lên file Google Sheets trên Drive
                conn.update(worksheet="KHO", data=updated_df)
                
                # Xóa bộ nhớ đệm ngay lập tức để app hiển thị vật tư mới luôn
                st.cache_data.clear()
                
                st.success(f"Đã đồng bộ vĩnh viễn sản phẩm '{name}' lên Google Drive!")
                st.rerun()
            else:
                st.warning("Vui lòng nhập tên sản phẩm!")

# --- TAB 2: LẬP BÁO GIÁ ---
with tab2:
    st.subheader("Thông tin khách hàng")
    c_name = st.text_input("Tên khách hàng / Đơn vị")
    c_phone = st.text_input("Số điện thoại")

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
            # Tìm giá của sản phẩm được chọn
            row_sp = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
            gia_goc = float(row_sp["Giá"])
            thanh_tien = sl * gia_goc * (1 + int(vat.replace("%",""))/100)

            st.session_state.quote.append({
                "Sản phẩm": sp_selected, "SL": sl, "Đơn giá": gia_goc, "VAT": vat, "Thành tiền": thanh_tien
            })
            st.toast("Đã thêm vào giỏ hàng báo giá!")
    else:
        st.warning("Không có sản phẩm nào trong kho để lập báo giá.")

    if st.session_state.quote:
        st.markdown("---")
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True, hide_index=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG THANH TOÁN: {tong:,.0f} VNĐ")

        # Chia 2 nút bấm nằm ngang hàng cho đẹp giao diện di động
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            # Mã hóa dữ liệu utf-8-sig để Excel trên máy tính đọc không bao giờ lỗi font tiếng Việt
            csv = df_q.to_csv(index=False).encode('utf-8-sig')
            # Tự động đặt tên file theo tên khách hàng nhập vào
            ten_file = f"Bao_Gia_{c_name}.csv" if c_name else "Bao_Gia_Daylight.csv"
            
            st.download_button(
                label="📥 TẢI FILE BÁO GIÁ",
                data=csv,
                file_name=ten_file,
                mime="text/csv",
                use_container_width=True
            )
            
        with col_btn2:
            if st.button("🗑️ XÓA HẾT DÒNG", use_container_width=True):
                st.session_state.quote = []
                st.rerun()

# --- TAB 3: HỢP ĐỒNG ---
with tab3:
    st.subheader("Soạn hợp đồng nhanh")
    if st.button("✨ AI TỰ ĐỘNG SOẠN THẢO"):
        content = f"CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nKhách hàng: {c_name.upper()}\nSĐT: {c_phone}\n..."
        st.text_area("Bản xem trước nội dung", value=content, height=300)
