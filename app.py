import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
import math
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
st.markdown("""<style>.main { background-color: #F1F5F9; } .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #1E3A8A; color: white; font-weight: bold; }</style>""", unsafe_allow_html=True)

st.title("☀️ DAYLIGHT VIETNAM")

# --- KẾT NỐI (LÀM MỚI SAU 300 GIÂY = 5 PHÚT) ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl=300).dropna(how="all")
except:
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state: st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

# --- TAB KHO ---
with tab1:
    st.dataframe(df_kho, use_container_width=True, hide_index=True)
    if st.expander("➕ Nhập hàng"):
        name = st.text_input("Tên thiết bị")
        q = st.number_input("Số lượng", min_value=1)
        p = st.number_input("Giá gốc", min_value=0)
        if st.button("Ghi dữ liệu"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.cache_data.clear(); st.rerun()

# --- TAB BÁO GIÁ ---
with tab2:
    c_name = st.text_input("Tên khách hàng")
    c_phone = st.text_input("SĐT")
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
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG: {tong:,.0f} VNĐ")
        
        def generate_pro_excel():
            wb = Workbook(); ws = wb.active
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1
            # Header
            ws['A1'] = COMPANY_NAME; ws['A1'].font = Font(bold=True, size=12, color="1E3A8A")
            ws['A2'] = f"Địa chỉ: {COMPANY_ADDR}"; ws['A3'] = f"MST: {COMPANY_MST}"
            ws.merge_cells('A5:G5'); ws['A5'] = "BẢNG BÁO GIÁ CHI TIẾT"; ws['A5'].font = Font(bold=True, size=16); ws['A5'].alignment = Alignment(horizontal='center')
            ws['A7'] = f"Kính gửi: {c_name.upper()}"; ws['A7'].font = Font(bold=True); ws['A8'] = f"SĐT: {c_phone}"
            # Bảng
            headers = ["STT", "Tên sản phẩm", "ĐVT", "SL", "Đơn giá", "VAT", "Thành tiền"]
            for c, h in enumerate(headers, 1):
                cell = ws.cell(row=11, column=c, value=h)
                cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.alignment = Alignment(horizontal='center')
                ws.column_dimensions[chr(64+c)].width = 15 if c > 2 else 25
            for i, r in enumerate(st.session_state.quote, 1):
                ws.append([i, r.get("Sản phẩm"), "Cái", r.get("SL"), r.get("Đơn giá"), r.get("VAT"), r.get("Thành tiền")])
            # Footer
            curr_r = len(st.session_state.quote) + 12
            ws.merge_cells(f'A{curr_r}:F{curr_r}'); ws.cell(row=curr_r, column=7, value=tong).font = Font(bold=True)
            ws.cell(row=curr_r+2, column=1, value="THÔNG TIN THANH TOÁN:").font = Font(bold=True)
            ws.cell(row=curr_r+3, column=1, value=BANK_STK); ws.cell(row=curr_r+4, column=1, value=BANK_BRANCH)
            out = io.BytesIO(); wb.save(out); return out.getvalue()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("📥 Excel A4", generate_pro_excel(), f"BaoGia_{c_name}.xlsx", use_container_width=True)
        with col2:
            if st.button("📸 Bill Zalo"): st.session_state.show_zalo = True
        with col3:
            if st.button("🗑️ Xóa hàng"):
                if st.session_state.quote: st.session_state.quote.pop()
                st.rerun()
        
        if st.session_state.get("show_zalo"):
            bill = f"<div style='border:1px solid #ccc; padding:15px; border-radius:10px;'><h3>{COMPANY_NAME}</h3><p>KH: {c_name}</p><hr>"+ "".join([f"<p>{r.get('Sản phẩm')} (VAT {r.get('VAT')}): {r.get('Thành tiền',0):,.0f} đ</p>" for r in st.session_state.quote]) + "<hr><b>TỔNG: " + f"{tong:,.0f} đ</b></div>"
            st.markdown(bill, unsafe_allow_html=True)
            if st.button("Đóng Bill"): st.session_state.show_zalo = False; st.rerun()

# --- TAB 3: HỢP ĐỒNG ---
with tab3:
    if st.button("✨ AI SOẠN HỢP ĐỒNG"):
        st.text_area("Nội dung", value=f"HỢP ĐỒNG KINH TẾ\nBên A: {c_name}", height=300)
