import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import io
from openpyxl import Workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH BẢO MẬT ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if not st.session_state.password_correct:
        st.title("🔒 HỆ THỐNG BẢO MẬT DAYLIGHT")
        pwd = st.text_input("Vui lòng nhập mật khẩu:", type="password")
        if st.button("Đăng nhập"):
            if pwd == "170401": 
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("Sai mật khẩu!")
        st.stop()
    return True

check_password()

# --- CẤU HÌNH CÔNG TY ---
COMPANY_NAME = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
COMPANY_MST = "2301380133"
COMPANY_ADDR = "Đông Lâu, Đại Đồng, Tiên Du, Bắc Ninh"
BANK_NAME_BENEFICIARY = "CÔNG TY TNHH DAYLIGHT VIỆT NAM"
BANK_STK = "Số tài khoản: 688 608 632 999"
BANK_BRANCH = "Ngân hàng: TMCP Công thương Việt Nam (Vietinbank) - CN KCN Tiên Sơn"

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

st.set_page_config(page_title="DAYLIGHT VIETNAM ERP", layout="centered")
st.title("☀️ DAYLIGHT VIETNAM")

# --- KẾT NỐI DỮ LIỆU ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_kho = conn.read(worksheet="KHO", ttl=300).dropna(how="all")
except:
    df_kho = pd.DataFrame(columns=["Tên sản phẩm", "ĐVT", "Tồn", "Giá"])

if 'quote' not in st.session_state: st.session_state.quote = []

tab1, tab2, tab3 = st.tabs(["📦 KHO HÀNG", "📄 LẬP BÁO GIÁ", "🤝 HỢP ĐỒNG"])

with tab1:
    st.dataframe(df_kho, use_container_width=True, hide_index=True)
    with st.expander("➕ Nhập hàng mới"):
        name = st.text_input("Tên thiết bị")
        q = st.number_input("Số lượng", min_value=1)
        p = st.number_input("Giá gốc", min_value=0)
        if st.button("Lưu vào kho"):
            new_row = pd.DataFrame([{"Tên sản phẩm": name, "Tồn": q, "Giá": p}])
            conn.update(worksheet="KHO", data=pd.concat([df_kho, new_row], ignore_index=True))
            st.cache_data.clear(); st.rerun()

with tab2:
    c_name = st.text_input("Tên khách hàng")
    c_phone = st.text_input("SĐT khách hàng")
    c_addr = st.text_input("Địa chỉ khách")
    sp_selected = st.selectbox("Chọn vật tư", df_kho["Tên sản phẩm"].tolist() if not df_kho.empty else [])
    sl = st.number_input("Số lượng bán", min_value=1)
    vat = st.selectbox("Thuế VAT", ["0%", "5%", "8%", "10%"])

    if st.button("Thêm vào báo giá"):
        row = df_kho[df_kho["Tên sản phẩm"] == sp_selected].iloc[0]
        st.session_state.quote.append({
            "Sản phẩm": sp_selected, "SL": sl, "Đơn giá": float(row["Giá"]), 
            "VAT": vat, "Thành tiền": sl * float(row["Giá"]) * (1 + int(vat.replace("%",""))/100)
        })
        st.rerun()

    if st.session_state.quote:
        df_q = pd.DataFrame(st.session_state.quote)
        st.dataframe(df_q, use_container_width=True)
        tong = sum(item["Thành tiền"] for item in st.session_state.quote)
        st.error(f"TỔNG CỘNG: {tong:,.0f} VNĐ")
        
        if st.button("📸 TẠO ẢNH BÁO GIÁ"):
            st.session_state.show_capture = True
        
        if st.session_state.get("show_capture"):
            st.markdown(f"""
            <div style='background-color: white; border: 2px solid #1E3A8A; padding: 20px; border-radius: 15px; color: #333;'>
                <h2 style='color: #1E3A8A; text-align: center;'>BÁO GIÁ DAYLIGHT</h2>
                <p><b>Khách:</b> {c_name}</p><hr>
                {''.join([f"<p>{r['Sản phẩm']} x {r['SL']} = <b>{r['Thành tiền']:,.0f} đ</b></p>" for r in st.session_state.quote])}
                <hr><h3 style='text-align: right; color: #E11D48;'>TỔNG: {tong:,.0f} đ</h3>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Đóng Ảnh"): st.session_state.show_capture = False; st.rerun()

        def generate_full_excel():
            wb = Workbook(); ws = wb.active
            ws.title = "BaoGia"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5)
            
            ws['A1'] = COMPANY_NAME; ws['A1'].font = Font(bold=True)
            ws['A2'] = f"Địa chỉ: {COMPANY_ADDR}"; ws['A3'] = f"MST: {COMPANY_MST}"
            ws.merge_cells('A5:G5'); ws['A5'] = "BẢNG BÁO GIÁ CHI TIẾT"; ws['A5'].font = Font(bold=True, size=16); ws['A5'].alignment = Alignment(horizontal='center')
            ws['A7'] = f"Kính gửi: {c_name.upper()}"; ws['A7'].font = Font(bold=True); ws['A8'] = f"SĐT: {c_phone}"; ws['A9'] = f"Địa chỉ: {c_addr}"
            
            headers = ["STT", "TÊN SẢN PHẨM / QUY CÁCH", "ĐVT", "SL", "ĐƠN GIÁ", "THUẾ VAT", "THÀNH TIỀN"]
            widths = [5, 35, 8, 7, 14, 9, 15]
            for c, (h, w) in enumerate(zip(headers, widths), 1):
                cell = ws.cell(row=11, column=c, value=h)
                cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.alignment = Alignment(horizontal='center'); cell.border = THIN_BORDER
                ws.column_dimensions[get_column_letter(c)].width = w
            
            for i, r in enumerate(st.session_state.quote, 1):
                row_data = [i, r.get("Sản phẩm"), "Cái", r.get("SL"), r.get("Đơn giá"), r.get("VAT"), r.get("Thành tiền")]
                for col, val in enumerate(row_data, 1):
                    c_cell = ws.cell(row=i+11, column=col, value=val)
                    c_cell.border = THIN_BORDER
            
            curr = len(st.session_state.quote) + 12
            ws.merge_cells(f'A{curr}:F{curr}'); ws.cell(row=curr, column=1, value="TỔNG CỘNG THANH TOÁN:").alignment = Alignment(horizontal='right')
            ws.cell(row=curr, column=7, value=tong).font = Font(bold=True); ws.cell(row=curr, column=7).fill = YELLOW_FILL; ws.cell(row=curr, column=7).border = THIN_BORDER
            
            f = curr + 2
            ws.cell(row=f, column=1, value="* Bảo hành theo tiêu chuẩn hãng.").font = Font(italic=True)
            ws.cell(row=f+1, column=1, value="THÔNG TIN THANH TOÁN:").font = Font(bold=True)
            ws.cell(row=f+2, column=1, value=f"Chủ TK: {BANK_NAME_BENEFICIARY}"); ws.cell(row=f+3, column=1, value=BANK_STK); ws.cell(row=f+4, column=1, value=BANK_BRANCH)
            ws.cell(row=f+6, column=2, value="NGƯỜI LẬP BIỂU").font = Font(bold=True); ws.cell(row=f+6, column=6, value="ĐẠI DIỆN CÔNG TY").font = Font(bold=True)
            
            out = io.BytesIO(); wb.save(out); return out.getvalue()

        st.download_button("📥 Excel A4 Chuẩn", generate_full_excel(), f"BaoGia_{c_name}.xlsx")
        if st.button("🗑️ Reset báo giá"): st.session_state.quote = []; st.rerun()

with tab3:
    if st.button("✨ AI SOẠN HỢP ĐỒNG"):
        st.text_area("Nội dung", value=f"HỢP ĐỒNG KINH TẾ\nBên A: {c_name}", height=300)
