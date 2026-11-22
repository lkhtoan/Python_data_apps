import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CẤU HÌNH TRANG WEB (Phải đặt ở dòng lệnh st đầu tiên)
# ==========================================
st.set_page_config(
    page_title="Hệ thống Quản lý & Kiểm kê Tài sản",
    page_icon="📊",
    layout="wide"
)

# ==========================================
# 2. HÀM HỖ TRỢ: TẠO DỮ LIỆU MẪU & XUẤT EXCEL CHUẨN
# ==========================================
def tao_du_lieu_mau():
    data = [
        {"Don_Vi": "Phòng Kế toán", "Ma_TS": "KT-01", "Ten_TS": "Máy tính Dell OptiPlex", "So_luong": 4, "Gia_tri": 64000000},
        {"Don_Vi": "Phòng Kế toán", "Ma_TS": "KT-02", "Ten_TS": "Máy in hóa đơn Epson", "So_luong": 2, "Gia_tri": 11000000},
        {"Don_Vi": "Phòng Tín dụng", "Ma_TS": "TD-01", "Ten_TS": "Laptop HP ProBook 450", "So_luong": 6, "Gia_tri": 114000000},
        {"Don_Vi": "Phòng Tín dụng", "Ma_TS": "TD-02", "Ten_TS": "Máy scan tài liệu Canon", "So_luong": 3, "Gia_tri": 25500000},
        {"Don_Vi": "PGD Ninh Kiều", "Ma_TS": "NK-01", "Ten_TS": "Máy đếm tiền Glory", "So_luong": 3, "Gia_tri": 18000000},
        {"Don_Vi": "PGD Cái Răng", "Ma_TS": "CR-01", "Ten_TS": "Bộ lưu điện UPS APC", "So_luong": 2, "Gia_tri": 9600000},
    ]
    return pd.DataFrame(data)

def xuat_excel_tren_ram(df_export):
    """Ghi DataFrame ra bộ nhớ đệm RAM (BytesIO) kèm định dạng OpenPyXL"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name="BaoCaoTaiSan", index=False)
        ws = writer.sheets["BaoCaoTaiSan"]

        # Định dạng Header thương hiệu
        header_fill = PatternFill(start_color="00685E", end_color="00685E", fill_type="solid")
        header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        thin_side = Side(border_style="thin", color="BFBFBF")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        ws.row_dimensions[1].height = 26
        for col_idx in range(1, len(df_export.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        # Định dạng dòng dữ liệu & số tiền
        for r in range(2, len(df_export) + 2):
            ws.row_dimensions[r].height = 20
            for c in range(1, len(df_export.columns) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                if df_export.columns[c - 1] in ["Gia_tri", "Nguyên giá", "Thành tiền"]:
                    cell.number_format = '#,##0" VNĐ"'

        # Auto-fit độ rộng cột
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buffer.seek(0)  # Đưa con trỏ đọc về đầu file ảo trong RAM
    return buffer

# ==========================================
# 3. GIAO DIỆN THANH BÊN (SIDEBAR)
# ==========================================
st.sidebar.header("📂 Nguồn Dữ Liệu")
uploaded_file = st.sidebar.file_uploader(
    "Tải lên file Excel kiểm kê (.xlsx)",
    type=["xlsx", "xls"]
)

use_sample = st.sidebar.checkbox("Dùng dữ liệu mẫu có sẵn để thử nghiệm", value=True)

# ==========================================
# 4. NỘI DUNG CHÍNH (MAIN AREA)
# ==========================================
st.title("📊 Ứng Dụng Tra Cứu & Xuất Báo Cáo Kiểm Kê Tài Sản")
st.markdown("Công cụ tương tác trực tiếp trên trình duyệt — **Không cần cài đặt cơ sở dữ liệu**.")

df = None
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.sidebar.success(f"Đã nạp: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Lỗi đọc file Excel: {e}")
elif use_sample:
    df = tao_du_lieu_mau()
    st.info("Đang hiển thị **dữ liệu mẫu**. Bạn có thể tải file Excel thật của cơ quan ở thanh bên trái.")

if df is not None and not df.empty:
    # --- Bộ lọc động trên Sidebar ---
    st.sidebar.divider()
    st.sidebar.header("🔍 Bộ Lọc Dữ Liệu")
    
    df_filtered = df.copy()
    if "Don_Vi" in df.columns:
        ds_don_vi = sorted(df["Don_Vi"].dropna().unique().tolist())
        chon_don_vi = st.sidebar.multiselect(
            "Chọn Đơn vị / Phòng ban:",
            options=ds_don_vi,
            default=ds_don_vi
        )
        df_filtered = df_filtered[df_filtered["Don_Vi"].isin(chon_don_vi)]

    # --- Khu vực thẻ chỉ số tổng hợp (KPI Metrics) ---
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Tổng số dòng tài sản", f"{len(df_filtered):,} mục")
    with col2:
        tong_sl = df_filtered["So_luong"].sum() if "So_luong" in df_filtered.columns else 0
        st.metric("Tổng số lượng thiết bị", f"{tong_sl:,} chiếc")
    with col3:
        tong_gt = df_filtered["Gia_tri"].sum() if "Gia_tri" in df_filtered.columns else 0
        st.metric("Tổng nguyên giá tài sản", f"{tong_gt:,.0f} VNĐ")

    st.divider()

    # --- Bảng xem trước tương tác ---
    st.subheader("📋 Bảng Dữ Liệu Chi Tiết")
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    # --- Nút tải báo cáo Excel đã trang trí về máy ---
    st.subheader("📥 Xuất Báo Cáo")
    excel_buffer = xuat_excel_tren_ram(df_filtered)
    st.download_button(
        label="⬇️ Tải xuống Báo cáo Excel chuẩn (OpenPyXL)",
        data=excel_buffer,
        file_name="Bao_cao_kiem_ke_web.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Vui lòng tải lên một file Excel hoặc tích chọn dùng dữ liệu mẫu ở thanh bên trái.")
