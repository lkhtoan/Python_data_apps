import io
import re
import zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CẤU HÌNH TRANG WEB & HẰNG SỐ
# ==========================================
st.set_page_config(
    page_title="Hệ thống Quản lý & Điều phối Dữ liệu Kiểm kê (Bài 10)",
    page_icon="🏢",
    layout="wide"
)

REQUIRED_COLS = {"Ma_TS", "Ten_TS", "So_luong", "Gia_tri"}

COLUMN_MAPPING = {
    "Mã TS": "Ma_TS",
    "Mã tài sản": "Ma_TS",
    "Tên TS": "Ten_TS",
    "Tên tài sản": "Ten_TS",
    "Số lượng": "So_luong",
    "Nguyên giá": "Gia_tri",
    "Giá trị": "Gia_tri",
    "Đơn vị": "Don_Vi",
}

# ==========================================
# 2. CÁC HÀM XỬ LÝ DỮ LIỆU & OPENPYXL TRÊN RAM
# ==========================================
@st.cache_data(show_spinner=False)
def tao_du_lieu_mau():
    """Tạo bộ dữ liệu mẫu chuẩn để người dùng trải nghiệm ngay cả khi chưa có file"""
    data = [
        {"Don_Vi": "Phòng Kế toán", "Ma_TS": "TS-001", "Ten_TS": "Máy tính để bàn Dell OptiPlex", "So_luong": 5, "Gia_tri": 85000000},
        {"Don_Vi": "Phòng Kế toán", "Ma_TS": "TS-002", "Ten_TS": "Máy in laser HP Pro", "So_luong": 2, "Gia_tri": 12000000},
        {"Don_Vi": "Phòng Tín dụng", "Ma_TS": "TS-003", "Ten_TS": "Laptop Lenovo ThinkPad", "So_luong": 8, "Gia_tri": 176000000},
        {"Don_Vi": "Phòng Tín dụng", "Ma_TS": "TS-004", "Ten_TS": "Máy scan tốc độ cao Fujitsu", "So_luong": 2, "Gia_tri": 28000000},
        {"Don_Vi": "PGD Ninh Kiều", "Ma_TS": "TS-005", "Ten_TS": "Máy đếm tiền Glory", "So_luong": 4, "Gia_tri": 24000000},
        {"Don_Vi": "PGD Ninh Kiều", "Ma_TS": "TS-006", "Ten_TS": "Hệ thống xếp hàng tự động", "So_luong": 1, "Gia_tri": 45000000},
        {"Don_Vi": "PGD Cái Răng", "Ma_TS": "TS-007", "Ten_TS": "Máy tính All-in-One HP", "So_luong": 4, "Gia_tri": 68000000},
        {"Don_Vi": "PGD Cái Răng", "Ma_TS": "TS-008", "Ten_TS": "Máy in sổ tiết kiệm Epson", "So_luong": 2, "Gia_tri": 19000000},
    ]
    df = pd.DataFrame(data)
    df["Nguon_File"] = "Du_lieu_mau_he_thong.xlsx"
    return df


def doc_va_chuan_hoa_file(uploaded_file):
    """Đọc 1 file Excel/CSV từ RAM, ánh xạ tên cột và đối soát Schema"""
    file_name = uploaded_file.name
    try:
        if file_name.lower().endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

        if df.empty:
            return None, "Cảnh báo", "File rỗng, không có dữ liệu."

        df.columns = [str(c).strip() for c in df.columns]
        df = df.rename(columns=COLUMN_MAPPING)

        missing_cols = REQUIRED_COLS - set(df.columns)
        if missing_cols:
            return None, "Lỗi cấu trúc", f"Thiếu cột bắt buộc: {', '.join(sorted(missing_cols))}"

        if "Don_Vi" not in df.columns:
            df.insert(0, "Don_Vi", file_name.rsplit(".", 1)[0])

        df["Don_Vi"] = df["Don_Vi"].fillna("Chưa phân loại").astype(str).str.strip()
        df["Nguon_File"] = file_name
        df["So_luong"] = pd.to_numeric(df["So_luong"], errors="coerce").fillna(0).astype(int)
        df["Gia_tri"] = pd.to_numeric(df["Gia_tri"], errors="coerce").fillna(0)

        return df, "Hợp lệ", f"Nạp thành công {len(df):,} dòng."
    except Exception as e:
        return None, "Lỗi đọc file", f"{type(e).__name__}: {e}"


def lam_sach_ten_file(ten_goc):
    """Loại bỏ ký tự cấm đặt tên file trên Windows và Linux"""
    ten_sach = re.sub(r'[\\/*?:"<>|]', "_", str(ten_goc).strip())
    return ten_sach if ten_sach else "Khong_xac_dinh"


def dinh_dang_sheet_openpyxl(ws, df_curr, is_data_sheet=True):
    """Áp dụng bộ định dạng chuẩn doanh nghiệp cho 1 worksheet OpenPyXL"""
    header_fill = PatternFill(start_color="00685E", end_color="00685E", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10)
    thin_side = Side(border_style="thin", color="BFBFBF")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.row_dimensions[1].height = 26
    for col_idx in range(1, len(df_curr.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for r in range(2, len(df_curr) + 2):
        ws.row_dimensions[r].height = 20
        for c in range(1, len(df_curr.columns) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            col_name = df_curr.columns[c - 1]
            if is_data_sheet and col_name == "Gia_tri":
                cell.number_format = '#,##0" VNĐ"'
            elif is_data_sheet and col_name == "So_luong":
                cell.number_format = "#,##0"

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)


def tao_excel_don_le_ram(df_data, sheet_name="Du_Lieu", df_log=None):
    """Tạo 1 file Excel hoàn chỉnh trong bộ nhớ RAM"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_data.to_excel(writer, sheet_name=sheet_name, index=False)
        dinh_dang_sheet_openpyxl(writer.sheets[sheet_name], df_data, is_data_sheet=True)

        if df_log is not None and not df_log.empty:
            df_log.to_excel(writer, sheet_name="Nhat_Ky_Doi_Soat", index=False)
            dinh_dang_sheet_openpyxl(writer.sheets["Nhat_Ky_Doi_Soat"], df_log, is_data_sheet=False)

    buffer.seek(0)
    return buffer


def tao_goi_zip_tach_don_vi(df_master, cot_tach="Don_Vi"):
    """
    Tách bảng tổng hợp theo từng đơn vị, định dạng OpenPyXL từng file
    và nén toàn bộ vào 1 file ZIP ngay trên RAM (không tạo file rác trên ổ cứng).
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for ten_dv, df_con in df_master.groupby(cot_tach):
            excel_buf = tao_excel_don_le_ram(df_con.reset_index(drop=True), sheet_name="Bao_Cao_Don_Vi")
            ten_file_an_toan = f"KiemKe_{lam_sach_ten_file(ten_dv)}.xlsx"
            zf.writestr(ten_file_an_toan, excel_buf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


# ==========================================
# 3. THANH ĐIỀU KHIỂN BÊN TRÁI (SIDEBAR)
# ==========================================
st.sidebar.title("⚙️ Điều Khiển Hệ Thống")
st.sidebar.header("1. Nguồn Dữ Liệu Đầu Vào")

uploaded_files = st.sidebar.file_uploader(
    "Kéo thả 1 hoặc nhiều file Excel / CSV:",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

use_sample = st.sidebar.checkbox(
    "Dùng dữ liệu mẫu khi chưa tải file",
    value=True
)

# ==========================================
# 4. NẠP & XỬ LÝ DỮ LIỆU TRUNG TÂM
# ==========================================
st.title("🏢 Trạm Điều Phối & Xử Lý Báo Cáo Kiểm Kê Toàn Diện")
st.markdown("Tích hợp **3 nghiệp vụ cốt lõi** trên một giao diện: **Hợp nhất đa file** $\rightarrow$ **Tra cứu & Lọc động** $\rightarrow$ **Tách báo cáo từng đơn vị (.ZIP)**.")

df_master = None
df_log = pd.DataFrame()

if uploaded_files:
    all_dfs = []
    logs = []
    for f in uploaded_files:
        df_item, status, msg = doc_va_chuan_hoa_file(f)
        logs.append({
            "Thời gian": datetime.now().strftime("%H:%M:%S"),
            "Tên file": f.name,
            "Trạng thái": status,
            "Chi tiết": msg
        })
        if df_item is not None:
            all_dfs.append(df_item)

    df_log = pd.DataFrame(logs)
    if all_dfs:
        df_master = pd.concat(all_dfs, ignore_index=True)
elif use_sample:
    df_master = tao_du_lieu_mau()
    st.info("💡 Đang chạy với **dữ liệu mẫu 4 đơn vị**. Bạn có thể tải file Excel/CSV thực tế lên ở thanh bên trái.")

# ==========================================
# 5. GIAO DIỆN 4 THẺ NGHIỆP VỤ (TABS)
# ==========================================
if df_master is not None and not df_master.empty:
    # Bộ lọc chung trên Sidebar
    st.sidebar.divider()
    st.sidebar.header("2. Bộ Lọc Toàn Cục")
    ds_don_vi = sorted(df_master["Don_Vi"].unique().tolist())
    chon_don_vi = st.sidebar.multiselect(
        "Chọn Đơn vị hiển thị:",
        options=ds_don_vi,
        default=ds_don_vi
    )
    tu_khoa = st.sidebar.text_input("Tìm kiếm nhanh (Mã hoặc Tên TS):", placeholder="VD: Dell, Máy in, TS-003...")

    # Áp dụng lọc
    df_filtered = df_master[df_master["Don_Vi"].isin(chon_don_vi)].copy()
    if tu_khoa.strip():
        kw = tu_khoa.strip()
        mask = (
            df_filtered["Ma_TS"].astype(str).str.contains(kw, case=False, na=False) |
            df_filtered["Ten_TS"].astype(str).str.contains(kw, case=False, na=False)
        )
        df_filtered = df_filtered[mask]

    # Hàng KPI tổng hợp
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Số đơn vị đang chọn", f"{df_filtered['Don_Vi'].nunique()} đơn vị")
    k2.metric("Tổng dòng tài sản", f"{len(df_filtered):,} mục")
    k3.metric("Tổng số lượng", f"{df_filtered['So_luong'].sum():,} chiếc")
    k4.metric("Tổng nguyên giá", f"{df_filtered['Gia_tri'].sum():,.0f} VNĐ")

    st.divider()

    tab_tong_hop, tab_tach_file, tab_thong_ke, tab_log = st.tabs([
        "📋 1. Tra cứu & Xuất Tổng hợp",
        "✂️ 2. Tách Báo cáo từng Đơn vị (.ZIP)",
        "📊 3. Phân tích & Biểu đồ",
        "🧾 4. Nhật ký Đối soát File"
    ])

    # --- TAB 1: TRA CỨU & XUẤT FILE TỔNG HỢP ---
    with tab_tong_hop:
        st.subheader("Bảng Dữ Liệu Chi Tiết Đã Lọc")
        st.dataframe(df_filtered, use_container_width=True, hide_index=True)

        if not df_filtered.empty:
            excel_master = tao_excel_don_le_ram(df_filtered, sheet_name="Tong_Hop_Kiem_Ke", df_log=df_log)
            st.download_button(
                label="⬇️ Tải Báo cáo Tổng hợp chuẩn OpenPyXL (.xlsx)",
                data=excel_master,
                file_name=f"BaoCao_TongHop_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # --- TAB 2: TÁCH FILE TỰ ĐỘNG THEO ĐƠN VỊ & ĐÓNG GÓI ZIP ---
    with tab_tach_file:
        st.subheader("Tự Động Tách Báo Cáo Riêng Cho Từng Đơn Vị")
        st.write("Hệ thống sẽ tự động chia nhỏ bảng dữ liệu hiện tại theo từng **Đơn vị**, định dạng chuẩn OpenPyXL cho từng file và nén thành **1 gói `.zip` duy nhất** để bạn gửi cho các đơn vị đối chiếu.")

        if not df_filtered.empty:
            ds_dv_se_tach = sorted(df_filtered["Don_Vi"].unique().tolist())
            st.write(f"**Danh sách {len(ds_dv_se_tach)} tệp Excel sẽ được tạo trong gói ZIP:**")
            preview_files = pd.DataFrame({
                "Đơn vị": ds_dv_se_tach,
                "Tên file xuất ra": [f"KiemKe_{lam_sach_ten_file(dv)}.xlsx" for dv in ds_dv_se_tach],
                "Số dòng dữ liệu": [len(df_filtered[df_filtered["Don_Vi"] == dv]) for dv in ds_dv_se_tach]
            })
            st.dataframe(preview_files, use_container_width=True, hide_index=True)

            zip_data = tao_goi_zip_tach_don_vi(df_filtered, cot_tach="Don_Vi")
            st.download_button(
                label=f"📦 Tải trọn bộ {len(ds_dv_se_tach)} file báo cáo đơn vị (.zip)",
                data=zip_data,
                file_name=f"BaoCao_Tach_DonVi_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip"
            )
        else:
            st.warning("Không có dữ liệu để tách file. Hãy kiểm tra lại bộ lọc ở thanh bên trái.")

    # --- TAB 3: THỐNG KÊ & BIỂU ĐỒ ---
    with tab_thong_ke:
        if not df_filtered.empty:
            df_group = df_filtered.groupby("Don_Vi", as_index=False).agg(
                So_danh_muc=("Ma_TS", "count"),
                Tong_so_luong=("So_luong", "sum"),
                Tong_nguyen_gia=("Gia_tri", "sum")
            ).sort_values(by="Tong_nguyen_gia", ascending=False)

            col_left, col_right = st.columns([1, 1])
            with col_left:
                st.subheader("Bảng tổng hợp theo Đơn vị")
                st.dataframe(df_group, use_container_width=True, hide_index=True)
            with col_right:
                st.subheader("Biểu đồ Nguyên giá theo Đơn vị")
                st.bar_chart(df_group.set_index("Don_Vi")["Tong_nguyen_gia"])
        else:
            st.info("Không có dữ liệu để vẽ biểu đồ.")

    # --- TAB 4: NHẬT KÝ ĐỐI SOÁT FILE ---
    with tab_log:
        if not df_log.empty:
            so_ok = (df_log["Trạng thái"] == "Hợp lệ").sum()
            so_err = len(df_log) - so_ok
            st.write(f"Kết quả nạp file: **{so_ok} hợp lệ** | **{so_err} lỗi/cảnh báo**")
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else:
            st.info("Bạn đang dùng dữ liệu mẫu hệ thống nên chưa có nhật ký tải file.")
else:
    st.warning("⚠️ Chưa có dữ liệu hợp lệ. Hãy tải file Excel/CSV lên hoặc bật chế độ **Dùng dữ liệu mẫu** ở thanh bên trái.")
