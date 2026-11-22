import io
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CẤU HÌNH TRANG WEB
# ==========================================
st.set_page_config(
    page_title="Hệ thống Hợp nhất & Đối soát Dữ liệu Kiểm kê",
    page_icon="🗂️",
    layout="wide"
)

# Bộ cột bắt buộc phải có trong mỗi file gửi lên
REQUIRED_COLS = {"Ma_TS", "Ten_TS", "So_luong", "Gia_tri"}

# Từ điển ánh xạ tên cột gần giống về tên chuẩn
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
# 2. HÀM HỖ TRỢ: ĐỌC & ĐỐI SOÁT TỪNG FILE TRÊN RAM
# ==========================================
def doc_va_kiem_tra_file(uploaded_file):
    """
    Đọc một file Excel hoặc CSV từ bộ nhớ RAM, chuẩn hóa cột và đối soát Schema.
    Trả về: (df_clean, trang_thai, thong_bao)
    """
    file_name = uploaded_file.name
    try:
        if file_name.lower().endswith(".csv"):
            # Thử đọc UTF-8-SIG trước, nếu lỗi encoding thì thử latin1
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

        if df.empty:
            return None, "Cảnh báo", "File không có dòng dữ liệu nào."

        # 1. Làm sạch khoảng trắng tên cột & ánh xạ tên chuẩn
        df.columns = [str(c).strip() for c in df.columns]
        df = df.rename(columns=COLUMN_MAPPING)

        # 2. Đối soát cột bắt buộc bằng phép trừ tập hợp (Set Difference)
        missing_cols = REQUIRED_COLS - set(df.columns)
        if missing_cols:
            return None, "Lỗi cấu trúc", f"Thiếu cột bắt buộc: {', '.join(sorted(missing_cols))}"

        # 3. Nếu file chưa có cột Don_Vi, tự trích xuất từ tên file (bỏ đuôi mở rộng)
        if "Don_Vi" not in df.columns:
            ten_dv_tu_file = file_name.rsplit(".", 1)[0]
            df.insert(0, "Don_Vi", ten_dv_tu_file)

        # 4. Gắn cột truy vết nguồn gốc file
        df["Nguon_File"] = file_name

        # 5. Ép kiểu số an toàn cho cột số lượng và giá trị
        df["So_luong"] = pd.to_numeric(df["So_luong"], errors="coerce").fillna(0).astype(int)
        df["Gia_tri"] = pd.to_numeric(df["Gia_tri"], errors="coerce").fillna(0)

        return df, "Hợp lệ", f"Nạp thành công {len(df):,} dòng dữ liệu."

    except Exception as e:
        return None, "Lỗi đọc file", f"Không thể mở file ({type(e).__name__}: {e})"


def tao_bo_file_mau_zip():
    """Tạo sẵn bộ 3 file mẫu (2 Excel, 1 CSV) đóng gói thành file ZIP trên RAM để người dùng tải về test"""
    import zipfile

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # File 1: Phòng Kế toán (.xlsx)
        df1 = pd.DataFrame([
            {"Ma_TS": "KT-01", "Ten_TS": "Máy tính Dell OptiPlex", "So_luong": 4, "Gia_tri": 64000000},
            {"Ma_TS": "KT-02", "Ten_TS": "Máy in hóa đơn Epson", "So_luong": 2, "Gia_tri": 11000000},
        ])
        b1 = io.BytesIO()
        df1.to_excel(b1, index=False, engine="openpyxl")
        zf.writestr("Phong_Ke_Toan.xlsx", b1.getvalue())

        # File 2: PGD Ninh Kiều (.csv - cố tình dùng tên cột tiếng Việt có dấu để test ánh xạ)
        df2 = pd.DataFrame([
            {"Mã tài sản": "NK-01", "Tên tài sản": "Máy đếm tiền Glory", "Số lượng": 3, "Nguyên giá": 18000000},
            {"Mã tài sản": "NK-02", "Tên tài sản": "Laptop HP ProBook", "Số lượng": 5, "Nguyên giá": 95000000},
        ])
        zf.writestr("PGD_Ninh_Kieu.csv", df2.to_csv(index=False, encoding="utf-8-sig"))

        # File 3: File lỗi thiếu cột (.xlsx - để test cơ chế chống sập & báo cáo lỗi)
        df3 = pd.DataFrame([
            {"Ma_TS": "ERR-01", "Ten_TS": "Tài sản thiếu cột giá trị"},
        ])
        b3 = io.BytesIO()
        df3.to_excel(b3, index=False, engine="openpyxl")
        zf.writestr("File_Loi_Thieu_Cot.xlsx", b3.getvalue())

    zip_buf.seek(0)
    return zip_buf


def xuat_bao_cao_tong_hop_ram(df_master, df_log):
    """Xuất file Excel gồm 2 sheet: Master_Data và Nhat_Ky_Doi_Soat"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_log.to_excel(writer, sheet_name="Nhat_Ky_Doi_Soat", index=False)

        header_fill = PatternFill(start_color="00685E", end_color="00685E", fill_type="solid")
        header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        thin_side = Side(border_style="thin", color="BFBFBF")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for sheet_name, df_curr in [("Master_Data", df_master), ("Nhat_Ky_Doi_Soat", df_log)]:
            ws = writer.sheets[sheet_name]
            ws.row_dimensions[1].height = 25
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
                    cell.border = border
                    if sheet_name == "Master_Data" and df_curr.columns[c - 1] == "Gia_tri":
                        cell.number_format = '#,##0" VNĐ"'

            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    buffer.seek(0)
    return buffer


# ==========================================
# 3. GIAO DIỆN THANH BÊN (SIDEBAR)
# ==========================================
st.sidebar.header("📂 Tải Lên Hàng Loạt File")
uploaded_files = st.sidebar.file_uploader(
    "Chọn nhiều file Excel (.xlsx) hoặc CSV (.csv):",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

st.sidebar.divider()
st.sidebar.subheader("🧪 Bộ Dữ Liệu Thử Nghiệm")
st.sidebar.caption("Tải bộ 3 file mẫu (gồm Excel, CSV và 1 file cố tình làm lỗi thiếu cột) về giải nén rồi kéo thả vào khung trên:")
st.sidebar.download_button(
    label="📦 Tải bộ 3 file mẫu (.zip)",
    data=tao_bo_file_mau_zip(),
    file_name="bo_file_test_batch_upload.zip",
    mime="application/zip"
)

# ==========================================
# 4. XỬ LÝ & HIỂN THỊ CHÍNH
# ==========================================
st.title("🗂️ Trạm Hợp Nhất & Đối Soát Báo Cáo Kiểm Kê Đa Định Dạng")
st.markdown("Hỗ trợ kéo thả **hàng loạt file Excel & CSV** cùng lúc — Tự động chuẩn hóa tên cột, lọc bỏ file lỗi và xuất báo cáo tổng hợp.")

if not uploaded_files:
    st.info("👈 Hãy chọn hoặc kéo thả **nhiều file Excel / CSV** vào khung bên trái để bắt đầu quá trình hợp nhất.")
    with st.expander("📌 Quy chuẩn cấu trúc cột bắt buộc"):
        st.write("Mỗi file tải lên cần có đủ 4 cột cốt lõi (chấp nhận tên tiếng Việt hoặc mã chuẩn):")
        st.code("Ma_TS (Mã tài sản) | Ten_TS (Tên tài sản) | So_luong (Số lượng) | Gia_tri (Nguyên giá)")
else:
    all_dfs = []
    audit_logs = []

    # Duyệt từng file trong danh sách người dùng tải lên
    for f in uploaded_files:
        df_item, status, msg = doc_va_kiem_tra_file(f)
        audit_logs.append({
            "Thời gian": datetime.now().strftime("%H:%M:%S"),
            "Tên file": f.name,
            "Trạng thái": status,
            "Chi tiết đối soát": msg
        })
        if df_item is not None:
            all_dfs.append(df_item)

    df_log = pd.DataFrame(audit_logs)
    so_file_hop_le = (df_log["Trạng thái"] == "Hợp lệ").sum()
    so_file_loi = len(df_log) - so_file_hop_le

    # Hiển thị tổng quan tình trạng nạp file
    c1, c2, c3 = st.columns(3)
    c1.metric("Tổng số file đã tải lên", f"{len(uploaded_files)} file")
    c2.metric("File hợp lệ đã gộp", f"{so_file_hop_le} file")
    c3.metric("File bị từ chối / lỗi", f"{so_file_loi} file", delta_color="inverse")

    if so_file_loi > 0:
        st.warning(f"⚠️ Phát hiện **{so_file_loi} file không đạt chuẩn** đã được tự động loại khỏi bảng tổng hợp. Xem chi tiết ở thẻ **Nhật ký đối soát**.")

    # Chia giao diện thành 3 thẻ (Tabs) chuyên nghiệp
    tab1, tab2, tab3 = st.tabs([
        "📊 Dữ liệu Tổng hợp (Master Data)",
        "📈 Thống kê theo Đơn vị",
        "🧾 Nhật ký Đối soát File"
    ])

    if all_dfs:
        df_master = pd.concat(all_dfs, ignore_index=True)

        with tab1:
            # Bộ lọc theo đơn vị ngay trong tab dữ liệu
            ds_dv = sorted(df_master["Don_Vi"].unique().tolist())
            chon_dv = st.multiselect("Lọc nhanh theo Đơn vị:", options=ds_dv, default=ds_dv)
            df_view = df_master[df_master["Don_Vi"].isin(chon_dv)]

            k1, k2, k3 = st.columns(3)
            k1.metric("Tổng số dòng tài sản", f"{len(df_view):,} mục")
            k2.metric("Tổng số lượng thiết bị", f"{df_view['So_luong'].sum():,} chiếc")
            k3.metric("Tổng nguyên giá", f"{df_view['Gia_tri'].sum():,.0f} VNĐ")

            st.dataframe(df_view, use_container_width=True, hide_index=True)

            st.divider()
            excel_out = xuat_bao_cao_tong_hop_ram(df_view, df_log)
            st.download_button(
                label="⬇️ Tải Báo cáo Tổng hợp + Nhật ký Đối soát (.xlsx)",
                data=excel_out,
                file_name=f"Master_KiemKe_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab2:
            df_summary = df_master.groupby("Don_Vi", as_index=False).agg(
                So_muc_TS=("Ma_TS", "count"),
                Tong_so_luong=("So_luong", "sum"),
                Tong_gia_tri=("Gia_tri", "sum")
            ).sort_values(by="Tong_gia_tri", ascending=False)

            st.subheader("Bảng tổng hợp theo từng Đơn vị")
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            st.subheader("Biểu đồ so sánh Nguyên giá Tài sản giữa các Đơn vị")
            st.bar_chart(df_summary.set_index("Don_Vi")["Tong_gia_tri"])

        with tab3:
            st.subheader("Kết quả kiểm tra cấu trúc từng tệp tải lên")
            st.dataframe(df_log, use_container_width=True, hide_index=True)
    else:
        with tab1:
            st.error("Không có file nào hợp lệ để tổng hợp dữ liệu.")
        with tab2:
            st.info("Chưa có dữ liệu thống kê.")
        with tab3:
            st.dataframe(df_log, use_container_width=True, hide_index=True)
