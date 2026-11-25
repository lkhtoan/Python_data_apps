import io
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. ĐỌC FILE CẤU HÌNH ĐỘNG (CONFIG.JSON)
# ==========================================
def get_app_dir():
    """Xác định thư mục gốc chứa ứng dụng (tương thích cả khi chạy .py lẫn đóng gói .exe)"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

APP_DIR = get_app_dir()
CONFIG_PATH = APP_DIR / "config.json"

DEFAULT_CONFIG = {
    "app_settings": {
        "page_title": "Hệ thống Kiểm kê Tài sản",
        "page_icon": "🏢",
        "banner_title": "🏢 Trạm Điều Phối & Xử Lý Báo Cáo Kiểm Kê",
        "default_sample": True
    },
    "schema": {
        "required_cols": ["Ma_TS", "Ten_TS", "So_luong", "Gia_tri"],
        "column_mapping": {}
    },
    "excel_style": {
        "header_bg_color": "00685E",
        "header_text_color": "FFFFFF",
        "font_name": "Arial",
        "currency_format": "#,##0\" VNĐ\""
    }
}

def load_config():
    """Đọc tham số từ config.json; tự tạo file mặc định nếu chưa tồn tại"""
    if not CONFIG_PATH.exists():
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Lỗi đọc file cấu hình config.json: {e}. Đang dùng cấu hình mặc định.")
        return DEFAULT_CONFIG

CONFIG = load_config()
APP_CFG = CONFIG["app_settings"]
SCHEMA_CFG = CONFIG["schema"]
STYLE_CFG = CONFIG["excel_style"]

REQUIRED_COLS = set(SCHEMA_CFG["required_cols"])
COLUMN_MAPPING = SCHEMA_CFG["column_mapping"]

# ==========================================
# 2. CẤU HÌNH TRANG WEB TỪ CONFIG
# ==========================================
st.set_page_config(
    page_title=APP_CFG["page_title"],
    page_icon=APP_CFG["page_icon"],
    layout="wide"
)

# ==========================================
# 3. CÁC HÀM XỬ LÝ DỮ LIỆU & OPENPYXL
# ==========================================
@st.cache_data(show_spinner=False)
def tao_du_lieu_mau():
    data = [
        {"Don_Vi": "Phòng Kế toán", "Ma_TS": "TS-001", "Ten_TS": "Máy tính để bàn Dell OptiPlex", "So_luong": 5, "Gia_tri": 85000000},
        {"Don_Vi": "Phòng Tín dụng", "Ma_TS": "TS-003", "Ten_TS": "Laptop Lenovo ThinkPad", "So_luong": 8, "Gia_tri": 176000000},
        {"Don_Vi": "PGD Ninh Kiều", "Ma_TS": "TS-005", "Ten_TS": "Máy đếm tiền Glory", "So_luong": 4, "Gia_tri": 24000000},
        {"Don_Vi": "PGD Cái Răng", "Ma_TS": "TS-007", "Ten_TS": "Máy tính All-in-One HP", "So_luong": 4, "Gia_tri": 68000000},
    ]
    df = pd.DataFrame(data)
    df["Nguon_File"] = "Du_lieu_mau_he_thong.xlsx"
    return df

def doc_va_chuan_hoa_file(uploaded_file):
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
    ten_sach = re.sub(r'[\\/*?:"<>|]', "_", str(ten_goc).strip())
    return ten_sach if ten_sach else "Khong_xac_dinh"

def dinh_dang_sheet_openpyxl(ws, df_curr, is_data_sheet=True):
    header_fill = PatternFill(
        start_color=STYLE_CFG["header_bg_color"],
        end_color=STYLE_CFG["header_bg_color"],
        fill_type="solid"
    )
    header_font = Font(
        name=STYLE_CFG["font_name"],
        size=10,
        bold=True,
        color=STYLE_CFG["header_text_color"]
    )
    body_font = Font(name=STYLE_CFG["font_name"], size=10)
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
                cell.number_format = STYLE_CFG["currency_format"]
            elif is_data_sheet and col_name == "So_luong":
                cell.number_format = "#,##0"

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

def tao_excel_don_le_ram(df_data, sheet_name="Du_Lieu", df_log=None):
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
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for ten_dv, df_con in df_master.groupby(cot_tach):
            excel_buf = tao_excel_don_le_ram(df_con.reset_index(drop=True), sheet_name="Bao_Cao_Don_Vi")
            zf.writestr(f"KiemKe_{lam_sach_ten_file(ten_dv)}.xlsx", excel_buf.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# ==========================================
# 4. GIAO DIỆN NGƯỜI DÙNG
# ==========================================
st.sidebar.title("⚙️ Điều Khiển Hệ Thống")
uploaded_files = st.sidebar.file_uploader(
    "Kéo thả file Excel / CSV:",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)
use_sample = st.sidebar.checkbox("Dùng dữ liệu mẫu", value=APP_CFG["default_sample"])

st.title(APP_CFG["banner_title"])

df_master = None
df_log = pd.DataFrame()

if uploaded_files:
    all_dfs, logs = [], []
    for f in uploaded_files:
        df_item, status, msg = doc_va_chuan_hoa_file(f)
        logs.append({"Thời gian": datetime.now().strftime("%H:%M:%S"), "Tên file": f.name, "Trạng thái": status, "Chi tiết": msg})
        if df_item is not None:
            all_dfs.append(df_item)
    df_log = pd.DataFrame(logs)
    if all_dfs:
        df_master = pd.concat(all_dfs, ignore_index=True)
elif use_sample:
    df_master = tao_du_lieu_mau()

if df_master is not None and not df_master.empty:
    ds_don_vi = sorted(df_master["Don_Vi"].unique().tolist())
    chon_don_vi = st.sidebar.multiselect("Chọn Đơn vị:", options=ds_don_vi, default=ds_don_vi)
    df_filtered = df_master[df_master["Don_Vi"].isin(chon_don_vi)].copy()

    k1, k2, k3 = st.columns(3)
    k1.metric("Tổng dòng tài sản", f"{len(df_filtered):,} mục")
    k2.metric("Tổng số lượng", f"{df_filtered['So_luong'].sum():,} chiếc")
    k3.metric("Tổng nguyên giá", f"{df_filtered['Gia_tri'].sum():,.0f} VNĐ")

    tab1, tab2, tab3 = st.tabs(["📋 Dữ liệu Tổng hợp", "✂️ Tách file ZIP", "⚙️ Xem Cấu hình hiện hành"])
    with tab1:
        st.dataframe(df_filtered, use_container_width=True, hide_index=True)
        if not df_filtered.empty:
            st.download_button(
                "⬇️ Tải Báo cáo Tổng hợp (.xlsx)",
                data=tao_excel_don_le_ram(df_filtered, "Tong_Hop", df_log),
                file_name=f"BaoCao_TongHop_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with tab2:
        if not df_filtered.empty:
            zip_data = tao_goi_zip_tach_don_vi(df_filtered)
            st.download_button(
                "📦 Tải gói báo cáo từng đơn vị (.zip)",
                data=zip_data,
                file_name=f"BaoCao_DonVi_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip"
            )
    with tab3:
        st.caption(f"Đang đọc cấu hình từ: `{CONFIG_PATH}`")
        st.json(CONFIG)