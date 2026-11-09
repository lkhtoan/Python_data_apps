from pathlib import Path
import pandas as pd

# 1. Đọc file Excel đã tạo ở bài trước
CURRENT_DIR = Path(__file__).parent
FILE_PATH = CURRENT_DIR / "bao_cao_giao_dich.xlsx"

if not FILE_PATH.exists():
    print(f"Chưa thấy file tại: {FILE_PATH}. Vui lòng chạy lại file tao_file_mau.py trước!")
    exit()

df = pd.read_excel(
    FILE_PATH,
    sheet_name="Du_lieu_2026",
    skiprows=2,
    dtype={"Ma_GD": str}
)

print("=" * 60)
print("DỮ LIỆU GỐC TRONG FILE:")
print(df)
print("=" * 60)


# ============================================================
# VÍ DỤ 1: ĐIỀU KIỆN ĐƠN
# ============================================================
# Lọc các giao dịch có số tiền >= 50,000,000 VNĐ
dk_don = df["So_tien"] >= 50_000_000
df_vd1 = df[dk_don]

print("\n--- VÍ DỤ 1: Điều kiện đơn (So_tien >= 50,000,000) ---")
print(df_vd1[["Ma_GD", "Phong_ban", "So_tien"]])


# ============================================================
# VÍ DỤ 2: ĐIỀU KIỆN PHỨC HỢP VỚI 'VÀ' (&)
# Lưu ý: Mỗi biểu thức bắt buộc bọc trong dấu ngoặc đơn ()
# ============================================================
# Giao dịch có So_tien >= 10,000,000 VÀ Trang_thai là "Hoàn thành"
dk_and = (df["So_tien"] >= 10_000_000) & (df["Trang_thai"] == "Hoàn thành")
df_vd2 = df[dk_and]

print("\n--- VÍ DỤ 2: Điều kiện VÀ (&) (Tien >= 10Tr VÀ Hoàn thành) ---")
print(df_vd2[["Ma_GD", "Phong_ban", "So_tien", "Trang_thai"]])


# ============================================================
# VÍ DỤ 3: ĐIỀU KIỆN PHỨC HỢP VỚI 'HOẶC' (|)
# ============================================================
# Giao dịch có Trang_thai là "Chờ duyệt" HOẶC "Từ chối"
dk_or = (df["Trang_thai"] == "Chờ duyệt") | (df["Trang_thai"] == "Từ chối")
df_vd3 = df[dk_or]

print("\n--- VÍ DỤ 3: Điều kiện HOẶC (|) (Chờ duyệt HOẶC Từ chối) ---")
print(df_vd3[["Ma_GD", "Phong_ban", "Trang_thai"]])


# ============================================================
# VÍ DỤ 4: LỌC THEO DANH SÁCH VỚI isin()
# Tương đương toán tử IN trong SQL hoặc chọn nhiều mục trong Filter Excel
# ============================================================
# Lấy các giao dịch thuộc phòng "Kế toán" hoặc "Tín dụng"
danh_sach_phong = ["Kế toán", "Tín dụng"]
dk_isin = df["Phong_ban"].isin(danh_sach_phong)
df_vd4 = df[dk_isin]

print("\n--- VÍ DỤ 4: Lọc theo danh sách với isin() (Kế toán hoặc Tín dụng) ---")
print(df_vd4[["Ma_GD", "Phong_ban", "So_tien"]])


# ============================================================
# VÍ DỤ 5: TÌM KIẾM CHUỖI VĂN BẢN VỚI str.contains()
# Tìm từ khóa không phân biệt hoa thường (case=False)
# ============================================================
# Tìm các dòng mà cột Phong_ban có chứa chữ "toán" (ví dụ: "Kế toán")
# na=False để nếu gặp ô trống sẽ không gây lỗi
dk_contains = df["Phong_ban"].str.contains("toán", case=False, na=False)
df_vd5 = df[dk_contains]

print("\n--- VÍ DỤ 5: Tìm chuỗi với str.contains() (Phong_ban chứa 'toán') ---")
print(df_vd5[["Ma_GD", "Phong_ban", "So_tien"]])


# ============================================================
# VÍ DỤ 6: KẾT HỢP NHIỀU KỸ THUẬT VÀ PHỦ ĐỊNH (~)
# ============================================================
# Lấy giao dịch:
# - Thuộc phòng "Kế toán" hoặc "Tín dụng" (isin)
# - Số tiền > 10 triệu
# - KHÔNG PHẢI trạng thái "Từ chối" (dùng dấu ~ để phủ định)
dk_tong_hop = (
    df["Phong_ban"].isin(["Kế toán", "Tín dụng"]) &
    (df["So_tien"] > 10_000_000) &
    (~(df["Trang_thai"] == "Từ chối"))
)

# Dùng .loc[điều_kiện, danh_sách_cột] để vừa lọc dòng vừa chọn cột
df_ket_qua_cuoi = df.loc[dk_tong_hop, ["Ma_GD", "Phong_ban", "So_tien", "Trang_thai"]]

print("\n--- VÍ DỤ 6: Tổng hợp (isin + Tien > 10Tr + KHÔNG BỊ Từ chối) ---")
print(df_ket_qua_cuoi)