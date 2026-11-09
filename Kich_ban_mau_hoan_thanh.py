import pandas as pd

# Giả lập một bảng dữ liệu thực tế
data = {
    "ma_gd": ["GD001", "GD002", "GD003", "GD004", "GD005", "GD006"],
    "phong_ban": ["Kế toán", "Tín dụng", "Kế toán", "Khách hàng DN", "Tín dụng", "Hành chính"],
    "so_tien": [15_000_000, 120_000_000, 5_000_000, 85_000_000, 45_000_000, 2_000_000],
    "trang_thai": ["Hoàn thành", "Chờ duyệt", "Hoàn thành", "Hoàn thành", "Từ chối", "Hoàn thành"]
}

df = pd.DataFrame(data)

print("--- 1. BẢNG DỮ LIỆU GỐC ---")
print(df)

# BÀI TOÁN: Lọc các giao dịch thuộc phòng "Kế toán" HOẶC "Tín dụng",
# có số tiền >= 10,000,000 và trạng thái là "Hoàn thành"
phong_muc_tieu = ["Kế toán", "Tín dụng"]

dieu_kien = (
    (df["phong_ban"].isin(phong_muc_tieu)) &
    (df["so_tien"] >= 10_000_000) &
    (df["trang_thai"] == "Hoàn thành")
)

# Áp dụng bộ lọc và chỉ trích xuất 2 cột cần báo cáo
df_ket_qua = df.loc[dieu_kien, ["ma_gd", "so_tien"]]

print("\n--- 2. KẾT QUẢ SAU LỌC ---")
print(df_ket_qua)
print(f"Tổng giá trị các món thỏa điều kiện: {df_ket_qua['so_tien'].sum():,} VNĐ")

# Xuất ra file Excel kết quả
# df_ket_qua.to_excel("bao_cao_rut_gon.xlsx", index=False)
df_ket_qua.to_excel("bao_cao_giao_dich.xlsx", index=False)