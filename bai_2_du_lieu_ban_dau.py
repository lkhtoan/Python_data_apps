import pandas as pd

# Dữ liệu mô phỏng giao dịch thực tế
data = {
    "Ma_GD": ["GD01", "GD02", "GD03", "GD04", "GD05", "GD06", "GD07", "GD08"],
    "Phong_ban": ["Kế toán", "Tín dụng", "Kế toán", "Khách hàng DN", "Tín dụng", "Kế toán", "Tín dụng", "Khách hàng DN"],
    "Trang_thai": ["Hoàn thành", "Chờ duyệt", "Hoàn thành", "Hoàn thành", "Từ chối", "Chờ duyệt", "Hoàn thành", "Hoàn thành"],
    "So_tien": [15_000_000, 120_000_000, 25_000_000, 80_000_000, 45_000_000, 10_000_000, 60_000_000, 150_000_000]
}

df = pd.DataFrame(data)
print("=== DỮ LIỆU BAN ĐẦU ===")
print(df)

# Tính tổng số tiền theo từng phòng ban
tong_theo_phong = df.groupby("Phong_ban")["So_tien"].sum().reset_index()

print("\n--- 1. TỔNG TIỀN THEO TỪNG PHÒNG BAN ---")
print(tong_theo_phong)