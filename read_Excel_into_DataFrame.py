import pandas as pd

# 1. Đọc sheet mặc định (sheet đầu tiên)
df = pd.read_excel("bao_cao_giao_dich.xlsx")

# 2. Các tham số thực tế quan trọng thường dùng:
df = pd.read_excel(
    "bao_cao_giao_dich.xlsx",
    sheet_name="Du_lieu_2026",   # Tên sheet hoặc chỉ số (0, 1, 2)
    skiprows=2,                  # Bỏ qua 2 dòng tiêu đề phụ/trống ở đầu file
    usecols=["Ma_GD", "Ngay", "Phong_ban", "So_tien", "Trang_thai"], # Chỉ đọc cột cần thiết (tiết kiệm RAM)
    dtype={"Ma_GD": str}         # Ép mã số thành chuỗi để không bị mất số 0 ở đầu (00123 -> "00123")
)