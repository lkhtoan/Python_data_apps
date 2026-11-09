import pandas as pd

# Dữ liệu giả lập đúng các cột của bài học
du_lieu = {
    "Ma_GD": ["00101", "00102", "00103", "00104", "00105"],
    "Ngay": ["2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04", "2026-10-05"],
    "Phong_ban": ["Kế toán", "Tín dụng", "Kế toán", "Khách hàng DN", "Tín dụng"],
    "So_tien": [15000000, 120000000, 5000000, 85000000, 45000000],
    "Trang_thai": ["Hoàn thành", "Chờ duyệt", "Hoàn thành", "Hoàn thành", "Từ chối"],
}
df_mau = pd.DataFrame(du_lieu)

# Ghi ra file Excel với đúng cấu trúc:
# - Tên file: bao_cao_giao_dich.xlsx
# - Tên sheet: Du_lieu_2026
# - Bắt đầu từ dòng số 3 (để mô phỏng skiprows=2)
with pd.ExcelWriter("bao_cao_giao_dich.xlsx", engine="openpyxl") as writer:
    df_mau.to_excel(
        writer,
        sheet_name="Du_lieu_2026",
        startrow=2,  # Dòng 1 và 2 để trống -> code đọc cần skiprows=2
        index=False,
    )

print("Đã tạo thành công file 'bao_cao_giao_dich.xlsx' chuẩn cấu trúc!")