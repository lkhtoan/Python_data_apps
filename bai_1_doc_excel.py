from pathlib import Path
import pandas as pd

# Xác định đường dẫn file nằm cùng thư mục với script đang chạy
thu_muc_hien_tai = Path(__file__).parent
duong_dan_file = thu_muc_hien_tai / "bao_cao_giao_dich.xlsx"

# Kiểm tra file có tồn tại trước khi đọc để tránh văng lỗi đột ngột
if not duong_dan_file.exists():
    print(f"Lỗi: Không tìm thấy file tại {duong_dan_file}")
else:
    # Đọc file Excel đúng các tham số chuẩn
    df = pd.read_excel(
        duong_dan_file,
        sheet_name="Du_lieu_2026",  # Đúng tên sheet đã tạo
        skiprows=2,  # Bỏ qua 2 dòng trống ở trên cùng
        usecols=[
            "Ma_GD",
            "Ngay",
            "Phong_ban",
            "So_tien",
            "Trang_thai",
        ],  # Chọn đúng các cột
        dtype={
            "Ma_GD": str
        },  # Giữ nguyên số 0 ở đầu mã (ví dụ: "00101" không bị biến thành 101)
    )

    print("--- ĐỌC FILE EXCEL THÀNH CÔNG ---")
    print(df)
    print("\nKiểu dữ liệu của từng cột:")
    print(df.dtypes)