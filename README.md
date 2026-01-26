# Hệ Thống Soạn Hàng Gimme

Ứng dụng web để tổng hợp và lọc đơn hàng từ file Excel/CSV.

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy local

```bash
streamlit run app.py
```

## Deploy lên Streamlit Cloud (Khuyến nghị)

1. Đẩy code lên GitHub
2. Truy cập [share.streamlit.io](https://share.streamlit.io)
3. Đăng nhập bằng GitHub
4. Chọn repository và file `app.py`
5. Click "Deploy"

## Tính năng

- Đọc file Excel/CSV
- Tách SKU, màu sắc, size từ dữ liệu
- Tổng hợp đơn hàng
- Xuất file Word với định dạng đẹp
- Hiển thị số trang tự động

