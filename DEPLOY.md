# Hướng Dẫn Deploy Ứng Dụng Streamlit

## ⚠️ Lưu ý về Vercel

**Vercel không phù hợp để deploy Streamlit** vì:
- Vercel chủ yếu dùng cho static sites và serverless functions
- Streamlit cần chạy như một Python server liên tục
- Vercel có timeout giới hạn cho serverless functions

## ✅ Khuyến nghị: Streamlit Cloud (Miễn phí, Dễ nhất)

### Bước 1: Chuẩn bị code
1. Đảm bảo có file `requirements.txt`
2. Đẩy code lên GitHub repository

### Bước 2: Deploy lên Streamlit Cloud
1. Truy cập: https://share.streamlit.io
2. Đăng nhập bằng tài khoản GitHub
3. Click "New app"
4. Chọn:
   - **Repository**: Repository chứa code
   - **Branch**: main (hoặc branch bạn muốn)
   - **Main file path**: `app.py`
5. Click "Deploy"
6. Đợi vài phút để build và deploy

### Bước 3: Sử dụng
- Ứng dụng sẽ có URL dạng: `https://your-app-name.streamlit.app`
- Bạn có thể share URL này cho người khác

---

## 🚀 Các nền tảng khác (Nếu cần)

### Option 2: Railway.app
1. Đăng ký tại https://railway.app
2. Kết nối GitHub repository
3. Tạo service mới từ GitHub
4. Thêm biến môi trường (nếu cần)
5. Deploy tự động

### Option 3: Render.com
1. Đăng ký tại https://render.com
2. Tạo "Web Service" mới
3. Kết nối GitHub repository
4. Cấu hình:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`
5. Deploy

### Option 4: Heroku (Có phí sau free tier)
1. Cài Heroku CLI
2. Tạo file `Procfile` với nội dung:
   ```
   web: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
   ```
3. Deploy bằng Heroku CLI

---

## 📝 File cần thiết để deploy

- ✅ `app.py` - File chính
- ✅ `requirements.txt` - Dependencies
- ✅ `.gitignore` - Bỏ qua file không cần thiết
- ✅ `README.md` - Tài liệu (tùy chọn)

---

## 🔧 Troubleshooting

### Lỗi import module
- Kiểm tra `requirements.txt` có đầy đủ packages
- Đảm bảo tên package đúng

### Lỗi khi chạy
- Kiểm tra log trong dashboard của platform
- Đảm bảo Python version phù hợp (thường 3.8+)

### File upload không hoạt động
- Streamlit Cloud hỗ trợ upload file, nhưng có giới hạn kích thước
- Kiểm tra giới hạn của platform bạn dùng

