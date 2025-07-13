# 🔗 Google Sheets Integration Setup

## 📋 Tổng quan

Google Sheets integration cho phép bạn:
- ✅ **Đọc dữ liệu trực tiếp** từ Google Sheets thay vì upload file Excel
- ✅ **Lưu check-in real-time** lên Google Sheets
- ✅ **Không cần download file** - mọi thứ sync tự động
- ✅ **Chia sẻ dễ dàng** với team qua Google Sheets

## 🔧 Cách Setup (2 phương án)

⚠️ **Lưu ý**: Google Sheets API chỉ chạy ở server-side (API routes) để bảo mật credentials.

### Phương án 1: Service Account (Khuyến nghị)

1. **Tạo Google Cloud Project:**
   - Vào [Google Cloud Console](https://console.cloud.google.com/)
   - Tạo project mới hoặc chọn project có sẵn
   - Enable Google Sheets API

2. **Tạo Service Account:**
   - Vào IAM & Admin > Service Accounts
   - Create Service Account
   - Download JSON key file

3. **Cấu hình .env:**
   ```bash
   GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@project.iam.gserviceaccount.com
   GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nYour key\n-----END PRIVATE KEY-----"
   ```

4. **Share Google Sheets:**
   - Mở Google Sheets của bạn
   - Share với email của Service Account
   - Cấp quyền "Editor"

### Phương án 2: OAuth 2.0 (Đơn giản hơn)

1. **Tạo OAuth Credentials:**
   - Vào Google Cloud Console
   - APIs & Services > Credentials
   - Create OAuth 2.0 Client ID

2. **Get Refresh Token:**
   - Sử dụng Google OAuth Playground
   - Authorize với Google Sheets scope
   - Exchange for refresh token

3. **Cấu hình .env:**
   ```bash
   GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
   GOOGLE_CLIENT_SECRET=your-client-secret
   GOOGLE_REFRESH_TOKEN=your-refresh-token
   ```

## 📊 Cách sử dụng

1. **Chuẩn bị Google Sheets:**
   ```
   | Tên        | Email           | Mã QR    | ... |
   |------------|-----------------|----------|-----|
   | Nguyễn A   | a@example.com   | QR123    | ... |
   | Trần B     | b@example.com   | QR456    | ... |
   ```

2. **Kết nối trong app:**
   - Copy URL của Google Sheets
   - Paste vào "Google Sheets URL"
   - Nhập tên sheet (mặc định: Sheet1)
   - Click "Connect to Google Sheets"

3. **Check-in tự động:**
   - Mỗi lần check-in sẽ tự động lưu vào cột "Checked-In At"
   - Không cần export file
   - Real-time sync với team

## 🔐 Bảo mật

- ✅ **Service Account**: An toàn nhất cho production
- ✅ **OAuth**: Tiện lợi cho development
- ⚠️ **Private key**: Không commit vào Git
- ⚠️ **Refresh token**: Bảo mật như password

## 🔧 Troubleshooting

**Lỗi "Authentication failed":**
- Check lại credentials trong .env
- Verify service account có quyền trên Google Sheets

**Lỗi "Sheet not found":**
- Check tên sheet chính xác
- Verify Google Sheets URL đúng format

**Lỗi "Permission denied":**
- Share Google Sheets với service account email
- Cấp quyền "Editor" hoặc "Owner"

## 📋 Ví dụ Google Sheets URL

```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
```

App sẽ tự extract ID: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms`

## 🎯 Lợi ích so với Excel Upload

| Feature | Excel Upload | Google Sheets |
|---------|--------------|---------------|
| **Real-time sync** | ❌ | ✅ |
| **Team collaboration** | ❌ | ✅ |
| **Auto backup** | ❌ | ✅ |
| **Mobile friendly** | ❌ | ✅ |
| **No file management** | ❌ | ✅ |
| **Formatting preserved** | ✅ | ✅ |

## 🚀 Next Steps

1. Setup Google API credentials
2. Configure .env variables  
3. Test connection với sample Google Sheets
4. Training team sử dụng new workflow

Happy check-in! 🎉