# 🚀 Service Account Quick Setup (5 phút)

## Ưu điểm:
- ✅ Không cần OAuth consent screen
- ✅ Không bị giới hạn test users
- ✅ Hoạt động ngay lập tức
- ✅ Phù hợp cho production

## Các bước:

### 1. Tạo Service Account
1. Vào [Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Click **"+ CREATE SERVICE ACCOUNT"**
3. Điền:
   - Service account name: `ticketcheck-sheets`
   - Service account ID: (tự động)
4. Click **"CREATE AND CONTINUE"**
5. Skip phần roles (click "CONTINUE")
6. Click **"DONE"**

### 2. Tạo Key
1. Click vào service account vừa tạo
2. Tab **"KEYS"** > **"ADD KEY"** > **"Create new key"**
3. Chọn **JSON**
4. **Download file JSON**

### 3. Copy thông tin từ JSON
Mở file JSON vừa download, tìm 2 thông tin:
```json
{
  "client_email": "ticketcheck-sheets@your-project.iam.gserviceaccount.com",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBg...\n-----END PRIVATE KEY-----\n"
}
```

### 4. Update .env.local
```bash
# Comment out OAuth credentials
# GOOGLE_CLIENT_ID=...
# GOOGLE_CLIENT_SECRET=...
# GOOGLE_REFRESH_TOKEN=...

# Add Service Account credentials
GOOGLE_SERVICE_ACCOUNT_EMAIL=ticketcheck-sheets@your-project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBg...\n-----END PRIVATE KEY-----\n"
```

### 5. Share Google Sheets với Service Account
1. Mở Google Sheets của bạn
2. Click **"Share"** button
3. Paste email của service account
4. Chọn **"Editor"** permission
5. Click **"Send"**

### 6. Restart Next.js
```bash
# Ctrl+C
npm run dev
```

### 7. Test
- Paste Google Sheets URL vào app
- Click "Connect to Google Sheets"
- ✅ Success!

## Troubleshooting:
- **"Not found"**: Check đã share sheet với service account email chưa
- **"Permission denied"**: Ensure role là "Editor" không phải "Viewer"
- **"Invalid key"**: Check copy đúng format, giữ nguyên \n trong private key