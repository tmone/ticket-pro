# 🚀 Google Sheets API - Quick Setup Guide

## 🎯 Cách nhanh nhất để test (OAuth 2.0)

### Bước 1: Tạo Google Cloud Project
1. Mở [Google Cloud Console](https://console.cloud.google.com/)
2. Click **"Create Project"** hoặc chọn project có sẵn
3. Đặt tên project (VD: "TicketCheck Pro")

### Bước 2: Enable Google Sheets API
1. Trong project, vào **"APIs & Services" > "Library"**
2. Tìm **"Google Sheets API"**
3. Click **"Enable"**

### Bước 3: Tạo OAuth 2.0 Credentials
1. Vào **"APIs & Services" > "Credentials"**
2. Click **"+ CREATE CREDENTIALS"** > **"OAuth client ID"**
3. Nếu chưa config consent screen:
   - Choose **"External"**
   - Điền app name và email
   - Add test users (email của bạn)
   - Save
4. Application type: **"Web application"**
5. Add redirect URI: `https://developers.google.com/oauthplayground`
6. **Copy Client ID và Client Secret**

### Bước 4: Get Refresh Token
1. Mở [OAuth 2.0 Playground](https://developers.google.com/oauthplayground/)
2. Click ⚙️ (Settings) > Check **"Use your own OAuth credentials"**
3. Paste **Client ID** và **Client Secret**
4. Ở Step 1: 
   - Tìm **"Google Sheets API v4"**
   - Check: `https://www.googleapis.com/auth/spreadsheets`
   - Click **"Authorize APIs"**
5. Login với Google account
6. Ở Step 2: Click **"Exchange authorization code for tokens"**
7. **Copy Refresh Token**

### Bước 5: Update .env.local
```bash
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
GOOGLE_REFRESH_TOKEN=your-refresh-token
```

### Bước 6: Restart Next.js
```bash
# Ctrl+C để stop server
npm run dev
```

## ✅ Test Connection

1. **Tạo Google Sheets mới**:
   - Headers ở row 1: `Tên | Email | Mã QR | Phone`
   - Add vài rows data test

2. **Copy Google Sheets URL**:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit
   ```

3. **Trong app**:
   - Paste URL vào "Google Sheets URL"
   - Click "Connect to Google Sheets"
   - ✅ Nếu thành công: Sẽ load data từ sheet

## 🔧 Troubleshooting

### ❌ "Google API credentials not configured"
→ Check `.env.local` có đúng format không
→ Restart Next.js server

### ❌ "Invalid credentials"  
→ Check Client ID, Secret, Refresh Token
→ Regenerate refresh token nếu expired

### ❌ "Insufficient permissions"
→ Check đã enable Google Sheets API chưa
→ Check scope có `spreadsheets` không

## 📋 Service Account (Production)

Cho production app, nên dùng Service Account:

1. **Create Service Account**:
   - APIs & Services > Credentials
   - Create credentials > Service account
   - Download JSON key

2. **Extract từ JSON**:
   ```json
   {
     "client_email": "xxx@xxx.iam.gserviceaccount.com",
     "private_key": "-----BEGIN PRIVATE KEY-----\n..."
   }
   ```

3. **Update .env.local**:
   ```bash
   GOOGLE_SERVICE_ACCOUNT_EMAIL=xxx@xxx.iam.gserviceaccount.com
   GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n..."
   ```

4. **Share Google Sheets** với service account email

## 🎉 Done!

Sau khi setup xong, app sẽ:
- ✅ Load data từ Google Sheets
- ✅ Auto-save check-ins real-time
- ✅ No file download needed!

Need help? Check console logs for detailed errors.