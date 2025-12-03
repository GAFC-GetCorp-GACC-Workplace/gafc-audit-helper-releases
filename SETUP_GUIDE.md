# Hướng Dẫn Setup 2 Repos (Private + Public)

## 📚 Tổng Quan

Hệ thống sử dụng **2 GitHub repositories**:

1. **Private Repo** (`gafc-audit-helper`) - Source code đầy đủ
2. **Public Repo** (`gafc-audit-helper-releases`) - Binary + scripts cho users

---

## 🚀 Setup Bước 1: Tạo Private Repo

### 1.1. Tạo GitHub Repo

1. Vào https://github.com/new
2. Điền:
   ```
   Repository name: gafc-audit-helper
   Visibility: ⦿ Private
   Description: GAFC Audit Helper - Private Source Code
   ☐ Initialize with README (không tick)
   ```
3. Click **Create repository**

### 1.2. Push Code Private Repo

```powershell
# Trong thư mục E:\audit\GAFC_Audit_Helper_Release
cd E:\audit\GAFC_Audit_Helper_Release

# Initialize git (nếu chưa có)
git init

# Add remote (thay YOUR_USERNAME bằng GitHub username của bạn)
git remote add origin https://github.com/YOUR_USERNAME/gafc-audit-helper.git

# Add all files
git add .

# Commit
git commit -m "Initial commit - Private source code"

# Push
git branch -M main
git push -u origin main
```

---

## 🌐 Setup Bước 2: Tạo Public Repo

### 2.1. Tạo GitHub Repo

1. Vào https://github.com/new
2. Điền:
   ```
   Repository name: gafc-audit-helper-releases
   Visibility: ⦿ Public
   Description: GAFC Audit Helper - Excel Add-in for Accounting Automation
   ☐ Initialize with README (không tick)
   ```
3. Click **Create repository**

### 2.2. Sync Files sang Public Repo

```powershell
# Chạy script sync từ private repo
cd E:\audit\GAFC_Audit_Helper_Release
.\scripts\sync_to_public.ps1
```

Script sẽ tự động copy:
- ✅ File XLAM
- ✅ Scripts user (install, update, setup_auto_update, etc.)
- ✅ Manifest JSON
- ✅ Documentation
- ✅ Tạo installer ZIP package

### 2.3. Push Public Repo

```powershell
# Vào thư mục public repo
cd E:\audit\GAFC_Audit_Helper_Release_Public

# Initialize git
git init

# Add remote (thay YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/gafc-audit-helper-releases.git

# Add files
git add .

# Commit
git commit -m "Initial release - v1.0.0"

# Push
git branch -M main
git push -u origin main
```

---

## ⚙️ Bước 3: Cấu Hình URLs

Sửa `YOUR_USERNAME` thành GitHub username thực của bạn trong các file sau:

### File 1: Private Repo - `scripts/update_audit_helper.ps1`

Dòng 4 đã được update sẵn:
```powershell
$ManifestUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/gafc-audit-helper-releases/main/releases/audit_tool.json"
```

### File 2: Private Repo - `releases/audit_tool.json`

Dòng 3 đã được update sẵn:
```json
"download_url": "https://github.com/YOUR_USERNAME/gafc-audit-helper-releases/releases/download/v1.0.0/gafc_audit_helper.xlam"
```

### File 3: Public Repo - `README.md`

Tìm và thay tất cả `YOUR_USERNAME`:
```markdown
https://github.com/YOUR_USERNAME/gafc-audit-helper-releases/releases
```

**Cách nhanh:** Find & Replace trong VS Code
```
Find: YOUR_USERNAME
Replace: your-actual-username
```

---

## 🎯 Bước 4: Tạo Release Đầu Tiên

### Option A: Dùng Script Tự Động (Recommend)

```powershell
# Trong private repo
cd E:\audit\GAFC_Audit_Helper_Release

# Build XLAM trong Excel trước (import code, save)

# Chạy script tạo release
.\scripts\create_release.ps1 -Version "1.0.0" -Message "Initial release"
```

Script sẽ:
1. ✅ Tính SHA256 hash
2. ✅ Update manifest
3. ✅ Commit changes
4. ✅ Create git tag
5. ✅ Hỏi push lên GitHub
6. ✅ Hỏi tạo GitHub Release

### Option B: Manual

#### 4.1. Sync files sang public repo

```powershell
cd E:\audit\GAFC_Audit_Helper_Release
.\scripts\sync_to_public.ps1
```

#### 4.2. Tính SHA256

```powershell
cd E:\audit\GAFC_Audit_Helper_Release_Public
$hash = (Get-FileHash "gafc_audit_helper.xlam" -Algorithm SHA256).Hash.ToLower()
Write-Host "SHA256: $hash"
```

Copy hash này.

#### 4.3. Update manifest

Sửa file `E:\audit\GAFC_Audit_Helper_Release_Public\releases\audit_tool.json`:
```json
{
  "sha256": "paste-hash-here"
}
```

#### 4.4. Commit & Push Public Repo

```powershell
cd E:\audit\GAFC_Audit_Helper_Release_Public
git add .
git commit -m "Release v1.0.0"
git tag v1.0.0
git push origin main
git push origin v1.0.0
```

#### 4.5. Tạo GitHub Release

1. Vào https://github.com/YOUR_USERNAME/gafc-audit-helper-releases/releases/new
2. Điền:
   ```
   Tag: v1.0.0 (chọn tag vừa tạo)
   Title: Release v1.0.0
   Description: Initial release
   ```
3. **Upload files**:
   - `gafc_audit_helper.xlam`
   - `gafc_audit_helper_installer.zip`
4. Click **Publish release**

---

## 🔄 Workflow Release Version Mới

### Khi có code mới:

```powershell
# 1. Edit code trong Excel
# 2. Save XLAM
# 3. Chạy trong private repo:
cd E:\audit\GAFC_Audit_Helper_Release
.\scripts\create_release.ps1 -Version "1.0.1" -Message "Fix bug ABC"

# 4. Sync sang public repo:
.\scripts\sync_to_public.ps1

# 5. Push public repo:
cd E:\audit\GAFC_Audit_Helper_Release_Public
git add .
git commit -m "Release v1.0.1"
git tag v1.0.1
git push origin main --tags

# 6. Tạo GitHub Release trên public repo (upload XLAM + ZIP)
```

---

## 🧪 Testing

### Test Update Script

```powershell
# Trong public repo
cd E:\audit\GAFC_Audit_Helper_Release_Public\scripts
.\update_audit_helper.ps1
```

Kết quả mong đợi:
```
Downloading version 1.0.0 ...
SHA256 verified successfully.
Installing version 1.0.0 ...
✓ Updated successfully
```

---

## 📝 Checklist Hoàn Chỉnh

- [ ] Private repo created và pushed
- [ ] Public repo created và pushed
- [ ] Đã thay `YOUR_USERNAME` trong tất cả files
- [ ] Đã sync files sang public repo
- [ ] Đã tính SHA256 và update manifest
- [ ] Đã tạo GitHub Release v1.0.0
- [ ] Đã test update script
- [ ] Users có thể download từ public repo

---

## 🔐 Security Notes

✅ **Private repo bảo vệ:**
- Source code VBA
- Build scripts
- License server secrets

✅ **Public repo chỉ expose:**
- Binary XLAM (vẫn có thể decompile nhưng khó hơn)
- User scripts (không có logic nhạy cảm)
- Documentation

---

## 📞 Next Steps

Sau khi setup xong:

1. **Chia sẻ link public repo** với users:
   ```
   https://github.com/YOUR_USERNAME/gafc-audit-helper-releases
   ```

2. **Hướng dẫn users cài đặt**:
   - Download installer ZIP từ Releases
   - Chạy install script
   - Setup auto-update

3. **Monitor**:
   - Check GitHub Release downloads
   - Monitor license activation từ server
   - Check update logs từ users

---

**Hoàn thành!** 🎉
