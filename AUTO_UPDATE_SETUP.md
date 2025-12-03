# Hướng Dẫn Cấu Hình Auto-Update cho GAFC Audit Helper

## 📋 Tổng Quan

Hệ thống auto-update tự động kiểm tra và cài đặt phiên bản mới mỗi 12 giờ khi Excel không chạy.

## 🔧 Các Bước Cấu Hình

### Bước 1: Setup GitHub Repository

1. **Tạo GitHub Repository** (public hoặc private)

2. **Upload file XLAM lên GitHub Releases**:
   ```
   - Vào repository → Releases → Create new release
   - Tag: v1.0.0
   - Upload file: chuyen_dau_ki_v9.xlam
   - Publish release
   ```

3. **Lấy download URL**:
   ```
   https://github.com/YOUR_ORG/YOUR_REPO/releases/download/v1.0.0/chuyen_dau_ki_v9.xlam
   ```

### Bước 2: Tính SHA256 Hash

Chạy PowerShell để tính hash của file XLAM:

```powershell
Get-FileHash "chuyen_dau_ki_v9.xlam" -Algorithm SHA256 | Select-Object -ExpandProperty Hash
```

Copy giá trị hash (ví dụ: `abc123def456...`)

### Bước 3: Cập Nhật Manifest File

Sửa file `releases/audit_tool.json`:

```json
{
  "latest": "1.0.0",
  "download_url": "https://github.com/YOUR_ORG/YOUR_REPO/releases/download/v1.0.0/chuyen_dau_ki_v9.xlam",
  "sha256": "abc123def456...",  // ← Paste hash vào đây
  "release_date": "2025-12-03",
  "release_notes": "Initial release with license validation"
}
```

### Bước 4: Upload Manifest lên GitHub

**Option A: Commit trực tiếp vào main branch**
```bash
git add releases/audit_tool.json
git commit -m "Update manifest for v1.0.0"
git push origin main
```

**Option B: Sử dụng GitHub Raw URL**
- Upload file `audit_tool.json` vào repository
- URL sẽ là: `https://raw.githubusercontent.com/YOUR_ORG/YOUR_REPO/main/releases/audit_tool.json`

### Bước 5: Cấu Hình Script Update

Sửa file `scripts/update_audit_helper.ps1`, dòng 4:

```powershell
$ManifestUrl = "https://raw.githubusercontent.com/YOUR_ORG/YOUR_REPO/main/releases/audit_tool.json"
```

Thay `YOUR_ORG` và `YOUR_REPO` bằng tên thực tế.

### Bước 6: Cài Đặt Auto-Update trên Client

**Chạy script setup** (với quyền admin nếu cần):

```powershell
cd E:\audit\v9\GAFC_Audit_Helper_Release\scripts
.\setup_auto_update.ps1
```

Script sẽ:
- Tạo Windows Scheduled Task
- Chạy mỗi 12 giờ
- Chỉ update khi Excel đóng
- Log vào `%TEMP%\gafc_update.log`

## 📝 Cấu Hình Nâng Cao

### Thay Đổi Tần Suất Check Update

Sửa file `setup_auto_update.ps1`, dòng 8:

```powershell
$UpdateInterval = 12  # ← Đổi thành 6, 24, 48, etc.
```

### Silent Mode (Không hiện output)

Sửa file `update_audit_helper.ps1`, dòng 6:

```powershell
$SilentMode = $true  # ← Đổi từ $false sang $true
```

## 🔄 Workflow Phát Hành Version Mới

### Khi có version mới (ví dụ v1.1.0):

1. **Build file XLAM mới** với code mới
   - **Tên file luôn cố định**: `gafc_audit_helper.xlam` (không thay đổi)
   - Version chỉ lưu trong metadata/manifest

2. **Tính SHA256 hash**:
   ```powershell
   Get-FileHash "gafc_audit_helper.xlam" -Algorithm SHA256
   ```

3. **Tạo GitHub Release mới**:
   - Tag: `v1.1.0`
   - Upload file `gafc_audit_helper.xlam`
   - Copy download URL

4. **Cập nhật manifest** (`releases/audit_tool.json`):
   ```json
   {
     "latest": "1.1.0",
     "download_url": "https://github.com/.../v1.1.0/gafc_audit_helper.xlam",
     "sha256": "abc123def456...",
     "release_date": "2025-12-10",
     "release_notes": "Bug fixes and improvements"
   }
   ```

5. **Commit và push manifest**:
   ```bash
   git add releases/audit_tool.json
   git commit -m "Release v1.1.0"
   git push
   ```

6. **Chờ auto-update chạy** (hoặc test ngay):
   ```powershell
   .\scripts\update_audit_helper.ps1
   ```

### Lưu ý:
- ✅ Tên file **luôn giữ nguyên** `gafc_audit_helper.xlam`
- ✅ Script tự động **replace** file cũ bằng file mới
- ✅ File cũ được **backup** thành `.bak` trước khi update
- ✅ Version tracking trong manifest field `"latest"`

## 🧪 Testing

### Test Manual Update

```powershell
.\scripts\update_audit_helper.ps1
```

Kết quả mong đợi:
- Nếu Excel đang chạy → "Excel is running. Skipping update."
- Nếu đã có version mới → "Already on latest version"
- Nếu có update → Download và cài đặt tự động

### Check Log File

```powershell
Get-Content "$env:TEMP\gafc_update.log" -Tail 20
```

### Verify Scheduled Task

```powershell
Get-ScheduledTask -TaskName "GAFC Audit Helper Auto Update"
```

## 🗑️ Gỡ Bỏ Auto-Update

```powershell
.\scripts\remove_auto_update.ps1
```

## 🔐 Bảo Mật

- ✅ SHA256 verification - Đảm bảo file không bị giả mạo
- ✅ Backup tự động - File cũ được backup trước khi update
- ✅ Check Excel running - Không update khi đang dùng
- ✅ Network check - Chỉ chạy khi có mạng

## ❓ Troubleshooting

### Update không chạy?

1. Check scheduled task:
   ```powershell
   Get-ScheduledTask -TaskName "GAFC Audit Helper Auto Update" | Get-ScheduledTaskInfo
   ```

2. Check log file:
   ```powershell
   Get-Content "$env:TEMP\gafc_update.log"
   ```

3. Chạy manual để debug:
   ```powershell
   .\scripts\update_audit_helper.ps1
   ```

### Manifest URL không accessible?

- Kiểm tra repository là public hoặc có token access
- Test URL trực tiếp trong browser
- Check firewall/proxy settings

### File bị locked khi update?

- Đảm bảo đóng tất cả instance của Excel
- Check Task Manager → Kill process `EXCEL.EXE` nếu cần

## 📞 Support

Nếu gặp vấn đề, check log file và GitHub Issues.
