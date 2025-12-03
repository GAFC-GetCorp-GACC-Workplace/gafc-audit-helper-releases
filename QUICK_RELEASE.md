# Quick Release Guide

## 🚀 Cách Release Nhanh Nhất

### Option 1: Double-Click (Windows CMD)

1. **Double-click** file `release.bat`
2. Nhập version (ví dụ: `1.0.1`)
3. Nhập message (hoặc Enter để skip)
4. Xong! Chờ GitHub Actions chạy

### Option 2: Git Bash

```bash
./release.sh
```

Rồi làm theo hướng dẫn trên màn hình.

### Option 3: PowerShell

```powershell
.\scripts\create_release.ps1 -Version "1.0.1" -Message "Bug fixes"
```

---

## 📝 Workflow Tự Động

Sau khi chạy script, GitHub Actions sẽ tự động:

1. ✅ Tính SHA256 hash của XLAM file
2. ✅ Update manifest (`releases/audit_tool.json`)
3. ✅ Clone public repo
4. ✅ Copy XLAM, scripts, README
5. ✅ Tạo installer ZIP package
6. ✅ Commit và push vào public repo
7. ✅ Tạo GitHub Release với files:
   - `gafc_audit_helper.xlam`
   - `gafc_audit_helper_installer.zip`
8. ✅ Release notes tự động với SHA256

**Xem tiến độ tại:**
https://github.com/muaroi2002/gafc-audit-helper/actions

**Kết quả release tại:**
https://github.com/muaroi2002/gafc-audit-helper-releases/releases

---

## ⚠️ Lưu Ý Trước Khi Release

### Checklist:

- [ ] File `gafc_audit_helper.xlam` đã build mới nhất (Save trong Excel)
- [ ] Code VBA đã test kỹ
- [ ] `DEV_ALLOW_BYPASS = False` trong `modLicenseAudit.bas`
- [ ] Version number tăng so với version trước
- [ ] Đã commit tất cả changes

### Nếu Release Lỗi:

**Xóa tag và thử lại:**

```bash
# Xóa local tag
git tag -d v1.0.1

# Xóa remote tag
git push origin :refs/tags/v1.0.1

# Xóa release trên GitHub (nếu đã tạo)
gh release delete v1.0.1 --yes

# Tạo lại
./release.sh
```

---

## 🔧 Troubleshooting

### Script không chạy được

**Windows CMD:**
```cmd
# Chạy trực tiếp
release.bat
```

**Git Bash:**
```bash
# Make executable
chmod +x release.sh

# Run
./release.sh
```

**PowerShell:**
```powershell
# Cho phép chạy scripts
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# Chạy
.\scripts\create_release.ps1 -Version "1.0.1"
```

### Workflow thất bại

Kiểm tra:
1. Secret `PUBLIC_REPO_TOKEN` đã add chưa?
2. Token còn hạn chưa?
3. File `gafc_audit_helper.xlam` có trong repo chưa?

Xem logs chi tiết:
https://github.com/muaroi2002/gafc-audit-helper/actions

---

## 🎯 Version Numbering

Format: `MAJOR.MINOR.PATCH`

- **MAJOR** (1.0.0 → 2.0.0): Breaking changes
- **MINOR** (1.0.0 → 1.1.0): New features
- **PATCH** (1.0.0 → 1.0.1): Bug fixes

Ví dụ:
- `1.0.1` - Fix lỗi nhỏ
- `1.1.0` - Thêm tính năng mới
- `2.0.0` - Thay đổi lớn (breaking changes)

---

**Chúc mừng bạn đã setup xong hệ thống release tự động! 🎉**
