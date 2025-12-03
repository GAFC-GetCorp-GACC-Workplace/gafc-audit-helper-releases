# Release Workflow - Hướng Dẫn Đơn Giản

## 🎯 Tóm Tắt

Bạn **CHỈ** cần làm việc với **repo private** (`gafc-audit-helper`).
Repo public (`gafc-audit-helper-releases`) sẽ **TỰ ĐỘNG** được cập nhật qua GitHub Actions.

---

## 📋 Quy Trình Release Mới

### Bước 1: Setup One-Time (Chỉ Làm 1 Lần)

1. **Tạo Personal Access Token**
   - Truy cập: https://github.com/settings/tokens
   - Click "Generate new token (classic)"
   - Tên: `GAFC Release Automation`
   - Scopes cần chọn:
     - ✅ `repo` (full control)
     - ✅ `workflow`
   - Copy token (lưu lại an toàn)

2. **Add Secret vào Private Repo**
   - Vào: https://github.com/muaroi2002/gafc-audit-helper/settings/secrets/actions
   - Click "New repository secret"
   - Name: `PUBLIC_REPO_TOKEN`
   - Value: [Paste token vừa tạo]
   - Click "Add secret"

3. **Tạo Public Repo** (nếu chưa có)
   ```powershell
   # Trên GitHub, tạo repo mới:
   # - Name: gafc-audit-helper-releases
   # - Visibility: Public
   # - DON'T initialize with README (để trống)
   ```

### Bước 2: Release Mỗi Lần Có Phiên Bản Mới

Sau khi setup xong, **MỖI LẦN** bạn muốn release version mới:

```powershell
# 1. Đảm bảo file XLAM đã build xong
# (Mở Excel, save gafc_audit_helper.xlam)

# 2. Commit changes trong private repo
cd E:\audit\GAFC_Audit_Helper_Release
git add .
git commit -m "Update to v1.0.1"

# 3. Tạo tag và push
git tag v1.0.1
git push origin main
git push origin v1.0.1

# 4. XEM MAGIC XẢY RA! 🎉
# - Vào GitHub Actions của private repo
# - Workflow tự động chạy
# - Public repo tự động cập nhật
# - Release tự động được tạo
```

**Chỉ vậy thôi!** Không cần chạy script gì thêm.

---

## 🔄 Workflow Tự Động Sẽ Làm Gì?

Khi bạn push tag `v*.*.*`, GitHub Actions tự động:

1. ✅ Tính SHA256 của XLAM file
2. ✅ Update `releases/audit_tool.json` với version mới
3. ✅ Clone public repo
4. ✅ Copy files cần thiết:
   - `gafc_audit_helper.xlam`
   - `releases/audit_tool.json`
   - Scripts: install, update, uninstall, setup_auto_update, remove_auto_update
   - `README.md`
5. ✅ Tạo `gafc_audit_helper_installer.zip`
6. ✅ Commit và push vào public repo
7. ✅ Tạo GitHub Release trong public repo với:
   - XLAM file
   - Installer ZIP
   - Release notes với SHA256

---

## 📁 Cấu Trúc Repo

### Private Repo (gafc-audit-helper)
```
gafc-audit-helper/
├── .github/workflows/release.yml  ← Workflow tự động
├── gafc_audit_helper.xlam         ← Build file này trong Excel
├── extracted_clean/               ← Source code VBA
├── releases/audit_tool.json       ← Manifest (auto-updated)
└── scripts/                       ← Tất cả scripts
```

### Public Repo (gafc-audit-helper-releases) - TỰ ĐỘNG
```
gafc-audit-helper-releases/
├── gafc_audit_helper.xlam              ← Auto-synced
├── gafc_audit_helper_installer.zip    ← Auto-generated
├── releases/audit_tool.json            ← Auto-synced
├── scripts/                            ← Auto-synced
│   ├── install_audit_helper.ps1
│   ├── update_audit_helper.ps1
│   ├── uninstall_audit_helper.ps1
│   ├── setup_auto_update.ps1
│   └── remove_auto_update.ps1
└── README.md                           ← Auto-synced
```

---

## 🐛 Troubleshooting

### Workflow thất bại?

**Kiểm tra:**
1. `PUBLIC_REPO_TOKEN` secret đã add chưa?
2. Token có đủ permissions (`repo` + `workflow`)?
3. Public repo đã tạo chưa?
4. File `gafc_audit_helper.xlam` có trong private repo chưa?

**Xem logs:**
- Vào: https://github.com/muaroi2002/gafc-audit-helper/actions
- Click vào workflow run thất bại
- Xem output từng step

### Muốn test không tạo release thật?

Tạo tag test:
```powershell
git tag v0.0.1-test
git push origin v0.0.1-test
```

Sau đó xóa:
```powershell
gh release delete v0.0.1-test --yes
git tag -d v0.0.1-test
git push origin :refs/tags/v0.0.1-test
```

---

## 📊 Version Numbering

Sử dụng Semantic Versioning:
- `v1.0.0` - Major release (breaking changes)
- `v1.1.0` - Minor release (new features)
- `v1.0.1` - Patch release (bug fixes)

---

## ✅ Checklist Trước Khi Release

- [ ] Code VBA đã update và test
- [ ] File XLAM đã build (Save trong Excel)
- [ ] `DEV_ALLOW_BYPASS = False` trong modLicenseAudit.bas
- [ ] Version number đã tăng (trong tag)
- [ ] Commit message rõ ràng
- [ ] Token secret đã setup (chỉ lần đầu)

---

**Lưu ý:** Sau khi workflow chạy xong, kiểm tra:
1. Public repo releases: https://github.com/muaroi2002/gafc-audit-helper-releases/releases
2. Manifest URL: https://raw.githubusercontent.com/muaroi2002/gafc-audit-helper-releases/main/releases/audit_tool.json
3. Test auto-update: `.\scripts\update_audit_helper.ps1`
