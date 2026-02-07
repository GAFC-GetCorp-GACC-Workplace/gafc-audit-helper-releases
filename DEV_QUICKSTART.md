# Developer Quick Start

## Setup Password (Chỉ cần làm 1 lần)

1. Mở file `vba_password.txt`
2. Xóa dấu `#` ở dòng cuối
3. Thay `YOUR_PASSWORD_HERE` bằng password của bạn
4. Lưu file

**Ví dụ:**
```
# VBA Password Configuration
# YOUR_PASSWORD_HERE  ← Xóa dòng này

MyPassword123  ← Thêm password vào đây
```

## Build Commands

```bash
# Development build (KHÔNG khóa VBA - dùng khi đang code)
python rebuild_xlam_dev.py

# Production build (CÓ khóa VBA - dùng khi deploy)
python rebuild_xlam_release.py
```

**Script tự động:**
- Đọc password từ `vba_password.txt`
- Unlock VBA (nếu file source đã khóa)
- Import modules mới
- Lock VBA lại với password (production mode)
- (Tùy chọn) `--unviewable` nếu muốn khóa không thể mở bằng password

## Unlock VBA để xem code

Nếu bạn cần xem lại VBA code từ file đã build:

```bash
python unlock_vba.py gafc_audit_helper_new.xlam
```

Sau đó mở VBA Editor và nhập password từ `vba_password.txt` để xem code.

## Workflow khi đang develop

1. **Chỉnh sửa code** trong `extracted_clean/*.bas`
2. **Build dev version**: `python rebuild_xlam_dev.py`
3. **Test** file `gafc_audit_helper_new_dev.xlam`
4. **Lặp lại** bước 1-3 cho đến khi OK
5. **Build production**: `python rebuild_xlam_release.py` (chỉ cần 1 lần cuối)

## Extract code từ XLAM về .bas

Nếu cần extract code từ file XLAM ra file .bas:

```bash
python extract_modules.py
```

## Tại sao cần dev mode?

- ✅ **Tiết kiệm thời gian**: Không cần nhập password mỗi lần build
- ✅ **Dễ debug**: VBA code không bị khóa, mở được VBA Editor
- ✅ **Nhanh hơn**: Bỏ qua bước lock/unlock VBA

⚠️ **LƯU Ý**: File build bằng `--dev` KHÔNG nên deploy cho user (vì không có bảo mật)

## Chi tiết đầy đủ

Xem file [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md)
