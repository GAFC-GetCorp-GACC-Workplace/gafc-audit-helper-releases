# Build Instructions - VBA Project with Encrypted Password Protection

## Tổng quan

Scripts `rebuild_xlam_dev.py` và `rebuild_xlam_release.py` hỗ trợ build dev/prod riêng biệt.

## Bảo mật nâng cao

Password được mã hóa bằng:
- ✅ **Fernet encryption** (AES 128-bit) - nếu có `cryptography` package
- ✅ **Machine-specific key** - dựa trên hardware ID của máy tính
- ✅ **Base64 obfuscation** - fallback nếu không có cryptography

➡️ Password chỉ hoạt động trên máy tính đã lưu, không thể dùng trên máy khác!

## Cài đặt (khuyến nghị)

Để có mã hóa mạnh nhất, cài đặt thêm:
```bash
pip install cryptography
```

Nếu không cài, script vẫn chạy nhưng chỉ dùng obfuscation (bảo mật thấp hơn).

## Cách sử dụng

### TL;DR - Quick Start

```bash
# Development (khi đang code, không khóa VBA)
python rebuild_xlam_dev.py

# Production (build để deploy, có khóa VBA)
python rebuild_xlam_release.py
```

### Lần đầu tiên (Production build)

1. Chạy script build:
   ```bash
   python rebuild_xlam_release.py
   ```

2. Script sẽ hỏi password:
   ```
   VBA Project Password Configuration:
   Enter password to lock VBA project (or press Enter to skip locking)
   Password: ********
   ```

3. Nhập password bạn muốn dùng để khóa VBA project

4. Script sẽ hỏi có lưu password không:
   ```
   Save encrypted password to build_config.dat for future builds? (y/n):
   ```
   - Chọn **y**: Password được **MÃ HÓA** và lưu vào `build_config.dat`
   - Chọn **n**: Mỗi lần build phải nhập lại password

### Các lần sau

- Nếu đã lưu password vào `build_config.dat`:
  ```bash
  python rebuild_xlam_release.py
  ```
  Script sẽ tự động giải mã và dùng password, không cần nhập lại

- **LƯU Ý**: File `build_config.dat` chỉ hoạt động trên máy tính đã tạo ra nó (vì key phụ thuộc hardware ID)

### Build không cần password (Development mode)

**Cách 1: Development build (khuyến nghị khi đang code)**
```bash
python rebuild_xlam_dev.py
```
- Không hỏi password, tự động skip khóa VBA
- Tiện lợi khi đang develop và cần debug code

**Cách 2: Manual skip**
Nếu muốn build file KHÔNG khóa password (để test hoặc debug):
- Khi script hỏi password, nhấn **Enter** (bỏ trống)
- File output sẽ không bị khóa VBA project

## Quy trình tự động

Script sẽ tự động:

1. ✅ Copy file source → file output
2. ✅ Mở file output trong Excel
3. ✅ **Unlock VBA project** (nếu file source đã bị khóa)
4. ✅ Xóa các module cũ
5. ✅ Import các module mới từ `extracted_clean/`
6. ✅ Cập nhật version từ `releases/audit_tool.json`
7. ✅ **Lock VBA project với password**
8. ✅ Lưu và đóng file

## Bảo mật

### Các lớp bảo mật:

1. **Mã hóa AES** (nếu có cryptography package)
   - Sử dụng Fernet encryption (AES 128-bit CBC + HMAC)
   - Key được sinh từ hardware ID (machine-specific)

2. **Machine binding**
   - Password được mã hóa với key phụ thuộc vào:
     - Hostname (tên máy tính)
     - MAC address (địa chỉ phần cứng)
   - Không thể copy file config sang máy khác để dùng

3. **Git protection**
   - Files `build_config.dat` và `.build_key` đã thêm vào `.gitignore`
   - Password **KHÔNG BAO GIỜ** được commit lên Git

4. **VBA Project Lock**
   - Code VBA được khóa bằng password trong Excel
   - Không thể xem/chỉnh sửa code mà không biết password

## Lưu ý

- Nếu file source (`gafc_audit_helper.xlam`) đã bị khóa password, script sẽ tự động unlock trước khi xử lý
- DO NOT use release output (gafc_audit_helper_new.xlam) as template; keep gafc_audit_helper.xlam unlocked
- Password phải **giống nhau** cho cả unlock và lock
- Nếu quên password, bạn cần unlock thủ công trong Excel trước khi chạy script

## Output

File kết quả: `gafc_audit_helper_new.xlam`
- VBA project đã được lock với password
- Sẵn sàng để phân phối cho người dùng
