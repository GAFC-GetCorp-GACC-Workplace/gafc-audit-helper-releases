# Bảo mật VBA Project - Hướng dẫn chi tiết

## Tổng quan các biện pháp bảo mật

### 1. Mã hóa Password (trong Python script)

**Cách hoạt động:**
- Password VBA được mã hóa trước khi lưu vào `build_config.dat`
- Sử dụng **Fernet encryption** (AES 128-bit CBC mode + HMAC)
- Key mã hóa được sinh từ **hardware ID** của máy tính

**Ưu điểm:**
- ✅ Password không lưu dạng plain text
- ✅ Không thể copy file config sang máy khác
- ✅ Không thể reverse engineer dễ dàng

**Hạn chế:**
- ⚠️ Nếu có quyền truy cập máy tính, vẫn có thể chạy script để decrypt
- ⚠️ Đây là bảo vệ lớp script Python, KHÔNG phải lớp Excel

### 2. VBA Project Password Protection (trong Excel)

**Cách hoạt động:**
- Excel tự mã hóa VBA code bằng password
- Không thể xem/chỉnh sửa code trong VBA Editor mà không có password

**Ưu điểm:**
- ✅ Bảo vệ trực tiếp trong file Excel
- ✅ Ngăn người dùng xem source code VBA
- ✅ Chuẩn bảo mật của Microsoft Office

**Hạn chế:**
- ⚠️ Password protection của VBA **KHÔNG PHẢI LÀ MÃ HÓA MẠNH**
- ⚠️ Có nhiều công cụ crack VBA password (VBA Password Remover, etc.)
- ⚠️ Chỉ ngăn người dùng thông thường, không ngăn được hacker chuyên nghiệp

### 3. Các biện pháp bảo mật bổ sung

#### 3.1. Code Obfuscation (Làm rối code)
**Không được implement trong project này**

Nếu muốn bảo mật cao hơn, có thể:
- Đổi tên biến thành a1, a2, b1, b2...
- Xóa comments
- Ghép nhiều lệnh vào 1 dòng
- Sử dụng công cụ VBA obfuscator

**Nhược điểm:** Code khó maintain sau này

#### 3.2. Compile VBA to DLL
**Không được implement trong project này**

- Chuyển logic quan trọng sang DLL (C++/C#)
- VBA chỉ gọi DLL function
- DLL compiled, khó reverse engineer hơn

**Nhược điểm:**
- Phức tạp, khó triển khai
- Cần phân phối thêm file DLL

#### 3.3. License Key System
**ĐÃ ĐƯỢC IMPLEMENT** (modLicenseAudit.bas)

- Kiểm tra license trước khi chạy
- Có thể kết hợp với server validation

**Ưu điểm:**
- ✅ Kiểm soát ai được dùng tool
- ✅ Có thể thu hồi license từ xa

#### 3.4. Digital Signature
**Không được implement trong project này**

- Ký số file Excel bằng certificate
- Đảm bảo file không bị chỉnh sửa

**Cách thực hiện:**
1. Mua/tạo code signing certificate
2. Sign file XLAM trong Excel: File → Info → Protect Workbook → Add Digital Signature

## So sánh mức độ bảo mật

| Biện pháp | Mức độ | Ngăn ai? | Độ khó crack |
|-----------|---------|----------|--------------|
| VBA Password Lock | ⭐⭐ | User thông thường | Dễ (có tool free) |
| Encrypted Password Storage | ⭐⭐⭐ | Người copy config file | Trung bình |
| VBA Obfuscation | ⭐⭐⭐ | Người đọc code | Trung bình |
| Compile to DLL | ⭐⭐⭐⭐ | Reverse engineer | Khó |
| License Server + DLL | ⭐⭐⭐⭐⭐ | Mọi người | Rất khó |

## Khuyến nghị cho project này

### Mức bảo mật hiện tại: ⭐⭐⭐ (Trung bình - Tốt)

**Đã có:**
- ✅ VBA Password Lock (ngăn user thông thường)
- ✅ Encrypted password storage (bảo vệ config file)
- ✅ License system (kiểm soát người dùng)
- ✅ Git protection (không commit password)

**Phù hợp cho:**
- Môi trường doanh nghiệp nội bộ
- Bảo vệ IP khỏi người dùng thông thường
- Ngăn chỉnh sửa code tùy tiện

**KHÔNG phù hợp nếu:**
- Cần bảo vệ khỏi hacker chuyên nghiệp
- Cần bảo mật military-grade
- Source code là tài sản cực kỳ giá trị

## Nâng cao bảo mật (tùy chọn)

Nếu cần bảo mật cao hơn, thực hiện theo thứ tự:

### Bước 1: Code Obfuscation (Khuyến nghị cao)
```vba
' Trước:
Dim connectionString As String
connectionString = "Server=myServer;Database=myDB"

' Sau obfuscate:
Dim a1 As String: a1 = "Server=myServer;Database=myDB"
```

### Bước 2: Digital Signature (Dễ thực hiện)
- Mua code signing certificate (~$100-300/năm)
- Hoặc tạo self-signed certificate (free, nhưng không trusted)

### Bước 3: Move sensitive logic to DLL (Phức tạp)
- Viết C# DLL cho các phần quan trọng
- VBA chỉ gọi DLL functions

### Bước 4: Implement online license validation (Rất phức tạp)
- Tạo web service kiểm tra license
- VBA phải online để hoạt động

## Kết luận

**Bảo mật hiện tại là ĐỦ DÙNG** cho hầu hết trường hợp:
- ✅ Ngăn 95% người dùng thông thường xem/chỉnh sửa code
- ✅ Bảo vệ password không bị commit nhầm lên Git
- ✅ Kiểm soát license người dùng

**Lưu ý quan trọng:**
> Không có hệ thống bảo mật nào là 100%. VBA code trong Excel có thể bị crack nếu người tấn công có đủ kỹ năng và thời gian. Mục tiêu là làm CHO VIỆc CRACK TRỞ NÊN KHÓ KHĂN VÀ TỐN THỜI GIAN đến mức không đáng giá.
