# 🔍 RegXHelper — COM DLL for Regular Expressions in VBA/VB6

**RegXHelper** là thư viện DLL miễn phí, viết bằng C++ Builder, cung cấp giao diện COM tương thích với VB6/VBA để xử lý biểu thức chính quy (Regular Expressions) một cách mạnh mẽ và dễ sử dụng.

📦 Không cần .NET, không cần cài đặt phức tạp — chỉ cần đăng ký DLL và sử dụng ngay trong Excel, Word, Access, hoặc bất kỳ môi trường COM nào.

📁 Repo này có kèm theo file Excel ví dụ (`RegXHelper_Demo.xlsm`) để bạn dễ dàng thử nghiệm các hàm.
---

## 📌 Tính năng chính

| Thuộc tính / Hàm | Mô tả |
|------------------|-------|
| `Pattern`        | 🔧 Định nghĩa biểu thức chính quy |
| `Global`         | 🌍 Tìm tất cả kết quả trong chuỗi |
| `IgnoreCase`     | 🔠 Bỏ qua phân biệt chữ hoa/thường |
| `Execute()`      | 📋 Trả về danh sách các chuỗi khớp |
| `Replace()`      | ✂️ Thay thế toàn bộ chuỗi khớp |
| `FirstMatch()`   | 🎯 Lấy chuỗi khớp đầu tiên |
| `ReplaceFirst()` | ✨ Thay thế chuỗi khớp đầu tiên |
| `Split()`        | 🔪 Tách chuỗi theo biểu thức chính quy |

---
## 🧠 Về nền tảng kỹ thuật 

Thư viện **RegXHelper** được phát triển bằng **C++ Builder**, kế thừa trực tiếp từ thư viện **System.RegularExpressions** — một nền tảng mạnh mẽ, kết tinh từ kiến thức và kinh nghiệm của rất nhiều lập trình viên trong cộng đồng Delphi/C++ Builder qua nhiều thế hệ.

RegXHelper không tái tạo lại hệ thống biểu thức chính quy. Thay vào đó, nó đóng vai trò như một lớp **wrapper trung gian**, giúp các ngôn ngữ như **VBA**, **VB6**, hoặc các môi trường **COM** có thể dễ dàng tiếp cận và sử dụng các tính năng biểu thức chính quy một cách trực quan, hiệu quả và thân thiện hơn.

> Tôi không viết lại biểu thức chính quy — tôi chỉ mở một cánh cửa để người dùng VBA có thể bước vào thế giới đó dễ dàng hơn.

---

## 🧠 Technical Foundation 

RegXHelper is developed using **C++ Builder**, built directly on top of the native **System.RegularExpressions** library — a powerful and mature engine shaped by the knowledge and contributions of many developers in the Delphi/C++ Builder community over generations.

This project does not attempt to reinvent regular expressions. Instead, it serves as a clean and efficient **wrapper layer**, allowing languages like **VBA**, **VB6**, and other **COM environments** to easily access and utilize regular expression features in a more intuitive and user-friendly way.

> I didn’t rewrite regular expressions — I simply opened a door so VBA users can walk into that world more easily.
---
## 🧪 Ví dụ sử dụng trong VBA (có sẵn trong file Excel)

```vb
Sub DemoRegX()
    Dim re As Object
    Set re = CreateObject("RegXHelper.RegExp")

    re.Pattern = "\d+"
    re.Global = True
    re.IgnoreCase = True

    Dim result
    result = re.Replace("Giá: 123 và 456", "###")

    MsgBox result  ' Kết quả: "Giá: ### và ###"
End Sub
```

📌 File Excel mẫu có các macro minh họa cho:
- Tách chuỗi
- Đếm số match
- Trích xuất nhóm tên
- Thay thế chuỗi khớp

---

## 🧠 Các hàm nâng cao (mang tính tượng trưng)

| Hàm | Mô tả |
|-----|------|
| `MatchCount()` | 🔢 Đếm số chuỗi khớp |
| `GetMatch()` | 🔍 Lấy chuỗi khớp theo chỉ số |
| `GroupByName()` | 🧩 Lấy nội dung nhóm tên |
| `GroupByIndex()` | 🧮 Lấy nội dung nhóm theo chỉ số |
| `GetGroups()` | 📚 Trả về tất cả nhóm trong một match |
| `GetMatchPosition()` | 📍 Vị trí chuỗi khớp |
| `HasMatch()` | ✅ Kiểm tra có khớp hay không |
| `MatchLengths()` | 📏 Độ dài các chuỗi khớp |
| `MatchValues()` | 🧾 Danh sách giá trị khớp |
| `GroupNames()` | 🏷️ Danh sách tên nhóm |
| `GroupExists()` | 🔎 Kiểm tra nhóm tên có tồn tại |
| `EscapePattern()` | 🛡️ Chuẩn hóa biểu thức |
| `IsValidPattern()` | ✔️ Kiểm tra biểu thức hợp lệ |
| `GetMatches()` | 📦 Trả về danh sách match |
| `GetMatchJSON()` | 🧪 Match dưới dạng JSON |
| `GetAllMatchesJSON()` | 📊 Tất cả match dưới dạng JSON |
| `SplitToArray()` | 🧨 Tách chuỗi thành mảng |

---

## 🛠 Cài đặt

1. Đăng ký DLL bằng `regsvr32`:
   ```cmd
   regsvr32 RegXHelper.dll
   ```

2. Tạo đối tượng trong VBA:
   ```vb
   Set re = CreateObject("RegXHelper.RegExp")
   ```

---

## 📁 Mã nguồn

Thư viện được viết bằng C++ Builder, sử dụng `System.RegularExpressions` và giao diện COM thông qua `TCppAutoObject<IRegExp>`. Bạn có thể mở rộng thêm các hàm như:

- `ReplaceExactLength`
- `SplitToDictionary`
- `GetMatchJSON`
- `GroupExists`
---
## 👤 Giới thiệu tác giả

**Tác giả:** Kieu Manh  
📧 Email: kieumanh366377@gmail.com  
📱 Phone: 0929.278.279 | 0929.278.379  
🛠 Phiên bản: 1.0 — Cập nhật lần cuối: 2025-09-27

> Tôi không viết lại biểu thức chính quy — tôi chỉ mở một cánh cửa để người dùng VBA có thể bước vào thế giới đó dễ dàng hơn.
---

## ❤️ Đóng góp

Nếu bạn thấy thư viện hữu ích, hãy ⭐ star repo hoặc chia sẻ cho cộng đồng VBA/VB6 nhé!
