# Lavasa AI PowerPoint Add-in

## Cách cài đặt:

1. Copy toàn bộ các file:
    - manifest.xml
    - taskpane.html
    - taskpane.css
    - taskpane.js
2. Gom chúng lại thành 1 thư mục (ví dụ: Lavasa-AI-PPT-Addin)
3. Nén thư mục thành 1 file .zip
4. Mở PowerPoint > Tab Developer > Add-ins > Sideload Add-in
5. Chọn tới file manifest.xml

## Yêu cầu:

- PowerPoint Office 365 hoặc 2021 trở lên
- Có API Key OpenAI để nhập vào Add-in

## Ghi chú:

- Nếu cần thêm model AI khác như Claude, Gemini => sửa taskpane.js theo hướng dẫn.
- Nếu dùng localhost test ➔ update SourceLocation trong manifest.xml cho đúng.

Happy Presenting! 🚀
