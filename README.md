# Công cụ tạo slide từ ảnh (OCR)

## Yêu cầu
- Python 3.10+ (đã cài sẵn trong môi trường)
- Không cần quyền root. Nếu không cài được Tesseract, script sẽ dùng fallback (tạo slide ảnh gốc) hoặc EasyOCR nếu có.

## Cài đặt thư viện Python
Trong môi trường bị quản lý (PEP 668), venv có thể không khả dụng. Bạn có thể thử:

```bash
# Nếu tạo được venv:
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Nếu không tạo được venv, có thể buộc pip dùng cờ (cân nhắc rủi ro):
python3 -m pip install --break-system-packages -r requirements.txt
```

## Cách dùng
```bash
python3 tools/make_slides.py "yeucau.jpg" "đề bài.jpg" \
  --output baithi.pptx \
  --lang vie \
  --fallback-lang eng \
  --title-from first-line \
  --wide \
  --image-fallback
```

### Tùy chọn thường dùng
- **--title-from**: first-line|filename — Lấy tiêu đề từ dòng OCR đầu hoặc tên file.
- **--no-preprocess**: Tắt tiền xử lý ảnh nếu OCR bị sai do threshold.
- **--font-name**, **--title-size**, **--bullet-size**, **--accent-color**: Tùy chỉnh phông và màu.
- **--image-fallback**: Nếu OCR không ra chữ, chèn ảnh gốc làm một slide.

## Ghi chú
- Để OCR tiếng Việt tốt, nên cài `tesseract-ocr` và `tesseract-ocr-vie` (cần quyền root). Trong môi trường không có quyền, script vẫn chạy với fallback. 
