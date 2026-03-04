# Hotel OCR — Deploy lên Render.com

## Cấu trúc thư mục

```
hotel-ocr/
├── app.py              ← Flask backend
├── ocrspace.py         ← Logic OCR gốc
├── requirements.txt
├── render.yaml
├── tinh.txt
├── xa.txt
├── country.txt
└── static/
    └── index.html      ← Frontend
```

---

## Hướng dẫn deploy lên Render.com (miễn phí)

### Bước 1 — Tạo GitHub repo

1. Vào https://github.com/new
2. Tạo repo mới, ví dụ: `hotel-ocr`
3. **Không** tick "Add README"

### Bước 2 — Push code lên GitHub

Mở PowerShell trong thư mục project:

```powershell
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/TEN_BAN/hotel-ocr.git
git push -u origin main
```

### Bước 3 — Deploy lên Render

1. Vào https://render.com → **Sign up** (dùng GitHub account)
2. Click **"New +"** → **"Web Service"**
3. Chọn repo `hotel-ocr` → **Connect**
4. Render tự đọc `render.yaml`, chỉ cần click **"Create Web Service"**
5. Chờ ~3-5 phút build xong

### Bước 4 — Dùng thôi

URL dạng: `https://hotel-ocr.onrender.com`

---

## Lưu ý Render Free Tier

| | Free |
|---|---|
| RAM | 512 MB |
| CPU | Shared |
| Sleep | Sau 15 phút không dùng (cold start ~30s) |
| Bandwidth | 100 GB/tháng |

- **opencv-python-headless** thay vì `opencv-python` để tiết kiệm RAM
- **2 workers** Gunicorn là đủ cho free tier
- Timeout 120s để OCR API không bị cut

## Chạy local

```powershell
pip install -r requirements.txt
python app.py
# Mở http://localhost:8888
```
