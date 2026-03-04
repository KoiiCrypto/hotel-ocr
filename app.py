#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — Hotel OCR · Flask backend cho Render.com
Không ghi file trên server: mọi export trả thẳng về browser.
"""

import os, io, re, csv, json, uuid, tempfile, traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_file, Response
from PIL import Image, ImageOps

# ── Đọc lookup files ngay khi import ─────────────────────────────────────────
_DIR = Path(__file__).parent

# Import toàn bộ logic từ ocrspace.py
import sys
sys.path.insert(0, str(_DIR))

from ocrspace import (
    ocr_via_api, read_qr, parse_doc, merge_qr_ocr,
    parse_flexible_date, export_xml as _export_xml,
    match_country, parse_address_segment,
    _TINH_NORM, OUTPUT_CSV, CSV_HEADERS,
)
import cv2, numpy as np

# ── Flask app ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder=None)
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024  # 30 MB max upload

# ── Đọc HTML frontend (nhúng sẵn) ────────────────────────────────────────────
_HTML_PATH = _DIR / "static" / "index.html"

# ══════════════════════════════════════════════════════════════════════════════
#  ENHANCED QR READER
# ══════════════════════════════════════════════════════════════════════════════

def _read_qr_enhanced(img_bgr):
    """
    Thử nhiều kỹ thuật để đọc QR — tăng tỉ lệ thành công với ảnh DT.
    Thứ tự: OpenCV detector → wechat → preprocessed variants
    """
    import ocrspace as oc

    if img_bgr is None:
        return None

    # ── Thử 1: OpenCV QRCodeDetector thường (nhanh nhất) ──
    result = oc.read_qr(img_bgr)
    if result:
        return result

    h, w = img_bgr.shape[:2]

    # ── Thử 2: WeChatQRCode (chính xác hơn, cần module extra) ──
    try:
        detector = cv2.wechat_qrcode_WeChatQRCode()
        texts, _ = detector.detectAndDecode(img_bgr)
        for t in (texts or []):
            if t and "|" in t:
                parsed = oc.parse_qr_text(t) if hasattr(oc, "parse_qr_text") else None
                if parsed:
                    return parsed
    except Exception:
        pass

    # ── Thử 3: Các biến thể tiền xử lý ảnh ──
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    variants = []

    # 3a. Sharpen
    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    variants.append(("sharpen", cv2.filter2D(gray, -1, kernel)))

    # 3b. CLAHE tăng contrast
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    variants.append(("clahe", clahe.apply(gray)))

    # 3c. Adaptive threshold
    variants.append(("thresh", cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)))

    # 3d. Upscale 1.5x nếu ảnh nhỏ
    if max(h, w) < 1500:
        up = cv2.resize(gray, (int(w*1.5), int(h*1.5)), interpolation=cv2.INTER_CUBIC)
        variants.append(("upscale", up))

    # 3e. Invert (QR nền tối chữ sáng)
    variants.append(("invert", cv2.bitwise_not(gray)))

    qr_det = cv2.QRCodeDetector()
    for name, variant in variants:
        # convert back to BGR nếu cần
        if len(variant.shape) == 2:
            v_bgr = cv2.cvtColor(variant, cv2.COLOR_GRAY2BGR)
        else:
            v_bgr = variant
        text, _, _ = qr_det.detectAndDecode(v_bgr)
        if text and "|" in text:
            app.logger.info(f"QR decoded via variant: {name}")
            parsed = oc.parse_qr_text(text) if hasattr(oc, "parse_qr_text") else None
            if not parsed:
                # fallback: gọi read_qr trên ảnh variant
                parsed = oc.read_qr(v_bgr)
            if parsed:
                return parsed

    return None


# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return send_file(str(_HTML_PATH))


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    """Nhận ảnh → QR + OCR + parse → trả JSON (không ghi disk)."""
    if "image" not in request.files:
        return jsonify(error="Thiếu file ảnh"), 400

    f        = request.files["image"]
    checkin  = request.form.get("checkin",  "")
    checkout = request.form.get("checkout", "")
    phong    = request.form.get("phong",    "")
    lydo     = request.form.get("lydo",     "Du Lịch")
    apikey   = request.form.get("apikey",   "")

    # Override API key nếu user nhập
    import ocrspace as oc
    if apikey:
        oc.OCR_API_KEY = apikey

    # Lưu ảnh vào temp file (cần cho cv2 và ocr_via_api)
    ext  = Path(f.filename or "img.jpg").suffix or ".jpg"
    tmp  = Path(tempfile.gettempdir()) / (str(uuid.uuid4()) + ext)
    try:
        f.save(str(tmp))

        # ── Bước 1: load ảnh gốc full-res để đọc QR ─────────────────────────
        tmp_jpg = None
        img_full = None
        try:
            pil_orig = Image.open(str(tmp))
            pil_orig = ImageOps.exif_transpose(pil_orig)       # fix xoay EXIF
            if pil_orig.mode not in ("RGB",):
                pil_orig = pil_orig.convert("RGB")
            # Để đọc QR: dùng ảnh full-res, chỉ upscale nếu quá nhỏ
            w0, h0 = pil_orig.size
            qr_pil = pil_orig
            if max(w0, h0) < 1000:                             # ảnh nhỏ → upscale
                scale = 1000 / max(w0, h0)
                qr_pil = pil_orig.resize((int(w0*scale), int(h0*scale)), Image.LANCZOS)
            img_full = cv2.cvtColor(np.array(qr_pil), cv2.COLOR_RGB2BGR)
        except Exception as e:
            app.logger.warning(f"Load full-res failed: {e}")

        # ── Bước 2: resize + compress để gửi OCR API ─────────────────────────
        tmp_jpg = Path(tempfile.gettempdir()) / (str(uuid.uuid4()) + ".jpg")
        try:
            MAX_DIM = 2400
            w, h = pil_orig.size
            pil_ocr = pil_orig
            if max(w, h) > MAX_DIM:
                ratio = MAX_DIM / max(w, h)
                pil_ocr = pil_orig.resize((int(w*ratio), int(h*ratio)), Image.LANCZOS)
            pil_ocr.save(str(tmp_jpg), "JPEG", quality=88, optimize=True)
            ocr_path = str(tmp_jpg)
        except Exception as pil_err:
            app.logger.warning(f"PIL resize failed: {pil_err}, dùng file gốc")
            ocr_path = str(tmp)
            tmp_jpg = None

        # img_orig cho cv2 dùng ảnh đã resize (đủ cho OCR text)
        img_orig = cv2.imread(ocr_path) if ocr_path != str(tmp) else img_full
        if img_orig is None:
            return jsonify(error="Không đọc được ảnh — định dạng không hỗ trợ"), 400

        method_parts = []

        # 1. QR — dùng ảnh FULL RES để tăng độ chính xác
        qr_img = img_full if img_full is not None else img_orig
        qr_info = _read_qr_enhanced(qr_img)
        if qr_info:
            method_parts.append("QR")

        # 2. OCR
        lines = oc.ocr_via_api(ocr_path)
        if not lines:
            return jsonify(error="OCR không trả về kết quả — kiểm tra API key"), 422
        method_parts.append("OCR")

        # 3. Parse + merge
        ocr_info = oc.parse_doc(lines)
        final    = oc.merge_qr_ocr(qr_info, ocr_info)
        mrz      = final.pop("_mrz", {})
        sources  = final.pop("_sources", {})

        if any(mrz.values()):
            method_parts.append("MRZ")

        # 4. Ngày check-in / check-out
        ci = parse_flexible_date(checkin,  is_checkout=False) if checkin  else ""
        co = parse_flexible_date(checkout, is_checkout=True)  if checkout else ""
        final["tu_ngay"]        = (ci + " 15:00:00") if ci else ""
        final["den_ngay"]       = (co + " 11:00:00") if co else ""
        if phong: final["ten_phong_khoa"] = phong
        if lydo:  final["ly_do_luu_tru"]  = lydo

        return jsonify(
            data      = final,
            sources   = sources,
            mrz       = {k: v for k, v in mrz.items() if v},
            raw_lines = lines,
            method    = "+".join(method_parts) or "OCR",
        )

    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify(error=str(e)), 500

    finally:
        try: tmp.unlink()
        except: pass
        try:
            if 'tmp_jpg' in dir() and tmp_jpg and tmp_jpg.exists(): tmp_jpg.unlink()
        except: pass


@app.route("/api/export/xml", methods=["POST"])
def api_export_xml():
    """Nhận JSON rows → xuất pp.xml trả thẳng về browser."""
    rows = request.get_json(silent=True)
    if not rows:
        return jsonify(error="Không có dữ liệu"), 400

    passport_rows = [r for r in rows if _is_passport(r)]
    if not passport_rows:
        return jsonify(error="Không có khách Hộ chiếu"), 404

    xml_str = _build_xml(passport_rows)
    return Response(
        xml_str.encode("utf-8"),
        mimetype="application/xml",
        headers={"Content-Disposition": 'attachment; filename="pp.xml"'}
    )


@app.route("/api/export/csv", methods=["POST"])
def api_export_csv():
    """Nhận JSON rows → xuất output.csv trả về browser."""
    rows = request.get_json(silent=True)
    if not rows:
        return jsonify(error="Không có dữ liệu"), 400

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_HEADERS, extrasaction="ignore")
    writer.writeheader()
    for i, row in enumerate(rows, 1):
        row["STT"] = i
        writer.writerow(row)

    csv_bytes = ("\ufeff" + buf.getvalue()).encode("utf-8")
    return Response(
        csv_bytes,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="output.csv"'}
    )


# ══════════════════════════════════════════════════════════════════════════════
#  XML BUILDER (không cần temp file)
# ══════════════════════════════════════════════════════════════════════════════

def _is_passport(row):
    loai = row.get("Loại giấy tờ (*)", "").lower()
    return "chiếu" in loai or "passport" in loai

def _safe(v):
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", str(v or "")).strip()

def _strip_time(v):
    return v.split(" ")[0] if v else ""

def _gender_xml(v):
    v = v.strip().upper()
    if v in ("NAM", "M", "MALE", "1"): return "M"
    if v in ("NỮ", "NU", "F", "FEMALE", "0"): return "F"
    return v or "M"

def _build_xml(rows):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<KHAI_BAO_TAM_TRU>"]
    for i, r in enumerate(rows, 1):
        def t(tag, val):
            return f"        <{tag}>{_safe(val)}</{tag}>"
        lines += [
            "    <THONG_TIN_KHACH>",
            t("so_thu_tu",       i),
            t("ho_ten",          r.get("Họ và tên (*)", "").upper()),
            t("ngay_sinh",       r.get("Ngày, tháng, năm sinh (*)", "")),
            t("ngay_sinh_dung_den", "D"),
            t("gioi_tinh",       _gender_xml(r.get("Giới tính (*)", ""))),
            t("ma_quoc_tich",    r.get("Quốc tịch (*)", "").upper()),
            t("so_ho_chieu",     r.get("Số giấy tờ (*)", "")),
            t("so_phong",        r.get("Tên phòng/Khoa (*)", "")),
            t("ngay_den",        _strip_time(r.get("Thời gian lưu trú (từ ngày) (*)", ""))),
            t("ngay_di_du_kien", _strip_time(r.get("Thời gian lưu trú (đến ngày)", ""))),
            t("ngay_tra_phong",  ""),
            "    </THONG_TIN_KHACH>",
        ]
    lines.append("</KHAI_BAO_TAM_TRU>")
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8888))
    app.run(host="0.0.0.0", port=port, debug=False)
