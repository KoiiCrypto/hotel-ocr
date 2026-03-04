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

        # ── Preprocess: fix EXIF rotation + resize ảnh DT lớn ──────────────
        tmp_jpg = Path(tempfile.gettempdir()) / (str(uuid.uuid4()) + ".jpg")
        try:
            pil_img = Image.open(str(tmp))
            # Fix EXIF orientation (ảnh chụp dọc trên DT hay bị xoay)
            pil_img = ImageOps.exif_transpose(pil_img)
            # Convert HEIC/WEBP/PNG → RGB JPEG
            if pil_img.mode in ("RGBA", "P", "LA"):
                pil_img = pil_img.convert("RGB")
            elif pil_img.mode != "RGB":
                pil_img = pil_img.convert("RGB")
            # Resize nếu quá lớn (DT chụp 12MP+ = ~4000px)
            MAX_DIM = 2400
            w, h = pil_img.size
            if max(w, h) > MAX_DIM:
                ratio = MAX_DIM / max(w, h)
                pil_img = pil_img.resize((int(w*ratio), int(h*ratio)), Image.LANCZOS)
            # Save as JPEG quality 88 — giảm từ 10MB xuống ~1MB
            pil_img.save(str(tmp_jpg), "JPEG", quality=88, optimize=True)
            ocr_path = str(tmp_jpg)
        except Exception as pil_err:
            app.logger.warning(f"PIL preprocess failed: {pil_err}, dùng file gốc")
            ocr_path = str(tmp)
            tmp_jpg = None

        img_orig = cv2.imread(ocr_path)
        if img_orig is None:
            return jsonify(error="Không đọc được ảnh — định dạng không hỗ trợ"), 400

        method_parts = []

        # 1. QR
        qr_info = oc.read_qr(img_orig)
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
