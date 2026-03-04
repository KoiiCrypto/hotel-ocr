#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
doc_cccd_local.py - Tích hợp OCR + Xuất file
──────────────────────────────────────────────────────────────
Đọc CCCD/Hộ chiếu → Xuất:
  - Việt Nam (CCCD/CMND) → Excel (.xlsx)
  - Người nước ngoài (Passport) → XML (.xml)

Cài đặt:
    pip install requests opencv-python numpy pyzbar pillow pandas openpyxl

Cách dùng:
    python doc_cccd_local.py <ảnh>          # 1 file
    python doc_cccd_local.py -d <thư_mục>   # toàn bộ ảnh trong thư mục

    # Với tham số:
    python doc_cccd_local.py cccd.jpg --checkin 0302 --checkout 10 --phong "Phòng 101" --lydo "Du lịch"

    # Chỉ xuất file (đã có output.csv):
    python doc_cccd_local.py --export-only
"""

import sys
import re
import csv
import cv2
import base64
import requests
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from xml.etree import ElementTree as ET
from xml.dom import minidom

# ── QR backend ────────────────────────────────────────────────────────────────
QR_BACKEND = "opencv"
try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    QR_BACKEND = "pyzbar"
except ImportError:
    pass

# ── OCR.space config ──────────────────────────────────────────────────────────
OCR_API_KEY = "K84757334988957"
OCR_API_URL = "https://api.ocr.space/parse/image"
OCR_ATTEMPTS = [
    {"OCREngine": 2, "language": "auto"},
    {"OCREngine": 3, "language": "auto"},
]

OUTPUT_CSV = "output.csv"

# ── Đường dẫn lookup files ────────────────────────────────────────────────
_DIR = Path(__file__).parent
TINH_FILE    = _DIR / "tinh.txt"
XA_FILE      = _DIR / "xa.txt"
COUNTRY_FILE = _DIR / "country.txt"

# ── Excel Template ────────────────────────────────────────────────────────────
EXCEL_TEMPLATE = _DIR / "file-mau-import-thong-bao-luu-tru.xlsx"


# ══════════════════════════════════════════════════════════════════════════════
#  Helper Functions - File Lookup
# ══════════════════════════════════════════════════════════════════════════════

def _load_lines(path: Path) -> list[str]:
    if not path.exists():
        print(f"⚠️  Không tìm thấy lookup file: {path}")
        return []
    with open(path, encoding="utf-8", errors="replace") as f:
        return [ln.strip() for ln in f if ln.strip()]


def _load_country(path: Path) -> dict[str, str]:
    name2code: dict[str, str] = {}
    if not path.exists():
        print(f"⚠️  Không tìm thấy lookup file: {path}")
        return name2code
    with open(path, encoding="utf-8", errors="replace") as f:
        for ln in f:
            ln = ln.strip()
            if not ln:
                continue
            parts = ln.rsplit(",", 1)
            if len(parts) == 2:
                name, code = parts[0].strip().strip('"'), parts[1].strip()
                name2code[name.lower()] = code
                name2code[code.lower()] = code

    ALIASES = {
        "republic of korea":       "KOR", "korea, republic of":      "KOR",
        "south korea":             "KOR", "democratic people's republic of korea": "PRK",
        "north korea":             "PRK", "vietnam":                 "VNM",
        "viet nam":                "VNM", "việt nam":                "VNM",
        "socialist republic of viet nam": "VNM", "united states":           "USA",
        "united states of america": "USA", "usa":                     "USA",
        "united kingdom":          "GBR", "uk":                      "GBR",
        "great britain":           "GBR", "china":                   "CHN",
        "people's republic of china": "CHN", "taiwan":                  "TWN",
        "russia":                  "RUS", "iran":                    "IRN",
        "syria":                   "SYR", "tanzania":                "TZA",
        "czech republic":          "CZE", "czechia":                 "CZE",
        "laos":                    "LAO", "cambodia":                "KHM",
        "myanmar":                 "MMR", "burma":                   "MMR",
    }
    for alias, code in ALIASES.items():
        name2code[alias] = code

    return name2code


_TINH_LIST    = _load_lines(TINH_FILE)
_XA_LIST      = _load_lines(XA_FILE)
_COUNTRY_MAP  = _load_country(COUNTRY_FILE)

_TINH_NORM = {}
for t in _TINH_LIST:
    key = re.sub(r"^(thành phố|tỉnh)\s*", "", t, flags=re.I).strip().lower()
    _TINH_NORM[key] = t
    _TINH_NORM[t.lower()] = t

_XA_PREFIXES = r"^(phường|xã|thị trấn|thị xã|đặc khu)\s*"
_XA_NORM = {}
for x in _XA_LIST:
    _XA_NORM[x.lower()] = x
    stripped = re.sub(_XA_PREFIXES, "", x, flags=re.I).strip().lower()
    if stripped:
        _XA_NORM[stripped] = x

_TINH_KEYS_SORTED = sorted(_TINH_NORM, key=len, reverse=True)
_XA_KEYS_SORTED   = sorted(_XA_NORM,   key=len, reverse=True)


def _build_regex():
    return {
        'tinh':   re.compile(r'^(tỉnh|thành phố|tp\.?)\s+', re.I),
        'huyen':  re.compile(r'^(quận|huyện|thị xã|tx\.?)\s+', re.I),
        'xa':     re.compile(r'^(phường|xã|thị trấn|p\.)\s+', re.I),
        'diachi': re.compile(r'^(tổ\s+\d|khóm|khu phố|thôn|ấp|đội|số\s+\d|ngõ|ngách|\d+/)', re.I),
    }

_LEVEL_RE = _build_regex()


def _detect_level(token: str):
    t = token.strip()
    for lvl, rx in _LEVEL_RE.items():
        if rx.match(t):
            return lvl
    return None


def _strip_admin_prefix(token: str) -> str:
    t = token.strip()
    for rx in [_LEVEL_RE['tinh'], _LEVEL_RE['huyen'], _LEVEL_RE['xa']]:
        t = rx.sub('', t).strip()
    return t


def _wb_search(key: str, text: str) -> bool:
    NW = r'(?<![\w\u00C0-\u024F\u1EA0-\u1EF9])'
    NWE = r'(?![\w\u00C0-\u024F\u1EA0-\u1EF9])'
    try:
        return bool(re.search(NW + re.escape(key) + NWE, text, re.I | re.U))
    except re.error:
        return key in text


def match_tinh(text: str) -> str:
    tl = text.lower().strip()
    if tl in _TINH_NORM:
        return _TINH_NORM[tl]
    for key in _TINH_KEYS_SORTED:
        if len(key) < 3:
            continue
        if _wb_search(key, tl):
            return _TINH_NORM[key]
    return ""


def match_xa_exact(name: str) -> str:
    pl = name.lower().strip()
    pl_s = re.sub(_XA_PREFIXES, "", pl, flags=re.I).strip()

    if pl   in _XA_NORM: return _XA_NORM[pl]
    if pl_s in _XA_NORM: return _XA_NORM[pl_s]

    for key in _XA_KEYS_SORTED:
        if len(key.split()) >= 2 and len(key) >= 5:
            if _wb_search(key, pl):
                return _XA_NORM[key]
    return ""


def parse_address_segment(addr: str) -> dict:
    result = {'tinh_tp': '', 'quan_huyen': '', 'phuong_xa': '', 'dia_chi_chi_tiet': ''}
    tokens = [t.strip() for t in addr.split(',') if t.strip()]
    if not tokens: return result

    assigned = [None] * len(tokens)

    for i, tok in enumerate(tokens):
        lvl = _detect_level(tok)
        if lvl: assigned[i] = lvl

    for i, tok in enumerate(tokens):
        if assigned[i] is not None: continue
        name = _strip_admin_prefix(tok)
        if match_tinh(name):    assigned[i] = 'tinh';  continue
        if match_xa_exact(name): assigned[i] = 'xa';   continue
        assigned[i] = '?'

    for i, a in enumerate(assigned):
        if a == 'tinh' and i > 0 and assigned[i-1] == '?':
            assigned[i-1] = 'huyen'

    for i, a in enumerate(assigned):
        if a == '?': assigned[i] = 'diachi'

    diachi_parts = []
    for tok, lvl in zip(tokens, assigned):
        name = _strip_admin_prefix(tok)
        if lvl == 'tinh':
            result['tinh_tp']    = match_tinh(name) or tok
        elif lvl == 'huyen':
            result['quan_huyen'] = tok
        elif lvl == 'xa':
            result['phuong_xa']  = match_xa_exact(name) or tok
        else:
            diachi_parts.append(tok)
    result['dia_chi_chi_tiet'] = ', '.join(diachi_parts)
    return result


def match_country(text: str) -> str:
    if not text.strip():
        return ""
    tl = text.strip()

    if tl.lower() in _COUNTRY_MAP:
        return _COUNTRY_MAP[tl.lower()]

    for code in re.findall(r'(?<![A-Z])([A-Z]{3})(?![A-Z])', tl):
        if code.lower() in _COUNTRY_MAP:
            return _COUNTRY_MAP[code.lower()]

    tl_lower = tl.lower()
    for key in sorted(_COUNTRY_MAP, key=len, reverse=True):
        if len(key.split()) >= 2 and len(key) >= 6:
            try:
                NW  = r'(?<![\w\u00C0-\u024F\u1EA0-\u1EF9])'
                NWE = r'(?![\w\u00C0-\u024F\u1EA0-\u1EF9])'
                if re.search(NW + re.escape(key) + NWE, tl_lower, re.I | re.U):
                    return _COUNTRY_MAP[key]
            except re.error:
                if key in tl_lower:
                    return _COUNTRY_MAP[key]

    if tl_lower in _COUNTRY_MAP:
        return _COUNTRY_MAP[tl_lower]

    return ""


# ══════════════════════════════════════════════════════════════════════════════
#  Date Input Parser (linh hoạt cho check-in/check-out)
# ══════════════════════════════════════════════════════════════════════════════

def parse_flexible_date(user_input: str, is_checkout: bool = False) -> str:
    """Parse ngày nhập linh hoạt."""
    if not user_input:
        return ""

    user_input = user_input.strip()
    now = datetime.now()
    current_year = now.year
    current_month = now.month

    # Format: dd/mm/yyyy
    match = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$', user_input)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
            return f"{day:02d}/{month:02d}/{year}"

    # Format: dd/mm/yy
    match = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$', user_input)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        year = 2000 + year if year < 30 else 1900 + year
        if 1 <= day <= 31 and 1 <= month <= 12:
            return f"{day:02d}/{month:02d}/{year}"

    # Format: ddmmyyyy liền
    match = re.match(r'^(\d{2})(\d{2})(\d{4})$', user_input)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
            return f"{day:02d}/{month:02d}/{year}"

    # Format: ddmmyy liền
    match = re.match(r'^(\d{2})(\d{2})(\d{2})$', user_input)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        year = 2000 + year if year < 30 else 1900 + year
        if 1 <= day <= 31 and 1 <= month <= 12:
            return f"{day:02d}/{month:02d}/{year}"

    # Format: mmdd (0302 → 03/02/năm hiện tại)
    match = re.match(r'^(\d{2})(\d{2})$', user_input)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        if 1 <= day <= 31 and 1 <= month <= 12:
            if is_checkout and month < current_month:
                year = current_year + 1
            else:
                year = current_year
            return f"{day:02d}/{month:02d}/{year}"

    # Format: chỉ ngày (1-31)
    if user_input.isdigit():
        day = int(user_input)
        if 1 <= day <= 31:
            if is_checkout and day < now.day:
                if current_month == 12:
                    month = 1
                    year = current_year + 1
                else:
                    month = current_month + 1
                    year = current_year
            else:
                month = current_month
                year = current_year
            return f"{day:02d}/{month:02d}/{year}"

    return ""


def ask_checkin_date(is_checkout: bool = False) -> str:
    """Hỏi ngày check-in hoặc check-out từ user."""
    label = "CHECK-OUT" if is_checkout else "CHECK-IN"
    default_time = "11:00:00" if is_checkout else "15:00:00"

    print(f"\n{'='*50}")
    print(f"📅 NHẬP NGÀY {label}")
    print(f"{'='*50}")
    print(f"  Định dạng hỗ trợ:")
    print(f"    • Ngày đơn (1-31) → ngày trong tháng hiện tại")
    print(f"    • ThángNgày (0302) → 03/02/{datetime.now().year}")
    print(f"    • NgàyThángNăm (03022026) → 03/02/2026")
    print(f"    • NgàyThángNăm (03/02/2026) → 03/02/2026")
    print(f"\n  Giờ mặc định: {default_time}")
    print(f"{'-'*50}")

    while True:
        user_input = input(f"  ➤ Nhập ngày {label.lower()}: ").strip()

        if not user_input:
            print("  ⚠️  Vui lòng nhập ngày!")
            continue

        result = parse_flexible_date(user_input, is_checkout)

        if result:
            print(f"  ✅ Đã parse: {result}")
            return result
        else:
            print(f"  ⚠️  Không nhận diện được định dạng: '{user_input}'")
            print(f"     Vui lòng nhập lại (ví dụ: 15, 0302, 03022026, 03/02/2026)")


def ask_room_number() -> str:
    """Hỏi số phòng từ user."""
    print(f"\n{'='*50}")
    print(f"🏠 NHẬP SỐ PHÒNG")
    print(f"{'='*50}")

    while True:
        user_input = input(f"  ➤ Nhập số phòng: ").strip()

        if user_input:
            print(f"  ✅ Số phòng: {user_input}")
            return user_input
        else:
            print("  ⚠️  Vui lòng nhập số phòng!")


# ══════════════════════════════════════════════════════════════════════════════
#  CSV Header (19 cột)
# ══════════════════════════════════════════════════════════════════════════════

CSV_HEADERS = [
    "STT",
    "Họ và tên (*)",
    "Ngày, tháng, năm sinh (*)",
    "Giới tính (*)",
    "Quốc gia (*)",
    "Quốc tịch (*)",
    "Loại giấy tờ (*)",
    "Tên giấy tờ (*)",
    "Số giấy tờ (*)",
    "Số điện thoại",
    "Loại cư trú (*)",
    "Tỉnh/TP (*)",
    "Quận/Huyện (*)",
    "Phường/Xã/Đặc khu (*)",
    "Địa chỉ chi tiết (*)",
    "Thời gian lưu trú (từ ngày) (*)",
    "Thời gian lưu trú (đến ngày)",
    "Lý do lưu trú (*)",
    "Tên phòng/Khoa (*)",
]


# ══════════════════════════════════════════════════════════════════════════════
#  OCR.SPACE API
# ══════════════════════════════════════════════════════════════════════════════

def ocr_via_api(image_path: str) -> list[str]:
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    ext = Path(image_path).suffix.lower()
    mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                ".png": "image/png", ".webp": "image/webp",
                ".bmp": "image/bmp", ".tiff": "image/tiff"}
    mime = mime_map.get(ext, "image/jpeg")

    for attempt in OCR_ATTEMPTS:
        eng  = attempt["OCREngine"]
        lang = attempt["language"]
        print(f"📡 OCR.space Engine {eng}, language={lang} ...")

        payload = {
            "apikey":            OCR_API_KEY,
            "language":          lang,
            "OCREngine":         eng,
            "base64Image":       f"data:{mime};base64,{img_b64}",
            "isOverlayRequired": False,
            "detectOrientation": True,
            "scale":             True,
            "isTable":           False,
        }

        try:
            resp = requests.post(OCR_API_URL, data=payload, timeout=60)
            resp.raise_for_status()
            result = resp.json()
        except requests.exceptions.RequestException as e:
            print(f"   ⚠️  Kết nối lỗi: {e}")
            continue

        if result.get("IsErroredOnProcessing"):
            errs = result.get("ErrorMessage", ["Unknown error"])
            print(f"   ⚠️  API lỗi: {errs} → thử cách tiếp theo...")
            continue

        lines = []
        for page in result.get("ParsedResults", []):
            raw_text = page.get("ParsedText", "")
            for line in raw_text.splitlines():
                line = line.strip()
                if line:
                    lines.append(line)

        print(f"   ✅ Thành công! {len(lines)} dòng text.\n")

        print("─" * 55)
        print("🔎 RAW OCR RESPONSE:")
        for page_i, page in enumerate(result.get("ParsedResults", [])):
            print(f"\n  [Page {page_i + 1}]")
            raw_text = page.get("ParsedText", "")
            for line in raw_text.splitlines():
                if line.strip():
                    print(f"  │ {line}")
        print("─" * 55 + "\n")

        return lines

    print("❌ Tất cả phương thức OCR đều thất bại.")
    return []


# ══════════════════════════════════════════════════════════════════════════════
#  QR CODE
# ══════════════════════════════════════════════════════════════════════════════

def read_qr(img_orig: np.ndarray) -> dict | None:
    raw = _decode_qr(img_orig)

    if not raw:
        gray = cv2.cvtColor(img_orig, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 0, 255,
                                  cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        raw = _decode_qr(cv2.cvtColor(thresh, cv2.COLOR_GRAY2BGR))

    return _parse_qr(raw) if raw else None


def _decode_qr(img: np.ndarray) -> str | None:
    if QR_BACKEND == "pyzbar":
        try:
            for obj in pyzbar_decode(img):
                if obj.type in ("QRCODE", "QR"):
                    return obj.data.decode("utf-8", errors="ignore")
        except Exception:
            pass

    try:
        texts, _ = cv2.wechat_qrcode_WeChatQRCode().detectAndDecode(img)
        if texts:
            return texts[0]
    except Exception:
        pass

    data, _, _ = cv2.QRCodeDetector().detectAndDecode(img)
    return data if data else None


def _parse_qr(raw: str) -> dict:
    info: dict = {}
    parts = raw.split("|")
    if len(parts) >= 6:
        keys = ["_qr_so_cccd", "_qr_so_cmnd_cu", "ho_va_ten",
                "ngay_sinh", "gioi_tinh", "_qr_dia_chi", "_qr_ngay_cap"]
        for i, k in enumerate(keys):
            if i < len(parts) and parts[i].strip():
                info[k] = parts[i].strip()
        info["so_giay_to"]   = info.get("_qr_so_cccd", "")
        info["loai_giay_to"] = "CCCD"
        raw_dob = info.get("ngay_sinh", "")
        if raw_dob:
            m_dob = re.match(r'^(\d{2})(\d{2})(\d{4})$', raw_dob)
            if m_dob:
                d, mo, y = int(m_dob.group(1)), int(m_dob.group(2)), int(m_dob.group(3))
                if 1 <= d <= 31 and 1 <= mo <= 12 and 1900 <= y <= 2100:
                    info["ngay_sinh"] = f"{d:02d}/{mo:02d}/{y}"
    else:
        try:
            import json
            info.update(json.loads(raw))
        except Exception:
            pass
    return info


# ══════════════════════════════════════════════════════════════════════════════
#  MRZ Parser
# ══════════════════════════════════════════════════════════════════════════════

def _norm_mrz_line(raw: str) -> str:
    s = re.sub(r'[くく〈〈＜＜«»]', '<', raw)
    s = re.sub(r'[^ -~]', '', s)
    return s.strip()


def _mrz_name_parts(name_field: str) -> tuple:
    def _fmt(s):
        return re.sub(r'\s+', ' ', s.replace('<', ' ')).strip().title()

    parts = re.split(r'<{2,}', name_field)
    if len(parts) >= 2:
        lastname  = _fmt(parts[0])
        firstname = _fmt(' '.join(p for p in parts[1:] if p.strip('<')))
    else:
        lastname  = _fmt(name_field)
        firstname = ''
    return lastname, firstname


def _yymmdd_to_date(s: str) -> str:
    if not (s and len(s) == 6 and s.isdigit()):
        return ''
    yy, mm, dd = int(s[:2]), int(s[2:4]), int(s[4:6])
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return ''
    yyyy = 2000 + yy if yy <= 30 else 1900 + yy
    return f"{dd:02d}/{mm:02d}/{yyyy}"


def _is_mrz_candidate(norm: str) -> bool:
    if len(norm) < 20:
        return False
    if not re.match(r'^[A-Z0-9<]+$', norm):
        return False
    has_filler   = norm.count('<') >= 1
    pure_alphnum = (re.match(r'^[A-Z0-9]+$', norm)
                    and bool(re.search(r'\d{6}', norm[10:])))
    return has_filler or bool(pure_alphnum)


def parse_mrz(lines: list) -> dict:
    result = {
        'mrz_passport_no': '',
        'mrz_country':     '',
        'mrz_lastname':    '',
        'mrz_firstname':   '',
        'mrz_fullname':    '',
        'mrz_dob':         '',
        'mrz_gender':      '',
        'mrz_expiry':      '',
    }

    normed     = [_norm_mrz_line(l) for l in lines]
    candidates = [n for n in normed if _is_mrz_candidate(n)]

    if not candidates:
        return result

    line1 = next((c for c in candidates
                  if (c.startswith('P')
                      and len(c) > 5
                      and re.match(r'^P[A-Z<][A-Z]{3}', c)
                      and '<<' in c)), '')

    line2 = ''
    for c in candidates:
        if c == line1:
            continue
        ln = c.replace('<', '')
        if len(ln) >= 15 and re.search(r'\d{6}', ln[10:] if len(ln) > 10 else ln):
            line2 = c
            break
    if not line2:
        for c in candidates:
            if c != line1 and len(c.replace('<','')) >= 10:
                line2 = c
                break

    if line1:
        country = line1[2:5]
        if re.match(r'^[A-Z0-9]{3}$', country):
            result['mrz_country'] = country
        name_field = line1[5:]
        last, first = _mrz_name_parts(name_field)
        result['mrz_lastname']  = last
        result['mrz_firstname'] = first
        result['mrz_fullname']  = (last + ' ' + first).strip() if first else last
    else:
        for c in normed:
            if c == line2:
                continue
            if len(c) > 8 and c.startswith('P') and '<<' in c:
                country = c[2:5]
                if re.match(r'^[A-Z0-9]{3}$', country):
                    result['mrz_country'] = country
                last, first = _mrz_name_parts(c[5:])
                result['mrz_lastname']  = last
                result['mrz_firstname'] = first
                result['mrz_fullname']  = (last + ' ' + first).strip() if first else last
                break

    if line2:
        ln_raw = line2
        ln     = line2.replace('<', '')

        result['mrz_passport_no'] = ln[:9] if len(ln) >= 9 else ln

        found_dob = False

        for offset in range(3):
            pos = 9 + offset
            if pos + 9 > len(ln):
                break
            cc  = ln[pos:pos+3]
            dob = ln[pos+3:pos+9]
            if not re.match(r'^[A-Z0-9]{3}$', cc):
                continue
            dob_fmt = _yymmdd_to_date(dob)
            if not dob_fmt:
                continue
            if not result['mrz_country']:
                result['mrz_country'] = cc
            result['mrz_dob'] = dob_fmt
            g_pos   = pos + 10
            exp_pos = pos + 11
            if g_pos < len(ln):
                g = ln[g_pos]
                result['mrz_gender'] = 'Nam' if g == 'M' else ('Nữ' if g == 'F' else '')
            if exp_pos + 6 <= len(ln):
                result['mrz_expiry'] = _yymmdd_to_date(ln[exp_pos:exp_pos+6])
            found_dob = True
            break

        if not found_dob:
            m = re.search(r'<{2,}(\d{6})(\d)([MFX])', ln_raw)
            if m:
                dob_fmt = _yymmdd_to_date(m.group(1))
                if dob_fmt:
                    result['mrz_dob']    = dob_fmt
                    result['mrz_gender'] = 'Nam' if m.group(3) == 'M' else (
                                           'Nữ'  if m.group(3) == 'F' else '')
                    m_no = re.match(r'^([A-Z0-9]{7,9})', ln_raw)
                    if m_no and not result['mrz_passport_no']:
                        result['mrz_passport_no'] = m_no.group(1)
                    exp_raw = ln_raw[m.end():]
                    m_exp = re.search(r'(\d{6})', exp_raw)
                    if m_exp:
                        result['mrz_expiry'] = _yymmdd_to_date(m_exp.group(1))

    return result


# ══════════════════════════════════════════════════════════════════════════════
#  Date Parser
# ══════════════════════════════════════════════════════════════════════════════

MONTHS = {
    "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
    "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
    "1월":1,"2월":2,"3월":3,"4월":4,"5월":5,"6월":6,
    "7월":7,"8월":8,"9월":9,"10월":10,"11월":11,"12월":12,
    "tháng 1":1,"tháng 2":2,"tháng 3":3,"tháng 4":4,
    "tháng 5":5,"tháng 6":6,"tháng 7":7,"tháng 8":8,
    "tháng 9":9,"tháng 10":10,"tháng 11":11,"tháng 12":12,
}

def parse_date(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip()

    cleaned = re.sub(r'\d+월\s*/\s*', '', raw).strip()
    cleaned = re.sub(r'\s+', ' ', cleaned)

    def _try(s: str) -> str:
        s = s.strip()
        m = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$', s)
        if m:
            return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
        m = re.match(r'^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$', s)
        if m:
            return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}"
        m = re.match(r'^(\d{2})(\d{2})(\d{4})$', s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= d <= 31 and 1 <= mo <= 12 and 1900 <= y <= 2100:
                return f"{d:02d}/{mo:02d}/{y}"
        m = re.match(r'^(\d{4})(\d{2})(\d{2})$', s)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= d <= 31 and 1 <= mo <= 12 and 1900 <= y <= 2100:
                return f"{d:02d}/{mo:02d}/{y}"
        m = re.match(r'^(\d{1,2})\s+([A-Za-z월]+)\s+(\d{2,4})$', s)
        if m:
            mon = m.group(2).lower()
            mo  = MONTHS.get(mon[:3]) or MONTHS.get(mon)
            if mo:
                y = int(m.group(3))
                if y < 100: y += 2000 if y < 30 else 1900
                return f"{int(m.group(1)):02d}/{mo:02d}/{y}"
        m = re.match(r'^([A-Za-z]+)\s+(\d{1,2})[,\s]+(\d{2,4})$', s)
        if m:
            mo = MONTHS.get(m.group(1).lower()[:3])
            if mo:
                y = int(m.group(3))
                if y < 100: y += 2000 if y < 30 else 1900
                return f"{int(m.group(2)):02d}/{mo:02d}/{y}"
        m = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$', s)
        if m:
            y = int(m.group(3))
            y += 2000 if y < 30 else 1900
            return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{y}"
        return ""

    return _try(cleaned) or _try(raw) or raw


# ══════════════════════════════════════════════════════════════════════════════
#  Parser
# ══════════════════════════════════════════════════════════════════════════════

def parse_doc(lines: list[str]) -> dict:
    text = "\n".join(lines)
    info: dict = {}

    _NV_LABEL = re.compile(
        r'^(date\s+of\s+birth|date\s+of\s+issue|date\s+of\s+expir|date\s+de'
        r'|lahir|tgl\.?\s*lahir|tgl\s+lahir|ngày\s*sinh|birth|expir'
        r'|passport\s*no\.?|passportno\.?|no\.?\s*passeport|no\.?\s*paspor'
        r'|full\s*name|nama\s*lengkap|surname|forename|given\s*name'
        r'|nationality|kewarganegaraan|country\s*code|kode\s*negara'
        r'|sex\s*/|kelamin|sexe|gneas)',
        re.I
    )

    def _is_label(v: str) -> bool:
        return bool(_NV_LABEL.match(v.strip()))

    def nv(kw: str, n: int = 1) -> str | None:
        for i, line in enumerate(lines):
            if not re.search(re.escape(kw), line, re.I):
                continue
            if ":" in line:
                v = line.split(":", 1)[-1].strip()
                if v and not _is_label(v):
                    return v
            for j in range(1, n + 1):
                if i + j >= len(lines):
                    break
                nxt = lines[i + j].strip()
                if nxt and not _is_label(nxt):
                    return nxt
        return None

    # Loại giấy tờ
    if re.search(r"hộ\s*chiếu|passport", text, re.I):
        info["loai_giay_to"] = "Hộ chiếu"
    elif re.search(r"căn\s*cước công dân|cccd", text, re.I):
        info["loai_giay_to"] = "CCCD"
    elif re.search(r"chứng minh nhân dân|cmnd", text, re.I):
        info["loai_giay_to"] = "CMND"
    else:
        info["loai_giay_to"] = ""
    info["ten_giay_to"] = ""

    # MRZ
    mrz = parse_mrz(lines)

    # Số giấy tờ
    HC_RE = re.compile(r'\b([A-Z]{1,2}[A-Z0-9]{6,8})\b')

    def _extract_doc_no(s: str) -> str:
        m = HC_RE.search(s)
        return m.group(1) if m else s.strip()

    so_gt_raw = (
        nv('no. paspor') or nv('no.paspor') or nv('no.paspor/')
        or nv('passport no.') or nv('passport no') or nv('passport number')
        or nv('여권번호')
        or nv('pas uti')
        or nv('số/ no') or nv('số/no')
        or ""
    )
    _SO_NOISE = re.compile(
        r'^(passportno\.?|passport\s*no\.?|no\.?\s*passeport'
        r'|no\.?\s*paspor|passportno\s*/|passeport\s*no'
        r'|r/passport\s*no|no,?\s*passeport)$', re.I)
    if so_gt_raw and _SO_NOISE.match(so_gt_raw.strip()):
        so_gt_raw = ""

    if so_gt_raw:
        info['so_giay_to'] = _extract_doc_no(so_gt_raw)
    else:
        m12  = re.search(r'\b(\d{12})\b', text)
        m_hc = HC_RE.search(text)
        m9   = re.search(r'\b(\d{9})\b', text)
        if m12:    info['so_giay_to'] = m12.group(1)
        elif m_hc: info['so_giay_to'] = m_hc.group(1)
        elif m9:   info['so_giay_to'] = m9.group(1)
        else:      info['so_giay_to'] = ''

    # Họ và tên
    is_passport = info.get('loai_giay_to') == 'Hộ chiếu'

    ho_ten = ''
    if is_passport and mrz.get('mrz_fullname'):
        ho_ten = mrz['mrz_fullname']
    else:
        ho_ten = (nv('họ và tên') or nv('full name')
                  or nv('nama lengkap | full') or nv('nama lengkap')
                  or "")
        if not ho_ten:
            surname   = (nv('surname') or nv('suname') or nv('성')
                         or nv('sloinne')
                         or nv('last name') or nv('family name') or nv('nom') or '')
            givenname = (nv('given names') or nv('given name') or nv('이름')
                         or nv('이틀') or nv('first name')
                         or nv('tusainm')
                         or nv('forename') or nv('prenom') or '')
            def _clean_latin(s: str) -> str:
                return re.sub(r'[^\x20-\x7EÀ-ɏ]', '', s).strip()
            surname   = _clean_latin(surname)
            givenname = _clean_latin(givenname)
            ho_ten    = ' '.join(p for p in [surname, givenname] if p)

    ho_ten = ho_ten.strip()
    if ho_ten and ho_ten == ho_ten.upper():
        ho_ten = ho_ten.title()
    info['ho_va_ten'] = ho_ten

    # Ngày sinh
    raw_date = (
        nv("ngày.*sinh") or nv("date of birth") or nv("date ol birth")
        or nv("청년들일") or nv("생년월일")
        or nv("tgl: lahir") or nv("tgl. lahir") or nv("tgl lahir")
        or nv("date de naissance") or nv("data de nascimento")
        or nv("date of birth/date") or nv("naissance")
        or ""
    )

    if not raw_date:
        for pat in [
            r'\b(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4})\b',
            r'\b(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})\b',
            r'\b(\d{1,2}\s+\d+월\s*/\s*[A-Za-z]+\s+\d{4})\b',
            r'\b(\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{2,4})\b',
            r'\b((?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2}[,\s]+\d{2,4})\b',
        ]:
            found = re.findall(pat, text, re.I)
            if found:
                raw_date = found[0]
                break

    if is_passport and mrz.get('mrz_dob'):
        raw_date = mrz['mrz_dob']
    elif not raw_date:
        raw_date = mrz.get('mrz_dob', '')

    info['ngay_sinh'] = parse_date(raw_date) or mrz.get('mrz_dob', '')

    # Giới tính
    gt_raw = (nv("giới tính") or nv("sex") or nv("성별") or nv("gender") or "").strip()

    def _parse_gender(s: str) -> str:
        s = s.strip().upper()
        if s in ("M", "MALE", "NAM", "男", "남"):          return "Nam"
        if s in ("F", "FEMALE", "NỮ", "NU", "女", "여"):    return "Nữ"
        if s in ("X", "D", "NB"):                            return s
        if s in ("W", "WEIBLICH"):                           return "Nữ"
        if s in ("MÄNNLICH", "MANNLICH"):                    return "Nam"
        if s in ("H", "HOMME", "MASCULIN", "MASCULINO"):    return "Nam"
        if s in ("FEMME", "FÉMININ", "FEMENINO", "FEMENINO"): return "Nữ"
        if s in ("ذكر", "ZAKR"):   return "Nam"
        if s in ("أنثى", "UNTHA"): return "Nữ"
        if re.search(r'\bnam\b|\bmale\b|\bm\b|남성', s, re.I):   return "Nam"
        if re.search(r'\bnữ\b|\bnu\b|\bfemale\b|\bf\b|여성', s, re.I): return "Nữ"
        return ""

    gt = _parse_gender(gt_raw)

    if not gt:
        for line in lines:
            ll = line.lower()
            if re.search(r'giới\s*tính|\bsex\b|성별|gender', ll):
                gt = _parse_gender(re.sub(r'.*(giới\s*tính|sex|성별|gender)\s*[:/]?\s*', '', line, flags=re.I).strip())
                if gt:
                    break

    if not gt:
        if re.search(r'\bnữ\b|\bfemale\b|\bféminin\b|\b(f)\b(?!ull|irst|am)', text, re.I):
            gt = "Nữ"
        elif re.search(r'\bmale\b|\bmasculin\b|\b(m)\b(?!r|s|in|ax|ay)', text, re.I):
            gt = "Nam"

    if is_passport and mrz.get('mrz_gender') and not gt:
        gt = mrz['mrz_gender']

    info['gioi_tinh'] = gt

    # Quốc tịch & Quốc gia
    if info.get("loai_giay_to") in ("CCCD", "CMND"):
        info["quoc_tich"] = "VNM"
        info["quoc_gia"]  = "VNM"
    else:
        qt_raw = (nv("quốc tịch") or nv("nationality") or nv("육해")
                  or nv("국가코드") or nv("country code")
                  or nv("kewarganegaraan") or nv("nationalite")
                  or nv("intacht/natio") or nv("intacht")
                  or mrz.get('mrz_country') or "")

        if not qt_raw:
            codes = re.findall(r'\b([A-Z]{3})\b', text)
            for c in codes:
                if c.lower() in _COUNTRY_MAP:
                    qt_raw = c
                    break

        if not qt_raw:
            if re.search(r"việt nam|viet nam", text, re.I):
                qt_raw = "Viet Nam"
            else:
                qt_raw = text

        code = match_country(qt_raw)
        if not code and qt_raw.strip():
            code = qt_raw.strip().upper()[:3]

        info["quoc_tich"] = code
        info["quoc_gia"]  = code

    # Địa chỉ
    ADDR_HARD_STOP = ["họ và tên", "ngày sinh", "giới tính", "quốc tịch",
                      "quê quán", "có giá trị", "đặc điểm", "date of birth",
                      "full name", "nationality", "place of origin"]
    ADDR_SKIP = ["date of expiry", "date of issue", "independence", "freedom",
                 "socialist republic", "citizen identity", "republic of viet"]
    addr_lines = []
    in_addr = False
    for line in lines:
        ll = line.lower()
        if "thường trú" in ll or "place of residence" in ll:
            in_addr = True
            if ":" in line:
                after = line.split(":", 1)[1].strip()
                if after:
                    addr_lines.append(after)
            continue
        if in_addr:
            if any(kw in ll for kw in ADDR_HARD_STOP):
                break
            if any(kw in ll for kw in ADDR_SKIP):
                continue
            if line.strip():
                addr_lines.append(line.strip())
            if len(addr_lines) >= 3:
                break

    addr_full = ""
    if addr_lines:
        combined = []
        for ln in addr_lines:
            if combined and not combined[-1].endswith(','):
                combined.append(',')
            combined.append(ln)
        addr_full = ' '.join(combined).replace(' ,', ',').strip().strip(',')

    if not addr_full:
        addr_full = nv("thường trú", 2) or nv("place of residence", 2) or ""

    addr_parsed = parse_address_segment(addr_full)
    info["tinh_tp"]          = addr_parsed["tinh_tp"]
    info["quan_huyen"]       = addr_parsed["quan_huyen"]
    info["phuong_xa"]        = addr_parsed["phuong_xa"]
    info["dia_chi_chi_tiet"] = addr_parsed["dia_chi_chi_tiet"]

    info.setdefault("so_dien_thoai",  "")
    info.setdefault("loai_cu_tru",    "Tạm trú")
    info.setdefault("tu_ngay",        "")
    info.setdefault("den_ngay",       "")
    info.setdefault("ly_do_luu_tru",  "")
    info.setdefault("ten_phong_khoa", "")

    info['_mrz'] = mrz

    return info


def merge_qr_ocr(qr: dict | None, ocr: dict) -> dict:
    return merge_all(qr, ocr)


def merge_all(qr: dict | None, ocr: dict) -> dict:
    mrz    = ocr.pop('_mrz', {})
    merged = dict(ocr)
    sources: dict[str, str] = {}
    is_passport = ocr.get('loai_giay_to') == 'Hộ chiếu'

    def _pick(key: str, qr_val: str, mrz_val: str, ocr_val: str) -> tuple[str, str]:
        if qr_val:   return qr_val,  'QR'
        if is_passport:
            if key == 'ho_va_ten':
                if mrz_val and _mrz_name_ok(mrz_val): return mrz_val, 'MRZ'
                if ocr_val: return ocr_val, 'OCR'
            elif key == 'so_giay_to':
                if ocr_val and mrz_val:
                    if mrz_val.startswith(ocr_val):
                        return ocr_val, 'OCR'
                    return mrz_val, 'MRZ'
                if mrz_val: return mrz_val, 'MRZ'
                if ocr_val: return ocr_val, 'OCR'
            else:
                if mrz_val: return mrz_val, 'MRZ'
                if ocr_val: return ocr_val, 'OCR'
        else:
            if ocr_val: return ocr_val, 'OCR'
            if mrz_val: return mrz_val, 'MRZ'
        return '', ''

    def _mrz_name_ok(name: str) -> bool:
        if not name:
            return False
        alpha = sum(1 for c in name if c.isalpha() or c == ' ')
        return alpha / len(name) >= 0.7

    FIELD_MAP = {
        'ho_va_ten':  ('ho_va_ten', 'mrz_fullname',    'ho_va_ten'),
        'ngay_sinh':  ('ngay_sinh', 'mrz_dob',         'ngay_sinh'),
        'gioi_tinh':  ('gioi_tinh', 'mrz_gender',      'gioi_tinh'),
        'so_giay_to': ('so_giay_to','mrz_passport_no', 'so_giay_to'),
        'quoc_tich':  ('',          'mrz_country',     'quoc_tich'),
        'quoc_gia':   ('',          'mrz_country',     'quoc_gia'),
    }

    for out_key, (qk, mk, ok) in FIELD_MAP.items():
        val, src = _pick(
            out_key,
            qr.get(qk, '')  if qr else '',
            mrz.get(mk, ''),
            ocr.get(ok, ''),
        )
        if val:
            if out_key in ('ngay_sinh',) and val:
                val = parse_date(val) or val
            if out_key in ('quoc_tich', 'quoc_gia') and val:
                if not re.match(r'^[A-Z]{2,3}$', val.upper().strip()):
                    val = ''
                else:
                    val = val.upper().strip()
            if val:
                merged[out_key] = val
                sources[out_key] = src

    if qr:
        for k in ['loai_giay_to', 'tinh_tp', 'quan_huyen', 'phuong_xa',
                  'dia_chi_chi_tiet', 'so_dien_thoai']:
            if qr.get(k) and not merged.get(k):
                merged[k] = qr[k]
                sources[k] = 'QR'

    merged['_sources'] = sources
    return merged


# ══════════════════════════════════════════════════════════════════════════════
#  CSV
# ══════════════════════════════════════════════════════════════════════════════

def _next_stt(csv_path: str) -> int:
    p = Path(csv_path)
    if not p.exists():
        return 1
    with open(p, "r", encoding="utf-8-sig") as f:
        n = sum(1 for _ in csv.reader(f))
    return max(1, n)


def append_csv(data: dict, csv_path: str):
    p = Path(csv_path)
    need_header = not p.exists()
    stt = _next_stt(csv_path)

    row = [
        stt,
        data.get("ho_va_ten", ""),
        data.get("ngay_sinh", ""),
        data.get("gioi_tinh", ""),
        data.get("quoc_gia", ""),
        data.get("quoc_tich", ""),
        data.get("loai_giay_to", ""),
        data.get("ten_giay_to", ""),
        data.get("so_giay_to", ""),
        data.get("so_dien_thoai", ""),
        data.get("loai_cu_tru", ""),
        data.get("tinh_tp", ""),
        data.get("quan_huyen", ""),
        data.get("phuong_xa", ""),
        data.get("dia_chi_chi_tiet", ""),
        data.get("tu_ngay", ""),
        data.get("den_ngay", ""),
        data.get("ly_do_luu_tru", ""),
        data.get("ten_phong_khoa", ""),
    ]

    with open(p, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if need_header:
            writer.writerow(CSV_HEADERS)
        writer.writerow(row)

    tag = "✅ Tạo mới" if need_header else "✅ Cập nhật"
    print(f"\n{tag} → {p.resolve()}  (STT {stt})")


# ══════════════════════════════════════════════════════════════════════════════
#  Export Functions (XML + Excel)
# ══════════════════════════════════════════════════════════════════════════════

def _safe_text(val: str) -> str:
    """Loại bỏ ký tự điều khiển XML không hợp lệ."""
    return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', val or '').strip()


def _strip_time(val: str) -> str:
    """Loại bỏ phần giờ khỏi chuỗi ngày."""
    if not val:
        return ""
    parts = val.strip().split(' ')
    if len(parts) >= 2 and ':' in parts[1]:
        return parts[0]
    return val.strip()


def _gender_to_xml(val: str) -> str:
    v = val.strip().upper()
    if v in ("NAM", "M", "MALE", "1"):
        return "M"
    if v in ("NỮ", "NU", "F", "FEMALE", "0"):
        return "F"
    return v or "M"


def _is_passport(row: dict) -> bool:
    loai = row.get("Loại giấy tờ (*)", "").strip().lower()
    return "chiếu" in loai or "passport" in loai or "hộ chiếu" in loai


def _is_vietnam(row: dict) -> bool:
    qt   = row.get("Quốc tịch (*)", "").strip().upper()
    loai = row.get("Loại giấy tờ (*)",   "").strip().upper()
    is_vn   = qt in ("VNM", "VN", "VIETNAM", "VIỆT NAM", "VIET NAM")
    is_cccd = any(k in loai for k in ("CCCD", "CMND", "CĂN CƯỚC", "CHỨNG MINH"))
    return is_vn and is_cccd


def _pretty_xml(root: ET.Element) -> str:
    raw = ET.tostring(root, encoding="unicode")
    dom = minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="    ", encoding=None)
    lines = pretty.split("\n")
    if lines[0].startswith("<?xml"):
        lines = lines[1:]
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + "\n".join(lines)


def export_xml(rows: list[dict], out_path: str) -> int:
    """Xuất XML cho khách Hộ chiếu."""
    passport_rows = [r for r in rows if _is_passport(r)]

    if not passport_rows:
        print("ℹ️  Không có khách Hộ chiếu trong CSV → bỏ qua pp.xml")
        return 0

    root = ET.Element("KHAI_BAO_TAM_TRU")

    for idx, row in enumerate(passport_rows, start=1):
        guest = ET.SubElement(root, "THONG_TIN_KHACH")

        def sub(tag: str, val: str):
            el = ET.SubElement(guest, tag)
            el.text = _safe_text(val) if val.strip() else None

        sub("so_thu_tu",         str(idx))
        sub("ho_ten",            row.get("Họ và tên (*)",    "").upper())
        sub("ngay_sinh",         row.get("Ngày, tháng, năm sinh (*)", ""))
        sub("ngay_sinh_dung_den","D")
        sub("gioi_tinh",         _gender_to_xml(row.get("Giới tính (*)", "")))
        sub("ma_quoc_tich",      row.get("Quốc tịch (*)", "").upper())
        sub("so_ho_chieu",       row.get("Số giấy tờ (*)",     ""))
        sub("so_phong",          row.get("Tên phòng/Khoa (*)",     ""))
        # XML không có giờ
        sub("ngay_den",          _strip_time(row.get("Thời gian lưu trú (từ ngày) (*)",   "")))
        sub("ngay_di_du_kien",   _strip_time(row.get("Thời gian lưu trú (đến ngày)",  "")))
        sub("ngay_tra_phong",    "")

    xml_str = _pretty_xml(root)
    Path(out_path).write_text(xml_str, encoding="utf-8")

    print(f"✅ pp.xml  → {out_path}  ({len(passport_rows)} khách hộ chiếu)")
    return len(passport_rows)


def export_excel(rows: list[dict], out_path: str) -> int:
    """Xuất Excel cho khách Việt Nam (CCCD/CMND)."""
    vn_rows = [r for r in rows if _is_vietnam(r)]

    if not vn_rows:
        print("ℹ️  Không có khách Việt Nam trong CSV → bỏ qua Excel")
        return 0

    # Header cho Excel (19 cột chính + 2 cột phụ)
    excel_headers = [
        "STT", "Họ và tên (*)", "Ngày, tháng, năm sinh (*)", "Giới tính (*)",
        "Quốc gia (*)", "Quốc tịch (*)", "Loại giấy tờ (*)", "Tên giấy tờ (*)",
        "Số giấy tờ (*)", "Số điện thoại", "Loại cư trú (*)", "Tỉnh/TP (*)",
        "Quận/Huyện_cũ (*)", "Phường/Xã/ Đặc khu (*)", "Địa chỉ chi tiết (*)",
        "Thời gian lưu trú \n(từ ngày) (*)", "Thời gian lưu trú \n(đến ngày)",
        "Lý do lưu trú (*)", "Tên phòng / Khoa (*)"
    ]

    data_rows = []
    for idx, row in enumerate(vn_rows, start=1):
        data_rows.append([
            idx,
            row.get("Họ và tên (*)",     ""),
            row.get("Ngày, tháng, năm sinh (*)",  ""),
            row.get("Giới tính (*)",  ""),
            row.get("Quốc gia (*)",   ""),
            row.get("Quốc tịch (*)", ""),
            row.get("Loại giấy tờ (*)",    ""),
            row.get("Tên giấy tờ (*)", ""),
            row.get("Số giấy tờ (*)",      ""),
            row.get("Số điện thoại",        ""),
            row.get("Loại cư trú (*)",""),
            row.get("Tỉnh/TP (*)",    ""),
            row.get("Quận/Huyện (*)", ""),
            row.get("Phường/Xã/Đặc khu (*)", ""),
            row.get("Địa chỉ chi tiết (*)", ""),
            row.get("Thời gian lưu trú (từ ngày) (*)",    ""),
            row.get("Thời gian lưu trú (đến ngày)",   ""),
            row.get("Lý do lưu trú (*)",      ""),
            row.get("Tên phòng/Khoa (*)", ""),
        ])

    df = pd.DataFrame(data_rows, columns=excel_headers)

    # Ghi Excel
    df.to_excel(out_path, index=False, engine='openpyxl')

    print(f"✅ Excel → {out_path}  ({len(vn_rows)} khách Việt Nam)")
    return len(vn_rows)


def read_csv(path: str) -> list[dict]:
    """Đọc output.csv."""
    p = Path(path)
    if not p.exists():
        print(f"❌ Không tìm thấy: {path}")
        return []

    rows = []
    with open(p, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(dict(row))

    print(f"📂 Đọc {len(rows)} dòng từ {p.name}")
    return rows


def export_all(csv_path: str = OUTPUT_CSV):
    """Xuất tất cả file (XML + Excel) từ CSV."""
    rows = read_csv(csv_path)

    if not rows:
        print("⚠️  CSV rỗng.")
        return

    # Đếm thống kê
    pp_count = sum(1 for r in rows if _is_passport(r))
    vn_count = sum(1 for r in rows if _is_vietnam(r))

    from collections import Counter
    qt_count = Counter(
        r.get("Quốc tịch (*)", "?").upper()
        for r in rows if _is_passport(r)
    )

    print()
    print("─" * 45)
    print(f"  Tổng cộng              : {len(rows):>4} người")
    print(f"  ├─ Hộ chiếu (→ pp.xml) : {pp_count:>4} người")
    if qt_count:
        for qt, cnt in qt_count.most_common(10):
            print(f"  │    {qt:<6}            : {cnt:>4}")
    print(f"  ├─ CCCD/CMND VN (→ xlsx): {vn_count:>4} người")
    print("─" * 45)
    print()

    # Xuất XML (passport)
    xml_count = export_xml(rows, "pp.xml")

    # Xuất Excel (Việt Nam)
    excel_count = export_excel(rows, "vietnam.xlsx")

    print()
    if xml_count == 0 and excel_count == 0:
        print("⚠️  Không có dữ liệu nào được xuất.")
    else:
        print("✅ Hoàn thành xuất file!")


# ══════════════════════════════════════════════════════════════════════════════
#  Main
# ══════════════════════════════════════════════════════════════════════════════

IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def parse_args():
    args = {
        "image":      None,
        "dir":        None,
        "output":     OUTPUT_CSV,
        "checkin":    "",
        "checkout":   "",
        "phong":      "",
        "lydo":       "Du Lịch",
        "export_only": False,
    }
    argv = sys.argv[1:]
    i = 0
    while i < len(argv):
        a = argv[i]
        if   a in ("-d", "--dir")    and i + 1 < len(argv): args["dir"]      = argv[i+1]; i += 2
        elif a == "--output"         and i + 1 < len(argv): args["output"]   = argv[i+1]; i += 2
        elif a == "--checkin"        and i + 1 < len(argv): args["checkin"]  = argv[i+1]; i += 2
        elif a == "--checkout"       and i + 1 < len(argv): args["checkout"] = argv[i+1]; i += 2
        elif a == "--phong"          and i + 1 < len(argv): args["phong"]    = argv[i+1]; i += 2
        elif a == "--lydo"           and i + 1 < len(argv): args["lydo"]     = argv[i+1]; i += 2
        elif a == "--export-only"                          : args["export_only"] = True; i += 1
        elif not a.startswith("-"):
            args["image"] = a; i += 1
        else:
            i += 1
    return args


LABEL_MAP = {
    "ho_va_ten":        "Họ và tên",
    "ngay_sinh":        "Ngày sinh",
    "gioi_tinh":        "Giới tính",
    "quoc_gia":         "Quốc gia",
    "quoc_tich":        "Quốc tịch",
    "loai_giay_to":     "Loại giấy tờ",
    "ten_giay_to":      "Tên giấy tờ",
    "so_giay_to":       "Số giấy tờ",
    "so_dien_thoai":    "Số điện thoại",
    "loai_cu_tru":      "Loại cư trú",
    "tinh_tp":          "Tỉnh/TP",
    "quan_huyen":       "Quận/Huyện",
    "phuong_xa":        "Phường/Xã",
    "dia_chi_chi_tiet": "Địa chỉ chi tiết",
    "tu_ngay":          "Từ ngày",
    "den_ngay":         "Đến ngày",
    "ly_do_luu_tru":    "Lý do lưu trú",
    "ten_phong_khoa":   "Tên phòng/Khoa",
}


def process_one(image_path: str, args: dict) -> bool:
    """Xử lý 1 file ảnh."""
    p = Path(image_path)
    if not p.exists():
        print(f"⚠️  Không tìm thấy: {image_path}")
        return False

    img_orig = cv2.imread(image_path)
    if img_orig is None:
        print(f"⚠️  Không đọc được ảnh: {p.name}")
        return False

    h, w = img_orig.shape[:2]
    print(f"\n{'═'*55}")
    print(f"📄 {p.name}  ({w}×{h}px)")
    print(f"{'═'*55}")

    # 1. QR
    print(f"⬛ Quét QR... (backend: {QR_BACKEND})")
    qr_info = read_qr(img_orig)
    print("✅ Tìm thấy QR!" if qr_info else "⚪ Không có QR.")

    # 2. OCR
    lines = ocr_via_api(image_path)
    if not lines:
        print("⚠️  OCR không trả về kết quả.")
        return False

    print("\n─ TEXT OCR RAW " + "─"*40)
    for line in lines:
        print(f"  {line}")

    # 3. Parse + merge
    ocr_info = parse_doc(lines)
    final    = merge_qr_ocr(qr_info, ocr_info)

    # Xử lý ngày check-in
    if args["checkin"]:
        checkin_date = args["checkin"].strip()
    else:
        checkin_date = ask_checkin_date(is_checkout=False)

    if " " in checkin_date:
        final["tu_ngay"] = checkin_date
    else:
        final["tu_ngay"] = f"{checkin_date} 15:00:00"

    # Xử lý ngày check-out
    if args["checkout"]:
        checkout_date = args["checkout"].strip()
    else:
        checkout_date = ask_checkin_date(is_checkout=True)

    if " " in checkout_date:
        final["den_ngay"] = checkout_date
    else:
        final["den_ngay"] = f"{checkout_date} 11:00:00"

    # Xử lý số phòng
    if args["phong"]:
        final["ten_phong_khoa"] = args["phong"]
    else:
        final["ten_phong_khoa"] = ask_room_number()

    # Xử lý lý do lưu trú
    if args["lydo"]:
        final["ly_do_luu_tru"] = args["lydo"]

    # 4. In kết quả
    sources = final.pop('_sources', {})
    print("\n─ KẾT QUẢ " + "─"*45)
    for k, lbl in LABEL_MAP.items():
        v   = final.get(k, "")
        src = sources.get(k, "")
        src_tag = f" [{src}]" if src else ""
        print(f"  {'✅' if v else '⚪'} {lbl:<28}: {v}{src_tag}")

    if final.get('loai_giay_to') == 'Hộ chiếu':
        mrz_raw = {k: v for k, v in (final.get('_mrz') or {}).items() if v}
        if mrz_raw:
            print("\n  ── MRZ raw ──")
            for k, v in mrz_raw.items():
                print(f"     {k:<22}: {v}")

    # 5. Ghi CSV
    append_csv(final, args["output"])
    return True


def main():
    args = parse_args()

    # Chế độ chỉ xuất file (đã có output.csv)
    if args["export_only"]:
        print("=" * 45)
        print("  XUẤT FILE (XML + Excel)")
        print("=" * 45)
        export_all(args["output"])
        return

    if not args["image"] and not args["dir"]:
        print("Cách dùng:")
        print("  python doc_cccd_local.py <ảnh>          # 1 file")
        print("  python doc_cccd_local.py -d <thư_mục>   # toàn bộ ảnh trong thư mục")
        print("  python doc_cccd_local.py --export-only  # chỉ xuất file từ output.csv")
        print()
        print("Tuỳ chọn:")
        print("  --output   <file.csv>")
        print("  --checkin  <dd/mm/yyyy> (hoặc dd/mm/yyyy hh:mm:ss)")
        print("  --checkout <dd/mm/yyyy> (hoặc dd/mm/yyyy hh:mm:ss)")
        print("  --phong    <tên phòng>")
        print("  --lydo     <lý do lưu trú>")
        print()
        print("Nếu không nhập --checkin/--checkout/--phong, chương trình sẽ hỏi tương tác.")
        print("Định dạng ngày linh hoạt:")
        print("  • Ngày đơn (1-31) → ngày trong tháng hiện tại")
        print("  • ThángNgày (0302) → 03/02/năm hiện tại")
        print("  • NgàyThángNăm (03022026) → 03/02/2026")
        print("  • NgàyThángNăm (03/02/2026) → 03/02/2026")
        sys.exit(1)

    # Xây dựng danh sách file
    if args["dir"]:
        d = Path(args["dir"])
        if not d.is_dir():
            print(f"Lỗi: '{args['dir']}' không phải thư mục.")
            sys.exit(1)
        image_files = sorted(
            f for f in d.iterdir()
            if f.is_file() and f.suffix.lower() in IMG_EXTS
        )
        if not image_files:
            print(f"⚠️  Không tìm thấy ảnh nào trong '{args['dir']}'")
            sys.exit(1)
        print(f"📁 Thư mục: {d}  →  {len(image_files)} ảnh tìm thấy")
        print(f"📊 CSV output: {args['output']}\n")
    else:
        image_files = [Path(args["image"])]

    # Xử lý từng file
    ok_count = 0
    for img_path in image_files:
        if process_one(str(img_path), args):
            ok_count += 1

    # Tổng kết
    if len(image_files) > 1:
        print(f"\n{'═'*55}")
        print(f"✅ Hoàn thành: {ok_count}/{len(image_files)} ảnh")
        print(f"📊 Đã ghi vào: {args['output']}")

        # Hỏi có xuất file không
        print(f"\n{'─'*50}")
        export_choice = input("  ➤ Xuất file XML + Excel? (Y/n): ").strip().lower()
        if export_choice in ("", "y", "yes", "c", "có"):
            export_all(args["output"])
    else:
        # Với 1 file, tự động xuất
        print(f"\n{'─'*50}")
        export_choice = input("  ➤ Xuất file XML + Excel? (Y/n): ").strip().lower()
        if export_choice in ("", "y", "yes", "c", "có"):
            export_all(args["output"])


if __name__ == "__main__":
    main()