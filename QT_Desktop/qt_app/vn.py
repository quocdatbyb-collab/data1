"""Tiện ích tiếng Việt: chuyển mã TCVN3, đọc số thành chữ, xử lý ngày và số."""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata

# ---------------------------------------------------------------------------
# TCVN3 (ABC) -> Unicode
# Dữ liệu TCVN3 khi giải mã theo latin-1 cho ra các ký tự như "C¸ nh©n".
# ---------------------------------------------------------------------------
_TCVN3_PAIRS = {
    0xA1: "Ă", 0xA2: "Â", 0xA3: "Ê", 0xA4: "Ô", 0xA5: "Ơ", 0xA6: "Ư", 0xA7: "Đ",
    0xA8: "ă", 0xA9: "â", 0xAA: "ê", 0xAB: "ô", 0xAC: "ơ", 0xAD: "ư", 0xAE: "đ",
    0xB5: "à", 0xB6: "ả", 0xB7: "ã", 0xB8: "á", 0xB9: "ạ",
    0xBB: "ằ", 0xBC: "ẳ", 0xBD: "ẵ", 0xBE: "ắ", 0xC6: "ặ",
    0xC7: "ầ", 0xC8: "ẩ", 0xC9: "ẫ", 0xCA: "ấ", 0xCB: "ậ",
    0xCC: "è", 0xCE: "ẻ", 0xCF: "ẽ", 0xD0: "é", 0xD1: "ẹ",
    0xD2: "ề", 0xD3: "ể", 0xD4: "ễ", 0xD5: "ế", 0xD6: "ệ",
    0xD7: "ì", 0xD8: "ỉ", 0xDC: "ĩ", 0xDD: "í", 0xDE: "ị",
    0xDF: "ò", 0xE1: "ỏ", 0xE2: "õ", 0xE3: "ó", 0xE4: "ọ",
    0xE5: "ồ", 0xE6: "ổ", 0xE7: "ỗ", 0xE8: "ố", 0xE9: "ộ",
    0xEA: "ờ", 0xEB: "ở", 0xEC: "ỡ", 0xED: "ớ", 0xEE: "ợ",
    0xEF: "ù", 0xF1: "ủ", 0xF2: "ũ", 0xF3: "ú", 0xF4: "ụ",
    0xF5: "ừ", 0xF6: "ử", 0xF7: "ữ", 0xF8: "ứ", 0xF9: "ự",
    0xFA: "ỳ", 0xFB: "ỷ", 0xFC: "ỹ", 0xFD: "ý", 0xFE: "ỵ",
}
TCVN3_TABLE = {code: ch for code, ch in _TCVN3_PAIRS.items()}
# Ký tự chỉ xuất hiện trong văn bản TCVN3 (không có trong tiếng Việt Unicode)
_TCVN3_MARKERS = set("¸µ¶·¹¨©ª«¬®§»¼½¾ÆÇÈÉËÌÎÏÐÑÒÓÔÕÖ×ØÜÞß")
_VN_UNICODE_RE = re.compile("[ĂăĐđƠơƯưẠ-ỹ]")


def looks_like_tcvn3(text: str) -> bool:
    """Đoán văn bản có phải TCVN3 (đã giải mã latin-1) hay không."""
    sample = text[:200000]
    if _VN_UNICODE_RE.search(sample):
        return False
    hits = sum(1 for ch in sample if ch in _TCVN3_MARKERS)
    return hits >= 3


def tcvn3_to_unicode(text: str) -> str:
    return text.translate(TCVN3_TABLE)


def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.replace("đ", "d").replace("Đ", "D")


def norm_key(text) -> str:
    """Chuẩn hoá chuỗi để so sánh: bỏ dấu, chữ thường, bỏ khoảng trắng thừa."""
    return " ".join(strip_accents(text).lower().split())


# ---------------------------------------------------------------------------
# Số và ngày
# ---------------------------------------------------------------------------
_NUM_RE = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$")


def to_number(value, default=None):
    """Chuyển chuỗi số ("1,234.56", "'1234") thành float/int. Không phải số -> default."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace("'", "").replace('"', "").replace(" ", "")
    if s in ("", "-", "+", "."):
        return default
    if not _NUM_RE.match(s):
        return default
    s = s.replace(",", "")
    try:
        f = float(s)
    except ValueError:
        return default
    if f.is_integer() and "." not in s and abs(f) < 1e15:
        return int(f)
    return f


def num0(value) -> float:
    n = to_number(value, 0)
    return n if n is not None else 0


def is_number(value) -> bool:
    return to_number(value) is not None


def parse_date(value):
    """Nhận yyyymmdd, dd/mm/yyyy, yyyy-mm-dd ... -> datetime.date. Không được -> None."""
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    s = str(value).strip().replace("'", "")
    if not s:
        return None
    try:
        if re.fullmatch(r"\d{8}", s):
            return _dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:\s.*)?", s)
        if m:
            return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        m = re.fullmatch(r"(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})(?:[\sT].*)?", s)
        if m:
            return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None
    return None


def date_or_text(value):
    """Trả về date nếu đọc được, ngược lại giữ nguyên chuỗi."""
    d = parse_date(value)
    if d is not None:
        return d
    return "" if value is None else str(value).strip()


def fmt_date(d) -> str:
    if isinstance(d, (_dt.date, _dt.datetime)):
        return d.strftime("%d/%m/%Y")
    return "" if d is None else str(d)


def date_full(d) -> str:
    """13 tháng 12 năm 2024"""
    return "%02d tháng %02d năm %04d" % (d.day, d.month, d.year)


def today_line() -> str:
    t = _dt.date.today()
    return "Ngày %02d tháng %02d năm %04d" % (t.day, t.month, t.year)


# ---------------------------------------------------------------------------
# Đọc số thành chữ
# ---------------------------------------------------------------------------
_DIGITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]


def _read3(group: str, first: bool) -> str:
    tram, chuc, dv = (int(c) for c in group)
    words = []
    if not (first and tram == 0):
        words += [_DIGITS[tram], "trăm"]
    if chuc == 0 and dv == 0:
        return " ".join(words)
    if chuc == 0:
        if words:
            words.append("lẻ")
    elif chuc == 1:
        words.append("mười")
    else:
        words += [_DIGITS[chuc], "mươi"]
    if dv == 1:
        words.append("mốt" if chuc >= 2 else "một")
    elif dv == 5:
        words.append("lăm" if chuc != 0 else "năm")
    elif dv == 4 and chuc >= 2:
        words.append("tư")
    elif dv != 0:
        words.append(_DIGITS[dv])
    return " ".join(words)


def read_integer(n: int, thousand_word: str = "nghìn") -> str:
    if n == 0:
        return "không"
    neg = n < 0
    s = str(abs(int(n)))
    while len(s) % 3:
        s = "0" + s
    groups = [s[i:i + 3] for i in range(0, len(s), 3)]
    units = ["", thousand_word, "triệu", "tỷ"]
    total = len(groups)
    parts = []
    for idx, g in enumerate(groups):
        k = total - idx - 1
        if g == "000":
            continue
        text = _read3(g, idx == 0)
        unit_words = []
        kk = k
        if kk % 3:
            unit_words.append(units[kk % 3])
        unit_words += ["tỷ"] * (kk // 3)
        parts.append((text + " " + " ".join(unit_words)).strip())
    out = " ".join(parts)
    out = re.sub(r"\s+", " ", out).strip()
    return ("âm " + out) if neg else out


_CCY_UNITS = {
    "VND": ("đồng", "hào"),
    "USD": ("đô la Mỹ", "cents"),
    "EUR": ("euro", "cents"),
    "GBP": ("bảng Anh", "pennies"),
}


def number_to_words(value, ccy: str = "", thousand_word: str = "nghìn") -> str:
    """Đọc số tiền thành chữ. Ví dụ: 1250000, VND -> 'Một triệu hai trăm năm mươi nghìn đồng'."""
    n = to_number(value)
    if n is None:
        return ""
    unit, sub = _CCY_UNITS.get(str(ccy).strip().upper(), ("", ""))
    neg = n < 0
    n = abs(n)
    integer = int(n)
    frac = int(round((n - integer) * 100))
    if frac == 100:
        integer += 1
        frac = 0
    text = read_integer(integer, thousand_word)
    if frac:
        frac_text = read_integer(frac, thousand_word)
        if unit:
            text = "%s %s và %s %s" % (text, unit, frac_text, sub)
        else:
            text = "%s phẩy %s" % (text, frac_text)
    elif unit:
        text = "%s %s" % (text, unit)
    if neg:
        text = "âm " + text
    text = text.strip()
    return text[:1].upper() + text[1:]


def col_letter(idx: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letter: str) -> int:
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1
