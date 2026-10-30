"""Đọc file CSV xuất từ IPCAS: tự nhận mã ký tự, dấu phân cách, ghép nhiều file."""

from __future__ import annotations

import csv
import glob
import io
import os
import re
import sys
import unicodedata

from . import vn

ENCODING_AUTO = "auto"
ENCODINGS = [
    (ENCODING_AUTO, "Tự động nhận dạng"),
    ("unicode", "Unicode (UTF-8 / UTF-16)"),
    ("tcvn3", "TCVN3 (ABC)"),
    ("cp1258", "Windows-1258"),
]


class SourceData:
    """Dữ liệu đọc từ 1 hoặc nhiều file cùng mẫu."""

    def __init__(self):
        self.files = []          # danh sách file đã đọc
        self.header = []         # dòng tiêu đề (file đầu tiên)
        self.rows = []           # các dòng dữ liệu (list[str])
        self.encoding = ""       # mã ký tự đã nhận dạng
        self.delimiter = ","
        self.warnings = []

    def __len__(self):
        return len(self.rows)


def _decode(raw: bytes, mode: str = ENCODING_AUTO):
    """Trả về (text, tên mã ký tự)."""
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        text, name = raw.decode("utf-16"), "UTF-16"
    elif raw.startswith(b"\xef\xbb\xbf"):
        text, name = raw[3:].decode("utf-8", errors="replace"), "UTF-8"
    elif len(raw) > 4 and raw[1:2] == b"\x00" and raw[3:4] == b"\x00":
        text, name = raw.decode("utf-16-le", errors="replace"), "UTF-16 LE"
    else:
        try:
            text, name = raw.decode("utf-8"), "UTF-8"
        except UnicodeDecodeError:
            if mode == "cp1258":
                text, name = raw.decode("cp1258", errors="replace"), "Windows-1258"
            else:
                latin = raw.decode("latin-1")
                if mode == "tcvn3" or (mode == ENCODING_AUTO and vn.looks_like_tcvn3(latin)):
                    return vn.tcvn3_to_unicode(latin), "TCVN3"
                text, name = raw.decode("cp1258", errors="replace"), "Windows-1258"
    if mode == "tcvn3" or (mode == ENCODING_AUTO and vn.looks_like_tcvn3(text)):
        return vn.tcvn3_to_unicode(text), name + " (TCVN3)"
    return unicodedata.normalize("NFC", text), name


def _sniff_delimiter(first_line: str) -> str:
    counts = {}
    for d in (",", "\t", ";", "|"):
        try:
            counts[d] = len(next(csv.reader([first_line], delimiter=d)))
        except (StopIteration, csv.Error):
            counts[d] = 0
    best = max(counts, key=lambda k: counts[k])
    return best if counts[best] > 1 else ","


def _clean(field: str) -> str:
    return field.replace('"', "").strip()


def _intern(field: str) -> str:
    """Làm sạch ô; chuỗi ngắn hay lặp lại (ngày, loại tiền, mã) được dùng chung để tiết kiệm bộ nhớ."""
    v = field.replace('"', "").strip() if ('"' in field or field[:1] == " " or field[-1:] == " ") else field
    return sys.intern(v) if len(v) <= 12 else v


def read_text_rows(text: str, delimiter: str | None = None):
    """Tách văn bản CSV thành (header, rows, delimiter)."""
    text = text.lstrip("\r\n\ufeff")
    if not text.strip():
        return [], [], delimiter or ","
    nl = text.find("\n")
    first = text if nl < 0 else text[:nl]
    delim = delimiter or _sniff_delimiter(first.rstrip("\r"))
    reader = csv.reader(io.StringIO(text, newline=None), delimiter=delim)
    all_rows = []
    try:
        for row in reader:
            if not row:
                continue
            row = [_intern(c) for c in row]
            if any(row):
                all_rows.append(row)
    except csv.Error:
        # File có dấu nháy lỗi: tách thủ công từng dòng
        all_rows = [[_clean(c) for c in ln.split(delim)] for ln in text.splitlines() if ln.strip()]
    if not all_rows:
        return [], [], delim
    return all_rows[0], all_rows[1:], delim


def repair_rows(rows, min_cols: int):
    """Ghép các dòng bị ngắt giữa chừng (ít cột hơn quy định) với dòng kế tiếp."""
    if not min_cols:
        return rows
    out = []
    i = 0
    n = len(rows)
    while i < n:
        row = list(rows[i])
        while len(row) < min_cols and i + 1 < n and len(rows[i + 1]) < min_cols:
            nxt = rows[i + 1]
            row[-1] = row[-1] + nxt[0] if row else nxt[0]
            row.extend(nxt[1:])
            i += 1
        out.append(row)
        i += 1
    return out


def find_parts(path: str):
    """Tìm tất cả file cùng mẫu: 5500_rt13.csv, 5500_rt13_1.csv, 5500_rt13_2.csv..."""
    folder, name = os.path.split(path)
    stem = os.path.splitext(name)[0]
    m = re.match(r"^(\d{4}_[A-Za-z0-9]+)", stem)
    prefix = m.group(1) if m else stem
    parts = []
    for p in glob.glob(os.path.join(folder or ".", "*")):
        base = os.path.basename(p)
        if not base.lower().endswith((".csv", ".txt")):
            continue
        b_stem = os.path.splitext(base)[0]
        if b_stem.lower() == prefix.lower() or re.match(re.escape(prefix) + r"[_\- (]", b_stem, re.I):
            parts.append(p)
    if path not in parts and os.path.exists(path):
        parts.append(path)
    return sorted(set(parts), key=lambda s: s.lower())


def read_sources(paths, encoding: str = ENCODING_AUTO, min_cols: int = 0, limit: int = 0) -> SourceData:
    """Đọc và ghép nhiều file. Dòng đầu mỗi file là tiêu đề và được bỏ qua."""
    data = SourceData()
    encs = []
    for p in paths:
        with open(p, "rb") as f:
            raw = f.read()
        text, enc = _decode(raw, encoding)
        del raw
        encs.append(enc)
        header, rows, delim = read_text_rows(text)
        del text
        if not data.files:
            data.header = header
            data.delimiter = delim
        data.files.append(p)
        data.rows.extend(rows)
        if limit and len(data.rows) >= limit:
            break
    data.encoding = ", ".join(sorted(set(encs)))
    data.rows = repair_rows(data.rows, min_cols)
    if limit:
        data.rows = data.rows[:limit]
    if min_cols:
        short = sum(1 for r in data.rows if len(r) < min_cols)
        if short:
            data.warnings.append("%d dòng có ít hơn %d cột (đã giữ nguyên, cột thiếu để trống)" % (short, min_cols))
    return data


def parse_filename(path: str):
    """'5500_rt13_1.csv' -> ('5500', 'rt13'). Không đúng mẫu -> ('', '')."""
    stem = os.path.splitext(os.path.basename(path))[0]
    m = re.match(r"^(\d{4})_([A-Za-z0-9]+)", stem)
    if not m:
        return "", ""
    return m.group(1), m.group(2).lower()
