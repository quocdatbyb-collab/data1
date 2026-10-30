"""Cấu hình ứng dụng (file QT.ini đặt cạnh chương trình) và ngữ cảnh khi tạo báo cáo."""

from __future__ import annotations

import configparser
import datetime as _dt
import json
import os
import sys

from . import vn
from . import forms

DEFAULT_BRANCHES = [
    ("5500", "Tỉnh Bình Dương"), ("5501", "Tp Thủ Dầu Một"), ("5502", "Tx Tân Uyên"), ("5503", "Tx Bến Cát"),
    ("5504", "Huyện Dầu Tiếng"), ("5505", "Huyện Phú Giáo"), ("5506", "Tân Phước Khánh"), ("5507", "Sở Sao"),
]

SIGNERS = [("lap_bieu", "Người lập biểu"), ("tp_tin_dung", "Trưởng phòng Tín dụng"),
           ("tp_ke_toan", "Trưởng phòng Kế toán"), ("tp_ttqt", "Trưởng phòng TTQT"), ("giam_doc", "Giám đốc")]

DEFAULTS = {
    "ngay_qt": "31/12/%d" % (_dt.date.today().year if _dt.date.today().month == 12 else _dt.date.today().year - 1),
    "loai_qt": "QT",
    "ty_gia_usd": "23560",
    "kho_giay": "A4",
    "doc_so": "nghìn",
    "ma_ky_tu": "auto",
    "co_chu": "10",
    "thu_muc_ket_qua": "",
    "kem_rt06": "1",
    "kem_rt02": "1",
    "kiem_tra_gl": "1",
    "mo_file": "1",
    "bo_qua_da_co": "1",
    "uu_tien_mau_tuy_chinh": "0",
    "chi_cn_danh_muc": "0",
}


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def data_dir() -> str:
    """Thư mục lưu cấu hình/mẫu: cạnh chương trình (dùng USB được), không ghi được thì dùng AppData."""
    if os.environ.get("QT_DATA_DIR"):
        os.makedirs(os.environ["QT_DATA_DIR"], exist_ok=True)
        return os.environ["QT_DATA_DIR"]
    base = app_dir()
    try:
        test = os.path.join(base, ".qt_write_test")
        with open(test, "w") as f:
            f.write("ok")
        os.remove(test)
        return base
    except OSError:
        root = os.environ.get("APPDATA") or os.path.expanduser("~")
        path = os.path.join(root, "QT_QuyetToan")
        os.makedirs(path, exist_ok=True)
        return path


def _read_ini_text(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1258", errors="replace")


def _flag(v) -> bool:
    return str(v).strip() in ("1", "true", "True", "yes", "co", "có")


class Settings:
    def __init__(self, path=None):
        self.dir = data_dir()
        self.path = path or os.path.join(self.dir, "QT.ini")
        self.colmap_path = os.path.join(self.dir, "cot_du_lieu.json")
        self.values = dict(DEFAULTS)
        self.form = {"fCSV": "", "Toantinh": "0", "Tomau": "1", "Kiemtra": "0"}
        self.signers = {k: "" for k, _ in SIGNERS}
        self.branches = list(DEFAULT_BRANCHES)
        self.colmaps = {}
        self.load()

    # ----------------------------------------------------------------- I/O
    def load(self):
        if os.path.exists(self.path):
            cp = configparser.ConfigParser(interpolation=None, strict=False)
            cp.optionxform = str
            try:
                cp.read_string(_read_ini_text(self.path))
            except configparser.Error:
                cp = None
            if cp is not None:
                if cp.has_section("Form"):
                    for k, v in cp.items("Form"):
                        self.form[k] = v
                if cp.has_section("CaiDat"):
                    for k, v in cp.items("CaiDat"):
                        self.values[k] = v
                if cp.has_section("NguoiKy"):
                    for k, v in cp.items("NguoiKy"):
                        self.signers[k] = v
                if cp.has_section("Brcd"):
                    br = [(k.strip(), v.strip()) for k, v in cp.items("Brcd") if k.strip()]
                    if br:
                        self.branches = br
        if os.path.exists(self.colmap_path):
            try:
                with open(self.colmap_path, "r", encoding="utf-8") as f:
                    self.colmaps = json.load(f)
            except (OSError, ValueError):
                self.colmaps = {}

    def save(self):
        cp = configparser.ConfigParser(interpolation=None)
        cp.optionxform = str
        cp["Form"] = {k: str(v) for k, v in self.form.items()}
        cp["CaiDat"] = {k: str(v) for k, v in self.values.items()}
        cp["NguoiKy"] = {k: str(v) for k, v in self.signers.items()}
        cp["Brcd"] = {k: v for k, v in self.branches}
        with open(self.path, "w", encoding="utf-8-sig") as f:
            cp.write(f)
        with open(self.colmap_path, "w", encoding="utf-8") as f:
            json.dump(self.colmaps, f, ensure_ascii=False, indent=1)

    # ------------------------------------------------------------- helpers
    def get(self, key):
        return self.values.get(key, DEFAULTS.get(key, ""))

    def flag(self, key) -> bool:
        return _flag(self.get(key))

    def set(self, key, value):
        if isinstance(value, bool):
            value = "1" if value else "0"
        self.values[key] = str(value)

    def branch_name(self, brcd) -> str:
        for k, v in self.branches:
            if str(k) == str(brcd):
                return v
        return ""

    def colmap(self, code):
        spec = forms.FORMS.get(code, {})
        base_code = spec.get("fields_from", code)
        m = forms.default_colmap(base_code)
        user = self.colmaps.get(base_code) or {}
        for k, v in user.items():
            try:
                m[k] = int(v)
            except (TypeError, ValueError):
                pass
        return m

    def qt_date(self):
        d = vn.parse_date(self.get("ngay_qt"))
        return d or _dt.date(_dt.date.today().year - 1, 12, 31)

    def usd_rate(self) -> float:
        return vn.to_number(self.get("ty_gia_usd"), 23560) or 23560

    def paper(self) -> int:
        return 8 if str(self.get("kho_giay")).upper() == "A3" else 9


class Context:
    """Thông tin dùng khi dựng một báo cáo."""

    def __init__(self, settings: Settings, brcd: str, options=None, custom_names=None):
        self.s = settings
        self.brcd = str(brcd)
        self.brname = settings.branch_name(brcd) or ""
        self.options = dict(options or {})
        d = settings.qt_date()
        self.date = d
        self.date_full = vn.date_full(d)
        self.date_short = d.strftime("%d/%m/%Y")
        self.ddmm = d.strftime("%d/%m")
        self.loai_qt = settings.get("loai_qt") or "QT"
        self.usd_rate = settings.usd_rate()
        self.thousand_word = settings.get("doc_so") or "nghìn"
        self.encoding = settings.get("ma_ky_tu") or "auto"
        self.paper = settings.paper()
        self.font_size = int(vn.to_number(settings.get("co_chu"), 10) or 10)
        self.tomau = bool(self.options.get("tomau", True))
        self.custom_names = custom_names or {}

    def opt(self, key) -> bool:
        if key in self.options:
            return bool(self.options[key])
        return self.s.flag(key)

    def colmap(self, code):
        return self.s.colmap(code)

    def signer(self, key) -> str:
        return (self.s.signers.get(key) or "").strip()

    def sheet_name(self, code) -> str:
        return "%s_%s" % (self.brcd, code)

    def branch_name(self, brcd) -> str:
        return self.s.branch_name(brcd) or str(brcd)

    def form_name(self, code) -> str:
        spec = forms.FORMS.get(forms.resolve(code))
        if spec:
            return spec["name"]
        return self.custom_names.get(code, code)
