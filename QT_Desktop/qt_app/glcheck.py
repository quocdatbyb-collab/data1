"""Đối chiếu số dư sao kê với cân đối tài khoản IPCAS (file *glcb41*) và lập Mẫu 30."""

from __future__ import annotations

import glob
import os
import re
from collections import OrderedDict

from . import vn
from .reader import read_sources
from .sheet import Col, Sheet, Diff, Sub, SumOf, bank_block, form_block, unit_line, signatures
from .forms import g, sig, SIG_LB, SIG_KT, SIG_GD


class GLBook:
    def __init__(self):
        self.files = []
        self.by_branch = {}   # (brcd, acct, ccy) -> balance
        self.by_acct = {}     # (acct, ccy) -> balance

    @property
    def empty(self):
        return not self.by_acct

    def lookup(self, brcd, acct, ccy):
        key = (str(brcd), str(acct).strip(), str(ccy).strip().upper())
        if key in self.by_branch:
            return self.by_branch[key]
        if key[0] and any(k[0] == key[0] for k in self._branches()):
            return self._prefix(self.by_branch, key, True)
        if key[1:] in self.by_acct:
            return self.by_acct[key[1:]]
        return self._prefix(self.by_acct, key[1:], False)

    def _branches(self):
        if not hasattr(self, "_brset"):
            self._brset = {(k[0],) for k in self.by_branch}
        return self._brset

    @staticmethod
    def _prefix(table, key, with_branch):
        """Tài khoản tổng hợp (vd 9711) = cộng các tài khoản chi tiết bắt đầu bằng số đó."""
        found = None
        for k, v in table.items():
            if with_branch:
                ok = k[0] == key[0] and k[2] == key[2] and k[1].startswith(key[1]) and k[1] != key[1]
            else:
                ok = k[1] == key[1] and k[0].startswith(key[0]) and k[0] != key[0]
            if ok:
                found = (found or 0) + v
        return found


def find_gl_files(folder: str):
    out = []
    for p in glob.glob(os.path.join(folder, "*")):
        b = os.path.basename(p).lower()
        if "glcb41" in b and b.endswith((".csv", ".txt", ".tsv")):
            out.append(p)
    return sorted(out)


def load_gl(folder: str, ctx) -> GLBook:
    book = GLBook()
    m = ctx.colmap("gl")
    for p in find_gl_files(folder):
        data = read_sources([p], ctx.encoding)
        mb = re.match(r"^(\d{4})", os.path.basename(p))
        brcd = mb.group(1) if mb else ""
        book.files.append(p)
        for row in data.rows:
            ccy = str(g(row, m["loai_tien"])).strip().upper()
            acct = str(g(row, m["tk"])).strip()
            if not acct:
                continue
            bal = abs(vn.num0(g(row, m["du_no"])) - vn.num0(g(row, m["du_co"])))
            if brcd:
                book.by_branch[(brcd, acct, ccy)] = bal
            book.by_acct[(acct, ccy)] = bal
    return book


def _aggregate(items):
    by_ccy = OrderedDict()
    by_acct = OrderedDict()
    for it in sorted(items, key=lambda x: (x.brcd, x.acct, x.ccy)):
        if not it.acct:
            continue
        k = (it.brcd, it.acct, it.ccy)
        a = by_ccy.setdefault(k, [0, 0])
        a[0] += it.amt or 0
        a[1] += it.amt_vnd or 0
        k2 = (it.brcd, it.acct)
        by_acct[k2] = by_acct.get(k2, 0) + (it.amt_vnd or 0)
    return by_ccy, by_acct


def build_check_sheets(ctx, items, book: GLBook, form_label: str):
    by_ccy, by_acct = _aggregate(items)

    # ---- Sheet chấm số dư từng loại tiền ----
    cols = [Col("Mã CN", "center", 8), Col("Tài khoản", "text", 14), Col("Loại tiền", "center", 9),
            Col("Cân đối IPCAS", "num2", 22), Col("Sao kê " + form_label, "num2", 22), Col("Chênh lệch", "num2", 18)]
    chk = Sheet("CheckGL", cols, header_row=3)
    chk.simple_header()
    chk.numbering = False
    chk.landscape = False
    for (brcd, acct, ccy), (amt, _vnd) in by_ccy.items():
        row = chk.add("data")
        row.set(0, brcd).set(1, acct).set(2, ccy).set(4, round(amt, 2))
        bal = book.lookup(brcd, acct, ccy)
        r = chk.next_row() - 1
        if bal is None:
            row.set(3, "?").set(5, "-")
            row.set_align(3, "center")
            row.set_align(5, "center")
        else:
            row.set(3, bal).set(5, Diff((3, r), (4, r)))
            if abs(bal - amt) > 0.005:
                row.kind = "sub2"
    chk.put(1, "A", "CHẤM SỐ DƯ TÀI KHOẢN GIỮA GL VÀ SAO KÊ", across_to="F", bold=True, size=12)
    if book.empty:
        chk.put(2, "A", "(Không tìm thấy file cân đối *glcb41* trong thư mục dữ liệu – cột Cân đối để '?')",
                across_to="F", italic=True)

    diffs = sum(1 for r in chk.rows if r.kind == "sub2")
    missing = sum(1 for r in chk.rows if r.cells[3] == "?")
    m30 = build_mau30(ctx, items, book, form_label)
    return [chk, m30], diffs, missing


FORM_NUMBERS = {"rt05": "05", "rt06": "06", "rt13": "13", "rt14": "14", "rt15a": "15a", "rt15b": "15b",
                "rt16": "16", "rt18": "18", "rt20a": "20a", "rt22": "22", "rt24": "24a", "rttt": "HTT"}

# Mẫu 30 áp dụng cho các mẫu (theo hướng dẫn QT 2025: 13, 15a, 15b, 18, 20a, 22, 23, 24)
MAU30_FORMS = ["rt13", "rt15a", "rt15b", "rt18", "rt20a", "rt22", "rt24", "rt14", "rt16"]

MAU30_NOTES = [
    "Ghi chú:",
    "     1. Chi nhánh loại I, II tổng hợp mẫu biểu này đối với các mẫu quyết toán: 13/QT, 15a/QT, 15b/QT, "
    "18/QT, 20a/QT, 22/QT, 23/QT, 24/QT",
    "     2. Cột 8: trong trường hợp không khớp đúng tại cột 7, chi nhánh phải nêu rõ nguyên nhân.",
    "     3. Chi nhánh loại I, II tổng hợp mẫu biểu này đối với các mẫu quyết toán và được lưu vào sheet 1 "
    "của file tổng hợp mẫu biểu quyết toán đó.",
]


def build_mau30(ctx, items, book: GLBook, form_code: str, branches=None, subtotals=None):
    """Mẫu 30 – Tổng hợp số liệu toàn chi nhánh (theo mẫu QT 2025).

    items: GLItem của một hoặc nhiều chi nhánh. branches: danh sách mã CN để giữ thứ tự (và hiện cả
    chi nhánh không có số liệu). subtotals: thêm dòng "Cộng TK" khi có từ 2 chi nhánh trở lên.
    """
    _by_ccy, by_acct = _aggregate(items)
    so = FORM_NUMBERS.get(form_code, form_code)
    cols = [Col("TT", "int", 5.3), Col("CHI NHÁNH TRỰC THUỘC", "text", 18),
            Col("MÃ CHI NHÁNH TRÊN HỆ THỐNG IPCAS", "center", 16), Col("TÀI KHOẢN (Cấp V)", "center", 13.3),
            Col("SỐ DƯ TÀI KHOẢN TRÊN CÂN ĐỐI IPCAS", "num", 20),
            Col("SỐ LIỆU TỔNG CỘNG CỦA TÀI KHOẢN TRÊN BÁO CÁO QUYẾT TOÁN\n CỦA CHI NHÁNH", "num", 30),
            Col("CHÊNH LỆCH", "num", 16), Col("NGUYÊN NHÂN", "text", 22)]
    m30 = Sheet("Mau 30", cols, header_row=9)
    m30.header_rows = 2
    m30.header_cells = [(9, 0, 10, 0, "TT"), (9, 1, 10, 1, "CHI NHÁNH TRỰC THUỘC"),
                        (9, 2, 10, 2, "MÃ CHI NHÁNH TRÊN HỆ THỐNG IPCAS"), (9, 3, 10, 3, "TÀI KHOẢN (Cấp V)"),
                        (9, 4, 9, 6, "TÌNH TRẠNG BÁO CÁO"), (10, 4, 10, 4, cols[4].header),
                        (10, 5, 10, 5, cols[5].header), (10, 6, 10, 6, "CHÊNH LỆCH"), (9, 7, 10, 7, "NGUYÊN NHÂN")]
    m30.numbering_labels = [1, 2, 3, 4, 5, 6, "7 = 5 - 6", 8]

    order = list(branches or [])
    for (b, _a) in by_acct:
        if b not in order:
            order.append(b)
    accts = sorted({a for (_b, a) in by_acct})
    multi = len(order) > 1 if subtotals is None else subtotals
    n = 0
    total_refs = {4: [], 5: [], 6: []}

    def gl_of(brcd, acct):
        return book.lookup(brcd, acct, "000")

    for acct in accts:
        start = m30.next_row()
        for brcd in order:
            if (brcd, acct) not in by_acct and not multi:
                continue
            n += 1
            row = m30.add("data")
            r = m30.next_row() - 1
            vnd = by_acct.get((brcd, acct), 0)
            row.set(0, n).set(1, ctx.branch_name(brcd)).set(2, brcd).set(3, acct).set(5, vnd)
            bal = gl_of(brcd, acct)
            if bal is None:
                row.set(4, "?").set(7, "Không có số dư trên cân đối" if not book.empty else "")
                row.set_align(4, "center")
            else:
                row.set(4, bal).set(6, Diff((4, r), (5, r)))
                if abs(bal - vnd) > 0.5:
                    row.kind = "sub1"
        end = m30.next_row() - 1
        if multi and end >= start:
            sub = m30.add("sub2").set(1, "Cộng TK " + acct).set(3, acct)
            for c in (4, 5, 6):
                sub.set(c, Sub(c, start, end))
                total_refs[c].append((c, m30.next_row() - 1))
    if multi and len(accts) > 1:
        tot = m30.add("total").set(1, "TỔNG CỘNG")
        for c in (4, 5, 6):
            tot.set(c, SumOf(total_refs[c]))

    bank_block(m30, ctx, col="A", star=False, across_to="C")
    m30.put(5, "A", "TỔNG HỢP SỐ LIỆU TOÀN CHI NHÁNH MẪU BIỂU QUYẾT TOÁN SỐ %s/%s" % (so, ctx.loai_qt),
            across_to="H", bold=True, size=14)
    m30.put(6, "A", "Tên mẫu biểu: " + ctx.form_name(form_code), across_to="H", size=12)
    m30.put(7, "A", "Ngày " + ctx.date_full, across_to="H", size=12, italic=True)
    form_block(m30, "H", ["1. Mẫu số 30/" + ctx.loai_qt, "2. CN loại I gửi kèm trong file mẫu biểu QT về TSC",
                          "3. Lưu tại chi nhánh"])
    unit_line(m30, 8, "H", "Đơn vị : VNĐ")
    signatures(m30, ctx, [sig("B", SIG_LB), sig("E", SIG_KT), sig("G", SIG_GD)], date_col="G")
    m30.footer["notes"] = MAU30_NOTES
    return m30
