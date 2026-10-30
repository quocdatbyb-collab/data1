"""Đối chiếu số dư sao kê với cân đối tài khoản IPCAS (file *glcb41*) và lập Mẫu 30."""

from __future__ import annotations

import glob
import os
import re
from collections import OrderedDict

from . import vn
from .reader import read_sources
from .sheet import Col, Sheet, Diff, bank_block, form_block, unit_line, signatures
from .forms import g, sig, SIG_LBANG, SIG_KT, SIG_GD


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
        return self.by_acct.get(key[1:])


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

    # ---- Mẫu 30 ----
    cols = [Col("TT", "int", 5.57), Col("CHI NHÁNH TRỰC THUỘC", "text", 14.71),
            Col("MÃ CHI NHÁNH TRÊN HỆ THỐNG IPCAS", "center", 15.14), Col("TÀI KHOẢN (Cấp V)", "center", 10.29),
            Col("SỐ DƯ TÀI KHOẢN TRÊN CÂN ĐỐI IPCAS", "num", 21.86),
            Col("SỐ LIỆU TỔNG CỘNG CỦA TÀI KHOẢN TRÊN BÁO CÁO QUYẾT TOÁN", "num", 21.86),
            Col("CHÊNH LỆCH", "num", 17), Col("NGUYÊN NHÂN", "text", 20)]
    m30 = Sheet("Mau 30", cols, header_row=9)
    m30.header_rows = 2
    m30.header_cells = [(9, 0, 10, 0, "TT"), (9, 1, 10, 1, "CHI NHÁNH TRỰC THUỘC"),
                        (9, 2, 10, 2, "MÃ CHI NHÁNH TRÊN HỆ THỐNG IPCAS"), (9, 3, 10, 3, "TÀI KHOẢN (Cấp V)"),
                        (9, 4, 9, 6, "TÌNH TRẠNG BÁO CÁO"), (10, 4, 10, 4, cols[4].header),
                        (10, 5, 10, 5, cols[5].header), (10, 6, 10, 6, "CHÊNH LỆCH"), (9, 7, 10, 7, "NGUYÊN NHÂN")]
    m30.numbering_labels = [1, 2, 3, 4, 5, 6, "7 = 5 - 6", 8]
    for n, ((brcd, acct), vnd) in enumerate(by_acct.items(), 1):
        row = m30.add("data")
        r = m30.next_row() - 1
        row.set(0, n).set(1, ctx.branch_name(brcd)).set(2, brcd).set(3, acct).set(5, vnd)
        bal = book.lookup(brcd, acct, "000")
        if bal is None:
            row.set(4, "?")
            row.set_align(4, "center")
        else:
            row.set(4, bal).set(6, Diff((4, r), (5, r)))
    bank_block(m30, ctx, col="A", across_to="C")
    m30.put(6, "A", "TỔNG HỢP SỐ LIỆU TOÀN CHI NHÁNH MẪU BIỂU QUYẾT TOÁN SỐ " + form_label.upper(),
            across_to="H", bold=True, size=16)
    m30.put(7, "A", "Tên mẫu biểu: " + ctx.form_name(form_label), across_to="H", size=12)
    m30.put(8, "A", "Ngày " + ctx.date_full, across_to="G", size=12, italic=True, bold=True)
    form_block(m30, "H", ["1. Mẫu số 30/" + ctx.loai_qt, "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(m30, 8, "H")
    signatures(m30, ctx, [sig("B", SIG_LBANG), sig("E", SIG_KT), sig("G", SIG_GD)])
    diffs = sum(1 for r in chk.rows if r.kind == "sub2")
    missing = sum(1 for r in chk.rows if r.cells[3] == "?")
    return [chk, m30], diffs, missing
