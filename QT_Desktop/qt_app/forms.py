"""Các mẫu biểu quyết toán có sẵn.

Mỗi mẫu đọc dữ liệu CSV theo bảng ánh xạ cột (có thể sửa trong phần Cài đặt)
và dựng ra một hoặc nhiều Sheet.
"""

from __future__ import annotations

from collections import OrderedDict

from . import vn
from .sheet import Col, Sheet, Sub, CountA, SumOf, bank_block, form_block, title_lines, unit_line, signatures

L = vn.col_index  # "A" -> 0

SIG_LB = ("LẬP BIỂU", "(Ký, ghi rõ họ tên, số ĐT liên hệ)", "lap_bieu")
SIG_LBANG = ("LẬP BẢNG", "(Ký, ghi rõ họ tên, số ĐT liên hệ)", "lap_bieu")
SIG_TD = ("TRƯỞNG PHÒNG TÍN DỤNG", "(Ký, ghi rõ họ tên)", "tp_tin_dung")
SIG_TD_S = ("TP TÍN DỤNG", "(Ký, ghi rõ họ tên)", "tp_tin_dung")
SIG_KT = ("TRƯỞNG PHÒNG KẾ TOÁN", "(Ký, ghi rõ họ tên)", "tp_ke_toan")
SIG_KT_S = ("TP KẾ TOÁN", "(Ký, ghi rõ họ tên)", "tp_ke_toan")
SIG_TTQT = ("TRƯỞNG PHÒNG TTQT", "(Ký, ghi rõ họ tên)", "tp_ttqt")
SIG_GD = ("GIÁM ĐỐC", "(Ký, đóng dấu, ghi rõ họ tên)", "giam_doc")


def sig(col, s):
    return (col, s[0], s[1], s[2])


class GLItem:
    __slots__ = ("brcd", "acct", "ccy", "amt", "amt_vnd")

    def __init__(self, brcd, acct, ccy, amt, amt_vnd):
        self.brcd, self.acct, self.ccy, self.amt, self.amt_vnd = brcd, acct, ccy, amt, amt_vnd


class Output:
    """Một file kết quả."""

    def __init__(self, code, sheets, gl_items=None, gl_label=None):
        self.code = code
        self.sheets = sheets
        self.gl_items = gl_items
        self.gl_label = gl_label or code


def g(row, idx):
    """Lấy cột thứ idx (đếm từ 1) của dòng; thiếu cột -> chuỗi rỗng."""
    if not idx:
        return ""
    i = int(idx) - 1
    if 0 <= i < len(row):
        return row[i]
    return ""


def cust_code(brcd, raw):
    raw = str(raw).strip()
    return str(brcd).strip() + raw[-9:]


def fc_fmt(row, col, ccy):
    if str(ccy).strip().upper() != "VND":
        row.set_fmt(col, "#,##0.00")


def _set_header_numbered(sheet: Sheet):
    sheet.simple_header()
    sheet.numbering = True


# ===========================================================================
# MẪU 05 – Kiểm kê hồ sơ, tài sản thế chấp, cầm cố
# ===========================================================================

RT05_SECTIONS = [
    (1, "I. TSTC của khách hàng là TCTD trực tiếp vay vốn:", "Cộng I"),
    (2, "II. TSTC của khách hàng là TCKT&CN trực tiếp vay vốn:", "Cộng II"),
    (3, "III. TSTC của đơn vị bảo lãnh cho bên thứ ba vay vốn:", "Cộng III"),
]


def _rt05_type(m, row):
    loai = vn.norm_key(g(row, m["loai_tc"]))
    tctd = str(g(row, m["ma_loai_kh"])).strip() == str(m.get("ma_tctd", "600"))
    if loai == "the chap":
        return 1 if tctd else 2
    return 3


def _rt05_records(ctx, data):
    m = ctx.colmap("rt05")
    recs = []
    for row in data.rows:
        recs.append({
            "type": _rt05_type(m, row),
            "cust": cust_code(g(row, m["ma_cn"]), g(row, m["ma_kh"])),
            "name": g(row, m["ten_kh"]),
            "so_hd": g(row, m["so_hd"]),
            "ngay_raw": str(g(row, m["ngay_hd"])).strip(),
            "ngay": vn.date_or_text(g(row, m["ngay_hd"])),
            "tong": vn.num0(g(row, m["tong"])),
            "bds": vn.num0(g(row, m["bds"])),
            "ds": vn.num0(g(row, m["ds"])),
            "gtcg": vn.num0(g(row, m["gtcg"])),
            "khac": vn.num0(g(row, m["ts_khac"])),
            "pm_co": "X" if str(g(row, m["pm_co"])).strip() else "",
            "pm_it": "X" if str(g(row, m["pm_it"])).strip() else "",
            "pm_khong": "X" if str(g(row, m["pm_khong"])).strip() else "",
        })
    recs.sort(key=lambda r: (r["type"], r["cust"], r["ngay_raw"]))
    return recs


def build_rt05(ctx, data):
    recs = _rt05_records(ctx, data)
    cols = [Col("Mã khách hàng", "text", 18), Col("Tên khách hàng"), Col("Số HĐ thế chấp/cầm cố"),
            Col("Ngày HĐ thế chấp/cầm cố", "date", 12)]
    cols += [Col(h, "num") for h in ("Tổng", "Bất ĐS", "Động sản", "GTCG", "TS khác")]
    cols += [Col(h, "mark", 11) for h in ("Hợp pháp, hợp lệ", "Khác", "Đầy đủ hồ sơ theo quy định",
                                          "Đang hoàn chỉnh thủ tục", "Không thể hoàn chỉnh được", "Khác",
                                          "Có khả năng phát mại", "Ít khả năng phát mại",
                                          "Không có khả năng phát mại")]
    sh = Sheet(ctx.sheet_name("rt05"), cols, header_row=8)
    sh.header_rows = 3
    R = 8
    hc = [
        (R, 0, R + 2, 0, "Mã khách hàng"), (R, 1, R + 2, 1, "Tên khách hàng"),
        (R, 2, R + 2, 2, "Số HĐ thế\nchấp/cầm cố"), (R, 3, R + 2, 3, "Ngày HĐ\nthế\nchấp/cầm cố"),
        (R, 4, R, 8, "Giá trị TSTC"),
        (R + 1, 4, R + 2, 4, "Tổng"), (R + 1, 5, R + 2, 5, "Bất ĐS"), (R + 1, 6, R + 2, 6, "Động sản"),
        (R + 1, 7, R + 2, 7, "GTCG"), (R + 1, 8, R + 2, 8, "TS khác"),
        (R, 9, R, 17, "Hiện trạng TSĐB"),
        (R + 1, 9, R + 1, 10, "Tính pháp lý"), (R + 1, 11, R + 1, 14, "Khả năng hoàn thiện\nhồ sơ tài sản"),
        (R + 1, 15, R + 1, 17, "Khả năng phát mại"),
    ]
    for c, t in zip(range(9, 18), ("Hợp pháp, hợp lệ", "Khác", "Đầy đủ hồ sơ theo quy định",
                                    "Đang hoàn chỉnh thủ tục", "Không thể hoàn chỉnh được", "Khác",
                                    "Có khả năng phát mại", "Ít khả năng phát mại", "Không có khả năng phát mại")):
        hc.append((R + 2, c, R + 2, c, t))
    sh.header_cells = hc

    tot_rows = []
    for t, label, sum_label in RT05_SECTIONS:
        sec = sh.add("section")
        sec.set(0, label)
        start = sh.next_row()
        items = [r for r in recs if r["type"] == t]
        i = 0
        while i < len(items):
            cust = items[i]["cust"]
            crow = sh.add("group")
            crow.set(0, cust).set(1, items[i]["name"])
            first = sh.next_row()
            n = 0
            while i < len(items) and items[i]["cust"] == cust:
                r = items[i]
                n += 1
                row = sh.add("data")
                row.set(0, ("TSBL %02d" if t == 3 else "TSTC %02d") % n)
                row.set(2, r["so_hd"]).set(3, r["ngay"])
                row.set(4, r["tong"]).set(5, r["bds"]).set(6, r["ds"]).set(7, r["gtcg"]).set(8, r["khac"])
                row.set(9, "X").set(11, "X")
                row.set(15, r["pm_co"]).set(16, r["pm_it"]).set(17, r["pm_khong"])
                i += 1
            last = sh.next_row() - 1
            for c in range(4, 9):
                crow.set(c, Sub(c, first, last))
        end = sh.next_row() - 1
        tr = sh.add("sub2")
        tr.set(0, sum_label)
        for c in range(4, 9):
            tr.set(c, Sub(c, start, end))
        for c in range(9, 18):
            tr.set(c, CountA(c, start, end))
        tot_rows.append(sh.next_row() - 1)
    total = sh.add("total")
    total.set(0, "Tổng cộng (I+II+III)")
    for c in range(4, 18):
        total.set(c, SumOf([(c, r) for r in tot_rows]))

    bank_block(sh, ctx, col="A", star=False, across_to="B")
    title_lines(sh, 5, "BÁO CÁO KIỂM KÊ HỒ SƠ, TÀI SẢN THẾ CHẤP, CẦM CỐ CỦA KHÁCH HÀNG",
                "Ngày " + ctx.date_full)
    form_block(sh, "R", ["1. Mẫu số 05/" + ctx.loai_qt, "2. Lưu tại Chi nhánh"])
    unit_line(sh, 7, "R")
    signatures(sh, ctx, [sig("A", SIG_LB), sig("D", SIG_TD), sig("H", SIG_KT), sig("O", SIG_GD)], date_col="O")
    return sh, recs


def build_rt06(ctx, data, recs=None):
    if recs is None:
        recs = _rt05_records(ctx, data)
    m = ctx.colmap("rt05")
    cols = [Col("Tên chỉ tiêu", "text", 48), Col("Giá trị TSBĐ\n(VND)", "num", 16.3)]
    cols += [Col(h, "num", 11) for h in ("Hợp pháp, hợp lệ", "Khác", "Đầy đủ hồ sơ theo quy định",
                                         "Đang hoàn chỉnh thủ tục", "Không thể hoàn chỉnh được", "Khác",
                                         "Có khả năng phát mại", "Ít có khả năng phát mại",
                                         "Không có khả năng phát mại")]
    sh = Sheet(ctx.sheet_name("rt06"), cols, header_row=8)
    sh.header_rows = 3
    R = 8
    hc = [(R, 0, R + 2, 0, "Tên chỉ tiêu"), (R, 1, R + 2, 1, "Giá trị TSBĐ\n(VND)"),
          (R, 2, R, 10, "HIỆN TRẠNG TSBĐ (SỐ LƯỢNG)"), (R + 1, 2, R + 1, 3, "Tính pháp lý"),
          (R + 1, 4, R + 1, 7, "Khả năng hoàn thiện\nhồ sơ tài sản"), (R + 1, 8, R + 1, 10, "Khả năng phát mại")]
    for c, t in zip(range(2, 11), [col.header for col in cols[2:]]):
        hc.append((R + 2, c, R + 2, c, t))
    sh.header_cells = hc

    # Cộng dồn theo loại TS: bất động sản / động sản / GTCG / khác
    acc = {}
    for row in data.rows:
        t = _rt05_type(m, row)
        kind = 3
        for k, key in enumerate(("bds", "ds", "gtcg", "ts_khac")):
            if vn.num0(g(row, m[key])) != 0:
                kind = k
        a = acc.setdefault((t, kind), [0.0] * 11)
        a[1] += vn.num0(g(row, m["tong"]))
        a[2] += 1
        a[4] += 1
        if str(g(row, m["pm_co"])).strip():
            a[8] += 1
        if str(g(row, m["pm_it"])).strip():
            a[9] += 1
        if str(g(row, m["pm_khong"])).strip():
            a[10] += 1
    names = ["Tài sản thế chấp bằng bất động sản", "Tài sản thế chấp bằng động sản",
             "Tài sản thế chấp bằng giấy tờ có giá", "Tài sản thế chấp khác"]
    sec_labels = ["I. TSTC của khách hàng là TCTD trực tiếp vay vốn:",
                  "II. TSTC của khách hàng là TCKT&CN trực tiếp vay vốn:",
                  "III. TSTC của đơn vị bảo lãnh cho bên thứ ba vay vốn"]
    tot_rows = []
    for t in (1, 2, 3):
        sh.add("section").set(0, sec_labels[t - 1])
        start = sh.next_row()
        for k in range(4):
            row = sh.add("data").set(0, names[k])
            a = acc.get((t, k))
            if a:
                for c in range(1, 11):
                    row.set(c, a[c] if c != 1 else a[1])
        end = sh.next_row() - 1
        tr = sh.add("sub2").set(0, "Cộng " + "I" * t if t < 3 else "Cộng III")
        for c in range(1, 11):
            tr.set(c, Sub(c, start, end))
        tot_rows.append(sh.next_row() - 1)
    total = sh.add("total").set(0, "TỔNG CỘNG (I + II + III)")
    for c in range(1, 11):
        total.set(c, SumOf([(c, r) for r in tot_rows]))

    bank_block(sh, ctx, col="A", star=False)
    title_lines(sh, 5, "BÁO CÁO TỔNG HỢP KIỂM KÊ HỒ SƠ, TÀI SẢN THẾ CHẤP, CẦM CỐ CỦA KHÁCH HÀNG",
                "Tại thời điểm " + ctx.date_full, size=14)
    form_block(sh, "K", ["1. Mẫu số 06/" + ctx.loai_qt, "2. CN loại I gửi báo cáo file, giấy về TSC",
                         "3. Lưu tại Chi nhánh"])
    unit_line(sh, 7, "K")
    signatures(sh, ctx, [sig("A", SIG_LB), sig("B", SIG_TD), sig("E", SIG_KT), sig("I", SIG_GD)], date_col="I")
    return sh


def run_rt05(ctx, data):
    sh, recs = build_rt05(ctx, data)
    outs = [Output("rt05", [sh])]
    if ctx.opt("kem_rt06"):
        outs.append(Output("rt06", [build_rt06(ctx, data, recs)]))
    return outs


def run_rt06(ctx, data):
    return [Output("rt06", [build_rt06(ctx, data)])]


# ===========================================================================
# MẪU 13, 14 – Sao kê tiền gửi
# ===========================================================================


def _rt1314_records(ctx, data, form):
    m = ctx.colmap(form)
    recs = []
    for row in data.rows:
        brcd = str(g(row, m["ma_cn"])).strip()
        recs.append({
            "brcd": brcd,
            "acct": str(g(row, m["tk_gl"])).strip(),
            "cust": cust_code(brcd, g(row, m["ma_kh"])),
            "name": g(row, m["ten_kh"]),
            "so_tk": str(g(row, m["so_tk"]))[1:],
            "so_so": g(row, m["so_so"]),
            "ccy": str(g(row, m["loai_tien"])).strip(),
            "nt": vn.num0(g(row, m["du_nt"])),
            "vnd": vn.num0(g(row, m["du_vnd"])),
            "ky_han": str(g(row, m["ky_han"])).replace("'", ""),
            "ngay_gui": vn.date_or_text(g(row, m["ngay_gui"])),
            "ngay_dh": vn.date_or_text(g(row, m["ngay_dh"])),
            "ls": vn.to_number(g(row, m["lai_suat"]), g(row, m["lai_suat"])),
            "ngay_tl": vn.date_or_text(g(row, m["ngay_tra_lai"])),
            "lai_tt": vn.num0(g(row, m["lai_tra_truoc"])),
            "lai_pt": vn.num0(g(row, m["lai_phai_tra"])),
            "tk_dc": str(g(row, m.get("tk_lai_du_chi", 0))).replace("'", ""),
        })
    return recs


def _rt1314_detail(sh, r, with_cust):
    row = sh.add("data")
    row.set(0, r["acct"])
    if with_cust:
        row.set(1, r["cust"]).set(2, r["name"])
    row.set(3, r["so_tk"]).set(4, r["so_so"]).set(5, r["ccy"]).set(6, r["nt"]).set(7, r["vnd"])
    row.set(8, r["ky_han"]).set(9, r["ngay_gui"]).set(10, r["ngay_dh"]).set(11, r["ls"])
    row.set(12, r["ngay_tl"]).set(13, r["lai_tt"]).set(14, r["lai_pt"])
    if sh.ncols > 15:
        row.set(15, r["tk_dc"])
    fc_fmt(row, 6, r["ccy"])
    return row


def build_rt1314(ctx, data, form):
    recs = _rt1314_records(ctx, data, form)
    is13 = form == "rt13"
    d = ctx.ddmm
    cols = [Col("Số hiệu tài khoản", "text"), Col("Mã khách hàng", "text"), Col("Tên khách hàng"),
            Col("Số tài khoản khách hàng", "text"), Col("Số sổ tiết kiệm", "text"), Col("Loại tiền tệ", "center"),
            Col("Số dư nguyên tệ tại thời điểm " + d, "num"), Col("Số dư quy đổi VNĐ tại thời điểm " + d, "num"),
            Col("Kỳ hạn (tháng)", "center"), Col("Ngày gửi", "date"), Col("Ngày đáo hạn", "date"),
            Col("Lãi suất (%/năm)", "rate"), Col("Ngày trả lãi cuối cùng", "date"),
            Col("Số lãi trả trước đến " + d, "num"), Col("Số lãi còn phải trả đến " + d, "num")]
    if is13:
        cols.append(Col("Tài khoản lãi dự chi", "text"))
    sh = Sheet(ctx.sheet_name(form), cols, header_row=9)
    _set_header_numbered(sh)
    sums_vnd = (7, 13, 14)

    if is13:
        recs.sort(key=lambda r: (r["acct"], r["ccy"], r["cust"]))
        i = 0
        while i < len(recs):
            acct = recs[i]["acct"]
            arow = sh.add("sub2").set(0, acct)
            arow.set_align(0, "left")
            block_start = sh.next_row()
            while i < len(recs) and recs[i]["acct"] == acct:
                ccy = recs[i]["ccy"]
                s = sh.next_row()
                while i < len(recs) and recs[i]["acct"] == acct and recs[i]["ccy"] == ccy:
                    _rt1314_detail(sh, recs[i], True)
                    i += 1
                e = sh.next_row() - 1
                sub = sh.add("sub1").set(0, "Cộng %s %s" % (acct, ccy))
                for c in (6,) + sums_vnd:
                    sub.set(c, Sub(c, s, e))
                fc_fmt(sub, 6, ccy)
            block_end = sh.next_row() - 1
            for c in sums_vnd:
                arow.set(c, Sub(c, block_start, block_end))
    else:
        recs.sort(key=lambda r: (r["cust"], r["ccy"], r["acct"]))
        i = 0
        while i < len(recs):
            cust = recs[i]["cust"]
            crow = sh.add("sub2").set(1, cust).set(2, recs[i]["name"])
            s = sh.next_row()
            while i < len(recs) and recs[i]["cust"] == cust:
                _rt1314_detail(sh, recs[i], False)
                i += 1
            e = sh.next_row() - 1
            for c in sums_vnd:
                crow.set(c, Sub(c, s, e))
    end = sh.next_row() - 1
    total = sh.add("total").set(0, "TỔNG CỘNG")
    for c in sums_vnd:
        total.set(c, Sub(c, sh.data_start, end))

    if is13:
        bank_block(sh, ctx, col="A", align="left", italic_name=True)
        sh.put(5, "C", "SAO KÊ CHI TIẾT TIỀN GỬI KHÁCH HÀNG", across_to="M", bold=True, size=16)
        sh.put(6, "C", "(TIỀN GỬI THANH TOÁN, TIỀN GỬI CÓ KỲ HẠN, TIỀN GỬI TIẾT KIỆM, KỲ PHIẾU, TRÁI PHIẾU)",
               across_to="M", size=12)
        sh.put(7, "C", "Ngày " + ctx.date_full, across_to="M", size=12, italic=True)
        form_block(sh, "P", ["1. Mẫu số 13/" + ctx.loai_qt, "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
        unit_line(sh, 8, "P")
        signatures(sh, ctx, [sig("B", SIG_LBANG), sig("F", SIG_KT), sig("M", SIG_GD)])
    else:
        bank_block(sh, ctx, col="B")
        title_lines(sh, 6, "SAO KÊ CHI TIẾT SỐ DƯ TIỀN GỬI, TIẾT KIỆM, KỲ PHIẾU, TRÁI PHIẾU\n"
                           "CÁC KHÁCH HÀNG CÓ SỐ DƯ TỪ 10 TỶ TRỞ LÊN", "Ngày " + ctx.date_full, sub_bold=True)
        form_block(sh, "O", ["1. Mẫu số 14/" + ctx.loai_qt, "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
        unit_line(sh, 8, "O")
        signatures(sh, ctx, [sig("B", SIG_LBANG), sig("E", SIG_KT), sig("L", SIG_GD)])
    gl = [GLItem(r["brcd"], r["acct"], r["ccy"], r["nt"], r["vnd"]) for r in recs]
    return sh, gl


def run_rt13(ctx, data):
    sh, gl = build_rt1314(ctx, data, "rt13")
    return [Output("rt13", [sh], gl)]


def run_rt14(ctx, data):
    sh, gl = build_rt1314(ctx, data, "rt14")
    return [Output("rt14", [sh], gl)]


# ===========================================================================
# MẪU 15a, 15b, 16 – Sao kê cho vay ; dữ liệu mẫu 02
# ===========================================================================


def _rt1516_cols(form, d):
    spec = [
        ("acct", "Số hiệu TK trên sổ cái", "text"), ("cust", "Mã khách hàng", "text"),
        ("name", "Tên khách hàng", "text"), ("ku", "Số khế ước/hợp đồng", "text"),
        ("ccy", "Loại tiền tệ", "center"), ("nhom", "Nhóm nợ", "center"),
        ("ma_nganh", "Mã ngành", "text"), ("ten_nganh", "Tên ngành", "text"),
        ("nguon_von", "Nguồn vốn", "text"), ("muc_dich", "Mục đích vay", "text"),
        ("nt", "Số dư nguyên tệ tại thời điểm " + d, "num"), ("vnd", "Số dư quy đổi VNĐ tại thời điểm " + d, "num"),
        ("ky_han", "Kỳ hạn (tháng)", "center"), ("ngay_gn", "Ngày giải ngân", "date"),
        ("ngay_dh", "Ngày đáo hạn", "date"), ("ls", "Lãi suất (%/năm)", "rate"),
        ("ngay_tl", "Ngày thu lãi cuối cùng", "date"), ("lai_tt", "Số lãi trả trước đến " + d, "num"),
        ("lai_dt", ("Số lãi phải thu đến hết " if form == "rt15b" else "Số lãi dự thu đến hết ") + d, "num"),
        ("tk3941", "Tài khoản lãi dự thu (TK 3941)", "text"),
    ]
    drop = set()
    if form == "rt15b":
        drop.add("lai_tt")
    if form == "rt16":
        drop.update(("ma_nganh", "ten_nganh", "tk3941"))
    return [s for s in spec if s[0] not in drop]


def build_rt1516(ctx, data, form):
    m = ctx.colmap(form)
    recs = []
    for row in data.rows:
        brcd = str(g(row, m["ma_cn"])).strip()
        recs.append({
            "brcd": brcd,
            "acct": str(g(row, m["tk_gl"])).strip(),
            "cust": cust_code(brcd, g(row, m["ma_kh"])),
            "cust_raw": str(g(row, m["ma_kh"])).strip(),
            "name": g(row, m["ten_kh"]),
            "ku": g(row, m["so_ku"]),
            "ccy": str(g(row, m["loai_tien"])).strip(),
            "nhom": str(g(row, m["nhom_no"])).replace("'", ""),
            "ma_nganh": g(row, m.get("ma_nganh", 0)),
            "ten_nganh": g(row, m.get("ten_nganh", 0)),
            "nguon_von": g(row, m.get("nguon_von", 0)),
            "muc_dich": g(row, m.get("muc_dich", 0)),
            "nt": vn.num0(g(row, m["du_nt"])),
            "vnd": vn.num0(g(row, m["du_vnd"])),
            "ky_han": str(g(row, m["ky_han"])).replace("'", ""),
            "ngay_gn": vn.date_or_text(g(row, m["ngay_gn"])),
            "ngay_dh": vn.date_or_text(g(row, m["ngay_dh"])),
            "ls": vn.to_number(g(row, m["lai_suat"]), g(row, m["lai_suat"])),
            "ngay_tl": vn.date_or_text(g(row, m["ngay_thu_lai"])),
            "lai_tt": vn.num0(g(row, m.get("lai_tra_truoc", 0))),
            "lai_dt": vn.num0(g(row, m["lai_du_thu"])),
            "tk3941": g(row, m.get("tk3941", 0)),
        })
    recs.sort(key=lambda r: (r["acct"], r["ccy"], r["nhom"], r["cust_raw"], r["ky_han"]))
    spec = _rt1516_cols(form, ctx.ddmm)
    keys = [s[0] for s in spec]
    pos = {k: i for i, k in enumerate(keys)}
    sh = Sheet(ctx.sheet_name(form), [Col(s[1], s[2]) for s in spec], header_row=9)
    _set_header_numbered(sh)
    vnd_sums = [pos[k] for k in ("vnd", "lai_tt", "lai_dt") if k in pos]
    i = 0
    while i < len(recs):
        acct = recs[i]["acct"]
        acct_start = sh.next_row()
        while i < len(recs) and recs[i]["acct"] == acct:
            ccy = recs[i]["ccy"]
            s = sh.next_row()
            while i < len(recs) and recs[i]["acct"] == acct and recs[i]["ccy"] == ccy:
                r = recs[i]
                row = sh.add("data")
                for k in keys:
                    row.set(pos[k], r[k])
                fc_fmt(row, pos["nt"], ccy)
                i += 1
            e = sh.next_row() - 1
            sub = sh.add("sub1").set(0, "Cộng %s %s" % (acct, ccy))
            for c in [pos["nt"]] + vnd_sums:
                sub.set(c, Sub(c, s, e))
            fc_fmt(sub, pos["nt"], ccy)
        acct_end = sh.next_row() - 1
        sub2 = sh.add("sub2").set(0, "Cộng %s" % acct)
        for c in vnd_sums:
            sub2.set(c, Sub(c, acct_start, acct_end))
    end = sh.next_row() - 1
    total = sh.add("total").set(0, "TỔNG CỘNG")
    for c in vnd_sums:
        total.set(c, Sub(c, sh.data_start, end))

    last = vn.col_letter(sh.ncols - 1)
    bank_block(sh, ctx, col="B")
    if form == "rt16":
        title = "SAO KÊ CHI TIẾT KHÁCH HÀNG CÓ DƯ NỢ TỪ 10 TỶ TRỞ LÊN"
        so = "16"
    elif form == "rt15a":
        title = "SAO KÊ CHI TIẾT TÀI KHOẢN CHO VAY KHÁCH HÀNG LÀ TỔ CHỨC, HỘ KINH DOANH VÀ CÁ NHÂN"
        so = "15a"
    else:
        title = ("SAO KÊ CHI TIẾT TÀI KHOẢN CHO VAY KHÁCH HÀNG LÀ TỔ CHỨC, HỘ KINH DOANH VÀ CÁ NHÂN\n"
                 "CÓ LÃI PHẢI THU HẠCH TOÁN NGOẠI BẢNG")
        so = "15b"
    title_lines(sh, 6, title, "Ngày " + ctx.date_full)
    form_block(sh, last, ["1. Mẫu số %s/%s" % (so, ctx.loai_qt), "2. CN loại I gửi file về TSC",
                          "3. Lưu tại Chi nhánh"])
    unit_line(sh, 8, last)
    n = sh.ncols
    signatures(sh, ctx, [sig(1, SIG_LB), sig(4, SIG_TD), sig(min(9, n - 5), SIG_KT), sig(n - 3, SIG_GD)])
    gl = [GLItem(r["brcd"], r["acct"], r["ccy"], r["nt"], r["vnd"]) for r in recs]
    return sh, gl


def build_rt02(ctx, data):
    """Dữ liệu phục vụ mẫu 02 (thư xác nhận dư nợ) từ file mẫu 15a."""
    m = ctx.colmap("rt02")
    personal = {"ca nhan", "ho gia dinh"}
    recs = []
    for row in data.rows:
        brcd = str(g(row, m["ma_cn"])).strip()
        loai_kh = str(g(row, m["loai_kh"])).strip()
        so_tk = str(g(row, m["so_hd"])).strip()
        if loai_kh[-2:] == "08" or loai_kh == "":
            ten_tk = "Thấu chi TKTG của khách hàng"
        else:
            ten_tk = {"1": "Nợ cho vay ngắn hạn", "2": "Nợ cho vay trung hạn",
                      "3": "Nợ cho vay dài hạn"}.get(so_tk[2:3], "")
        ccy = str(g(row, m["loai_tien"])).strip()
        nt = vn.num0(g(row, m["du_nt"]))
        vnd = vn.to_number(g(row, m["du_vnd"]))
        if vnd is None:
            vnd = nt * ctx.usd_rate if ccy.upper() == "USD" else nt
        recs.append({"makh": brcd + str(g(row, m["ma_kh"])).strip()[1:], "name": g(row, m["ten_kh"]),
                     "loai_kh": loai_kh, "so_tk": so_tk, "ten_tk": ten_tk, "nt": nt, "ccy": ccy, "vnd": vnd})
    tot = {}
    for r in recs:
        tot[r["makh"]] = tot.get(r["makh"], 0) + r["vnd"]
    keep = [r for r in recs if vn.norm_key(r["loai_kh"]) not in personal or tot[r["makh"]] >= 1e9]
    keep.sort(key=lambda r: (r["loai_kh"], r["makh"], r["so_tk"], r["ccy"]))
    agg = OrderedDict()
    for r in keep:
        k = (r["makh"], r["so_tk"], r["ccy"])
        if k in agg:
            agg[k]["nt"] += r["nt"]
        else:
            agg[k] = dict(r)
    cols = [Col("STT", "int", 6), Col("Mã KH", "text"), Col("Tên Kh"), Col("Loại KH"), Col("Số TK", "text"),
            Col("Tên TK"), Col("Số dư nguyên tệ", "num"), Col("Loại tiền", "center"), Col("Bằng chữ", "text", 60),
            Col("Địa chỉ")]
    sh = _list_sheet(ctx.sheet_name("rt02"), cols)
    for n, r in enumerate(agg.values(), 1):
        row = sh.add("data")
        for c, v in enumerate([n, r["makh"], r["name"], r["loai_kh"], r["so_tk"], r["ten_tk"], r["nt"],
                               r["ccy"], vn.number_to_words(r["nt"], r["ccy"], ctx.thousand_word), ""]):
            row.set(c, v)
        if r["nt"] != round(r["nt"]):
            row.set_fmt(6, "#,##0.00")
    return sh


def _list_sheet(name, cols):
    sh = Sheet(name, cols, header_row=1)
    sh.simple_header()
    sh.numbering = False
    return sh


def run_rt15(form):
    def run(ctx, data):
        sh, gl = build_rt1516(ctx, data, form)
        outs = [Output(form, [sh], gl)]
        if form == "rt15a" and ctx.opt("kem_rt02"):
            outs.append(Output("rt02", [build_rt02(ctx, data)]))
        return outs
    return run


def run_rt02(ctx, data):
    return [Output("rt02", [build_rt02(ctx, data)])]


# ===========================================================================
# MẪU 18 – Tài khoản 92: bảo lãnh, thư tín dụng
# ===========================================================================


def build_rt18(ctx, data):
    m = ctx.colmap("rt18")
    recs = []
    for row in data.rows:
        recs.append({
            "brcd": str(g(row, m["ma_cn"])).strip(),
            "acct": str(g(row, m["tk"])).strip(),
            "ccy": str(g(row, m["loai_tien"])).strip(),
            "cust": str(g(row, m["ma_kh"])).strip()[-9:],
            "name": g(row, m["ten_kh"]),
            "so_lc": str(g(row, m["so_lc"]))[1:],
            "ngay_ph": vn.date_or_text(g(row, m["ngay_ph"])),
            "ngay_dh": vn.date_or_text(g(row, m["ngay_dh"])),
            "loai": g(row, m["loai_bl"]),
            "nt": vn.num0(g(row, m["du_nt"])),
            "vnd": vn.num0(g(row, m["du_vnd"])),
            "kq": vn.num0(g(row, m["ky_quy"])),
            "hd_hm": g(row, m["so_hd_hm"]),
            "hm": vn.to_number(g(row, m["tien_hm"]), g(row, m["tien_hm"])),
        })
    recs.sort(key=lambda r: (r["brcd"], r["acct"], r["ccy"], r["cust"]))
    cols = [Col("Tài khoản", "text"), Col("Loại tiền tệ", "center"), Col("Mã khách hàng", "text"),
            Col("Tên khách hàng", "text", 0), Col("Số LC, hợp đồng bảo lãnh", "text"), Col("Ngày bắt đầu", "date"),
            Col("Ngày hết hiệu lực", "date"), Col("Loại bảo lãnh, LC", "center"), Col("Nguyên tệ", "num"),
            Col("Quy đổi", "num"), Col("Số tiền ký quỹ trên tài khoản 427 (quy đổi)", "num"),
            Col("Số hợp đồng hạn mức", "text"), Col("Số tiền\n(Quy đổi)", "num", 11.5)]
    sh = Sheet(ctx.sheet_name("rt18"), cols, header_row=9)
    sh.header_rows = 2
    R = 9
    hc = [(R, c, R + 1, c, cols[c].header) for c in range(8)]
    hc += [(R, 8, R, 9, "Số dư của LC, bảo lãnh trên tài khoản ngoại bảng ngày " + ctx.ddmm),
           (R + 1, 8, R + 1, 8, "Nguyên tệ"), (R + 1, 9, R + 1, 9, "Quy đổi"),
           (R, 10, R + 1, 10, cols[10].header),
           (R, 11, R, 12, "Hạn mức tín dụng cho khách hàng (nếu có)"),
           (R + 1, 11, R + 1, 11, "Số hợp đồng hạn mức"), (R + 1, 12, R + 1, 12, "Số tiền\n(Quy đổi)")]
    sh.header_cells = hc
    i = 0
    while i < len(recs):
        acct = recs[i]["acct"]
        arow = sh.add("sub2").set(0, "Cộng " + acct)
        a_start = sh.next_row()
        while i < len(recs) and recs[i]["acct"] == acct:
            ccy = recs[i]["ccy"]
            crow = sh.add("sub1").set(0, "Cộng %s %s" % (acct, ccy))
            s = sh.next_row()
            while i < len(recs) and recs[i]["acct"] == acct and recs[i]["ccy"] == ccy:
                r = recs[i]
                row = sh.add("data")
                for c, v in enumerate([r["acct"], r["ccy"], r["cust"], r["name"], r["so_lc"], r["ngay_ph"],
                                       r["ngay_dh"], r["loai"], r["nt"], r["vnd"], r["kq"], r["hd_hm"], r["hm"]]):
                    row.set(c, v)
                fc_fmt(row, 8, ccy)
                fc_fmt(row, 12, ccy)
                i += 1
            e = sh.next_row() - 1
            for c in (8, 9, 10, 12):
                crow.set(c, Sub(c, s, e))
            fc_fmt(crow, 8, ccy)
        a_end = sh.next_row() - 1
        for c in (9, 10):
            arow.set(c, Sub(c, a_start, a_end))
    end = sh.next_row() - 1
    total = sh.add("total").set(0, "CỘNG")
    for c in (9, 10):
        total.set(c, Sub(c, sh.data_start, end))
    if sh.rows:
        sh.col_widths[3] = max(27.5, min(48, max((len(str(r.cells[3] or "")) for r in sh.rows), default=0) + 2))
    sh.row_heights[9] = 62

    bank_block(sh, ctx, col="C")
    title_lines(sh, 6, "SAO KÊ CHI TIẾT TÀI KHOẢN 92: BẢO LÃNH VÀ THƯ TÍN DỤNG CHO KHÁCH HÀNG",
                "Ngày " + ctx.date_full)
    form_block(sh, "M", ["1. Mẫu số 18/" + ctx.loai_qt, "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(sh, 8, "M", "ĐVT: VND")
    signatures(sh, ctx, [sig("C", SIG_LB), sig("E", SIG_TD), sig("I", SIG_TTQT), sig("L", SIG_GD)])
    gl = [GLItem(r["brcd"], r["acct"], r["ccy"], r["nt"], r["vnd"]) for r in recs]
    return [Output("rt18", [sh], gl)]


# ===========================================================================
# MẪU 20a – Nợ được xử lý bằng nguồn dự phòng
# ===========================================================================


def build_rt20a(ctx, data):
    m = ctx.colmap("rt20a")
    d = ctx.date_short
    cols = [Col("STT", "int", 6), Col("Mã chi nhánh", "center"), Col("Tên chi nhánh"), Col("Mã khách hàng", "text"),
            Col("Tên khách hàng"), Col("Số hợp đồng tín dụng", "text"), Col("Số khế ước", "text"),
            Col("Lãi suất", "gen"), Col("Tài khoản (9711xx)", "text"), Col("Số tiền (Quy đổi VND)", "num"),
            Col("Tài khoản (9712xx)", "text"), Col("Số tiền (Quy đổi VND)", "num"), Col("Loại tài sản đảm bảo"),
            Col("Giá trị TSĐB", "num"), Col("Hiện trạng của tài sản đảm bảo"), Col("Giá trị thu hồi ước tính", "num")]
    sh = Sheet(ctx.sheet_name("rt20a"), cols, header_row=9)
    sh.header_rows = 2
    R = 9
    hc = [(R, c, R + 1, c, cols[c].header) for c in range(8)]
    hc += [(R, 8, R, 9, "Dư nợ gốc đã xử lý rủi ro tại ngày " + d), (R + 1, 8, R + 1, 8, "Tài khoản (9711xx)"),
           (R + 1, 9, R + 1, 9, "Số tiền (Quy đổi VND)"),
           (R, 10, R, 11, "Dư nợ lãi đã xử lý rủi ro tại ngày " + d), (R + 1, 10, R + 1, 10, "Tài khoản (9712xx)"),
           (R + 1, 11, R + 1, 11, "Số tiền (Quy đổi VND)"), (R, 12, R + 1, 12, "Loại tài sản đảm bảo"),
           (R, 13, R + 1, 13, "Giá trị tài sản đảm bảo tại ngày %s\n(hạch toán ngoại bảng)" % d),
           (R, 14, R + 1, 14, "Hiện trạng của tài sản đảm bảo"), (R, 15, R + 1, 15, "Giá trị thu hồi ước tính")]
    sh.header_cells = hc
    n = 0
    for row in data.rows:
        if vn.num0(g(row, m["loc_1"])) + vn.num0(g(row, m["loc_2"])) <= 0:
            continue
        n += 1
        r = sh.add("data")
        for c, v in enumerate([n, ctx.brcd, ctx.brname, g(row, m["ma_kh"]), g(row, m["ten_kh"]),
                               g(row, m["so_hd"]), g(row, m["so_ku"]),
                               vn.to_number(g(row, m["lai_suat"]), g(row, m["lai_suat"])), "",
                               vn.num0(g(row, m["goc"])), "", vn.num0(g(row, m["lai"]))]):
            r.set(c, v)
    end = sh.next_row() - 1
    total = sh.add("total").set(0, "CỘNG")
    for c in (9, 11):
        total.set(c, Sub(c, sh.data_start, end))

    bank_block(sh, ctx, col="C", italic_name=True)
    title_lines(sh, 6, "BÁO CÁO NỢ ĐƯỢC XỬ LÝ BẰNG NGUỒN DỰ PHÒNG", "ĐẾN " + ctx.date_full.upper())
    form_block(sh, "P", ["1. Mẫu số 20a/" + ctx.loai_qt, "2. NHNo loại 1, 2 gửi file về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(sh, 8, "P")
    signatures(sh, ctx, [sig("C", SIG_LB), sig("E", SIG_TD_S), sig("G", SIG_KT_S), sig("M", SIG_GD)])
    return [Output("rt20a", [sh])]


# ===========================================================================
# MẪU 22 – Doanh thu, chi phí chờ phân bổ
# ===========================================================================


def build_rt22(ctx, data):
    m = ctx.colmap("rt22")
    recs = []
    for row in data.rows:
        tk_tn = str(g(row, m["tk_tn"])).strip()
        sh_ = 1 if tk_tn[:1] == "7" else 0   # file doanh thu (TK 7xxx) có thêm 1 cột phía trước
        brcd = str(g(row, m["ma_cn"])).strip()
        recs.append({
            "brcd": brcd,
            "cust": cust_code(brcd, g(row, m["ma_kh"])),
            "name": g(row, m["ten_kh"]),
            "acct": str(g(row, m["tk_cho_pb"])).strip(),
            "tk_tn": tk_tn,
            "ref": g(row, m["so_tc"]),
            "ccy": str(g(row, m["loai_tien"])).strip(),
            "ngay_mo": vn.date_or_text(g(row, m["ngay_mo"])),
            "ngay_dh": vn.date_or_text(g(row, m["ngay_dh"])),
            "ngay_pb": vn.date_or_text(g(row, m["ngay_pb"] + sh_)),
            "ls": vn.to_number(g(row, m["lai_suat"] + sh_), g(row, m["lai_suat"] + sh_)),
            "goc": vn.num0(g(row, m["goc"] + sh_)),
            "con_nt": vn.num0(g(row, m["con_pb_nt"] + sh_)),
            "con_vnd": vn.num0(g(row, m["con_pb_vnd"] + sh_)),
        })
    cols = [Col("Mã khách hàng", "text"), Col("Tên khách hàng"),
            Col("Tài khoản doanh thu (hoặc chi phí chờ phân bổ)", "text"), Col("TK thu nhập (hoặc chi phí)", "text"),
            Col("Số tham chiếu", "text"), Col("Loại tiền", "center"), Col("Ngày mở", "date"),
            Col("Ngày đến hạn", "date"), Col("Ngày phân bổ", "date"), Col("Lãi suất", "num2"),
            Col("Số tiền gốc", "num"), Col("Số tiền còn phải phân bổ cho kỳ sau (nguyên tệ)", "num"),
            Col("Số tiền còn phải phân bổ cho kỳ sau (Quy đổi VND)", "num")]
    sh = Sheet(ctx.sheet_name("rt22"), cols, header_row=9)
    _set_header_numbered(sh)
    groups = OrderedDict()
    for r in recs:
        groups.setdefault("3" if r["acct"][:1] == "3" else "4", []).append(r)
    for letter, (key, items) in zip("ABCDEF", groups.items()):
        sec = sh.add("section").set(0, letter).set(1, "CHI PHÍ CHỜ PHÂN BỔ" if key == "3" else "DOANH THU CHỜ PHÂN BỔ")
        sec.set_align(0, "center")
        s = sh.next_row()
        for r in items:
            row = sh.add("data")
            for c, v in enumerate([r["cust"], r["name"], r["acct"], r["tk_tn"], r["ref"], r["ccy"], r["ngay_mo"],
                                   r["ngay_dh"], r["ngay_pb"], r["ls"], r["goc"], r["con_nt"], r["con_vnd"]]):
                row.set(c, v)
            fc_fmt(row, 11, r["ccy"])
        e = sh.next_row() - 1
        tot = sh.add("sub2").set(1, "CỘNG")
        tot.set_align(1, "center")
        for c in (10, 11, 12):
            tot.set(c, Sub(c, s, e))

    bank_block(sh, ctx, col="A", across_to="B")
    title_lines(sh, 6, "SAO KÊ CHI TIẾT SỐ DƯ TÀI KHOẢN DOANH THU VÀ CHI PHÍ CHỜ PHÂN BỔ", "Ngày " + ctx.date_full,
                sub_bold=True)
    form_block(sh, "M", ["1. Mẫu số 22/" + ctx.loai_qt, "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(sh, 8, "M")
    signatures(sh, ctx, [sig("B", SIG_LB), sig("E", SIG_KT), sig("K", SIG_GD)])
    gl = [GLItem(r["brcd"], r["acct"], r["ccy"], r["con_nt"], r["con_vnd"]) for r in recs]
    return [Output("rt22", [sh], gl)]


# ===========================================================================
# MẪU 24a – Phải thu, phải trả
# ===========================================================================


def build_rt24(ctx, data):
    m = ctx.colmap("rt24")
    recs = []
    for row in data.rows:
        recs.append({
            "brcd": str(g(row, m["ma_cn"])).strip(),
            "acct": str(g(row, m["tk"])).strip(),
            "ten_tk": g(row, m["ten_tk"]),
            "ma_dt": str(g(row, m["ma_dt"])).strip()[-9:],
            "ten_dt": g(row, m["ten_dt"]),
            "ly_do": g(row, m["ly_do"]),
            "ngay": vn.date_or_text("".join(str(g(row, m["ngay_ps"])).split())),
            "ccy": str(g(row, m["loai_tien"])).strip(),
            "nt": vn.num0(g(row, m["du_nt"])),
            "vnd": vn.num0(g(row, m["du_vnd"])),
        })
    recs.sort(key=lambda r: (r["brcd"], r["acct"], r["ccy"], r["ma_dt"]))
    cols = [Col("Tài khoản", "text"), Col("Tên tài khoản"), Col("Tên đối tượng phải thu/phải trả"),
            Col("Lý do hạch toán"), Col("Ngày/Tháng/\nNăm phát sinh", "date", 16), Col("Loại tiền", "center"),
            Col("Số dư tại ngày %s\n(nguyên tệ)" % ctx.ddmm, "num"),
            Col("Số dư tại ngày %s\n(quy đổi VNĐ)" % ctx.ddmm, "num"),
            Col("Dự phòng đã trích\nđến 31/12", "num"), Col("Ghi chú")]
    sh = Sheet(ctx.sheet_name("rt24"), cols, header_row=9)
    _set_header_numbered(sh)
    sh.add("section").set(0, "I. CÁC KHOẢN PHẢI THU")
    payable = False
    i = 0
    while i < len(recs):
        acct = recs[i]["acct"]
        if acct[:1] == "4" and not payable:
            sh.add("section").set(0, "II. CÁC KHOẢN PHẢI TRẢ")
            payable = True
        a_start = sh.next_row()
        while i < len(recs) and recs[i]["acct"] == acct:
            ccy = recs[i]["ccy"]
            s = sh.next_row()
            while i < len(recs) and recs[i]["acct"] == acct and recs[i]["ccy"] == ccy:
                r = recs[i]
                row = sh.add("data")
                for c, v in enumerate([r["acct"], r["ten_tk"], r["ten_dt"], r["ly_do"], r["ngay"], r["ccy"],
                                       r["nt"], r["vnd"]]):
                    row.set(c, v)
                fc_fmt(row, 6, ccy)
                i += 1
            e = sh.next_row() - 1
            sub = sh.add("sub1").set(0, "Cộng %s %s" % (acct, ccy))
            sub.set(6, Sub(6, s, e)).set(7, Sub(7, s, e))
            fc_fmt(sub, 6, ccy)
        a_end = sh.next_row() - 1
        sh.add("sub2").set(0, "Cộng " + acct).set(7, Sub(7, a_start, a_end))
    end = sh.next_row() - 1
    sh.add("total").set(0, "CỘNG").set(7, Sub(7, sh.data_start, end))

    bank_block(sh, ctx, col="A", across_to="B")
    title_lines(sh, 6, "SAO KÊ CHI TIẾT TÀI KHOẢN PHẢI THU, PHẢI TRẢ", "Ngày " + ctx.date_full)
    form_block(sh, "J", ["1. Mẫu số 24a/" + ctx.loai_qt, "2. CN loại I gửi file, giấy về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(sh, 8, "J")
    sh.row_heights[8] = 25
    signatures(sh, ctx, [sig("A", SIG_LB), sig("C", SIG_KT_S), sig("G", SIG_GD)])
    gl = [GLItem(r["brcd"], r["acct"], r["ccy"], r["nt"], r["vnd"]) for r in recs]
    return [Output("rt24", [sh], gl)]


# ===========================================================================
# DP – Danh sách xác nhận số dư tiền gửi tổ chức (MSIT81)
# ===========================================================================


def build_rtdp(ctx, data):
    m = ctx.colmap("rtdp")
    personal = {"ca nhan", "ho gia dinh"}
    recs = []
    for row in data.rows:
        if vn.norm_key(g(row, m["loai_kh"])) in personal:
            continue
        if vn.num0(g(row, m["loc_bang_0"])) != 0:
            continue
        ccy = str(g(row, m["loai_tien"])).strip()
        so_du = vn.num0(g(row, m["so_du"]))
        recs.append([str(g(row, m["ma_cn"])).strip() + str(g(row, m["ma_kh"])).strip(), g(row, m["ten_kh"]),
                     g(row, m["ma_sp"]), g(row, m["so_tk"]), g(row, m["ten_kh"]), so_du, ccy,
                     vn.number_to_words(so_du, ccy, ctx.thousand_word), g(row, m["dia_chi"])])
    recs.sort(key=lambda r: r[0])
    cols = [Col("STT", "int", 6), Col("Mã KH", "text"), Col("Tên Kh"), Col("Mã SP", "text"), Col("Số TK", "text"),
            Col("Tên TK"), Col("Số dư nguyên tệ", "num"), Col("Loại tiền", "center"), Col("Bằng chữ", "text", 60),
            Col("Địa chỉ")]
    sh = _list_sheet(ctx.sheet_name("rtdp"), cols)
    for n, r in enumerate(recs, 1):
        row = sh.add("data")
        row.set(0, n)
        for c, v in enumerate(r, 1):
            row.set(c, v)
        if r[5] != round(r[5]):
            row.set_fmt(6, "#,##0.00")
    return [Output("rtdp", [sh])]


# ===========================================================================
# HTT – Sao kê chi tiết số dư tài khoản dự thu cuối kỳ
# ===========================================================================


def build_rttt(ctx, data):
    m = ctx.colmap("rttt")
    cols = [Col("Mã khách hàng", "text"), Col("Tên khách hàng"), Col("Tài khoản dự thu", "text"),
            Col("Loại tiền", "center"), Col("Ngày mở", "date"), Col("Ngày đến hạn", "date"),
            Col("Ngày trả lãi cuối cùng", "date"), Col("Ngày tính dự thu", "date"), Col("Lãi suất", "num2"),
            Col("Số tiền gốc", "num"), Col("Số lãi cuối kỳ (nguyên tệ)", "num"), Col("Tỷ giá", "num2"),
            Col("Số lãi cuối kỳ (VND)", "num"), Col("Số lãi kỳ trước (VND)", "num")]
    sh = Sheet(ctx.sheet_name("rttt"), cols, header_row=8)
    _set_header_numbered(sh)
    for row in data.rows:
        brcd = str(g(row, m["ma_cn"])).strip()
        ccy = str(g(row, m["loai_tien"])).strip()
        r = sh.add("data")
        vals = [cust_code(brcd, g(row, m["ma_kh"])), g(row, m["ten_kh"]), g(row, m["tk_du_thu"]), ccy,
                vn.date_or_text(g(row, m["ngay_mo"])), vn.date_or_text(g(row, m["ngay_dh"])),
                vn.date_or_text(g(row, m["ngay_tra_lai"])), vn.date_or_text(g(row, m["ngay_du_thu"])),
                vn.to_number(g(row, m["lai_suat"]), g(row, m["lai_suat"])), vn.num0(g(row, m["goc"])),
                vn.num0(g(row, m["lai_nt"])), vn.to_number(g(row, m["ty_gia"]), g(row, m["ty_gia"])),
                vn.num0(g(row, m["lai_vnd"])), vn.num0(g(row, m["lai_ky_truoc"]))]
        for c, v in enumerate(vals):
            r.set(c, v)
        fc_fmt(r, 9, ccy)
        fc_fmt(r, 10, ccy)
    end = sh.next_row() - 1
    total = sh.add("total").set(0, "Tổng cộng")
    for c in (12, 13):
        total.set(c, Sub(c, sh.data_start, end))

    for i, t in enumerate(["NGÂN HÀNG NÔNG NGHIỆP", "VÀ PHÁT TRIỂN NÔNG THÔN VIỆT NAM", "******************"]):
        sh.put(i + 1, "A", t, bold=True, align="left")
    sh.put(4, "A", "Chi nhánh: " + ctx.brname, bold=True, italic=True, align="left")
    sh.put(5, "B", "SAO KÊ CHI TIẾT SỐ DƯ TÀI KHOẢN DỰ THU CUỐI KỲ", across_to="L", bold=True, size=16)
    sh.put(6, "B", "Ngày " + ctx.date_short, across_to="L", italic=True, size=12)
    form_block(sh, "N", ["1. Mẫu số HTT", "2. CN loại I gửi file về TSC", "3. Lưu tại Chi nhánh"])
    unit_line(sh, 7, "N")
    signatures(sh, ctx, [sig("B", SIG_LB), sig("F", SIG_KT), sig("L", SIG_GD)])
    return [Output("rttt", [sh])]


# ===========================================================================
# Danh mục mẫu và bảng ánh xạ cột mặc định
# ===========================================================================

# (khoá, số cột mặc định, mô tả)
_COMMON_1516 = [
    ("ma_cn", 2, "Mã chi nhánh"), ("nhom_no", 4, "Nhóm nợ"), ("tk_gl", 5, "Số hiệu TK sổ cái"),
    ("ma_kh", 6, "Mã khách hàng"), ("ten_kh", 8, "Tên khách hàng"), ("so_ku", 13, "Số khế ước (số giải ngân)"),
    ("loai_tien", 14, "Loại tiền"), ("du_nt", 16, "Số dư nguyên tệ"), ("du_vnd", 17, "Số dư quy đổi VND"),
    ("ky_han", 18, "Kỳ hạn (tháng)"), ("ngay_gn", 19, "Ngày giải ngân"), ("ngay_dh", 20, "Ngày đáo hạn"),
    ("lai_suat", 21, "Lãi suất"), ("ngay_thu_lai", 24, "Ngày thu lãi cuối cùng"),
]

FORMS = OrderedDict([
    ("rt05", {"name": "Mẫu 05 – Kiểm kê hồ sơ, tài sản thế chấp, cầm cố", "run": run_rt05, "min_cols": 18,
              "fields": [("ma_cn", 2, "Mã chi nhánh"), ("ma_kh", 3, "Mã khách hàng"), ("ten_kh", 4, "Tên khách hàng"),
                         ("loai_tc", 5, "Loại thế chấp/bảo lãnh (\"The chap\")"),
                         ("so_hd", 6, "Số HĐ thế chấp/bảo lãnh"), ("ngay_hd", 7, "Ngày HĐ (yyyymmdd)"),
                         ("tong", 10, "Số tiền quy đổi VND"), ("bds", 11, "Bất động sản"), ("ds", 12, "Động sản"),
                         ("gtcg", 13, "Chứng từ có giá"), ("ts_khac", 14, "Tài sản khác"),
                         ("pm_co", 15, "Có khả năng phát mại"), ("pm_it", 16, "Ít có khả năng phát mại"),
                         ("pm_khong", 17, "Không có khả năng phát mại"),
                         ("ma_loai_kh", 21, "Mã loại KH (600 = TCTD)")]}),
    ("rt06", {"name": "Mẫu 06 – Tổng hợp kiểm kê TSTC (dùng file mẫu 05)", "run": run_rt06, "min_cols": 18,
              "fields_from": "rt05"}),
    ("rt13", {"name": "Mẫu 13 – Sao kê chi tiết tiền gửi khách hàng", "run": run_rt13, "min_cols": 17,
              "fields": [("ma_cn", 2, "Mã chi nhánh"), ("tk_gl", 3, "Số hiệu TK sổ cái"), ("ma_kh", 4, "Mã khách hàng"),
                         ("ten_kh", 5, "Tên khách hàng"), ("so_tk", 6, "Số TK tiền gửi"), ("so_so", 7, "Số sổ"),
                         ("loai_tien", 8, "Loại tiền"), ("du_nt", 9, "Số dư nguyên tệ"),
                         ("du_vnd", 10, "Số dư quy đổi VND"), ("ky_han", 11, "Kỳ hạn"), ("ngay_gui", 12, "Ngày gửi"),
                         ("ngay_dh", 13, "Ngày đến hạn"), ("lai_suat", 14, "Lãi suất"),
                         ("ngay_tra_lai", 15, "Ngày trả lãi cuối cùng"), ("lai_tra_truoc", 16, "Số lãi trả trước"),
                         ("lai_phai_tra", 17, "Số lãi còn phải trả"), ("tk_lai_du_chi", 21, "TK lãi dự chi")]}),
    ("rt14", {"name": "Mẫu 14 – Sao kê KH có số dư tiền gửi từ 10 tỷ", "run": run_rt14, "min_cols": 17,
              "fields_from": "rt13"}),
    ("rt15a", {"name": "Mẫu 15a – Sao kê chi tiết tài khoản cho vay", "run": run_rt15("rt15a"), "min_cols": 41,
               "fields": _COMMON_1516 + [("lai_tra_truoc", 25, "Số lãi trả trước"), ("lai_du_thu", 26, "Số lãi dự thu"),
                                         ("ma_nganh", 35, "Mã ngành"), ("ten_nganh", 36, "Tên ngành"),
                                         ("nguon_von", 38, "Nguồn vốn"), ("muc_dich", 36, "Mục đích vay"),
                                         ("tk3941", 43, "TK lãi dự thu 3941")]}),
    ("rt15b", {"name": "Mẫu 15b – Cho vay có lãi phải thu ngoại bảng", "run": run_rt15("rt15b"), "min_cols": 41,
               "fields": _COMMON_1516 + [("lai_du_thu", 25, "Số lãi phải thu"), ("ma_nganh", 34, "Mã ngành"),
                                         ("nguon_von", 37, "Nguồn vốn"), ("muc_dich", 35, "Mục đích vay"),
                                         ("tk3941", 42, "TK lãi dự thu 3941")]}),
    ("rt16", {"name": "Mẫu 16 – Khách hàng có dư nợ từ 10 tỷ trở lên", "run": run_rt15("rt16"), "min_cols": 41,
              "fields": _COMMON_1516 + [("lai_tra_truoc", 25, "Số lãi trả trước"), ("lai_du_thu", 26, "Số lãi dự thu"),
                                        ("nguon_von", 37, "Nguồn vốn"), ("muc_dich", 35, "Mục đích vay")]}),
    ("rt02", {"name": "Dữ liệu mẫu 02 – Xác nhận dư nợ (dùng file mẫu 15a)", "run": run_rt02, "min_cols": 41,
              "fields": [("ma_cn", 2, "Mã chi nhánh"), ("ma_kh", 6, "Mã khách hàng"), ("loai_kh", 7, "Loại khách hàng"),
                         ("ten_kh", 8, "Tên khách hàng"), ("so_hd", 9, "Số khế ước/hợp đồng"),
                         ("loai_tien", 14, "Loại tiền"), ("du_nt", 16, "Số dư nguyên tệ"),
                         ("du_vnd", 17, "Số dư quy đổi VND")]}),
    ("rt18", {"name": "Mẫu 18 – Sao kê TK 92: bảo lãnh, thư tín dụng", "run": build_rt18, "min_cols": 16,
              "fields": [("ma_cn", 2, "Mã chi nhánh"), ("tk", 4, "Tài khoản"), ("loai_tien", 5, "Loại tiền LC-BL"),
                         ("ma_kh", 7, "Mã khách hàng"), ("ten_kh", 8, "Tên khách hàng"), ("so_lc", 9, "Số LC-BL"),
                         ("ngay_ph", 10, "Ngày phát hành"), ("ngay_dh", 11, "Ngày đến hạn"),
                         ("loai_bl", 12, "Loại LC-BL"), ("du_nt", 13, "Số dư"), ("du_vnd", 14, "Số dư quy đổi VND"),
                         ("ky_quy", 16, "Số dư ký quỹ"), ("so_hd_hm", 17, "Số phê duyệt (HĐ hạn mức)"),
                         ("tien_hm", 18, "Tiền phê duyệt")]}),
    ("rt20a", {"name": "Mẫu 20a – Nợ được xử lý bằng nguồn dự phòng", "run": build_rt20a, "min_cols": 8,
               "fields": [("ma_kh", 2, "Mã khách hàng"), ("ten_kh", 3, "Tên khách hàng"), ("so_ku", 4, "Số khế ước"),
                          ("so_hd", 5, "Số hợp đồng tín dụng"), ("goc", 12, "Dư nợ gốc đã XLRR"),
                          ("lai", 13, "Dư nợ lãi đã XLRR"), ("loc_1", 24, "Cột lọc 1 (lấy dòng có cột 1 + cột 2 > 0)"),
                          ("loc_2", 25, "Cột lọc 2"), ("lai_suat", 38, "Lãi suất")]}),
    ("rt22", {"name": "Mẫu 22 – Doanh thu, chi phí chờ phân bổ", "run": build_rt22, "min_cols": 20,
              "fields": [("ma_cn", 1, "Mã chi nhánh"), ("ma_kh", 3, "Mã khách hàng"), ("ten_kh", 4, "Tên khách hàng"),
                         ("tk_tn", 5, "TK thu nhập/chi phí"), ("tk_cho_pb", 7, "TK doanh thu/chi phí chờ phân bổ"),
                         ("so_tc", 9, "Số tham chiếu"), ("loai_tien", 10, "Loại tiền"), ("ngay_mo", 11, "Ngày mở"),
                         ("ngay_dh", 12, "Ngày đến hạn"),
                         ("ngay_pb", 13, "Ngày phân bổ (TK 7xxx tự cộng thêm 1 cột)"), ("lai_suat", 14, "Lãi suất"),
                         ("goc", 15, "Số tiền gốc"), ("con_pb_nt", 16, "Còn phải phân bổ (nguyên tệ)"),
                         ("con_pb_vnd", 18, "Còn phải phân bổ (quy đổi VND)")]}),
    ("rt24", {"name": "Mẫu 24a – Sao kê chi tiết tài khoản phải thu, phải trả", "run": build_rt24, "min_cols": 11,
              "fields": [("ma_cn", 2, "Mã chi nhánh"), ("tk", 3, "Tài khoản"), ("ten_tk", 4, "Tên tài khoản"),
                         ("ma_dt", 5, "Mã đối tượng"), ("ten_dt", 6, "Tên đối tượng"), ("ly_do", 7, "Lý do hạch toán"),
                         ("ngay_ps", 8, "Năm/ngày phát sinh"), ("loai_tien", 9, "Loại tiền"),
                         ("du_nt", 10, "Số dư nguyên tệ"), ("du_vnd", 11, "Số dư quy đổi VND")]}),
    ("rtdp", {"name": "DP – Danh sách xác nhận số dư tiền gửi tổ chức (MSIT81)", "run": build_rtdp, "min_cols": 45,
              "fields": [("ma_cn", 1, "Mã chi nhánh (ghép trước mã KH)"), ("ma_kh", 3, "Mã khách hàng"),
                         ("ten_kh", 4, "Tên khách hàng"), ("loai_tien", 6, "Loại tiền"), ("so_du", 7, "Số dư"),
                         ("so_tk", 9, "Số tài khoản"), ("dia_chi", 12, "Địa chỉ"),
                         ("loc_bang_0", 17, "Cột lọc (chỉ lấy dòng = 0)"), ("ma_sp", 19, "Mã sản phẩm"),
                         ("loai_kh", 22, "Loại khách hàng")]}),
    ("rttt", {"name": "HTT – Sao kê chi tiết số dư tài khoản dự thu cuối kỳ", "run": build_rttt, "min_cols": 20,
              "fields": [("ma_cn", 1, "Mã chi nhánh"), ("ma_kh", 3, "Mã khách hàng"), ("ten_kh", 4, "Tên khách hàng"),
                         ("tk_du_thu", 7, "Tài khoản dự thu"), ("loai_tien", 10, "Loại tiền"),
                         ("ngay_mo", 11, "Ngày mở"), ("ngay_dh", 12, "Ngày đến hạn"),
                         ("ngay_tra_lai", 13, "Ngày trả lãi cuối cùng"), ("ngay_du_thu", 14, "Ngày tính dự thu"),
                         ("lai_suat", 15, "Lãi suất"), ("goc", 16, "Số tiền gốc"), ("lai_nt", 17, "Lãi cuối kỳ (NT)"),
                         ("ty_gia", 18, "Tỷ giá"), ("lai_vnd", 19, "Lãi cuối kỳ (VND)"),
                         ("lai_ky_truoc", 20, "Lãi kỳ trước (VND)")]}),
])

# Tên file khác nhau nhưng cùng mẫu
ALIASES = {"rt24a": "rt24", "rt20": "rt20a", "rt15": "rt15a", "rtht": "rttt", "htt": "rttt", "rt22a": "rt22"}

GL_FIELDS = [("loai_tien", 1, "Loại tiền (000 = quy đổi)"), ("tk", 2, "Tài khoản"),
             ("du_no", 11, "Dư nợ cuối kỳ"), ("du_co", 12, "Dư có cuối kỳ")]


def resolve(code: str) -> str:
    code = (code or "").lower()
    return ALIASES.get(code, code)


def field_defs(code: str):
    spec = FORMS.get(code)
    if not spec:
        return []
    if "fields_from" in spec:
        return FORMS[spec["fields_from"]]["fields"]
    return spec["fields"]


def default_colmap(code: str):
    if code == "gl":
        return {k: v for k, v, _ in GL_FIELDS}
    return {k: v for k, v, _ in field_defs(code)}
