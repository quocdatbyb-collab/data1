"""Mô hình bảng tính và bộ ghi Excel (XlsxWriter) dùng chung cho mọi mẫu biểu."""

from __future__ import annotations

import datetime as _dt

import xlsxwriter

from . import AUTHOR, APP_NAME
from . import vn

# ---------------------------------------------------------------------------
# Công thức (giữ công thức sống trong Excel, đồng thời tính sẵn giá trị)
# ---------------------------------------------------------------------------


class Formula:
    def text(self) -> str:
        raise NotImplementedError

    def value(self, sheet: "Sheet"):
        raise NotImplementedError


class Sub(Formula):
    """=SUBTOTAL(9, X r1:X r2) – tự bỏ qua các dòng cộng con lồng bên trong."""

    def __init__(self, col: int, r1: int, r2: int):
        self.col, self.r1, self.r2 = col, r1, r2

    def text(self):
        c = vn.col_letter(self.col)
        if self.r2 < self.r1:
            return "=0"
        return "=SUBTOTAL(9,%s%d:%s%d)" % (c, self.r1, c, self.r2)

    def value(self, sheet):
        total = 0
        for r in range(self.r1, self.r2 + 1):
            v = sheet.get(r, self.col)
            if isinstance(v, Sub):
                continue
            if isinstance(v, Formula):
                v = sheet.eval(v)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                total += v
        return total


class CountA(Formula):
    def __init__(self, col: int, r1: int, r2: int):
        self.col, self.r1, self.r2 = col, r1, r2

    def text(self):
        c = vn.col_letter(self.col)
        if self.r2 < self.r1:
            return "=0"
        return "=COUNTA(%s%d:%s%d)" % (c, self.r1, c, self.r2)

    def value(self, sheet):
        n = 0
        for r in range(self.r1, self.r2 + 1):
            v = sheet.get(r, self.col)
            if v not in (None, ""):
                n += 1
        return n


class SumOf(Formula):
    """=SUM(E20,E40,E60)"""

    def __init__(self, refs):
        self.refs = list(refs)  # [(col, row)]

    def text(self):
        if not self.refs:
            return "=0"
        return "=SUM(%s)" % ",".join("%s%d" % (vn.col_letter(c), r) for c, r in self.refs)

    def value(self, sheet):
        total = 0
        for c, r in self.refs:
            v = sheet.eval(sheet.get(r, c))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                total += v
        return total


class Diff(Formula):
    """=A - B"""

    def __init__(self, a, b):
        self.a, self.b = a, b  # (col,row)

    def text(self):
        return "=%s%d-%s%d" % (vn.col_letter(self.a[0]), self.a[1], vn.col_letter(self.b[0]), self.b[1])

    def value(self, sheet):
        x = sheet.eval(sheet.get(self.a[1], self.a[0]))
        y = sheet.eval(sheet.get(self.b[1], self.b[0]))
        x = x if isinstance(x, (int, float)) else 0
        y = y if isinstance(y, (int, float)) else 0
        return x - y


# ---------------------------------------------------------------------------
# Mô hình
# ---------------------------------------------------------------------------

NUM_FORMATS = {
    "text": "@",
    "num": "#,##0",
    "num2": "#,##0.00",
    "rate": "#,##0.00",
    "date": "dd/mm/yyyy",
    "int": "0",
    "center": "@",
    "mark": "General",
    "gen": "General",
}

ROW_KINDS = ("data", "group", "section", "sub1", "sub2", "total")


class Col:
    def __init__(self, header: str, kind: str = "text", width: float = 0, align: str = ""):
        self.header = header
        self.kind = kind
        self.width = width
        self.align = align or ("center" if kind in ("date", "center", "int", "mark") else "")


class Row:
    __slots__ = ("cells", "kind", "fmt", "align")

    def __init__(self, ncols: int, kind: str = "data"):
        self.cells = [None] * ncols
        self.kind = kind
        self.fmt = None    # {col: num_format}
        self.align = None  # {col: align}

    def set(self, col, value):
        self.cells[col] = value
        return self

    def set_fmt(self, col, numfmt):
        if self.fmt is None:
            self.fmt = {}
        self.fmt[col] = numfmt

    def set_align(self, col, align):
        if self.align is None:
            self.align = {}
        self.align[col] = align


class Sheet:
    """Một sheet báo cáo: khối tiêu đề, bảng (header + dữ liệu), chữ ký."""

    def __init__(self, name: str, columns, header_row: int = 9):
        self.name = name[:31]
        self.columns = list(columns)
        self.ncols = len(self.columns)
        self.header_row = header_row          # dòng Excel (1-based) bắt đầu tiêu đề cột
        self.header_cells = []                # (r1, c1, r2, c2, text)
        self.header_rows = 1
        self.numbering = True                 # dòng đánh số thứ tự cột 1,2,3...
        self.numbering_labels = None          # nhãn đánh số tuỳ chọn
        self.rows = []                        # list[Row]
        self.cells = []                       # (row, col, value, style dict, across_to_col)
        self.merges = []                      # (r1,c1,r2,c2) ngoài bảng
        self.row_heights = {}
        self.col_widths = {}
        self.footer_gap = 1
        self.footer = None                    # dict cấu hình chữ ký
        self.landscape = True
        self.fit_width = True
        self.hide_zero = True
        self.freeze = True
        self.table_border = True
        self._cache = {}

    # ----- vị trí -----
    @property
    def data_start(self) -> int:
        return self.header_row + self.header_rows + (1 if self.numbering else 0)

    def next_row(self) -> int:
        """Dòng Excel của Row tiếp theo sẽ được thêm."""
        return self.data_start + len(self.rows)

    def add(self, kind: str = "data") -> Row:
        row = Row(self.ncols, kind)
        self.rows.append(row)
        return row

    def excel_row(self, row: Row) -> int:
        return self.data_start + self.rows.index(row)

    @property
    def table_end(self) -> int:
        return self.data_start + len(self.rows) - 1

    # ----- truy cập giá trị -----
    def get(self, r: int, c: int):
        idx = r - self.data_start
        if 0 <= idx < len(self.rows):
            cells = self.rows[idx].cells
            if 0 <= c < len(cells):
                return cells[c]
        return None

    def eval(self, v):
        if isinstance(v, Formula):
            key = id(v)
            if key not in self._cache:
                self._cache[key] = v.value(self)
            return self._cache[key]
        return v

    # ----- tiện ích dựng tiêu đề -----
    def put(self, r: int, c, value, across_to=None, **style):
        if isinstance(c, str):
            c = vn.col_index(c)
        if isinstance(across_to, str):
            across_to = vn.col_index(across_to)
        self.cells.append((r, c, value, style, across_to))

    def simple_header(self):
        """Tiêu đề 1 dòng lấy từ Col.header."""
        self.header_rows = 1
        self.header_cells = [(self.header_row, c, self.header_row, c, col.header)
                             for c, col in enumerate(self.columns)]


# ---------------------------------------------------------------------------
# Khối tiêu đề / chữ ký chuẩn Agribank
# ---------------------------------------------------------------------------


def bank_block(sheet: Sheet, ctx, col="A", start_row=1, star=True, across_to=None, show_code=True,
               bold_rows=None, italic_name=False, align="center"):
    lines = ["NGÂN HÀNG NÔNG NGHIỆP", "VÀ PHÁT TRIỂN NÔNG THÔN VIỆT NAM"]
    if star:
        lines.append("******************")
    if show_code:
        lines.append("Mã chi nhánh: %s" % ctx.brcd)
    lines.append("Tên chi nhánh: %s" % ctx.brname)
    for i, text in enumerate(lines):
        st = {"bold": True}
        if italic_name and text.startswith("Tên chi nhánh"):
            st["italic"] = True
        if align == "center":
            sheet.put(start_row + i, col, text, across_to=across_to, align="center", **st)
        else:
            sheet.put(start_row + i, col, text, align="left", **st)


def form_block(sheet: Sheet, col, lines, start_row=1):
    for i, text in enumerate(lines):
        sheet.put(start_row + i, col, text, align="right")


def title_lines(sheet: Sheet, row: int, title: str, subtitle: str = "", col="A", across_to=None,
                size=16, sub_size=12, sub_bold=False, extra=None):
    last = across_to if across_to is not None else sheet.ncols - 1
    sheet.put(row, col, title, across_to=last, bold=True, size=size, wrap="\n" in title)
    if "\n" in title:
        sheet.row_heights[row] = 42
    r = row + 1
    if extra:
        sheet.put(r, col, extra, across_to=last, size=sub_size)
        r += 1
    if subtitle:
        sheet.put(r, col, subtitle, across_to=last, italic=True, size=sub_size, bold=sub_bold)


def unit_line(sheet: Sheet, row: int, col, text="Đơn vị: VND"):
    sheet.put(row, col, text, align="right", italic=True)


def signatures(sheet: Sheet, ctx, items, date_col=None, gap=1):
    """items: [(col, chức danh, ghi chú, khoá tên người ký)]"""
    sheet.footer = {"items": items, "date_col": date_col, "gap": gap, "ctx": ctx}


# ---------------------------------------------------------------------------
# Ghi Excel
# ---------------------------------------------------------------------------


class _Formats:
    def __init__(self, wb, font_name="Times New Roman", font_size=10, tomau=True):
        self.wb = wb
        self.font_name = font_name
        self.font_size = font_size
        self.tomau = tomau
        self.cache = {}

    def get(self, **props):
        key = tuple(sorted((k, v) for k, v in props.items() if v not in (None, False, "")))
        fmt = self.cache.get(key)
        if fmt is None:
            d = {"font_name": self.font_name, "font_size": self.font_size, "valign": "vcenter"}
            for k, v in key:
                if k == "numfmt":
                    d["num_format"] = v
                elif k == "size":
                    d["font_size"] = v
                elif k == "color":
                    d["font_color"] = v
                elif k == "fill":
                    d["bg_color"] = v
                    d["pattern"] = 1
                elif k == "border":
                    d["border"] = 1
                elif k == "wrap":
                    d["text_wrap"] = True
                elif k == "across":
                    d["center_across"] = True
                elif k == "align":
                    d["align"] = v
                else:
                    d[k] = v
            if "across" in dict(key):
                d.pop("align", None)
            fmt = self.wb.add_format(d)
            self.cache[key] = fmt
        return fmt


ROW_STYLE_COLOR = {
    "data": {},
    "group": {"bold": True},
    "section": {"bold": True},
    "sub1": {"bold": True, "color": "#0000FF", "fill": "#F2F2F2"},
    "sub2": {"bold": True, "color": "#FF0000", "fill": "#D9D9D9"},
    "total": {"bold": True, "color": "#FF0000", "fill": "#BFBFBF"},
}
HEADER_FILL = "#D9D9D9"


def _row_style(kind, tomau):
    st = dict(ROW_STYLE_COLOR.get(kind, {}))
    if not tomau:
        st.pop("color", None)
        st.pop("fill", None)
    return st


def _display_len(v) -> int:
    if v is None:
        return 0
    if isinstance(v, (_dt.date, _dt.datetime)):
        return 10
    if isinstance(v, float):
        return len("{:,.2f}".format(v))
    if isinstance(v, int):
        return len("{:,}".format(v))
    if isinstance(v, Formula):
        return 0
    return len(str(v))


def _auto_widths(sheet: Sheet):
    widths = []
    sample = sheet.rows if len(sheet.rows) <= 5000 else sheet.rows[:3000] + sheet.rows[-500:]
    for c, col in enumerate(sheet.columns):
        if c in sheet.col_widths:
            widths.append(sheet.col_widths[c])
            continue
        if col.width:
            widths.append(col.width)
            continue
        mx = 0
        for row in sample:
            v = row.cells[c]
            if isinstance(v, Formula):
                v = sheet.eval(v)
            mx = max(mx, _display_len(v))
        head_words = max((len(w) for w in str(col.header).split()), default=6)
        w = max(mx + 2, head_words + 2, 8)
        widths.append(min(w, 48))
    return widths


def _write_value(ws, r0, c0, v, fmt, sheet):
    if v is None or v == "":
        ws.write_blank(r0, c0, None, fmt)
    elif isinstance(v, Formula):
        ws.write_formula(r0, c0, v.text(), fmt, sheet.eval(v))
    elif isinstance(v, bool):
        ws.write_string(r0, c0, "X" if v else "", fmt)
    elif isinstance(v, (int, float)):
        ws.write_number(r0, c0, v, fmt)
    elif isinstance(v, _dt.datetime):
        ws.write_datetime(r0, c0, v, fmt)
    elif isinstance(v, _dt.date):
        ws.write_datetime(r0, c0, _dt.datetime(v.year, v.month, v.day), fmt)
    else:
        ws.write_string(r0, c0, str(v), fmt)


def write_workbook(path: str, sheets, ctx, progress=None):
    """Ghi danh sách Sheet ra file .xlsx."""
    big = sum(len(s.rows) for s in sheets) > 20000
    wb = xlsxwriter.Workbook(path, {"constant_memory": big, "strings_to_numbers": False,
                                    "strings_to_formulas": False, "strings_to_urls": False})
    wb.set_properties({"title": sheets[0].name if sheets else "", "author": AUTHOR,
                       "company": "Agribank", "comments": "Tạo bởi %s - %s" % (APP_NAME, AUTHOR)})
    fm = _Formats(wb, font_size=getattr(ctx, "font_size", 10), tomau=ctx.tomau)
    for sheet in sheets:
        _write_sheet(wb, fm, sheet, ctx, progress)
    wb.close()


def _write_sheet(wb, fm, sheet: Sheet, ctx, progress=None):
    ws = wb.add_worksheet(sheet.name)
    tomau = ctx.tomau
    widths = _auto_widths(sheet)
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Gom tất cả ô cần ghi theo thứ tự dòng (bắt buộc với chế độ tiết kiệm bộ nhớ)
    events = {}  # row0 -> list of (col0, value, fmt)

    def add(r1, c0, v, fmt):
        events.setdefault(r1 - 1, []).append((c0, v, fmt))

    # 1. Ô tiêu đề tự do
    for (r, c, v, st, across_to) in sheet.cells:
        props = dict(st)
        if across_to is not None and across_to > c:
            props["across"] = True
            add(r, c, v, fm.get(**props))
            for cc in range(c + 1, across_to + 1):
                add(r, cc, None, fm.get(**props))
        else:
            add(r, c, v, fm.get(**props))

    # 2. Tiêu đề cột
    hdr_fill = HEADER_FILL if tomau else None
    hfmt = fm.get(bold=True, align="center", wrap=True, border=True, fill=hdr_fill)
    hr0 = sheet.header_row
    hr_end = hr0 + sheet.header_rows - 1
    covered = set()
    for (r1, c1, r2, c2, text) in sheet.header_cells:
        if r2 > r1 or c2 > c1:
            ws.merge.append([r1 - 1, c1, r2 - 1, c2])
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                covered.add((rr, cc))
                add(rr, cc, text if (rr == r1 and cc == c1) else None, hfmt)
    for rr in range(hr0, hr_end + 1):
        for cc in range(sheet.ncols):
            if (rr, cc) not in covered:
                add(rr, cc, None, hfmt)
    if sheet.numbering:
        nfmt = fm.get(align="center", border=True, fill=hdr_fill, numfmt="0", italic=True)
        labels = sheet.numbering_labels or list(range(1, sheet.ncols + 1))
        for cc in range(sheet.ncols):
            add(hr_end + 1, cc, labels[cc] if cc < len(labels) else None, nfmt)

    # 3. Chữ ký (sau bảng)
    if sheet.footer:
        _footer_cells(sheet, add, fm)

    heights = _header_heights(sheet, widths)
    heights.update(sheet.row_heights)
    for r, h in heights.items():
        ws.set_row(r - 1, h)
    for m in sheet.merges:
        ws.merge.append([m[0] - 1, m[1], m[2] - 1, m[3]])

    ds0 = sheet.data_start - 1
    de0 = ds0 + len(sheet.rows)

    def flush(lo, hi):
        for r0 in sorted(k for k in events if lo <= k < hi):
            for c0, v, fmt in sorted(events[r0], key=lambda e: e[0]):
                _write_value(ws, r0, c0, v, fmt, sheet)

    flush(-1, ds0)

    # 4. Dữ liệu – ghi thẳng từng dòng, định dạng được lưu đệm theo (loại dòng, cột)
    col_base = []
    for col in sheet.columns:
        col_base.append({"numfmt": NUM_FORMATS.get(col.kind, "General"), "align": col.align})
    border = sheet.table_border
    ncols = sheet.ncols
    cache = {}

    def fmt_for(kind, c, numfmt=None, align=None):
        key = (kind, c, numfmt, align)
        f = cache.get(key)
        if f is None:
            props = dict(col_base[c])
            props.update(_row_style(kind, tomau))
            if numfmt:
                props["numfmt"] = numfmt
            if align:
                props["align"] = align
            if border:
                props["border"] = True
            f = fm.get(**props)
            cache[key] = f
        return f

    total_rows = len(sheet.rows)
    extra = {k: v for k, v in events.items() if ds0 <= k < de0}
    for i, row in enumerate(sheet.rows):
        r0 = ds0 + i
        kind = row.kind
        rf, ra = row.fmt, row.align
        cells = row.cells
        for c in range(ncols):
            if rf is None and ra is None:
                f = fmt_for(kind, c)
            else:
                f = fmt_for(kind, c, rf.get(c) if rf else None, ra.get(c) if ra else None)
            _write_value(ws, r0, c, cells[c], f, sheet)
        if r0 in extra:
            for c0, v, f in sorted(extra[r0], key=lambda e: e[0]):
                if c0 >= ncols:
                    _write_value(ws, r0, c0, v, f, sheet)
        if progress and i % 5000 == 0:
            progress(i, total_rows)

    flush(de0, 10 ** 9)

    # 5. In ấn
    if sheet.hide_zero:
        ws.hide_zero()
    if sheet.freeze and sheet.rows:
        ws.freeze_panes(sheet.data_start - 1, 0)
    if sheet.landscape:
        ws.set_landscape()
    ws.set_paper(getattr(ctx, "paper", 9))
    ws.set_margins(left=0.45, right=0.45, top=0.5, bottom=0.5)
    ws.center_horizontally()
    if sheet.fit_width:
        ws.fit_to_pages(1, 0)
    ws.repeat_rows(sheet.header_row - 1, sheet.data_start - 2)
    ws.set_footer("&C&P/&N", {"margin": 0.3})
    ws.set_header("", {"margin": 0.3})


def _header_heights(sheet: Sheet, widths):
    """Ước lượng chiều cao dòng tiêu đề cột để chữ xuống dòng không bị che."""
    line_h = 13.5
    need = {}
    spans = []
    for (r1, c1, r2, c2, text) in sheet.header_cells:
        width = sum(widths[c] for c in range(c1, min(c2, len(widths) - 1) + 1)) or 8
        lines = 0
        for part in str(text or "").split("\n"):
            lines += max(1, -(-int(len(part) * 1.1) // max(int(width), 1)))
        spans.append((r1, r2, lines * line_h + 3))
    for r1, r2, h in spans:
        if r1 == r2:
            need[r1] = max(need.get(r1, 15), h)
    for r1, r2, h in spans:
        if r2 > r1:
            cur = sum(need.get(r, 15) for r in range(r1, r2 + 1))
            if cur < h:
                need[r2] = need.get(r2, 15) + (h - cur)
    return {r: min(h, 160) for r, h in need.items()}


def _footer_cells(sheet: Sheet, add, fm):
    f = sheet.footer
    ctx = f["ctx"]
    r = sheet.table_end + 1 + f.get("gap", 1)
    if f.get("date_col") is not None:
        dc = f["date_col"]
        dc = vn.col_index(dc) if isinstance(dc, str) else dc
        add(r, dc, vn.today_line(), fm.get(italic=True, align="center"))
        r += 1
    names_row = r + 5
    for item in f["items"]:
        col, title, note = item[0], item[1], item[2]
        key = item[3] if len(item) > 3 else ""
        c = vn.col_index(col) if isinstance(col, str) else col
        add(r, c, title, fm.get(bold=True, align="center"))
        add(r + 1, c, note, fm.get(italic=True, align="center"))
        name = ctx.signer(key) if key else ""
        if name:
            add(names_row, c, name, fm.get(bold=True, align="center"))
