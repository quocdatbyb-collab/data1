"""Mẫu biểu tuỳ chỉnh: người dùng tự thiết kế mẫu trong Excel rồi tải lên chương trình.

Cách hoạt động
--------------
* Sheet đầu tiên (hoặc sheet tên "MAU") là hình thức mẫu biểu: tiêu đề, tiêu đề cột,
  định dạng, chữ ký... thiết kế tuỳ ý trong Excel.
* Một DÒNG DỮ LIỆU chứa các ô có mã {c1}, {c2}... (cột thứ 1, 2... của file CSV)
  hoặc {Tên cột CSV}. Chương trình nhân dòng này cho từng dòng dữ liệu.
* Các dòng phía dưới dòng dữ liệu (dòng cộng, chữ ký) được đẩy xuống sau dữ liệu.
  Ô {TONG} sẽ thành công thức cộng cột tương ứng, {DEM} = số dòng dữ liệu.
* Mã chung dùng ở bất kỳ đâu: {MA_CN} {TEN_CN} {NGAY_QT} {NGAY_QT_CHU} {NGAY_LAP}
  {LOAI_QT} {MA_MAU} {TEN_MAU} {NAM} {DDMM}.
* Bộ lọc trong mã: {c3|phai:9} {c3|trai:4} {c6|bo:1} {c7|ngay} {c10|so} {c10|chu:c8} {c4|hoa}
  {c2}{c3|phai:9} (ghép chuỗi).
* Sheet "CAU_HINH" (tuỳ chọn): SAP_XEP, NHOM_THEO, CONG_COT, NHAN_NHOM, COT_NHAN, LOC, DONG_TIEU_DE.
"""

from __future__ import annotations

import copy
import datetime as _dt
import json
import os
import re
import shutil

from . import vn

PH_RE = re.compile(r"\{([^{}]+)\}")
GLOBAL_KEYS = ("MA_CN", "TEN_CN", "NGAY_QT", "NGAY_QT_CHU", "NGAY_LAP", "LOAI_QT", "MA_MAU", "TEN_MAU", "NAM",
               "DDMM")
CONFIG_SHEET = "CAU_HINH"
GUIDE_SHEET = "HUONG_DAN"


class TemplateError(Exception):
    pass


# ---------------------------------------------------------------------------
# Danh mục mẫu tuỳ chỉnh
# ---------------------------------------------------------------------------


class TemplateStore:
    def __init__(self, base_dir):
        self.dir = os.path.join(base_dir, "MauBieu")
        os.makedirs(self.dir, exist_ok=True)
        self.index_path = os.path.join(self.dir, "danhmuc.json")
        self.items = []
        self.load()

    def load(self):
        self.items = []
        if os.path.exists(self.index_path):
            try:
                with open(self.index_path, "r", encoding="utf-8") as f:
                    self.items = json.load(f)
            except (OSError, ValueError):
                self.items = []
        self.items = [t for t in self.items if os.path.exists(self.path_of(t))]

    def save(self):
        with open(self.index_path, "w", encoding="utf-8") as f:
            json.dump(self.items, f, ensure_ascii=False, indent=1)

    def path_of(self, item):
        return os.path.join(self.dir, item["file"])

    def get(self, code):
        code = (code or "").lower()
        for t in self.items:
            if t["code"] == code:
                return t
        return None

    def codes(self):
        return [t["code"] for t in self.items]

    def add(self, src_path, code, name, override=False):
        code = re.sub(r"[^a-z0-9]", "", code.lower())
        if not code:
            raise TemplateError("Mã mẫu chỉ gồm chữ không dấu và số, ví dụ: rt25")
        info = inspect_template(src_path)
        fname = "%s.xlsx" % code
        shutil.copyfile(src_path, os.path.join(self.dir, fname))
        self.items = [t for t in self.items if t["code"] != code]
        self.items.append({"code": code, "name": name or info.get("TEN_MAU") or code, "file": fname,
                           "override": bool(override), "added": _dt.date.today().strftime("%d/%m/%Y")})
        self.items.sort(key=lambda t: t["code"])
        self.save()
        return info

    def remove(self, code):
        t = self.get(code)
        if not t:
            return
        try:
            os.remove(self.path_of(t))
        except OSError:
            pass
        self.items = [x for x in self.items if x["code"] != t["code"]]
        self.save()

    def names(self):
        return {t["code"]: t["name"] for t in self.items}


# ---------------------------------------------------------------------------
# Đọc mẫu
# ---------------------------------------------------------------------------


def _load(path):
    import openpyxl
    try:
        return openpyxl.load_workbook(path)
    except Exception as e:  # noqa: BLE001
        raise TemplateError("Không mở được file mẫu: %s" % e)


def _main_sheet(wb):
    if "MAU" in wb.sheetnames:
        return wb["MAU"]
    for ws in wb.worksheets:
        if ws.title not in (CONFIG_SHEET, GUIDE_SHEET):
            return ws
    raise TemplateError("File mẫu không có sheet thiết kế")


def _is_data_token(tok: str) -> bool:
    head = tok.split("|")[0].strip()
    if head.upper() in GLOBAL_KEYS or head.upper() in ("TONG", "DEM") or head.upper().startswith("TONG:"):
        return False
    return True


def _find_data_row(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                toks = PH_RE.findall(cell.value)
                if any(_is_data_token(t) for t in toks):
                    return cell.row
    return None


def _read_config(wb):
    cfg = {}
    if CONFIG_SHEET in wb.sheetnames:
        ws = wb[CONFIG_SHEET]
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip().upper()
            val = row[1] if len(row) > 1 else None
            if not key or key.startswith("#") or val is None or str(val).strip() == "":
                continue
            if key == "LOC":
                cfg.setdefault("LOC", []).append(str(val).strip())
            else:
                cfg[key] = str(val).strip()
    return cfg


def inspect_template(path):
    wb = _load(path)
    ws = _main_sheet(wb)
    r = _find_data_row(ws)
    if r is None:
        raise TemplateError("Không tìm thấy dòng dữ liệu. Hãy đặt các mã {c1}, {c2}... vào một dòng của mẫu.")
    tokens = []
    for cell in ws[r]:
        if isinstance(cell.value, str):
            tokens += PH_RE.findall(cell.value)
    cfg = _read_config(wb)
    info = dict(cfg)
    info["data_row"] = r
    info["tokens"] = tokens
    info["sheet"] = ws.title
    return info


# ---------------------------------------------------------------------------
# Tính giá trị mã
# ---------------------------------------------------------------------------


def _header_index(header):
    idx = {}
    for i, h in enumerate(header or []):
        idx.setdefault(vn.norm_key(h), i + 1)
    return idx


def _col_of(name, hidx):
    name = name.strip()
    m = re.fullmatch(r"[cC](?:ot)?\s*:?\s*(\d+)", name)
    if m:
        return int(m.group(1))
    return hidx.get(vn.norm_key(name))


def _apply_filters(value, filters, row, hidx, ctx):
    out = value
    typed = None
    for f in filters:
        f = f.strip()
        name, _, arg = f.partition(":")
        name = vn.norm_key(name)
        if name in ("phai", "right"):
            out = str(out)[-int(arg or 1):]
        elif name in ("trai", "left"):
            out = str(out)[:int(arg or 1)]
        elif name in ("bo", "bodau", "skip"):
            out = str(out)[int(arg or 1):]
        elif name in ("ngay", "date"):
            d = vn.parse_date(out)
            if d:
                typed = d
                out = d
        elif name in ("so", "num", "number"):
            n = vn.to_number(out)
            typed = n if n is not None else 0
            out = typed
        elif name in ("chu", "bangchu", "words"):
            ccy = ""
            if arg:
                c = _col_of(arg, hidx)
                ccy = _field(row, c) if c else arg
            out = vn.number_to_words(out, ccy, ctx.thousand_word)
            typed = None
        elif name in ("hoa", "upper"):
            out = str(out).upper()
        elif name in ("thuong", "lower"):
            out = str(out).lower()
        elif name in ("trim", "gon"):
            out = str(out).strip()
    return out, typed


def _field(row, c):
    if c and 0 < c <= len(row):
        return row[c - 1]
    return ""


def _globals(ctx, code, name):
    return {"MA_CN": ctx.brcd, "TEN_CN": ctx.brname, "NGAY_QT": ctx.date_short, "NGAY_QT_CHU": ctx.date_full,
            "NGAY_LAP": vn.today_line(), "LOAI_QT": ctx.loai_qt, "MA_MAU": code, "TEN_MAU": name,
            "NAM": str(ctx.date.year), "DDMM": ctx.ddmm}


def _render_text(text, gl):
    def rep(m):
        k = m.group(1).strip().upper()
        return str(gl.get(k, m.group(0)))
    return PH_RE.sub(rep, text)


def _render_data_cell(text, row, hidx, ctx, stt, numeric_cell):
    toks = PH_RE.findall(text)
    only = len(toks) == 1 and PH_RE.fullmatch(text.strip()) is not None

    def eval_tok(tok):
        parts = tok.split("|")
        head = parts[0].strip()
        if head.upper() == "STT":
            return stt, stt
        c = _col_of(head, hidx)
        raw = _field(row, c) if c else ""
        return _apply_filters(raw, parts[1:], row, hidx, ctx)

    if only:
        val, typed = eval_tok(toks[0])
        if typed is not None:
            return typed
        if numeric_cell:
            n = vn.to_number(val)
            if n is not None:
                return n
        return val

    def rep(m):
        v, _ = eval_tok(m.group(1))
        if isinstance(v, (_dt.date, _dt.datetime)):
            return vn.fmt_date(v)
        return str(v)
    return PH_RE.sub(rep, text)


# ---------------------------------------------------------------------------
# Lọc / sắp xếp
# ---------------------------------------------------------------------------

_COND_RE = re.compile(r"^(.+?)\s*(>=|<=|<>|!=|=|>|<|\bchua\b|\bbatdau\b|\bkhongchua\b)\s*(.*)$", re.I)


def _expr_value(expr, row, hidx):
    terms = [t.strip() for t in expr.split("+")]
    if len(terms) > 1:
        return sum(vn.num0(_field(row, _col_of(t, hidx))) for t in terms)
    c = _col_of(terms[0], hidx)
    return _field(row, c) if c else terms[0]


def _make_filter(conds, hidx):
    checks = []
    for cond in conds:
        m = _COND_RE.match(cond.strip())
        if not m:
            continue
        left, op, right = m.group(1), m.group(2).lower(), m.group(3).strip().strip("'\"")
        checks.append((left, op, right))

    def ok(row):
        for left, op, right in checks:
            lv = _expr_value(left, row, hidx)
            ln, rn = vn.to_number(lv), vn.to_number(right)
            if op in (">", "<", ">=", "<=") and ln is not None and rn is not None:
                if not {">": ln > rn, "<": ln < rn, ">=": ln >= rn, "<=": ln <= rn}[op]:
                    return False
                continue
            ls, rs = vn.norm_key(lv), vn.norm_key(right)
            if ln is not None and rn is not None:
                eq = ln == rn
            else:
                eq = ls == rs
            if op == "=" and not eq:
                return False
            if op in ("<>", "!=") and eq:
                return False
            if op == "chua" and rs not in ls:
                return False
            if op == "khongchua" and rs in ls:
                return False
            if op == "batdau" and not ls.startswith(rs):
                return False
        return True
    return ok


def _cols_list(spec, hidx):
    out = []
    for part in re.split(r"[,;]", spec or ""):
        part = part.strip()
        if part:
            c = _col_of(part, hidx)
            if c:
                out.append(c)
    return out


# ---------------------------------------------------------------------------
# Xuất báo cáo từ mẫu
# ---------------------------------------------------------------------------


def render_template(path, out_path, ctx, data, code, name, progress=None):
    """Đổ dữ liệu vào mẫu và lưu ra out_path. Trả về (số dòng dữ liệu, bản xem trước)."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import CellRange
    from openpyxl.styles import Font, PatternFill

    wb = _load(path)
    ws = _main_sheet(wb)
    cfg = _read_config(wb)
    data_row = _find_data_row(ws)
    if data_row is None:
        raise TemplateError("Mẫu không có dòng dữ liệu {c1}, {c2}...")
    gl = _globals(ctx, code, name)
    hidx = _header_index(data.header)

    rows = list(data.rows)
    skip = int(vn.to_number(cfg.get("BO_DONG_DAU"), 0) or 0)
    if skip:
        rows = rows[skip:]
    if cfg.get("LOC"):
        flt = _make_filter(cfg["LOC"], hidx)
        rows = [r for r in rows if flt(r)]
    sort_cols = _cols_list(cfg.get("SAP_XEP", ""), hidx)
    group_cols = _cols_list(cfg.get("NHOM_THEO", ""), hidx)
    if group_cols:
        sort_cols = group_cols + [c for c in sort_cols if c not in group_cols]
    if sort_cols:
        rows.sort(key=lambda r: tuple(str(_field(r, c)) for c in sort_cols))

    max_col = ws.max_column
    tpl_vals, tpl_styles, tpl_numfmt = [], [], []
    for c in range(1, max_col + 1):
        cell = ws.cell(row=data_row, column=c)
        tpl_vals.append(cell.value)
        tpl_styles.append(copy.copy(cell._style) if cell.has_style else None)
        tpl_numfmt.append(cell.number_format or "General")
    data_height = ws.row_dimensions[data_row].height

    # Dòng phía dưới (cộng, chữ ký) – lưu lại rồi xoá để chèn dữ liệu
    footer = []
    last_row = ws.max_row
    for r in range(data_row + 1, last_row + 1):
        cells = []
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cells.append((cell.value, copy.copy(cell._style) if cell.has_style else None))
        footer.append((cells, ws.row_dimensions[r].height))
    footer_merges = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row > data_row:
            footer_merges.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
            ws.unmerge_cells(str(mr))
    for r in range(data_row + 1, last_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).value = None

    # Thay mã chung ở phần đầu
    for row in ws.iter_rows(min_row=1, max_row=data_row - 1):
        for cell in row:
            if isinstance(cell.value, str) and "{" in cell.value:
                cell.value = _render_text(cell.value, gl)

    sum_letters = [s.strip().upper() for s in re.split(r"[,;]", cfg.get("CONG_COT", "")) if s.strip()]
    label_col = vn.col_index(cfg.get("COT_NHAN", "A") or "A") + 1
    label_tpl = cfg.get("NHAN_NHOM", "Cộng {gia_tri}")
    bold = Font(bold=True)
    sub_font = Font(bold=True, color="0000FF") if ctx.tomau else bold
    sub_fill = PatternFill("solid", fgColor="F2F2F2") if ctx.tomau else None

    numeric_cols = set()
    for i, nf in enumerate(tpl_numfmt, 1):
        if nf not in ("General", "@") and "y" not in nf.lower() and "d" not in nf.lower():
            numeric_cols.add(i)

    r = data_row
    preview = []
    stt = 0
    group_start = r
    gkey_prev = None
    n = len(rows)

    def write_group_total(row_idx, start, end, key_text):
        lab = ws.cell(row=row_idx, column=label_col)
        lab.value = label_tpl.replace("{gia_tri}", key_text).replace("{GIA_TRI}", key_text)
        for c in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=c)
            if tpl_styles[c - 1] is not None:
                cell._style = copy.copy(tpl_styles[c - 1])
            cell.font = sub_font
            if sub_fill is not None:
                cell.fill = sub_fill
        for letter in sum_letters:
            col = vn.col_index(letter) + 1
            ws.cell(row=row_idx, column=col).value = "=SUBTOTAL(9,%s%d:%s%d)" % (letter, start, letter, end)

    for i, rec in enumerate(rows):
        gkey = tuple(str(_field(rec, c)) for c in group_cols) if group_cols else None
        if group_cols and gkey_prev is not None and gkey != gkey_prev:
            write_group_total(r, group_start, r - 1, " ".join(gkey_prev))
            r += 1
            group_start = r
        gkey_prev = gkey
        stt += 1
        vals = []
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if tpl_styles[c - 1] is not None:
                cell._style = copy.copy(tpl_styles[c - 1])
            v = tpl_vals[c - 1]
            if isinstance(v, str) and "{" in v:
                v = _render_data_cell(v, rec, hidx, ctx, stt, c in numeric_cols)
                if isinstance(v, str) and "{" in v:
                    v = _render_text(v, gl)
            cell.value = v
            vals.append(v)
        if data_height:
            ws.row_dimensions[r].height = data_height
        if len(preview) < 300:
            preview.append(vals)
        r += 1
        if progress and i % 2000 == 0:
            progress(i, n)
    if group_cols and gkey_prev is not None:
        write_group_total(r, group_start, r - 1, " ".join(gkey_prev))
        r += 1
    if not rows:
        for c in range(1, max_col + 1):
            ws.cell(row=data_row, column=c).value = None
        r = data_row + 1
    data_end = r - 1
    offset = r - (data_row + 1)

    # Dòng cộng / chữ ký
    for k, (cells, height) in enumerate(footer):
        rr = data_row + 1 + k + offset
        for c, (val, style) in enumerate(cells, 1):
            cell = ws.cell(row=rr, column=c)
            if style is not None:
                cell._style = style
            if isinstance(val, str) and "{" in val:
                def rep(m, col=c):
                    tok = m.group(1).strip()
                    up = tok.upper()
                    if up == "TONG" or up.startswith("TONG:"):
                        letter = tok.split(":", 1)[1].strip().upper() if ":" in tok else get_column_letter(col)
                        return "SUBTOTAL(9,%s%d:%s%d)" % (letter, data_row, letter, max(data_end, data_row))
                    if up == "DEM":
                        return str(stt)
                    return str(gl.get(up, m.group(0)))
                text = val.strip()
                if PH_RE.fullmatch(text) and PH_RE.fullmatch(text).group(1).strip().upper().startswith("TONG"):
                    val = "=" + PH_RE.sub(rep, text)
                elif PH_RE.fullmatch(text) and text.strip("{} ").upper() == "DEM":
                    val = stt
                else:
                    val = PH_RE.sub(rep, val)
                    if "SUBTOTAL(" in val:
                        val = "=" + val if not val.startswith("=") else val
            cell.value = val
        if height:
            ws.row_dimensions[rr].height = height
    for (r1, c1, r2, c2) in footer_merges:
        ws.merge_cells(str(CellRange(min_col=c1, min_row=r1 + offset, max_col=c2, max_row=r2 + offset)))

    for sname in (CONFIG_SHEET, GUIDE_SHEET):
        if sname in wb.sheetnames:
            del wb[sname]
    wb.properties.creator = "8702 - Nguyễn Quốc Đạt"
    try:
        wb.calculation.fullCalcOnLoad = True
    except AttributeError:
        pass
    ws.title = ("%s_%s" % (ctx.brcd, code))[:31]
    wb.save(out_path)
    headers = []
    if data_row > 1:
        for c in range(1, max_col + 1):
            v = ws.cell(row=data_row - 1, column=c).value
            headers.append("" if v is None else str(v))
    return stt, (headers, preview)


# ---------------------------------------------------------------------------
# Tạo mẫu khởi đầu từ file CSV
# ---------------------------------------------------------------------------


def create_starter(data, out_path, code="rtxx", title="TÊN MẪU BIỂU"):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    header = data.header or ["Cột %d" % (i + 1) for i in range(max((len(r) for r in data.rows[:50]), default=5))]
    ncol = len(header)
    sample = data.rows[:200]
    date_cols = []
    for i in range(ncol):
        vals = [str(r[i]).strip() for r in sample if i < len(r) and str(r[i]).strip()]
        date_cols.append(bool(vals) and all(vn.parse_date(v) is not None and len(v) in (8, 10) for v in vals))
    money_words = ("so du", "so tien", "du no", "lai", "goc", "gia tri", "amount", "bal", "quy doi", "nguyen te")
    numeric = []
    for i in range(ncol):
        vals = [str(r[i]).strip() for r in sample if i < len(r) and str(r[i]).strip()]
        nums = [vn.to_number(v) for v in vals if vn.is_number(v)]
        if not vals or len(nums) < 0.9 * len(vals):
            numeric.append(False)
            continue
        codes = any(v.isdigit() and (v.startswith("0") or len(v) >= 9) for v in vals)
        dates = all(len(v) == 8 and v.isdigit() and v[:2] in ("19", "20") for v in vals)
        hkey = vn.norm_key(header[i]) if i < len(header) else ""
        by_name = any(w in hkey for w in money_words)
        big = sorted(abs(n) for n in nums)[len(nums) // 2] >= 100000
        has_dec = any("." in v for v in vals)
        numeric.append(not codes and not dates and not date_cols[i] and (by_name or big or has_dec))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAU"
    font = Font(name="Times New Roman", size=10)
    bold = Font(name="Times New Roman", size=10, bold=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill("solid", fgColor="D9D9D9")
    last = get_column_letter(max(ncol, 3))

    def put(ref, val, f=font, align=None):
        c = ws[ref]
        c.value = val
        c.font = f
        if align:
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return c

    put("A1", "NGÂN HÀNG NÔNG NGHIỆP", bold)
    put("A2", "VÀ PHÁT TRIỂN NÔNG THÔN VIỆT NAM", bold)
    put("A3", "Mã chi nhánh: {MA_CN}", bold)
    put("A4", "Tên chi nhánh: {TEN_CN}", bold)
    put(last + "1", "1. Mẫu số ../{LOAI_QT}", font, "right")
    put(last + "2", "2. Lưu tại Chi nhánh", font, "right")
    ws.merge_cells("A6:%s6" % last)
    put("A6", title, Font(name="Times New Roman", size=16, bold=True), "center")
    ws.merge_cells("A7:%s7" % last)
    put("A7", "Ngày {NGAY_QT_CHU}", Font(name="Times New Roman", size=12, italic=True), "center")
    put(last + "8", "Đơn vị: VND", Font(name="Times New Roman", size=10, italic=True), "right")
    for i, h in enumerate(header, 1):
        col = get_column_letter(i)
        c = put("%s9" % col, h, bold, "center")
        c.border = border
        c.fill = fill
        c2 = put("%s10" % col, i, Font(name="Times New Roman", size=10, italic=True), "center")
        c2.border = border
        c2.fill = fill
        d = put("%s11" % col, "{c%d|ngay}" % i if date_cols[i - 1] else "{c%d}" % i)
        d.border = border
        if date_cols[i - 1]:
            d.number_format = "dd/mm/yyyy"
            d.alignment = Alignment(horizontal="center")
        elif numeric[i - 1]:
            d.number_format = "#,##0"
        else:
            d.number_format = "@"
        t = ws["%s12" % col]
        t.border = border
        t.font = bold
        if numeric[i - 1]:
            t.value = "{TONG}"
            t.number_format = "#,##0"
        ws.column_dimensions[col].width = min(max(len(str(h)) + 2, 10), 40)
    put("A12", "CỘNG", bold, "center")
    ws.row_dimensions[9].height = 45
    sig_cols = ["B", get_column_letter(max(2, ncol // 2)), get_column_letter(max(3, ncol - 1))]
    for col, (t, n) in zip(sig_cols, [("LẬP BIỂU", "(Ký, ghi rõ họ tên, số ĐT liên hệ)"),
                                      ("TRƯỞNG PHÒNG KẾ TOÁN", "(Ký, ghi rõ họ tên)"),
                                      ("GIÁM ĐỐC", "(Ký, đóng dấu, ghi rõ họ tên)")]):
        put("%s14" % col, t, bold, "center")
        put("%s15" % col, n, Font(name="Times New Roman", size=10, italic=True), "center")
    ws.print_title_rows = "9:10"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "A11"

    cfg = wb.create_sheet(CONFIG_SHEET)
    rows = [("THAM SỐ", "GIÁ TRỊ", "GIẢI THÍCH"),
            ("TEN_MAU", title, "Tên hiển thị của mẫu"),
            ("SAP_XEP", "", "Sắp xếp theo cột, ví dụ: c3, c5"),
            ("NHOM_THEO", "", "Nhóm và chèn dòng cộng theo cột, ví dụ: c3"),
            ("CONG_COT", "", "Các cột Excel cần cộng trong dòng cộng nhóm, ví dụ: G,H"),
            ("NHAN_NHOM", "Cộng {gia_tri}", "Nhãn dòng cộng nhóm"),
            ("COT_NHAN", "A", "Cột Excel ghi nhãn dòng cộng nhóm"),
            ("LOC", "", "Điều kiện lọc, ví dụ: c24+c25 > 0 ; c5 = The chap ; c22 <> Cá nhân (mỗi dòng 1 điều kiện)"),
            ("LOC", "", ""),
            ("BO_DONG_DAU", "", "Bỏ qua thêm N dòng dữ liệu đầu (ngoài dòng tiêu đề)")]
    for rr in rows:
        cfg.append(rr)
    for c in ("A", "B", "C"):
        cfg[c + "1"].font = Font(bold=True)
    cfg.column_dimensions["A"].width = 16
    cfg.column_dimensions["B"].width = 30
    cfg.column_dimensions["C"].width = 90

    g = wb.create_sheet(GUIDE_SHEET)
    guide = [
        "HƯỚNG DẪN THIẾT KẾ MẪU BIỂU",
        "1. Sửa sheet MAU tuỳ ý: tiêu đề, tiêu đề cột, định dạng, độ rộng cột, chữ ký...",
        "2. Dòng 11 là DÒNG DỮ LIỆU: {c5} = cột thứ 5 của file CSV. Có thể xoá/đổi thứ tự cột.",
        "   Có thể dùng tên cột: {Tên khách hàng}. Số thứ tự: {STT}.",
        "3. Bộ lọc: {c3|phai:9} lấy 9 ký tự cuối; {c3|trai:4}; {c6|bo:1} bỏ 1 ký tự đầu;",
        "   {c7|ngay} đổi yyyymmdd thành ngày; {c10|so} thành số; {c10|chu:c8} đọc số thành chữ theo loại tiền ở cột 8;",
        "   {c4|hoa} viết hoa; {c2}{c3|phai:9} ghép mã chi nhánh + mã khách hàng.",
        "4. Dòng dưới dữ liệu: {TONG} = cộng cột đó; {TONG:H} = cộng cột H; {DEM} = số dòng.",
        "5. Mã chung: {MA_CN} {TEN_CN} {NGAY_QT} {NGAY_QT_CHU} {NGAY_LAP} {LOAI_QT} {MA_MAU} {TEN_MAU} {NAM} {DDMM}.",
        "6. Sheet CAU_HINH: sắp xếp, nhóm/cộng nhóm, lọc dữ liệu.",
        "7. Lưu file, vào chương trình > Mẫu biểu tuỳ chỉnh > Tải lên mẫu mới. Đặt tên file dữ liệu: 5500_<mã mẫu>.csv",
        "Sheet CAU_HINH và HUONG_DAN tự động bị loại khỏi file kết quả.",
    ]
    for line in guide:
        g.append([line])
    g["A1"].font = Font(bold=True, size=12)
    g.column_dimensions["A"].width = 120
    wb.properties.creator = "8702 - Nguyễn Quốc Đạt"
    wb.save(out_path)
    return out_path
