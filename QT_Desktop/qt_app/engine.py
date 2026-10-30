"""Điều phối: tìm file dữ liệu, chạy mẫu biểu, ghi file kết quả."""

from __future__ import annotations

import glob
import os
import time
from collections import OrderedDict

from . import forms, glcheck, reader
from .custom import render_template
from .settings import Context
from .sheet import write_workbook, Formula

TEST_LIMIT = 50


class Job:
    def __init__(self, brcd, code, files, raw_code=None):
        self.brcd = brcd
        self.code = code
        self.raw_code = raw_code or code
        self.files = files
        self.status = ""
        self.results = []
        self.error = ""

    @property
    def folder(self):
        return os.path.dirname(self.files[0]) if self.files else ""


class Result:
    def __init__(self):
        self.paths = []
        self.messages = []
        self.preview = None   # (headers, rows)
        self.rows_read = 0


def form_label(code):
    spec = forms.FORMS.get(code)
    return spec["name"] if spec else code


def is_known(code, store=None) -> bool:
    code = forms.resolve(code)
    return code in forms.FORMS or bool(store and store.get(code))


def discover(folder, settings, store=None, only_branches=False):
    """Quét thư mục -> danh sách Job (gộp các file nhiều phần của cùng mẫu)."""
    groups = OrderedDict()
    branch_codes = {str(k) for k, _ in settings.branches}
    for p in sorted(glob.glob(os.path.join(folder, "*"))):
        if not p.lower().endswith((".csv", ".txt")):
            continue
        if glcheck.is_gl_name(p):
            continue
        brcd, raw = reader.parse_filename(p)
        if not brcd:
            continue
        if only_branches and brcd not in branch_codes:
            continue
        code = forms.resolve(raw)
        groups.setdefault((brcd, code, raw), []).append(p)
    jobs = []
    for (brcd, code, raw), files in groups.items():
        if not is_known(code, store):
            continue
        jobs.append(Job(brcd, code, files, raw))
    jobs.sort(key=lambda j: (j.brcd, list(forms.FORMS).index(j.code) if j.code in forms.FORMS else 99, j.code))
    return jobs


def job_from_file(path, brcd=None, code=None):
    b, raw = reader.parse_filename(path)
    brcd = brcd or b
    code = forms.resolve(code or raw)
    return Job(brcd, code, reader.find_parts(path), raw)


def output_dir_for(job, settings, override=None):
    d = override or settings.get("thu_muc_ket_qua")
    if not d:
        d = os.path.join(job.folder, "Result")
    os.makedirs(d, exist_ok=True)
    return d


def result_path(out_dir, brcd, code):
    return os.path.join(out_dir, "%s_%s.xlsx" % (brcd, code))


def _free_path(path):
    """Nếu file đang mở trong Excel (bị khoá) thì đặt tên khác."""
    if not os.path.exists(path):
        return path
    try:
        with open(path, "a"):
            pass
        return path
    except OSError:
        base, ext = os.path.splitext(path)
        i = 1
        while True:
            p = "%s (%d)%s" % (base, i, ext)
            if not os.path.exists(p):
                return p
            i += 1


def _preview_of(sheet, limit=300):
    headers = [c.header.replace("\n", " ") for c in sheet.columns]
    rows = []
    for row in sheet.rows[:limit]:
        vals = []
        for v in row.cells:
            if isinstance(v, Formula):
                v = sheet.eval(v)
            vals.append(v)
        rows.append(vals)
    return headers, rows


def check_columns(data, min_cols, job):
    """Báo lỗi rõ ràng khi file không đúng cấu trúc mẫu (vd chọn nhầm file cân đối)."""
    if not min_cols or not data.rows:
        return
    short = sum(1 for r in data.rows if len(r) < min_cols)
    if short > len(data.rows) * 0.3:
        ncol = max(len(r) for r in data.rows)
        hint = ""
        if ncol <= 15:
            hint = ("\nFile này có thể là FILE CÂN ĐỐI TÀI KHOẢN (glcb41). Hãy chọn nó ở ô 'File cân đối' "
                    "(hoặc đổi tên có chữ glcb41, vd %s_glcb41.txt) thay vì chọn làm dữ liệu mẫu." % job.brcd)
        raise ValueError("File %s chỉ có %d cột, mẫu %s cần ít nhất %d cột (%d/%d dòng thiếu cột).%s" % (
            ", ".join(os.path.basename(f) for f in data.files), ncol, job.code, min_cols, short,
            len(data.rows), hint))


class Runner:
    """Chạy job; log(msg) và progress(phần trăm, nội dung) là hàm gọi lại."""

    def __init__(self, settings, store=None, log=None, progress=None, cancel=None):
        self.settings = settings
        self.store = store
        self.log = log or (lambda m: None)
        self.progress = progress or (lambda p, t="": None)
        self.cancel = cancel
        self._gl_cache = {}

    def _gl(self, folder, ctx, files=None):
        key = (folder, tuple(files or ()))
        if key not in self._gl_cache:
            self._gl_cache[key] = glcheck.load_gl(folder, ctx, files)
        return self._gl_cache[key]

    def run(self, job: Job, options, out_dir=None) -> Result:
        res = Result()
        t0 = time.time()
        s = self.settings
        names = self.store.names() if self.store else {}
        ctx = Context(s, job.brcd, options, names)
        out_dir = output_dir_for(job, s, out_dir)
        limit = TEST_LIMIT if options.get("kiemtra") else 0
        tpl = self.store.get(job.code) if self.store else None
        use_custom = tpl is not None and (job.code not in forms.FORMS or tpl.get("override")
                                          or s.flag("uu_tien_mau_tuy_chinh"))

        self.progress(2, "Đang đọc dữ liệu %s_%s..." % (job.brcd, job.raw_code))
        if use_custom:
            data = reader.read_sources(job.files, ctx.encoding, 0, limit)
        else:
            spec = forms.FORMS[job.code]
            data = reader.read_sources(job.files, ctx.encoding, spec.get("min_cols", 0), limit)
            check_columns(data, spec.get("min_cols", 0), job)
        res.rows_read = len(data.rows)
        self.log("Đọc %d dòng từ %d file (mã ký tự: %s, phân cách: %s)%s" % (
            len(data.rows), len(data.files), data.encoding or "-", "TAB" if data.delimiter == "\t" else data.delimiter,
            " – CHẾ ĐỘ XEM THỬ %d dòng" % TEST_LIMIT if limit else ""))
        for w in data.warnings:
            self.log("  ! " + w)
        if not data.rows:
            raise ValueError("File không có dòng dữ liệu nào: " + ", ".join(os.path.basename(f) for f in job.files))

        if use_custom:
            path = _free_path(result_path(out_dir, job.brcd, job.code))
            self.progress(40, "Đang đổ dữ liệu vào mẫu tuỳ chỉnh...")
            n, prev = render_template(self.store.path_of(tpl), path, ctx, data, job.code, tpl["name"],
                                      progress=lambda i, t: self.progress(40 + 55 * i / max(t, 1), "Đang ghi..."))
            res.paths.append(path)
            res.preview = prev
            self.log("Mẫu tuỳ chỉnh '%s': %d dòng -> %s" % (tpl["name"], n, os.path.basename(path)))
            self.progress(100, "Xong")
            return res

        self.progress(20, "Đang xử lý %s..." % form_label(job.code))
        outs = forms.FORMS[job.code]["run"](ctx, data)
        data.rows = []  # giải phóng bộ nhớ trước khi ghi Excel
        for k, out in enumerate(outs):
            if self.cancel is not None and self.cancel.is_set():
                break
            sheets = list(out.sheets)
            if out.gl_items is not None and ctx.opt("kiem_tra_gl"):
                book = self._gl(job.folder, ctx, options.get("gl_files"))
                if book.files:
                    self.log("  Cân đối: " + ", ".join(os.path.basename(f) for f in book.files))
                extra, diffs, missing = glcheck.build_check_sheets(ctx, out.gl_items, book, out.gl_label)
                sheets += extra
                if book.empty:
                    self.log("  ! Không có file cân đối *glcb41* trong thư mục – sheet CheckGL/Mau 30 để trống cột cân đối")
                else:
                    if diffs:
                        self.log("  ! Có %d dòng chênh lệch giữa sao kê và cân đối (xem sheet CheckGL)" % diffs)
                    if missing:
                        self.log("  ! %d tài khoản/loại tiền không có trong file cân đối (đánh dấu '?')" % missing)
                    if not diffs and not missing:
                        self.log("  Đối chiếu cân đối GL: khớp")
            path = _free_path(result_path(out_dir, job.brcd, out.code))
            base = 30 + 65 * k / len(outs)
            span = 65 / len(outs)
            self.progress(base, "Đang ghi file %s..." % os.path.basename(path))
            write_workbook(path, sheets, ctx,
                           progress=lambda i, t, b=base, sp=span: self.progress(b + sp * i / max(t, 1), "Đang ghi Excel..."))
            res.paths.append(path)
            if res.preview is None and sheets:
                res.preview = _preview_of(sheets[0])
            self.log("Đã tạo %s (%d dòng)" % (os.path.basename(path), len(sheets[0].rows)))
        self.progress(100, "Xong (%.1f giây)" % (time.time() - t0))
        return res


def describe_files(paths):
    out = []
    for p in paths:
        try:
            size = os.path.getsize(p)
        except OSError:
            size = 0
        out.append("%s (%s)" % (os.path.basename(p), _human(size)))
    return out


def _human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return ("%d %s" % (n, unit)) if unit == "B" else ("%.1f %s" % (n, unit))
        n /= 1024.0
    return str(n)


def quick_info(path, encoding="auto"):
    """Đọc nhanh vài trăm dòng để hiển thị thông tin file."""
    data = reader.read_sources([path], encoding, 0, 300)
    return data


# ---------------------------------------------------------------------------
# Tổng hợp nhiều chi nhánh (Mẫu 30 + sheet chi tiết từng chi nhánh)
# ---------------------------------------------------------------------------


def discover_branches(folder, code, settings, store=None):
    """Các chi nhánh có file dữ liệu của mẫu `code` trong thư mục."""
    return [j for j in discover(folder, settings, store) if j.code == code]


def consolidate(settings, jobs, code, parent_brcd, options, out_dir=None, log=None, progress=None,
                cancel=None, detail=True):
    """Tạo file tổng hợp: sheet 1 = Mẫu 30 toàn chi nhánh, sheet CheckGL, rồi sheet chi tiết từng chi nhánh."""
    log = log or (lambda m: None)
    progress = progress or (lambda p, t="": None)
    if not jobs:
        raise ValueError("Không có chi nhánh nào để tổng hợp.")
    spec = forms.FORMS[code]
    parent = Context(settings, parent_brcd, options)
    all_items, details, branches = [], [], []
    gl_books = {}
    book = glcheck.GLBook()
    n = len(jobs)
    for k, job in enumerate(jobs):
        if cancel is not None and cancel.is_set():
            log("! Đã dừng theo yêu cầu.")
            break
        progress(5 + 60.0 * k / n, "Đang đọc %s_%s (%d/%d)..." % (job.brcd, job.raw_code, k + 1, n))
        ctx = Context(settings, job.brcd, options)
        limit = TEST_LIMIT if options.get("kiemtra") else 0
        data = reader.read_sources(job.files, ctx.encoding, spec.get("min_cols", 0), limit)
        check_columns(data, spec.get("min_cols", 0), job)
        outs = spec["run"](ctx, data)
        data.rows = []
        main = next((o for o in outs if o.code == code), outs[0])
        items = main.gl_items or []
        all_items += items
        branches.append(job.brcd)
        if detail:
            for sh in main.sheets:
                details.append(sh)
        if job.folder not in gl_books:
            gl_books[job.folder] = glcheck.load_gl(job.folder, ctx, options.get("gl_files"))
            b = gl_books[job.folder]
            book.files += b.files
            book.by_branch.update(b.by_branch)
            book.by_acct.update(b.by_acct)
        log("%s – %s: %d dòng, %d tài khoản" % (job.brcd, settings.branch_name(job.brcd) or "",
                                                 len(main.sheets[0].rows) if main.sheets else 0,
                                                 len({i.acct for i in items})))
    if book.empty:
        log("! Không có file cân đối *glcb41* – cột 'Số dư trên cân đối IPCAS' để '?'")
    elif not book.by_branch:
        log("! File cân đối không có mã chi nhánh ở đầu tên (vd 5501_glcb41.txt) – không tách được số dư "
            "theo từng chi nhánh")
    progress(70, "Đang lập Mẫu 30...")
    m30 = glcheck.build_mau30(parent, all_items, book, code, branches=branches)
    m30.name = "Mau 30"
    sheets = [m30]
    if options.get("kiem_tra_gl", True):
        chk, _diffs, _missing = glcheck.build_check_sheets(parent, all_items, book, code)
        sheets.append(chk[0])
    names = set(s.name for s in sheets)
    for sh in details:
        base = sh.name[:31]
        name, i = base, 1
        while name in names:
            i += 1
            name = ("%s_%d" % (base, i))[:31]
        sh.name = name
        names.add(name)
        sheets.append(sh)
    d = out_dir or settings.get("thu_muc_ket_qua") or os.path.join(jobs[0].folder, "Result")
    os.makedirs(d, exist_ok=True)
    path = _free_path(os.path.join(d, "%s_%s_TongHop.xlsx" % (parent_brcd, code)))
    progress(75, "Đang ghi file tổng hợp...")
    write_workbook(path, sheets, parent, progress=lambda i, t: progress(75 + 24.0 * i / max(t, 1), "Đang ghi Excel..."))
    lech = sum(1 for r in m30.rows if r.kind == "sub1")
    log("Đã tạo %s: Mẫu 30 gồm %d chi nhánh, %d dòng chênh lệch" % (os.path.basename(path), len(branches), lech))
    progress(100, "Xong")
    return path, _preview_of(m30)
