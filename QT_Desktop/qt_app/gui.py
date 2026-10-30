"""Giao diện desktop (Tkinter) cho QT Quyết Toán."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
import datetime as _dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from . import APP_NAME, APP_TITLE, AUTHOR, VERSION
from . import custom, engine, forms, reader, vn
from .settings import Settings, SIGNERS

# ---------------------------------------------------------------------------
# Màu sắc / phông chữ
# ---------------------------------------------------------------------------
C_PRIMARY = "#9D1C3B"     # đỏ mận Agribank
C_PRIMARY_DARK = "#7A1430"
C_ACCENT = "#0E7C3A"      # xanh lá
C_BG = "#F3F4F6"
C_CARD = "#FFFFFF"
C_TEXT = "#1F2937"
C_MUTED = "#6B7280"
C_BORDER = "#E5E7EB"
C_SIDEBAR = "#2B2F36"
C_SIDEBAR_HOVER = "#3A3F48"

FONT = "Segoe UI" if sys.platform.startswith("win") else "DejaVu Sans"


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        messagebox.showerror(APP_NAME, "Không mở được:\n%s\n\n%s" % (path, e))


def form_choices(store):
    items = [(code, "%s  —  %s" % (code, spec["name"])) for code, spec in forms.FORMS.items()]
    if store:
        for t in store.items:
            if t["code"] not in forms.FORMS:
                items.append((t["code"], "%s  —  %s (mẫu tuỳ chỉnh)" % (t["code"], t["name"])))
    return items


# ---------------------------------------------------------------------------
# Thành phần dùng chung
# ---------------------------------------------------------------------------


class Card(tk.Frame):
    def __init__(self, master, title="", **kw):
        super().__init__(master, bg=C_CARD, highlightbackground=C_BORDER, highlightthickness=1, **kw)
        if title:
            tk.Label(self, text=title, bg=C_CARD, fg=C_PRIMARY, font=(FONT, 11, "bold"), anchor="w").pack(
                fill="x", padx=14, pady=(10, 2))
        self.body = tk.Frame(self, bg=C_CARD)
        self.body.pack(fill="both", expand=True, padx=14, pady=(4, 12))


class LogBox(tk.Frame):
    def __init__(self, master, height=8):
        super().__init__(master, bg=C_CARD)
        self.text = tk.Text(self, height=height, wrap="word", bg="#FBFBFC", fg=C_TEXT, relief="flat",
                            font=("Consolas" if sys.platform.startswith("win") else "DejaVu Sans Mono", 9),
                            padx=8, pady=6, state="disabled")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.text.tag_configure("warn", foreground="#B45309")
        self.text.tag_configure("err", foreground="#B91C1C")
        self.text.tag_configure("ok", foreground=C_ACCENT)
        self.text.tag_configure("time", foreground=C_MUTED)

    def write(self, msg):
        tag = None
        s = msg.strip()
        if s.startswith("!") or "chênh lệch" in s.lower():
            tag = "warn"
        elif s.startswith("LỖI") or s.startswith("Lỗi"):
            tag = "err"
        elif s.startswith("Đã tạo") or s.startswith("HOÀN THÀNH") or "khớp" in s:
            tag = "ok"
        self.text.configure(state="normal")
        self.text.insert("end", _dt.datetime.now().strftime("%H:%M:%S  "), "time")
        self.text.insert("end", msg + "\n", tag)
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")


class Grid(tk.Frame):
    """Bảng dữ liệu có thanh cuộn."""

    def __init__(self, master, columns=(), height=10, show_headings=True):
        super().__init__(master, bg=C_CARD)
        self.tree = ttk.Treeview(self, columns=columns, show="headings" if show_headings else "", height=height)
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background="#F9FAFB")
        self.tree.tag_configure("sub1", foreground="#1D4ED8")
        self.tree.tag_configure("sub2", foreground="#B91C1C")
        self.tree.tag_configure("done", foreground=C_ACCENT)
        self.tree.tag_configure("error", foreground="#B91C1C")

    def set_columns(self, headers, widths=None):
        cols = ["c%d" % i for i in range(len(headers))]
        self.tree.configure(columns=cols)
        for i, h in enumerate(headers):
            self.tree.heading(cols[i], text=h)
            w = widths[i] if widths and i < len(widths) else max(80, min(260, len(str(h)) * 7))
            self.tree.column(cols[i], width=w, minwidth=40, stretch=False)

    def clear(self):
        self.tree.delete(*self.tree.get_children())


# ---------------------------------------------------------------------------
# Ứng dụng
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self.store = custom.TemplateStore(self.settings.dir)
        self.q = queue.Queue()
        self.busy = False
        self.cancel = threading.Event()

        self.title("%s %s — Tác giả: %s" % (APP_NAME, VERSION, AUTHOR))
        self.configure(bg=C_BG)
        geo = self.settings.get("cua_so") or "1180x760"
        try:
            self.geometry(geo)
        except tk.TclError:
            self.geometry("1180x760")
        self.minsize(980, 640)
        self._set_icon()
        self._styles()

        self.opt_tomau = tk.BooleanVar(value=_flag(self.settings.form.get("Tomau", "1")))
        self.opt_kiemtra = tk.BooleanVar(value=_flag(self.settings.form.get("Kiemtra", "0")))
        self.opt_gl = tk.BooleanVar(value=self.settings.flag("kiem_tra_gl"))
        self.opt_rt06 = tk.BooleanVar(value=self.settings.flag("kem_rt06"))
        self.opt_rt02 = tk.BooleanVar(value=self.settings.flag("kem_rt02"))
        self.opt_open = tk.BooleanVar(value=self.settings.flag("mo_file"))
        self.opt_skip = tk.BooleanVar(value=self.settings.flag("bo_qua_da_co"))
        self.opt_only_br = tk.BooleanVar(value=self.settings.flag("chi_cn_danh_muc"))

        self._layout()
        self.bind("<Escape>", lambda e: self.on_close())
        self.bind("<F5>", lambda e: self.pages["single"].run())
        self.bind("<Control-o>", lambda e: self.pages["single"].pick_file())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(120, self._poll)
        self.show("single")

    # ------------------------------------------------------------------ ui
    def _set_icon(self):
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        ico = os.path.join(base, "assets", "app.ico")
        png = os.path.join(base, "assets", "app.png")
        try:
            if sys.platform.startswith("win") and os.path.exists(ico):
                self.iconbitmap(ico)
            elif os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.iconphoto(True, self._icon_img)
        except tk.TclError:
            pass

    def _styles(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", font=(FONT, 10), background=C_CARD, foreground=C_TEXT)
        st.configure("TFrame", background=C_CARD)
        st.configure("TLabel", background=C_CARD, foreground=C_TEXT)
        st.configure("Muted.TLabel", background=C_CARD, foreground=C_MUTED, font=(FONT, 9))
        st.configure("TCheckbutton", background=C_CARD)
        st.map("TCheckbutton", background=[("active", C_CARD)])
        st.configure("TRadiobutton", background=C_CARD)
        st.configure("TButton", padding=(12, 6), background="#EEF0F3", bordercolor=C_BORDER)
        st.map("TButton", background=[("active", "#E2E5EA")])
        st.configure("Primary.TButton", font=(FONT, 11, "bold"), padding=(22, 10), background=C_PRIMARY,
                     foreground="white", bordercolor=C_PRIMARY_DARK)
        st.map("Primary.TButton", background=[("disabled", "#C9A3AE"), ("active", C_PRIMARY_DARK)],
               foreground=[("disabled", "#F5F5F5")])
        st.configure("Accent.TButton", padding=(14, 6), background=C_ACCENT, foreground="white", bordercolor=C_ACCENT)
        st.map("Accent.TButton", background=[("active", "#0B6430")])
        st.configure("Treeview", rowheight=24, fieldbackground=C_CARD, bordercolor=C_BORDER)
        st.configure("Treeview.Heading", font=(FONT, 9, "bold"), background="#EEF0F3", relief="flat")
        st.map("Treeview", background=[("selected", "#F4D7DE")], foreground=[("selected", C_TEXT)])
        st.configure("TNotebook", background=C_CARD, borderwidth=0)
        st.configure("TNotebook.Tab", padding=(14, 6), font=(FONT, 9, "bold"))
        st.map("TNotebook.Tab", background=[("selected", C_CARD)], foreground=[("selected", C_PRIMARY)])
        st.configure("Horizontal.TProgressbar", troughcolor="#E5E7EB", background=C_ACCENT, thickness=10)
        st.configure("TEntry", padding=4)
        st.configure("TCombobox", padding=4)

    def _layout(self):
        # Thanh tiêu đề
        top = tk.Frame(self, bg=C_PRIMARY, height=58)
        top.pack(side="top", fill="x")
        top.pack_propagate(False)
        tk.Label(top, text="  " + APP_TITLE, bg=C_PRIMARY, fg="white", font=(FONT, 15, "bold")).pack(
            side="left", padx=(10, 6))
        tk.Label(top, text="Agribank · Quyết toán niên độ · v" + VERSION, bg=C_PRIMARY, fg="#F3D3DB",
                 font=(FONT, 9)).pack(side="left", pady=(6, 0))
        tk.Label(top, text="Tác giả: " + AUTHOR + "  ", bg=C_PRIMARY, fg="white", font=(FONT, 10, "bold")).pack(
            side="right", padx=10)

        # Thanh trạng thái
        bar = tk.Frame(self, bg="#E9EBEF", height=30)
        bar.pack(side="bottom", fill="x")
        self.status = tk.Label(bar, text="Sẵn sàng", bg="#E9EBEF", fg=C_TEXT, anchor="w", font=(FONT, 9))
        self.status.pack(side="left", padx=10, fill="x", expand=True)
        self.progress = ttk.Progressbar(bar, length=260, mode="determinate", maximum=100)
        self.progress.pack(side="right", padx=10, pady=6)

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True)

        side = tk.Frame(body, bg=C_SIDEBAR, width=210)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        self.nav = {}
        items = [("single", "Tạo mẫu biểu"), ("batch", "Xử lý hàng loạt"),
                 ("templates", "Mẫu biểu tuỳ chỉnh"), ("columns", "Cột dữ liệu"),
                 ("settings", "Cài đặt"), ("help", "Hướng dẫn"), ("about", "Giới thiệu")]
        tk.Label(side, text="CHỨC NĂNG", bg=C_SIDEBAR, fg="#9CA3AF", font=(FONT, 8, "bold"), anchor="w").pack(
            fill="x", padx=18, pady=(16, 6))
        for key, text in items:
            b = tk.Label(side, text=text, bg=C_SIDEBAR, fg="#E5E7EB", font=(FONT, 10), anchor="w", padx=18, pady=10,
                         cursor="hand2")
            b.pack(fill="x")
            b.bind("<Button-1>", lambda e, k=key: self.show(k))
            b.bind("<Enter>", lambda e, w=b, k=key: w.configure(bg=C_SIDEBAR_HOVER) if self.current != k else None)
            b.bind("<Leave>", lambda e, w=b, k=key: w.configure(bg=C_SIDEBAR) if self.current != k else None)
            self.nav[key] = b
        tk.Label(side, text="Phím tắt:\nCtrl+O  Chọn file\nF5        Tạo mẫu biểu\nEsc       Thoát", bg=C_SIDEBAR,
                 fg="#9CA3AF", font=(FONT, 8), justify="left", anchor="w").pack(side="bottom", fill="x", padx=18,
                                                                              pady=14)

        self.content = tk.Frame(body, bg=C_BG)
        self.content.pack(side="left", fill="both", expand=True, padx=14, pady=12)
        self.current = None
        self.pages = {
            "single": SinglePage(self.content, self),
            "batch": BatchPage(self.content, self),
            "templates": TemplatePage(self.content, self),
            "columns": ColumnsPage(self.content, self),
            "settings": SettingsPage(self.content, self),
            "help": HelpPage(self.content, self),
            "about": AboutPage(self.content, self),
        }

    def show(self, key):
        if self.current:
            self.pages[self.current].pack_forget()
            self.nav[self.current].configure(bg=C_SIDEBAR, fg="#E5E7EB", font=(FONT, 10))
        self.current = key
        self.pages[key].pack(fill="both", expand=True)
        self.nav[key].configure(bg=C_PRIMARY, fg="white", font=(FONT, 10, "bold"))
        if hasattr(self.pages[key], "on_show"):
            self.pages[key].on_show()

    # ------------------------------------------------------------- options
    def options(self):
        return {"tomau": self.opt_tomau.get(), "kiemtra": self.opt_kiemtra.get(), "kiem_tra_gl": self.opt_gl.get(),
                "kem_rt06": self.opt_rt06.get(), "kem_rt02": self.opt_rt02.get()}

    def save_options(self):
        s = self.settings
        s.form["Tomau"] = "1" if self.opt_tomau.get() else "0"
        s.form["Kiemtra"] = "1" if self.opt_kiemtra.get() else "0"
        s.set("kiem_tra_gl", self.opt_gl.get())
        s.set("kem_rt06", self.opt_rt06.get())
        s.set("kem_rt02", self.opt_rt02.get())
        s.set("mo_file", self.opt_open.get())
        s.set("bo_qua_da_co", self.opt_skip.get())
        s.set("chi_cn_danh_muc", self.opt_only_br.get())

    def options_panel(self, parent, batch=False):
        f = tk.Frame(parent, bg=C_CARD)
        opts = [(self.opt_tomau, "Tô màu các dòng cộng"),
                (self.opt_kiemtra, "Chế độ xem thử (chỉ 50 dòng đầu)"),
                (self.opt_gl, "Đối chiếu cân đối GL (glcb41) & lập Mẫu 30"),
                (self.opt_rt06, "Tạo kèm Mẫu 06 khi làm Mẫu 05"),
                (self.opt_rt02, "Tạo kèm dữ liệu Mẫu 02 khi làm Mẫu 15a")]
        if batch:
            opts += [(self.opt_skip, "Bỏ qua mẫu đã có file kết quả"),
                     (self.opt_only_br, "Chỉ xử lý chi nhánh trong danh mục")]
        else:
            opts.append((self.opt_open, "Mở file Excel sau khi tạo xong"))
        for i, (var, text) in enumerate(opts):
            ttk.Checkbutton(f, text=text, variable=var).grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 24),
                                                             pady=3)
        return f

    # -------------------------------------------------------------- worker
    def start(self, target, *args):
        if self.busy:
            messagebox.showinfo(APP_NAME, "Chương trình đang xử lý, vui lòng chờ.")
            return False
        self.busy = True
        self.cancel.clear()
        self.save_options()
        try:
            self.settings.save()
        except OSError:
            pass

        def wrap():
            try:
                target(*args)
            except Exception as e:  # noqa: BLE001
                self.q.put(("error", "%s\n\n%s" % (e, traceback.format_exc(limit=3))))
            finally:
                self.q.put(("idle",))
        threading.Thread(target=wrap, daemon=True).start()
        return True

    def post(self, *msg):
        self.q.put(msg)

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    self.progress["value"] = msg[1]
                    if msg[2]:
                        self.status.configure(text=msg[2])
                elif kind == "status":
                    self.status.configure(text=msg[1])
                elif kind == "call":
                    msg[1](*msg[2:])
                elif kind == "error":
                    self.status.configure(text="Có lỗi xảy ra")
                    messagebox.showerror(APP_NAME, "Có lỗi khi xử lý:\n\n" + msg[1])
                elif kind == "idle":
                    self.busy = False
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def on_close(self):
        if self.busy and not messagebox.askyesno(APP_NAME, "Đang xử lý dở. Vẫn thoát chương trình?"):
            return
        self.save_options()
        try:
            self.settings.set("cua_so", self.geometry())
            self.settings.save()
        except OSError:
            pass
        self.destroy()


def _flag(v):
    return str(v).strip() in ("1", "true", "True")


# ---------------------------------------------------------------------------
# Trang: Tạo mẫu biểu (từng file)
# ---------------------------------------------------------------------------


class SinglePage(tk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, bg=C_BG)
        self.app = app
        self.job = None
        self.last_paths = []

        top = tk.Frame(self, bg=C_BG)
        top.pack(fill="x")
        c1 = Card(top, "1. Chọn file dữ liệu quyết toán (CSV xuất từ IPCAS)")
        c1.pack(side="left", fill="both", expand=True)
        b = c1.body
        row = tk.Frame(b, bg=C_CARD)
        row.pack(fill="x")
        self.path = tk.StringVar(value=app.settings.form.get("fCSV", ""))
        ttk.Entry(row, textvariable=self.path).pack(side="left", fill="x", expand=True, ipady=2)
        ttk.Button(row, text="Chọn file...", command=self.pick_file).pack(side="left", padx=(8, 0))
        info = tk.Frame(b, bg=C_CARD)
        info.pack(fill="x", pady=(10, 0))
        ttk.Label(info, text="Mã chi nhánh:").grid(row=0, column=0, sticky="w")
        self.brcd = tk.StringVar()
        e = ttk.Entry(info, textvariable=self.brcd, width=8)
        e.grid(row=0, column=1, sticky="w", padx=(6, 14))
        self.brcd.trace_add("write", lambda *a: self._update_brname())
        self.brname = ttk.Label(info, text="", foreground=C_PRIMARY)
        self.brname.grid(row=0, column=2, sticky="w")
        ttk.Label(info, text="Mẫu biểu:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.form_var = tk.StringVar()
        self.form_cb = ttk.Combobox(info, textvariable=self.form_var, state="readonly", width=70)
        self.form_cb.grid(row=1, column=1, columnspan=3, sticky="w", padx=(6, 0), pady=(8, 0))
        self.files_lbl = ttk.Label(info, text="", style="Muted.TLabel", wraplength=760, justify="left")
        self.files_lbl.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))
        self.refresh_forms()

        c2 = Card(self, "2. Tuỳ chọn & nơi lưu kết quả")
        c2.pack(fill="x", pady=(10, 0))
        app.options_panel(c2.body).pack(fill="x")
        out = tk.Frame(c2.body, bg=C_CARD)
        out.pack(fill="x", pady=(8, 0))
        ttk.Label(out, text="Lưu kết quả vào:").pack(side="left")
        self.out_dir = tk.StringVar(value=app.settings.get("thu_muc_ket_qua"))
        ttk.Entry(out, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out, text="Chọn...", command=self.pick_out).pack(side="left")
        ttk.Label(c2.body, text="Để trống = thư mục 'Result' nằm cùng thư mục file dữ liệu.",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        act = tk.Frame(self, bg=C_BG)
        act.pack(fill="x", pady=10)
        self.run_btn = ttk.Button(act, text="▶  TẠO MẪU BIỂU  (F5)", style="Primary.TButton", command=self.run)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(act, text="Mở file kết quả", command=self.open_result, state="disabled")
        self.open_btn.pack(side="left", padx=(10, 0))
        ttk.Button(act, text="Mở thư mục kết quả", command=self.open_folder).pack(side="left", padx=(8, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.log = LogBox(nb)
        self.preview = Grid(nb, height=8)
        nb.add(self.log, text="  Nhật ký xử lý  ")
        nb.add(self.preview, text="  Xem trước kết quả  ")
        self.nb = nb
        if self.path.get() and os.path.isfile(self.path.get()):
            self.load_file(self.path.get(), quiet=True)

    def on_show(self):
        self.refresh_forms()

    def refresh_forms(self):
        self.choices = form_choices(self.app.store)
        self.form_cb["values"] = [c[1] for c in self.choices]
        if self.job:
            self._select_form(self.job.code)

    def _select_form(self, code):
        for c, label in self.choices:
            if c == code:
                self.form_var.set(label)
                return
        self.form_var.set("")

    def selected_code(self):
        label = self.form_var.get()
        for c, lab in self.choices:
            if lab == label:
                return c
        return ""

    def _update_brname(self):
        name = self.app.settings.branch_name(self.brcd.get().strip())
        self.brname.configure(text=name or "(chưa có trong danh mục chi nhánh)")

    def pick_file(self):
        init = os.path.dirname(self.path.get()) if self.path.get() else ""
        p = filedialog.askopenfilename(title="Chọn file quyết toán", initialdir=init or None,
                                       filetypes=[("File quyết toán", "*.csv *.CSV *.txt"), ("Tất cả", "*.*")])
        if p:
            self.load_file(p)

    def load_file(self, p, quiet=False):
        self.path.set(p)
        self.app.settings.form["fCSV"] = p
        job = engine.job_from_file(p)
        self.job = job
        self.brcd.set(job.brcd)
        self._select_form(job.code)
        files = engine.describe_files(job.files)
        msg = "Sẽ ghép %d file: %s" % (len(files), "; ".join(files)) if len(files) > 1 else "File: " + files[0]
        if not job.code or not engine.is_known(job.code, self.app.store):
            msg += "\n⚠ Không nhận ra mẫu từ tên file (đúng dạng: 5500_rt13.csv). Hãy chọn mẫu ở ô trên."
        try:
            info = engine.quick_info(p, self.app.settings.get("ma_ky_tu"))
            msg += "\nMã ký tự: %s · Phân cách: %s · Số cột: %d" % (
                info.encoding, "TAB" if info.delimiter == "\t" else info.delimiter, len(info.header))
        except OSError as e:
            msg += "\nKhông đọc được file: %s" % e
        self.files_lbl.configure(text=msg)
        if not quiet:
            self.log.write("Đã chọn %s" % os.path.basename(p))

    def pick_out(self):
        d = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if d:
            self.out_dir.set(d)

    def run(self):
        p = self.path.get().strip()
        if not p or not os.path.isfile(p):
            messagebox.showwarning(APP_NAME, "Hãy chọn file dữ liệu CSV trước.")
            return
        code = self.selected_code()
        if not code:
            messagebox.showwarning(APP_NAME, "Hãy chọn mẫu biểu cần tạo.")
            return
        brcd = self.brcd.get().strip()
        if not brcd:
            messagebox.showwarning(APP_NAME, "Hãy nhập mã chi nhánh (4 số).")
            return
        job = engine.job_from_file(p, brcd, code)
        self.app.settings.set("thu_muc_ket_qua", self.out_dir.get().strip())
        self.log.write("─" * 60)
        self.log.write("Bắt đầu: %s_%s — %s" % (brcd, code, engine.form_label(code)))
        self.run_btn.state(["disabled"])
        opts = self.app.options()
        self.app.start(self._work, job, opts, self.out_dir.get().strip() or None)

    def _work(self, job, opts, out_dir):
        app = self.app
        runner = engine.Runner(app.settings, app.store,
                               log=lambda m: app.post("call", self.log.write, m),
                               progress=lambda pc, t="": app.post("progress", pc, t), cancel=app.cancel)
        try:
            res = runner.run(job, opts, out_dir)
        except Exception as e:  # noqa: BLE001
            app.post("call", self.log.write, "LỖI: %s" % e)
            app.post("call", self.run_btn.state, ["!disabled"])
            raise
        app.post("call", self._done, res)

    def _done(self, res):
        self.run_btn.state(["!disabled"])
        self.last_paths = res.paths
        self.open_btn.state(["!disabled"] if res.paths else ["disabled"])
        self.log.write("HOÀN THÀNH: " + ", ".join(os.path.basename(p) for p in res.paths))
        if res.preview:
            headers, rows = res.preview
            self.preview.set_columns(headers)
            self.preview.clear()
            for i, r in enumerate(rows):
                vals = []
                for v in r:
                    if isinstance(v, float):
                        vals.append("{:,.2f}".format(v))
                    elif isinstance(v, int) and not isinstance(v, bool):
                        vals.append("{:,}".format(v))
                    elif isinstance(v, (_dt.date, _dt.datetime)):
                        vals.append(vn.fmt_date(v))
                    else:
                        vals.append("" if v is None else str(v))
                self.preview.tree.insert("", "end", values=vals, tags=("odd",) if i % 2 else ())
        self.app.status.configure(text="Đã tạo %d file kết quả" % len(res.paths))
        if res.paths and self.app.opt_open.get():
            for p in res.paths:
                open_path(p)
        elif res.paths:
            if messagebox.askyesno(APP_NAME, "Đã tạo xong:\n%s\n\nMở file ngay?" % "\n".join(
                    os.path.basename(p) for p in res.paths)):
                for p in res.paths:
                    open_path(p)

    def open_result(self):
        for p in self.last_paths:
            if os.path.exists(p):
                open_path(p)

    def open_folder(self):
        d = self.out_dir.get().strip()
        if not d and self.last_paths:
            d = os.path.dirname(self.last_paths[0])
        if not d and self.path.get():
            d = os.path.join(os.path.dirname(self.path.get()), "Result")
        if d and os.path.isdir(d):
            open_path(d)
        else:
            messagebox.showinfo(APP_NAME, "Chưa có thư mục kết quả.")


# ---------------------------------------------------------------------------
# Trang: Xử lý hàng loạt (toàn tỉnh)
# ---------------------------------------------------------------------------


class BatchPage(tk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, bg=C_BG)
        self.app = app
        self.jobs = []
        c = Card(self, "Xử lý hàng loạt cả thư mục (toàn tỉnh / nhiều chi nhánh)")
        c.pack(fill="x")
        row = tk.Frame(c.body, bg=C_CARD)
        row.pack(fill="x")
        self.folder = tk.StringVar(value=app.settings.get("thu_muc_hang_loat"))
        ttk.Entry(row, textvariable=self.folder).pack(side="left", fill="x", expand=True, ipady=2)
        ttk.Button(row, text="Chọn thư mục...", command=self.pick).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Quét thư mục", style="Accent.TButton", command=self.scan).pack(side="left", padx=(8, 0))
        ttk.Label(c.body, text="Chương trình tìm mọi file dạng <mã CN>_<mẫu>.csv (vd 5501_rt13.csv, 5501_rt13_2.csv), "
                               "tự ghép các phần của cùng mẫu. Kết quả lưu vào thư mục Result.",
                  style="Muted.TLabel", wraplength=880, justify="left").pack(anchor="w", pady=(6, 6))
        app.options_panel(c.body, batch=True).pack(fill="x")

        self.log = LogBox(self, height=7)
        self.log.pack(side="bottom", fill="x", pady=(10, 0))
        mid = Card(self, "")
        mid.pack(fill="both", expand=True, pady=(10, 0))
        bar = tk.Frame(mid.body, bg=C_CARD)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Chọn tất cả", command=lambda: self.mark_all(True)).pack(side="left")
        ttk.Button(bar, text="Bỏ chọn", command=lambda: self.mark_all(False)).pack(side="left", padx=6)
        self.count_lbl = ttk.Label(bar, text="", style="Muted.TLabel")
        self.count_lbl.pack(side="left", padx=10)
        self.stop_btn = ttk.Button(bar, text="■ Dừng", command=self.stop, state="disabled")
        self.stop_btn.pack(side="right")
        self.run_btn = ttk.Button(bar, text="▶  CHẠY CÁC MẪU ĐÃ CHỌN", style="Primary.TButton", command=self.run)
        self.run_btn.pack(side="right", padx=8)
        self.grid_ = Grid(mid.body, height=10)
        self.grid_.pack(fill="both", expand=True)
        heads = ["Chọn", "Mã CN", "Tên chi nhánh", "Mẫu", "Tên mẫu biểu", "Số file", "Trạng thái", "Kết quả"]
        self.grid_.set_columns(heads, [52, 62, 150, 62, 330, 60, 120, 240])
        self.grid_.tree.bind("<Button-1>", self.on_click)
        self.grid_.tree.bind("<Double-1>", self.on_double)

    def pick(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các file quyết toán (VD: X:\\20241231\\)")
        if d:
            self.folder.set(d)
            self.scan()

    def scan(self):
        d = self.folder.get().strip()
        if not d or not os.path.isdir(d):
            messagebox.showwarning(APP_NAME, "Hãy chọn thư mục chứa dữ liệu.")
            return
        self.app.settings.set("thu_muc_hang_loat", d)
        self.jobs = engine.discover(d, self.app.settings, self.app.store, self.app.opt_only_br.get())
        out_dir = self.app.settings.get("thu_muc_ket_qua") or os.path.join(d, "Result")
        self.grid_.clear()
        for i, j in enumerate(self.jobs):
            exists = os.path.exists(engine.result_path(out_dir, j.brcd, j.code))
            j.checked = not (exists and self.app.opt_skip.get())
            j.status = "Đã có kết quả" if exists else "Chờ xử lý"
            name = engine.form_label(j.code)
            t = self.app.store.get(j.code)
            if t and j.code not in forms.FORMS:
                name = t["name"] + " (tuỳ chỉnh)"
            self.grid_.tree.insert("", "end", iid=str(i), values=(
                "☑" if j.checked else "☐", j.brcd, self.app.settings.branch_name(j.brcd), j.code, name,
                len(j.files), j.status, ""))
        self._count()
        self.log.write("Quét %s: tìm thấy %d mẫu cần xử lý" % (d, len(self.jobs)))
        if not self.jobs:
            messagebox.showinfo(APP_NAME, "Không tìm thấy file nào đúng dạng <mã CN>_<mẫu>.csv trong thư mục này.")

    def _count(self):
        n = sum(1 for j in self.jobs if getattr(j, "checked", False))
        self.count_lbl.configure(text="Đã chọn %d / %d mẫu" % (n, len(self.jobs)))

    def mark_all(self, val):
        for i, j in enumerate(self.jobs):
            j.checked = val
            self.grid_.tree.set(str(i), "c0", "☑" if val else "☐")
        self._count()

    def on_click(self, ev):
        if self.grid_.tree.identify_column(ev.x) != "#1":
            return
        iid = self.grid_.tree.identify_row(ev.y)
        if not iid:
            return
        j = self.jobs[int(iid)]
        j.checked = not getattr(j, "checked", False)
        self.grid_.tree.set(iid, "c0", "☑" if j.checked else "☐")
        self._count()

    def on_double(self, ev):
        iid = self.grid_.tree.identify_row(ev.y)
        if iid:
            j = self.jobs[int(iid)]
            for p in j.results:
                open_path(p)

    def stop(self):
        self.app.cancel.set()
        self.log.write("! Đang dừng sau mẫu hiện tại...")

    def run(self):
        todo = [(i, j) for i, j in enumerate(self.jobs) if getattr(j, "checked", False)]
        if not todo:
            messagebox.showinfo(APP_NAME, "Chưa chọn mẫu nào. Hãy bấm 'Quét thư mục' rồi tích chọn.")
            return
        self.run_btn.state(["disabled"])
        self.stop_btn.state(["!disabled"])
        self.app.start(self._work, todo, self.app.options())

    def _work(self, todo, opts):
        app = self.app
        out_dir = app.settings.get("thu_muc_ket_qua") or None
        runner = engine.Runner(app.settings, app.store, log=lambda m: app.post("call", self.log.write, "   " + m),
                               cancel=app.cancel)
        ok = err = 0
        total = len(todo)
        for n, (i, j) in enumerate(todo):
            if app.cancel.is_set():
                app.post("call", self.log.write, "! Đã dừng theo yêu cầu.")
                break
            app.post("call", self._set, i, "Đang xử lý...", "", None)
            app.post("progress", 100.0 * n / total, "Đang xử lý %d/%d: %s_%s" % (n + 1, total, j.brcd, j.code))
            app.post("call", self.log.write, "[%d/%d] %s_%s — %s" % (n + 1, total, j.brcd, j.code,
                                                                     engine.form_label(j.code)))
            try:
                res = runner.run(j, opts, out_dir)
                j.results = res.paths
                ok += 1
                app.post("call", self._set, i, "✔ Xong", ", ".join(os.path.basename(p) for p in res.paths), "done")
            except Exception as e:  # noqa: BLE001
                err += 1
                app.post("call", self._set, i, "✖ Lỗi", str(e), "error")
                app.post("call", self.log.write, "LỖI %s_%s: %s" % (j.brcd, j.code, e))
        app.post("progress", 100, "Hoàn thành: %d mẫu thành công, %d lỗi" % (ok, err))
        app.post("call", self._finish, ok, err)

    def _set(self, i, status, result, tag):
        iid = str(i)
        self.grid_.tree.set(iid, "c6", status)
        if result:
            self.grid_.tree.set(iid, "c7", result)
        self.grid_.tree.item(iid, tags=(tag,) if tag else ())
        self.grid_.tree.see(iid)

    def _finish(self, ok, err):
        self.run_btn.state(["!disabled"])
        self.stop_btn.state(["disabled"])
        self.log.write("HOÀN THÀNH: %d mẫu thành công, %d lỗi" % (ok, err))
        d = self.app.settings.get("thu_muc_ket_qua") or os.path.join(self.folder.get(), "Result")
        if messagebox.askyesno(APP_NAME, "Đã xử lý xong, mời anh/chị kiểm tra kết quả!\n\n"
                                         "Thành công: %d   Lỗi: %d\n\nMở thư mục kết quả?" % (ok, err)):
            if os.path.isdir(d):
                open_path(d)


# ---------------------------------------------------------------------------
# Trang: Mẫu biểu tuỳ chỉnh
# ---------------------------------------------------------------------------


class TemplatePage(tk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, bg=C_BG)
        self.app = app
        c = Card(self, "Mẫu biểu tuỳ chỉnh — tự thiết kế mẫu mới trong Excel, tải lên là dùng được")
        c.pack(fill="x")
        txt = ("Khi có mẫu biểu mới (hoặc mẫu cũ thay đổi hình thức), anh/chị không cần sửa chương trình:\n"
               "  ①  Bấm 'Tạo mẫu khởi đầu từ file CSV' → chương trình sinh sẵn file Excel có tiêu đề, cột dữ liệu {c1}, {c2}...\n"
               "  ②  Mở file đó, sửa tiêu đề/cột/định dạng/chữ ký cho đúng mẫu mới rồi lưu lại.\n"
               "  ③  Bấm 'Tải lên mẫu mới', đặt mã mẫu (vd: rt25). Từ đó file dữ liệu 5501_rt25.csv sẽ tự dùng mẫu này.\n"
               "Có thể dùng mẫu tải lên để THAY THẾ mẫu có sẵn (vd rt13) khi Hội sở thay đổi hình thức báo cáo.")
        ttk.Label(c.body, text=txt, justify="left", wraplength=900).pack(anchor="w")
        bar = tk.Frame(c.body, bg=C_CARD)
        bar.pack(fill="x", pady=(10, 0))
        ttk.Button(bar, text="① Tạo mẫu khởi đầu từ file CSV...", command=self.starter).pack(side="left")
        ttk.Button(bar, text="③ Tải lên mẫu mới (.xlsx)...", style="Accent.TButton", command=self.upload).pack(
            side="left", padx=8)
        ttk.Button(bar, text="Bảng mã {…}", command=self.codes_help).pack(side="left")

        lst = Card(self, "Danh sách mẫu đã tải lên")
        lst.pack(fill="both", expand=True, pady=(10, 0))
        self.grid_ = Grid(lst.body, height=10)
        self.grid_.pack(fill="both", expand=True)
        self.grid_.set_columns(["Mã mẫu", "Tên mẫu biểu", "Thay mẫu có sẵn", "Ngày tải lên", "File"],
                               [80, 380, 120, 110, 200])
        bar2 = tk.Frame(lst.body, bg=C_CARD)
        bar2.pack(fill="x", pady=(8, 0))
        ttk.Button(bar2, text="Mở file mẫu", command=self.open_tpl).pack(side="left")
        ttk.Button(bar2, text="Cập nhật (tải lại file đã sửa)...", command=self.reupload).pack(side="left", padx=8)
        ttk.Button(bar2, text="Lưu bản sao ra...", command=self.export).pack(side="left")
        ttk.Button(bar2, text="Bật/tắt thay mẫu có sẵn", command=self.toggle_override).pack(side="left", padx=8)
        ttk.Button(bar2, text="Xoá mẫu", command=self.delete).pack(side="right")
        self.refresh()

    def on_show(self):
        self.refresh()

    def refresh(self):
        self.app.store.load()
        self.grid_.clear()
        for t in self.app.store.items:
            builtin = t["code"] in forms.FORMS
            self.grid_.tree.insert("", "end", iid=t["code"], values=(
                t["code"], t["name"], ("Có" if t.get("override") else "Không") if builtin else "—",
                t.get("added", ""), t["file"]))

    def _sel(self):
        s = self.grid_.tree.selection()
        if not s:
            messagebox.showinfo(APP_NAME, "Hãy chọn một mẫu trong danh sách.")
            return None
        return self.app.store.get(s[0])

    def starter(self):
        p = filedialog.askopenfilename(title="Chọn file CSV dữ liệu của mẫu mới",
                                       filetypes=[("CSV", "*.csv *.CSV *.txt"), ("Tất cả", "*.*")])
        if not p:
            return
        _b, raw = reader.parse_filename(p)
        code = raw or "rtxx"
        title = simpledialog.askstring(APP_NAME, "Tên (tiêu đề) mẫu biểu:", initialvalue="SAO KÊ CHI TIẾT ...",
                                       parent=self)
        if title is None:
            return
        out = filedialog.asksaveasfilename(title="Lưu file mẫu khởi đầu", defaultextension=".xlsx",
                                           initialfile="Mau_%s.xlsx" % code, filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        try:
            data = reader.read_sources(reader.find_parts(p)[:1], self.app.settings.get("ma_ky_tu"), 0, 500)
            custom.create_starter(data, out, code, title or "TÊN MẪU BIỂU")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, "Không tạo được mẫu: %s" % e)
            return
        if messagebox.askyesno(APP_NAME, "Đã tạo file mẫu khởi đầu:\n%s\n\nMở file để chỉnh sửa ngay?" % out):
            open_path(out)

    def upload(self, preset=None):
        p = filedialog.askopenfilename(title="Chọn file mẫu biểu (.xlsx)", filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            info = custom.inspect_template(p)
        except custom.TemplateError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        base = os.path.splitext(os.path.basename(p))[0].lower().replace("mau_", "")
        code = preset["code"] if preset else simpledialog.askstring(
            APP_NAME, "Mã mẫu (dùng trong tên file dữ liệu, vd rt25 → 5501_rt25.csv):", initialvalue=base, parent=self)
        if not code:
            return
        code = forms.resolve(code.strip().lower())
        name = preset["name"] if preset else simpledialog.askstring(
            APP_NAME, "Tên mẫu biểu:", initialvalue=info.get("TEN_MAU", code), parent=self)
        if name is None:
            return
        override = False
        if code in forms.FORMS:
            override = preset.get("override", False) if preset else messagebox.askyesno(
                APP_NAME, "Mã '%s' trùng với mẫu có sẵn:\n%s\n\nDùng mẫu tải lên THAY THẾ mẫu có sẵn?" % (
                    code, forms.FORMS[code]["name"]))
        try:
            self.app.store.add(p, code, name, override)
        except custom.TemplateError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        self.refresh()
        self.app.pages["single"].refresh_forms()
        toks = ", ".join("{%s}" % t for t in info["tokens"][:12])
        messagebox.showinfo(APP_NAME, "Đã tải lên mẫu '%s' (%s).\n\nDòng dữ liệu: dòng %d\nCác mã: %s%s" % (
            name, code, info["data_row"], toks, " ..." if len(info["tokens"]) > 12 else ""))

    def reupload(self):
        t = self._sel()
        if t:
            self.upload(preset=t)

    def open_tpl(self):
        t = self._sel()
        if t:
            open_path(self.app.store.path_of(t))

    def export(self):
        t = self._sel()
        if not t:
            return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=t["file"],
                                           filetypes=[("Excel", "*.xlsx")])
        if out:
            import shutil
            shutil.copyfile(self.app.store.path_of(t), out)

    def toggle_override(self):
        t = self._sel()
        if not t:
            return
        if t["code"] not in forms.FORMS:
            messagebox.showinfo(APP_NAME, "Mẫu này không trùng mã với mẫu có sẵn.")
            return
        t["override"] = not t.get("override")
        self.app.store.save()
        self.refresh()

    def delete(self):
        t = self._sel()
        if t and messagebox.askyesno(APP_NAME, "Xoá mẫu '%s' (%s)?" % (t["name"], t["code"])):
            self.app.store.remove(t["code"])
            self.refresh()
            self.app.pages["single"].refresh_forms()

    def codes_help(self):
        messagebox.showinfo(APP_NAME, CODES_HELP)


CODES_HELP = """BẢNG MÃ DÙNG TRONG FILE MẪU

Dòng dữ liệu:
  {c5}            giá trị cột thứ 5 của file CSV
  {Tên khách hàng} giá trị theo tên cột trong dòng tiêu đề CSV
  {STT}           số thứ tự
  {c3|phai:9}     9 ký tự cuối      {c3|trai:4}   4 ký tự đầu
  {c6|bo:1}       bỏ 1 ký tự đầu    {c4|hoa}      viết HOA
  {c7|ngay}       yyyymmdd → ngày   {c10|so}      chuyển thành số
  {c10|chu:c8}    đọc số thành chữ theo loại tiền ở cột 8
  {c2}{c3|phai:9} ghép nhiều mã trong một ô

Dòng cộng / chữ ký (dưới dòng dữ liệu):
  {TONG}  cộng cột chứa ô này     {TONG:H}  cộng cột H     {DEM}  số dòng

Mã chung (đặt ở bất kỳ đâu):
  {MA_CN} {TEN_CN} {NGAY_QT} {NGAY_QT_CHU} {NGAY_LAP} {LOAI_QT}
  {MA_MAU} {TEN_MAU} {NAM} {DDMM}

Sheet CAU_HINH (tuỳ chọn): SAP_XEP, NHOM_THEO, CONG_COT, NHAN_NHOM, COT_NHAN, LOC, BO_DONG_DAU"""


# ---------------------------------------------------------------------------
# Trang: Cột dữ liệu (ánh xạ cột cho mẫu có sẵn)
# ---------------------------------------------------------------------------


class ColumnsPage(tk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, bg=C_BG)
        self.app = app
        self.headers = []
        c = Card(self, "Cột dữ liệu của các mẫu có sẵn")
        c.pack(fill="x")
        ttk.Label(c.body, justify="left", wraplength=900, text=(
            "Nếu năm nay IPCAS thay đổi thứ tự cột trong file xuất, chỉ cần sửa số cột tại đây (không phải sửa "
            "chương trình). Bấm 'Nạp file CSV để đối chiếu' để xem tên cột thực tế bên cạnh. Nháy đúp vào dòng "
            "để sửa số cột.")).pack(anchor="w")
        row = tk.Frame(c.body, bg=C_CARD)
        row.pack(fill="x", pady=(8, 0))
        ttk.Label(row, text="Mẫu:").pack(side="left")
        self.codes = [c for c, s in forms.FORMS.items() if "fields" in s] + ["gl"]
        labels = ["%s — %s" % (c, forms.FORMS[c]["name"]) for c in self.codes[:-1]] + ["gl — File cân đối glcb41"]
        self.labels = labels
        self.form_var = tk.StringVar(value=labels[0])
        cb = ttk.Combobox(row, textvariable=self.form_var, values=labels, state="readonly", width=60)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self.load())
        ttk.Button(row, text="Nạp file CSV để đối chiếu...", command=self.load_csv).pack(side="left", padx=(8, 0))

        body = Card(self, "")
        body.pack(fill="both", expand=True, pady=(10, 0))
        self.grid_ = Grid(body.body, height=14)
        self.grid_.pack(fill="both", expand=True)
        self.grid_.set_columns(["Trường", "Ý nghĩa", "Cột số", "Mặc định", "Tên cột trong file đối chiếu"],
                               [110, 330, 70, 80, 300])
        self.grid_.tree.bind("<Double-1>", self.edit)
        bar = tk.Frame(body.body, bg=C_CARD)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Button(bar, text="Lưu thay đổi", style="Accent.TButton", command=self.save).pack(side="left")
        ttk.Button(bar, text="Khôi phục mặc định mẫu này", command=self.reset).pack(side="left", padx=8)
        self.code = self.codes[0]
        self.work = {}
        self.load()

    def current_code(self):
        return self.codes[self.labels.index(self.form_var.get())]

    def _fields(self, code):
        if code == "gl":
            return forms.GL_FIELDS
        return forms.field_defs(code)

    def load(self):
        self.code = self.current_code()
        m = self.app.settings.colmap(self.code) if self.code != "gl" else dict(
            forms.default_colmap("gl"), **{k: int(v) for k, v in (self.app.settings.colmaps.get("gl") or {}).items()})
        self.work = dict(m)
        self.render()

    def render(self):
        self.grid_.clear()
        for i, (key, default, desc) in enumerate(self._fields(self.code)):
            col = self.work.get(key, default)
            hname = self.headers[col - 1] if self.headers and 0 < col <= len(self.headers) else ""
            self.grid_.tree.insert("", "end", iid=key, values=(key, desc, col, default, hname),
                                   tags=("sub2",) if col != default else (("odd",) if i % 2 else ()))

    def edit(self, ev):
        iid = self.grid_.tree.identify_row(ev.y)
        if not iid:
            return
        cur = self.work.get(iid)
        v = simpledialog.askinteger(APP_NAME, "Số thứ tự cột cho '%s' (đếm từ 1):" % iid, initialvalue=cur,
                                    minvalue=1, maxvalue=500, parent=self)
        if v:
            self.work[iid] = v
            self.render()

    def load_csv(self):
        p = filedialog.askopenfilename(title="Chọn file CSV để đối chiếu tên cột",
                                       filetypes=[("CSV", "*.csv *.CSV *.txt"), ("Tất cả", "*.*")])
        if not p:
            return
        data = reader.read_sources([p], self.app.settings.get("ma_ky_tu"), 0, 5)
        self.headers = data.header
        _b, raw = reader.parse_filename(p)
        code = forms.resolve(raw)
        if code in self.codes:
            self.form_var.set(self.labels[self.codes.index(code)])
            self.load()
        else:
            self.render()

    def save(self):
        defaults = {k: d for k, d, _ in self._fields(self.code)}
        diff = {k: v for k, v in self.work.items() if defaults.get(k) != v}
        if diff:
            self.app.settings.colmaps[self.code] = diff
        else:
            self.app.settings.colmaps.pop(self.code, None)
        self.app.settings.save()
        messagebox.showinfo(APP_NAME, "Đã lưu cấu hình cột cho mẫu %s." % self.code)
        self.render()

    def reset(self):
        self.work = {k: d for k, d, _ in self._fields(self.code)}
        self.render()


# ---------------------------------------------------------------------------
# Trang: Cài đặt
# ---------------------------------------------------------------------------


class SettingsPage(tk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, bg=C_BG)
        self.app = app
        s = app.settings
        canvas = tk.Canvas(self, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=C_BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units")
                        if self.winfo_ismapped() else None)

        self.vars = {}
        c1 = Card(inner, "Thông tin kỳ quyết toán")
        c1.pack(fill="x")
        g = c1.body
        self._entry(g, 0, "Ngày quyết toán (dd/mm/yyyy):", "ngay_qt", 14)
        self._entry(g, 1, "Loại báo cáo (QT / ...):", "loai_qt", 14)
        self._entry(g, 2, "Tỷ giá USD (dùng khi file thiếu cột quy đổi):", "ty_gia_usd", 14)

        c2 = Card(inner, "Trình bày & dữ liệu")
        c2.pack(fill="x", pady=(10, 0))
        g = c2.body
        self._combo(g, 0, "Khổ giấy in:", "kho_giay", ["A4", "A3"])
        self._combo(g, 1, "Cỡ chữ:", "co_chu", ["9", "10", "11", "12"])
        self._combo(g, 2, "Đọc số hàng nghìn:", "doc_so", ["nghìn", "ngàn"])
        self._combo(g, 3, "Mã ký tự file CSV:", "ma_ky_tu", [k for k, _ in reader.ENCODINGS],
                    labels={k: v for k, v in reader.ENCODINGS})
        self.v_prio = tk.BooleanVar(value=s.flag("uu_tien_mau_tuy_chinh"))
        ttk.Checkbutton(g, text="Luôn ưu tiên mẫu tuỳ chỉnh khi trùng mã với mẫu có sẵn",
                        variable=self.v_prio).grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))

        c3 = Card(inner, "Người ký (in sẵn họ tên dưới chữ ký – để trống nếu ký tay)")
        c3.pack(fill="x", pady=(10, 0))
        self.signer_vars = {}
        for i, (k, label) in enumerate(SIGNERS):
            ttk.Label(c3.body, text=label + ":").grid(row=i, column=0, sticky="w", pady=3)
            v = tk.StringVar(value=s.signers.get(k, ""))
            ttk.Entry(c3.body, textvariable=v, width=40).grid(row=i, column=1, sticky="w", padx=8, pady=3)
            self.signer_vars[k] = v

        c4 = Card(inner, "Danh mục chi nhánh")
        c4.pack(fill="x", pady=(10, 0))
        self.br = Grid(c4.body, height=8)
        self.br.pack(fill="x")
        self.br.set_columns(["Mã chi nhánh", "Tên chi nhánh"], [120, 420])
        self.br.tree.bind("<Double-1>", lambda e: self.br_edit())
        bb = tk.Frame(c4.body, bg=C_CARD)
        bb.pack(fill="x", pady=(6, 0))
        ttk.Button(bb, text="Thêm", command=self.br_add).pack(side="left")
        ttk.Button(bb, text="Sửa", command=self.br_edit).pack(side="left", padx=6)
        ttk.Button(bb, text="Xoá", command=self.br_del).pack(side="left")
        ttk.Button(bb, text="Nhập từ QT.ini cũ...", command=self.import_ini).pack(side="left", padx=6)
        self.branches = list(s.branches)
        self.br_render()

        act = tk.Frame(inner, bg=C_BG)
        act.pack(fill="x", pady=12)
        ttk.Button(act, text="LƯU CÀI ĐẶT", style="Primary.TButton", command=self.save).pack(side="left")
        ttk.Label(act, text="  File cấu hình: " + s.path, background=C_BG, foreground=C_MUTED).pack(side="left")

    def _entry(self, g, r, label, key, width):
        ttk.Label(g, text=label).grid(row=r, column=0, sticky="w", pady=3)
        v = tk.StringVar(value=self.app.settings.get(key))
        ttk.Entry(g, textvariable=v, width=width).grid(row=r, column=1, sticky="w", padx=8, pady=3)
        self.vars[key] = v

    def _combo(self, g, r, label, key, values, labels=None):
        ttk.Label(g, text=label).grid(row=r, column=0, sticky="w", pady=3)
        disp = [labels.get(x, x) for x in values] if labels else values
        cur = self.app.settings.get(key)
        v = tk.StringVar(value=labels.get(cur, cur) if labels else cur)
        ttk.Combobox(g, textvariable=v, values=disp, state="readonly", width=30).grid(row=r, column=1, sticky="w",
                                                                                     padx=8, pady=3)
        self.vars[key] = (v, dict(zip(disp, values)))

    def br_render(self):
        self.br.clear()
        for i, (k, v) in enumerate(self.branches):
            self.br.tree.insert("", "end", iid=str(i), values=(k, v), tags=("odd",) if i % 2 else ())

    def br_add(self):
        code = simpledialog.askstring(APP_NAME, "Mã chi nhánh (4 số):", parent=self)
        if not code:
            return
        name = simpledialog.askstring(APP_NAME, "Tên chi nhánh:", parent=self) or ""
        self.branches = [b for b in self.branches if b[0] != code.strip()] + [(code.strip(), name.strip())]
        self.branches.sort()
        self.br_render()

    def br_edit(self):
        s = self.br.tree.selection()
        if not s:
            return
        i = int(s[0])
        k, v = self.branches[i]
        name = simpledialog.askstring(APP_NAME, "Tên chi nhánh %s:" % k, initialvalue=v, parent=self)
        if name is not None:
            self.branches[i] = (k, name.strip())
            self.br_render()

    def br_del(self):
        s = self.br.tree.selection()
        if s and messagebox.askyesno(APP_NAME, "Xoá chi nhánh đã chọn?"):
            del self.branches[int(s[0])]
            self.br_render()

    def import_ini(self):
        p = filedialog.askopenfilename(title="Chọn file QT.ini của chương trình cũ", filetypes=[("INI", "*.ini")])
        if not p:
            return
        old = Settings(path=p)
        self.branches = list(old.branches)
        self.br_render()
        messagebox.showinfo(APP_NAME, "Đã nhập %d chi nhánh. Bấm 'Lưu cài đặt' để ghi lại." % len(self.branches))

    def save(self):
        s = self.app.settings
        for key, v in self.vars.items():
            if isinstance(v, tuple):
                var, mp = v
                s.set(key, mp.get(var.get(), var.get()))
            else:
                s.set(key, v.get().strip())
        if vn.parse_date(s.get("ngay_qt")) is None:
            messagebox.showwarning(APP_NAME, "Ngày quyết toán không hợp lệ (dd/mm/yyyy).")
            return
        s.set("uu_tien_mau_tuy_chinh", self.v_prio.get())
        for k, v in self.signer_vars.items():
            s.signers[k] = v.get().strip()
        s.branches = list(self.branches)
        try:
            s.save()
        except OSError as e:
            messagebox.showerror(APP_NAME, "Không lưu được cài đặt: %s" % e)
            return
        self.app.pages["single"]._update_brname()
        messagebox.showinfo(APP_NAME, "Đã lưu cài đặt.")


# ---------------------------------------------------------------------------
# Trang: Hướng dẫn / Giới thiệu
# ---------------------------------------------------------------------------

HELP = """HƯỚNG DẪN SỬ DỤNG

1. CHUẨN BỊ DỮ LIỆU
   • Xuất các file quyết toán từ IPCAS, đặt tên dạng  <mã CN>_<mẫu>.csv
     Ví dụ: 5500_rt05.csv, 5501_rt13.csv, 5501_rt22_1.csv, 5501_rt22_2.csv
   • Dữ liệu chia nhiều file (rt13_1, rt13_2...): chương trình tự tìm và ghép nối.
   • Tự nhận mã ký tự Unicode (UTF-8, UTF-16 "Unicode LE") và TCVN3 (ABC); tự nhận dấu phẩy hoặc TAB.
   • (Tuỳ chọn) Đặt file cân đối tài khoản *glcb41* (.csv/.txt) cùng thư mục để đối chiếu số dư
     → chương trình thêm sheet CheckGL và Mẫu 30 vào file kết quả.

2. TẠO TỪNG MẪU  (mục "Tạo mẫu biểu")
   • Chọn file → chương trình tự nhận mã chi nhánh và mẫu biểu (có thể chọn lại).
   • "Chế độ xem thử" = xuất nhanh 50 dòng đầu để kiểm tra trước.
   • Bấm TẠO MẪU BIỂU (F5) và chờ thông báo hoàn thành.

3. XỬ LÝ HÀNG LOẠT  (toàn tỉnh)
   • Chọn thư mục (VD: X:\\20241231\\) → Quét → tích chọn → Chạy.
   • Kết quả lưu vào thư mục Result; mẫu đã có kết quả có thể bỏ qua.

4. CÁC MẪU CÓ SẴN
   05, 06 (tự tạo kèm 05), 13, 14, 15a, 15b, 16, 02 (dữ liệu xác nhận dư nợ từ 15a),
   18, 20a, 22, 24a, DP (danh sách xác nhận số dư tiền gửi tổ chức – MSIT81),
   HTT (sao kê dự thu cuối kỳ), 30 (tổng hợp đối chiếu cân đối – tự tạo).

5. LƯU Ý MẪU 22
   • Vấn tin màn hình GLST34: (1) Unearned Income + (2) Prepaid Expenses.
   • Chọn Export → Save as type: CSV with headers + Encoding: Unicode LE.
   • Lưu (1), (2) với tên: xxxx_rt22_1.CSV ; xxxx_rt22_2.CSV
   • Chọn 1 trong 2 file, chương trình tự tìm file còn lại.

6. MẪU MỚI / MẪU THAY ĐỔI
   • Mục "Mẫu biểu tuỳ chỉnh": tự thiết kế mẫu trong Excel và tải lên.
   • Mục "Cột dữ liệu": sửa số thứ tự cột khi IPCAS đổi cấu trúc file.

7. CÀI ĐẶT
   • Ngày quyết toán, loại báo cáo, tỷ giá, khổ giấy, người ký, danh mục chi nhánh.

Kết quả là file Excel (.xlsx) đã định dạng sẵn: tiêu đề, cột, dòng cộng (công thức SUBTOTAL),
in khổ ngang vừa 1 trang chiều rộng, lặp tiêu đề mỗi trang, đánh số trang.
Không cần cài Microsoft Excel để tạo file.

CHÚC QUYẾT TOÁN VUI VẺ!"""


class HelpPage(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=C_BG)
        c = Card(self, "Hướng dẫn sử dụng")
        c.pack(fill="both", expand=True)
        t = tk.Text(c.body, wrap="word", bg=C_CARD, relief="flat", font=(FONT, 10), padx=6, pady=4)
        sb = ttk.Scrollbar(c.body, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.insert("1.0", HELP)
        t.configure(state="disabled")
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")


class AboutPage(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=C_BG)
        c = Card(self, "")
        c.pack(fill="both", expand=True)
        box = tk.Frame(c.body, bg=C_CARD)
        box.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(box, text=APP_TITLE, bg=C_CARD, fg=C_PRIMARY, font=(FONT, 20, "bold")).pack()
        tk.Label(box, text="%s · Phiên bản %s" % (APP_NAME, VERSION), bg=C_CARD, fg=C_MUTED,
                 font=(FONT, 11)).pack(pady=(4, 18))
        tk.Label(box, text="Tác giả", bg=C_CARD, fg=C_MUTED, font=(FONT, 10)).pack()
        tk.Label(box, text=AUTHOR, bg=C_CARD, fg=C_TEXT, font=(FONT, 16, "bold")).pack(pady=(2, 18))
        tk.Label(box, justify="center", bg=C_CARD, fg=C_TEXT, font=(FONT, 10), text=(
            "Ứng dụng hỗ trợ lập mẫu biểu quyết toán niên độ từ dữ liệu IPCAS.\n"
            "Chuyển đổi và nâng cấp từ chương trình QT2024 (AutoIt).\n"
            "Xuất Excel trực tiếp – không cần cài Microsoft Excel.")).pack()


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    app = App()
    app.mainloop()
