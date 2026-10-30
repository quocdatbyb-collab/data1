"""Chụp màn hình các trang giao diện (chạy dưới Xvfb)."""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from qt_app import gui  # noqa: E402

out = sys.argv[1]
sample = sys.argv[2]
os.makedirs(out, exist_ok=True)
app = gui.App()
app.geometry("1180x760+0+0")
steps = []


def shot(name):
    app.update()
    subprocess.call(["import", "-window", "root", os.path.join(out, name + ".png")])


def s1():
    app.pages["single"].load_file(os.path.join(sample, "5501_rt13_1.csv"))
    app.update()
    app.opt_open.set(False)
    shot("01_single")
    app.pages["single"].run()


def s2():
    if app.busy:
        app.after(500, s2)
        return
    shot("02_single_done")
    app.pages["single"].nb.select(1)
    shot("03_preview")
    app.show("batch")
    app.pages["batch"].folder.set(sample)
    app.pages["batch"].scan()
    shot("04_batch")
    app.show("templates")
    shot("05_templates")
    app.show("columns")
    shot("06_columns")
    app.show("settings")
    shot("07_settings")
    app.show("about")
    shot("08_about")
    app.destroy()


def close_dialogs():
    # tự đóng messagebox (nếu có)
    for w in app.winfo_children():
        pass


import tkinter.messagebox as mb  # noqa: E402
mb.askyesno = lambda *a, **k: False
mb.showinfo = lambda *a, **k: None
app.after(800, s1)
app.after(1500, s2)
app.mainloop()
print("done")
