"""Chạy toàn bộ mẫu trên dữ liệu giả lập và kiểm tra file kết quả."""

import os
import shutil
import sys
import tempfile
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from qt_app import engine, custom  # noqa: E402
from qt_app.settings import Settings  # noqa: E402

work = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="qt_test_")
samples = os.path.join(work, "samples")
if os.path.exists(samples):
    shutil.rmtree(samples)
os.system('%s "%s" "%s"' % (sys.executable, os.path.join(HERE, "make_samples.py"), samples))

s = Settings(path=os.path.join(work, "QT.ini"))
s.colmap_path = os.path.join(work, "cot.json")
s.set("ngay_qt", "31/12/2024")
s.signers["giam_doc"] = "Nguyễn Văn Giám Đốc"
store = custom.TemplateStore(work)

# mẫu tuỳ chỉnh rt99 từ file khởi đầu
data = engine.quick_info(os.path.join(samples, "5501_rt99.csv"))
starter = os.path.join(work, "mau_rt99.xlsx")
custom.create_starter(data, starter, "rt99", "SAO KÊ THỬ NGHIỆM MẪU TUỲ CHỈNH")
import openpyxl  # noqa: E402
wb = openpyxl.load_workbook(starter)
cfg = wb["CAU_HINH"]
for row in cfg.iter_rows(min_row=2):
    if row[0].value == "NHOM_THEO":
        row[1].value = "c3"
    if row[0].value == "CONG_COT":
        row[1].value = "J,K"
wb.save(starter)
store.add(starter, "rt99", "Mẫu thử nghiệm")

jobs = engine.discover(samples, s, store)
print("Tìm thấy %d job:" % len(jobs))
for j in jobs:
    print("  ", j.brcd, j.code, [os.path.basename(f) for f in j.files])

out = os.path.join(work, "Result")
runner = engine.Runner(s, store, log=lambda m: print("    " + m), cancel=threading.Event())
opts = {"tomau": True, "kiemtra": False, "kem_rt06": True, "kem_rt02": True, "kiem_tra_gl": True}
paths = []
ok = True
for j in jobs:
    print("== %s_%s" % (j.brcd, j.code))
    try:
        r = runner.run(j, opts, out)
        paths += r.paths
    except Exception:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        ok = False
print("\nKết quả:")
for p in paths:
    print("  ", os.path.basename(p), os.path.getsize(p))
print("OK" if ok else "CÓ LỖI")
print("WORK", work)
