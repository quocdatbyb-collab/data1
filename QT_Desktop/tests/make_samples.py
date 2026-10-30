"""Sinh dữ liệu giả lập (không phải dữ liệu thật) để kiểm thử các mẫu biểu."""

import os
import random
import sys

random.seed(7)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "samples")
os.makedirs(OUT, exist_ok=True)

NAMES = ["Nguyễn Văn An", "Trần Thị Bình", "Công ty TNHH Minh Phát", "Lê Hoàng Cường", "HTX Nông nghiệp Phú An",
         "Phạm Thị Dung", "Công ty CP Gỗ Bình Dương", "Võ Minh Em"]
TCVN3 = {"à": "µ", "á": "¸", "ạ": "¹", "ả": "¶", "ã": "·", "â": "©", "ấ": "Ê", "ầ": "Ç", "ậ": "Ë", "ă": "¨",
         "đ": "®", "ê": "ª", "ệ": "Ö", "ễ": "Ô", "ố": "è", "ộ": "é", "ơ": "¬", "ư": "\xad", "ị": "Þ", "ì": "×",
         "ọ": "ä", "ọ": "ä", "ú": "ó", "ụ": "ô", "ủ": "ñ", "ứ": "ø", "ữ": "÷", "ó": "ã", "ò": "ß", "í": "Ý",
         "ồ": "å", "ổ": "æ", "ờ": "ê", "ợ": "î", "ớ": "í", "ế": "Õ", "ề": "Ò", "ể": "Ó", "ỉ": "Ø", "ỏ": "á",
         "Đ": "§", "ẩ": "È", "ặ": "Æ", "ắ": "¾", "ằ": "»", "ơ": "¬", "ô": "«", "Ô": "¤", "ỗ": "ç", "ỳ": "ú",
         "ý": "ý", "ẹ": "Ñ", "é": "Ð", "è": "Ì", "ù": "ï", "ũ": "ò", "õ": "â", "ĩ": "Ü", "ỹ": "ü", "ừ": "õ"}


def to_tcvn3(s):
    return "".join(TCVN3.get(ch, ch) for ch in s)


def write(name, header, rows, enc="utf-8", delim=","):
    path = os.path.join(OUT, name)
    lines = [delim.join(header)]
    for r in rows:
        cells = []
        for v in r:
            v = str(v)
            if delim == "," and ("," in v or " " in v):
                v = '"%s"' % v
            cells.append(v)
        lines.append(delim.join(cells))
    text = "\r\n".join(lines) + "\r\n"
    if enc == "tcvn3":
        data = to_tcvn3(text).encode("latin-1")
    elif enc == "utf-16":
        data = text.encode("utf-16")
    else:
        data = text.encode(enc)
    with open(path, "wb") as f:
        f.write(data)
    return path


def cust():
    return "550%010d" % random.randint(1, 2000)


def date():
    return "20%02d%02d%02d" % (random.randint(18, 24), random.randint(1, 12), random.randint(1, 28))


# ---- rt05 (21 cột) ----
rows = []
for i in range(40):
    c = random.choice(range(8))
    loai = random.choice(["The chap", "The chap", "Bao lanh"])
    val = random.randint(1, 50) * 10000000
    kinds = [0, 0, 0, 0]
    kinds[random.randint(0, 3)] = val
    pm = ["", "", ""]
    pm[random.randint(0, 2)] = "X"
    rows.append(["20241213", "5501", "550%010d" % (c + 1), NAMES[c], loai, "HD%05d" % i, date(), "VND", val, val]
                + kinds + pm + ["4211", "C%03d" % c, NAMES[c], random.choice(["600", "100", "200"])])
write("5501_rt05.csv", ["Ngay", "Ma CN", "Ma KH", "Ten KH", "Loai TC", "So HD", "Ngay HD", "Loai tien", "So tien",
                        "Quy doi", "BDS", "DS", "CTCG", "TS khac", "Co KN", "It KN", "Khong KN", "TK", "Ma chu",
                        "Ten chu", "Loai KH"], rows)

# ---- rt13 (21 cột), chia 2 file ----
def rt13_rows(n):
    out = []
    for i in range(n):
        ccy = random.choice(["VND", "VND", "USD"])
        nt = random.randint(1, 900) * (100000 if ccy == "VND" else 1) + (0.5 if ccy == "USD" else 0)
        vnd = nt if ccy == "VND" else round(nt * 25450)
        acct = random.choice(["4211", "4221", "4232"])
        c = random.randint(0, 7)
        out.append(["20241213", "5501", acct, "550%010d" % (c + 1), NAMES[c],
                    "1%012d" % random.randint(1, 10 ** 9), "SO%04d" % i, ccy, nt, vnd, "'12", date(), date(), "4.5",
                    date(), random.randint(0, 50000), random.randint(0, 200000), "DP", "1", "TR", "'4913%s" % acct[-2:]])
    return out


h13 = ["Ngay", "Ma CN", "TK", "Ma KH", "Ten KH", "So TK", "So so", "Loai tien", "Du NT", "Du VND", "Ky han",
       "Ngay gui", "Ngay DH", "LS", "Ngay TL", "Lai TT", "Lai PT", "DPTPCD", "ACCTSEQ", "TRCD", "TK lai DC"]
r13 = rt13_rows(120)
write("5501_rt13_1.csv", h13, r13[:70], enc="utf-16")
write("5501_rt13_2.csv", h13, r13[70:], enc="utf-16")
write("5501_rt14.csv", h13, rt13_rows(30), enc="tcvn3")

# ---- GL (glcb41, TAB) ----
glrows = []
sums = {}
for r in r13:
    sums[(r[2], r[7])] = sums.get((r[2], r[7]), 0) + r[8]
    sums[(r[2], "000")] = sums.get((r[2], "000"), 0) + r[9]
for (acct, ccy), v in sums.items():
    if acct == "4232" and ccy == "VND":
        v += 1000  # tạo chênh lệch
    glrows.append([ccy, acct] + ["0"] * 8 + ["0", str(round(v, 2))])
write("5501_glcb41.txt", ["CCY", "ACCT"] + ["C%d" % i for i in range(3, 13)], glrows, delim="\t")

# ---- rt15a (43 cột) ----
def rt15_rows(n, ncols=43):
    out = []
    for i in range(n):
        ccy = random.choice(["VND", "VND", "USD"])
        nt = random.randint(1, 500) * (1000000 if ccy == "VND" else 10)
        vnd = nt if ccy == "VND" else nt * 25450
        c = random.randint(0, 7)
        row = [""] * ncols
        row[1] = "5501"; row[2] = "PGD1"; row[3] = "'%d" % random.randint(1, 5)
        row[4] = random.choice(["2111", "2121", "2131"]); row[5] = "550%010d" % (c + 1)
        row[6] = random.choice(["Cá nhân", "Hộ gia đình", "Doanh nghiệp", "Cá nhân08"]); row[7] = NAMES[c]
        row[8] = "55%d%07d" % (random.randint(1, 3), i); row[12] = "LD%07d" % i; row[13] = ccy
        row[15] = nt; row[16] = vnd; row[17] = "'12"; row[18] = date(); row[19] = date(); row[20] = "9.5"
        row[23] = date(); row[24] = random.randint(0, 9000); row[25] = random.randint(0, 90000)
        row[33] = "A01"; row[34] = "Trồng trọt"; row[35] = "Nông nghiệp"; row[36] = "NV"; row[37] = "Vốn TW"
        row[41] = "394101"; row[42] = "394102"
        out.append(row)
    return out


h15 = ["C%d" % i for i in range(1, 44)]
write("5501_rt15a.csv", h15, rt15_rows(60))
write("5501_rt15b.csv", h15, rt15_rows(20))
write("5501_rt16.csv", h15, rt15_rows(15))

# ---- rt18 (19 cột) ----
rows = []
for i in range(25):
    ccy = random.choice(["VND", "USD"])
    nt = random.randint(1, 100) * (10000000 if ccy == "VND" else 100)
    c = random.randint(0, 7)
    rows.append(["20241213", "5501", "PGD", random.choice(["9211", "9212", "9231"]), ccy, "DN", "550%010d" % (c + 1),
                 NAMES[c], "BL%05d" % i, date(), date(), random.choice(["BLTT", "BLDT"]), nt,
                 nt if ccy == "VND" else nt * 25450, ccy, random.randint(0, 9) * 1000000, "HM%03d" % c,
                 random.randint(1, 9) * 100000000, "USER"])
write("5501_rt18.csv", ["C%d" % i for i in range(1, 20)], rows)

# ---- rt20a (TAB, 38 cột) ----
rows = []
for i in range(15):
    row = [""] * 38
    c = random.randint(0, 7)
    row[0] = "5501"; row[1] = "550%010d" % (c + 1); row[2] = NAMES[c]; row[3] = "KU%05d" % i; row[4] = "HD%05d" % i
    row[11] = random.randint(1, 90) * 1000000; row[12] = random.randint(0, 9) * 100000
    row[23] = random.choice([0, 1, 5]); row[24] = random.choice([0, 2]); row[37] = "12.5"
    rows.append(row)
write("5501_rt20a.csv", ["C%d" % i for i in range(1, 39)], rows, delim="\t")

# ---- rt22: 2 file (1) doanh thu chờ PB 488, TK 7xxx có thêm 1 cột ; (2) chi phí chờ PB 388 ----
def rt22_rows(n, revenue):
    out = []
    for i in range(n):
        c = random.randint(0, 7)
        ccy = random.choice(["VND", "USD"])
        row = ["5501", "x", "550%010d" % (c + 1), NAMES[c], ("7011" if revenue else "8011"), "y",
               ("4881" if revenue else "3881"), "z", "REF%04d" % i, ccy, "15/01/2024", "15/01/2025"]
        rest = ["31/12/2024", "8.0", random.randint(1, 90) * 1000000, random.randint(1, 50) * 100000, "25450",
                random.randint(1, 50) * 100000]
        if revenue:
            rest = ["extra"] + rest
        out.append(row + rest + [""] * 3)
    return out


write("5501_rt22_1.csv", ["C%d" % i for i in range(1, 22)], rt22_rows(8, True), enc="utf-16")
write("5501_rt22_2.csv", ["C%d" % i for i in range(1, 22)], rt22_rows(6, False), enc="utf-16")

# ---- rt24a (11 cột) ----
rows = []
for i in range(30):
    ccy = random.choice(["VND", "USD"])
    acct = random.choice(["3612", "3614", "4591", "4599"])
    nt = random.randint(1, 100) * (100000 if ccy == "VND" else 10)
    rows.append(["20241213", "5501", acct, "TK %s" % acct, "DT%09d" % random.randint(1, 50), random.choice(NAMES),
                 "Tạm ứng công tác", date(), ccy, nt, nt if ccy == "VND" else nt * 25450])
write("5501_rt24a.csv", ["C%d" % i for i in range(1, 12)], rows)

# ---- rtdp (45 cột, TCVN3) ----
rows = []
for i in range(20):
    row = ["0"] * 45
    c = random.randint(0, 7)
    ccy = random.choice(["VND", "USD"])
    row[0] = "5501"; row[2] = "%09d" % (c + 1); row[3] = NAMES[c]; row[5] = ccy
    row[6] = random.randint(1, 500) * 100000 + (0.25 if ccy == "USD" else 0); row[8] = "1%012d" % i
    row[11] = "Thủ Dầu Một, Bình Dương"; row[16] = random.choice(["0", "0", "5"]); row[18] = "SP%02d" % c
    row[21] = random.choice(["Cá nhân", "Hộ gia đình", "Doanh nghiệp", "Tổ chức"])
    rows.append(row)
write("5501_rtdp.csv", ["C%d" % i for i in range(1, 46)], rows, enc="tcvn3")

# ---- rttt (20 cột) ----
rows = []
for i in range(12):
    c = random.randint(0, 7)
    ccy = random.choice(["VND", "USD"])
    rows.append(["5501", "x", "550%010d" % (c + 1), NAMES[c], "y", "z", "3941%02d" % c, "a", "b", ccy, date(), date(),
                 date(), date(), "9.5", random.randint(1, 90) * 1000000, random.randint(1, 90) * 1000, "25450",
                 random.randint(1, 90) * 1000, random.randint(1, 90) * 1000])
write("5501_rttt.csv", ["C%d" % i for i in range(1, 21)], rows)

# ---- mẫu tuỳ chỉnh rt99 (cùng cấu trúc rt24a) ----
write("5501_rt99.csv", ["Ngày", "Mã CN", "Tài khoản", "Tên TK", "Mã ĐT", "Tên đối tượng", "Lý do", "Ngày PS",
                        "Loại tiền", "Số dư NT", "Số dư VND"], rows=[r for r in []] or [
    ["20241213", "5501", "3612", "TK 3612", "DT1", "Nguyễn Văn An", "Tạm ứng", "20240105", "VND", "1500000", "1500000"],
    ["20241213", "5501", "3612", "TK 3612", "DT2", "Trần Thị Bình", "Tạm ứng", "20240312", "VND", "2500000", "2500000"],
    ["20241213", "5501", "3614", "TK 3614", "DT3", "Lê Hoàng Cường", "Ứng trước", "20240520", "USD", "120.5", "3066725"],
])
print("Đã tạo dữ liệu mẫu tại", OUT)
