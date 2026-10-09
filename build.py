"""Fetch TWSE 證券商成交金額表 (monthly broker turnover) for a rolling 10 years
and build site/index.html. Only broker-level rows are kept: code ending '*'
(brokerage total) and 'T' (proprietary desk); branch rows are dropped."""
import io, json, os, zipfile, datetime, urllib.request
import xlrd

BASE = "https://www.twse.com.tw/staticFiles/inspection/inspection/03/003/{ym}_C03003.zip"
YEARS_BACK = 10

def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read()

def months():
    today = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=8)
    y, m = today.year, today.month
    out = []
    for _ in range(YEARS_BACK * 12 + 1):          # up to current month inclusive
        out.append(f"{y}{m:02d}")
        m -= 1
        if m == 0: y -= 1; m = 12
    return sorted(out)

def parse(xls_bytes, ym):
    book = xlrd.open_workbook(file_contents=xls_bytes)
    sh = book.sheet_by_index(0)
    rows = []
    for i in range(3, sh.nrows):
        code = str(sh.cell_value(i, 0)).strip()
        if not code or not (code.endswith("*") or code.endswith("T")):
            continue
        name = str(sh.cell_value(i, 1)).strip()
        amt, pct, yamt, ypct = (sh.cell_value(i, c) for c in (3, 4, 5, 6))
        try:
            amt, pct, yamt, ypct = int(float(amt)), float(pct), int(float(yamt)), float(ypct)
        except (TypeError, ValueError):
            continue
        kind = "B" if code.endswith("*") else "P"
        rows.append([code[:3], name, ym, kind, amt, round(pct, 4), yamt, round(ypct, 4)])
    return rows

all_rows, got = [], []
for ym in months():
    try:
        z = zipfile.ZipFile(io.BytesIO(fetch(BASE.format(ym=ym))))
        xls = z.read(z.namelist()[0])
    except Exception as e:                        # month not published yet (or very old gap)
        print(f"skip {ym}: {str(e)[:60]}")
        continue
    r = parse(xls, ym)
    if r:
        all_rows += r; got.append(ym)

if len(got) < 12:
    raise SystemExit("Too little data fetched - aborting so the old site stays up")

tw_today = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=8)).date().isoformat()
data = json.dumps({"rows": all_rows, "fetched": tw_today}, ensure_ascii=False, separators=(",", ":"))
template = open("template.html", encoding="utf-8").read()
os.makedirs("site", exist_ok=True)
open("site/index.html", "w", encoding="utf-8").write(template.replace("__DATA__", data))
print(f"Built site/index.html: {len(all_rows)} rows, {got[0]} -> {got[-1]} ({len(got)} months)")
