from __future__ import annotations

import html
import json
import re
import sys
import unicodedata
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_FILE = ROOT / "dashboard.html"
INDEX_FILE = ROOT / "index.html"
ERROR_FILE = ROOT / "گزارش خطاها.txt"

MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
]
PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def text_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def latin(value) -> str:
    return text_value(value).translate(PERSIAN_DIGITS)


def normalized(value: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text_value(value)).replace("ي", "ی").replace("ك", "ک")).lower()


def column_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    total = 0
    for char in letters.group(0) if letters else "A":
        total = total * 26 + ord(char) - 64
    return total


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values = []
    for item in root.findall("m:si", NS):
        values.append("".join(node.text or "" for node in item.findall(".//m:t", NS)))
    return values


def first_sheet(archive: zipfile.ZipFile) -> tuple[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = workbook.find("m:sheets/m:sheet", NS)
    if sheet is None:
        raise ValueError("فایل اکسل هیچ شیتی ندارد.")
    rel_id = sheet.attrib[DOC_REL]
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in relationships.findall("r:Relationship", REL_NS):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib.get("Target")
            break
    if not target:
        raise ValueError("مسیر شیت اول پیدا نشد.")
    target_path = PurePosixPath(target.lstrip("/"))
    if not str(target_path).startswith("xl/"):
        target_path = PurePosixPath("xl") / target_path
    return sheet.attrib.get("name", "Sheet1"), str(target_path)


def read_sheet(path: Path) -> tuple[str, dict[int, dict[int, object]]]:
    with zipfile.ZipFile(path) as archive:
        strings = shared_strings(archive)
        sheet_name, sheet_path = first_sheet(archive)
        root = ET.fromstring(archive.read(sheet_path))
        rows: dict[int, dict[int, object]] = defaultdict(dict)
        for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
            ref = cell.attrib.get("r", "A1")
            row_match = re.search(r"\d+", ref)
            row_num = int(row_match.group(0)) if row_match else 1
            col_num = column_number(ref)
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.findall(".//m:is/m:t", NS))
            else:
                raw = cell.findtext("m:v", default="", namespaces=NS)
                if cell_type == "s" and raw:
                    value = strings[int(raw)]
                elif cell_type in {"str", "b"}:
                    value = raw
                elif raw == "":
                    value = ""
                else:
                    try:
                        value = float(raw)
                    except ValueError:
                        value = raw
            rows[row_num][col_num] = value
    return sheet_name, rows


def find_label(rows: dict[int, dict[int, object]], labels: set[str]):
    targets = {normalized(label) for label in labels}
    for row_num in sorted(rows):
        if row_num > 10:
            break
        for col_num, value in rows[row_num].items():
            if normalized(text_value(value).replace(":", "")) in targets:
                return rows[row_num].get(col_num + 1, "")
    return ""


def header_columns(rows: dict[int, dict[int, object]]) -> tuple[int, dict[str, int]]:
    for row_num in sorted(rows):
        texts = {col: normalized(text_value(value)) for col, value in rows[row_num].items()}
        if not any("کدشاخص" in value for value in texts.values()):
            continue
        if not any("ارزیابی" in value for value in texts.values()):
            continue
        columns: dict[str, int] = {}
        for col, value in texts.items():
            if "کدشاخص" in value:
                columns["code"] = col
            elif "گروه" in value:
                columns["group"] = col
            elif value == "شاخص" or ("شاخص" in value and "کد" not in value and "گروه" not in value):
                columns["indicator"] = col
            elif "ارزیابی" in value:
                columns["score"] = col
            elif "وضعیت" in value or value == "بررسی":
                columns["status"] = col
            elif "توضیح" in value:
                columns["comment"] = col
            elif "زمان" in value or "تاریخ" in value:
                columns["date"] = col
            elif "فروشنده" in value or "ارزیاب" in value:
                columns["evaluator"] = col
            elif "سرپرست" in value:
                columns["supervisor"] = col
        required = {"code", "group", "indicator", "score"}
        if required.issubset(columns):
            return row_num, columns
    raise ValueError("سطر عنوان جدول شاخص‌ها پیدا نشد.")


def score_value(value):
    clean = latin(value).strip()
    if clean in {"0", "0.0"}:
        return 0
    if clean in {"1", "1.0"}:
        return 1
    return None


def derive_period(month_value, year_value) -> tuple[str, str]:
    month = text_value(month_value)
    year = latin(year_value)
    year_match = re.search(r"(?:13|14)\d{2}", year)
    year = year_match.group(0) if year_match else ""
    return month, year


def derive_audit_date(date_value) -> tuple[str, str]:
    """Return a stable period key and a readable date from one Excel date field."""
    if isinstance(date_value, (int, float)) and date_value > 0:
        audit_date = datetime(1899, 12, 30) + timedelta(days=float(date_value))
        value = audit_date.strftime("%Y-%m-%d")
        return value, value

    clean = latin(date_value).strip().replace(".", "/").replace("-", "/")
    match = re.search(r"((?:13|14|19|20)\d{2})/(\d{1,2})/(\d{1,2})", clean)
    if not match:
        return "", ""
    year, month, day = (int(part) for part in match.groups())
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return "", ""
    value = f"{year:04d}-{month:02d}-{day:02d}"
    return value, value.replace("-", "/")


def parse_audit(path: Path) -> dict:
    sheet_name, rows = read_sheet(path)
    header_row, columns = header_columns(rows)
    store = text_value(find_label(rows, {"نام فروشگاه", "فروشگاه"}))
    audit_date_raw = find_label(rows, {"تاریخ آدیت", "تاریخ"})
    month_raw = find_label(rows, {"ماه"})
    year_raw = find_label(rows, {"سال"})
    evaluator = text_value(find_label(rows, {"ارزیاب", "فروشنده"}))
    supervisor = text_value(find_label(rows, {"سرپرست"}))

    evaluations = []
    for row_num in range(header_row + 1, max(rows.keys(), default=header_row) + 1):
        row = rows.get(row_num, {})
        code_text = text_value(row.get(columns["code"], ""))
        indicator = text_value(row.get(columns["indicator"], ""))
        if code_text.startswith("جمع") or indicator.startswith("جمع"):
            break
        if not code_text or not indicator:
            continue
        score = score_value(row.get(columns["score"]))
        status = text_value(row.get(columns.get("status", -1), ""))
        comment = text_value(row.get(columns.get("comment", -1), ""))
        if not evaluator:
            evaluator = text_value(row.get(columns.get("evaluator", -1), ""))
        if not supervisor:
            supervisor = text_value(row.get(columns.get("supervisor", -1), ""))
        evaluations.append({
            "code": latin(code_text),
            "group": text_value(row.get(columns["group"], "")),
            "indicator": indicator,
            "score": score,
            "status": status,
            "comment": comment,
        })
        if len(evaluations) == 55:
            break

    period_key, audit_date = derive_audit_date(audit_date_raw)
    month, year = derive_period(month_raw, year_raw)
    if not period_key and month and year:
        month_index = MONTHS.index(month) + 1
        period_key = f"{year}-{month_index:02d}"
        audit_date = f"{month} {year}"
    if not store:
        raise ValueError("هدر «نام فروشگاه» داخل اکسل تکمیل نشده است.")
    if not period_key:
        raise ValueError("هدر «تاریخ آدیت» داخل اکسل تکمیل نشده یا فرمت تاریخ معتبر نیست.")
    if len(evaluations) != 55:
        raise ValueError(f"تعداد ردیف‌های شاخص {len(evaluations)} است؛ باید دقیقاً ۵۵ ردیف باشد.")
    incomplete = [index + 1 for index, item in enumerate(evaluations) if item["score"] not in {0, 1}]
    if incomplete:
        preview = "، ".join(str(number) for number in incomplete[:8])
        raise ValueError(f"امتیاز ردیف‌های {preview} کامل نیست؛ همه امتیازها باید صفر یا یک باشند.")

    groups: dict[str, dict] = {}
    issues = []
    for item in evaluations:
        group = groups.setdefault(item["code"], {
            "code": item["code"], "name": item["group"], "total": 0, "passed": 0,
        })
        group["total"] += 1
        group["passed"] += item["score"]
        if item["score"] == 0:
            issues.append({
                "code": item["code"], "group": item["group"], "indicator": item["indicator"],
                "status": item["status"], "comment": item["comment"] or "بدون توضیح",
            })

    passed = sum(item["score"] for item in evaluations)
    return {
        "id": f"{normalized(store)}:{period_key}",
        "storeName": store,
        "month": month,
        "year": year,
        "periodKey": period_key,
        "auditDate": audit_date,
        "evaluator": evaluator,
        "supervisor": supervisor,
        "total": 55,
        "passed": passed,
        "readiness": round(passed / 55 * 100, 2),
        "groups": list(groups.values()),
        "issues": issues,
        "evaluations": evaluations,
        "sourceFile": path.name,
        "sourceMtime": path.stat().st_mtime,
    }


def collect_data() -> tuple[list[dict], list[str]]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    warnings = []
    periods_by_key: dict[str, dict] = {}
    files = sorted(
        [path for path in DATA_DIR.iterdir() if path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")],
        key=lambda path: path.name.lower(),
    )
    for path in files:
        try:
            audit = parse_audit(path)
            previous = periods_by_key.get(audit["id"])
            if previous and previous["sourceMtime"] > audit["sourceMtime"]:
                continue
            if previous:
                warnings.append(f"{path.name}: اطلاعات فروشگاه و دوره تکراری بود و نسخه جدیدتر جایگزین شد.")
            periods_by_key[audit["id"]] = audit
        except Exception as exc:
            warnings.append(f"{path.name}: {exc}")

    stores_map: dict[str, dict] = {}
    for audit in periods_by_key.values():
        store_key = normalized(audit["storeName"])
        store = stores_map.setdefault(store_key, {
            "id": f"store-{len(stores_map) + 1}",
            "name": audit["storeName"],
            "periods": [],
        })
        audit.pop("sourceMtime", None)
        store["periods"].append(audit)
    stores = sorted(stores_map.values(), key=lambda item: normalized(item["name"]))
    for index, store in enumerate(stores, start=1):
        store["id"] = f"store-{index}"
        store["periods"].sort(key=lambda item: item["periodKey"], reverse=True)
    return stores, warnings


def build_html(stores: list[dict], warnings: list[str]) -> str:
    payload = json.dumps({
        "stores": stores,
        "warnings": warnings,
        "generatedAt": datetime.now().strftime("%Y/%m/%d - %H:%M"),
    }, ensure_ascii=False).replace("</", "<\\/")
    template = r'''<!doctype html>
<html lang="fa" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>داشبورد مدیریتی آدیت فروشگاه‌ها</title>
  <style>
    :root{--navy:#0b3558;--blue:#1c6e9e;--purple:#6d38a0;--green:#2f9e72;--red:#c95050;--ink:#183044;--muted:#6e8190;--line:#dce5eb;--panel:#fff;--bg:#f3f7f9}
    *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Tahoma,"Segoe UI",sans-serif}button,select{font:inherit}
    .shell{width:min(1440px,calc(100% - 32px));margin:auto;padding:22px 0 44px}.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:17px 20px;border-radius:18px;background:#fff;box-shadow:0 10px 32px rgba(11,53,88,.08)}
    .brand{display:flex;align-items:center;gap:12px}.mark{width:45px;height:45px;display:grid;place-items:center;border-radius:14px;background:linear-gradient(135deg,var(--navy),var(--purple));color:#fff;font-weight:900}.brand strong{display:block}.brand small,.muted{color:var(--muted)}
    .stamp{padding:9px 12px;border-radius:10px;background:#e9f5ef;color:#237657;font-size:12px}.hero{display:grid;grid-template-columns:1fr auto;gap:28px;align-items:center;padding:40px 8px 28px}.hero h1{margin:8px 0 10px;font-size:clamp(24px,3vw,42px);line-height:1.45}.hero p{margin:0;color:var(--muted);line-height:2}.eyebrow{color:var(--blue);font-size:12px;font-weight:800}
    .hero-score{width:150px;height:150px;border-radius:50%;display:grid;place-items:center;background:conic-gradient(var(--green) calc(var(--score)*1%),#dfe9ed 0);position:relative}.hero-score:after{content:"";position:absolute;inset:14px;border-radius:50%;background:var(--bg)}.hero-score div{z-index:1;text-align:center}.hero-score strong{display:block;font-size:28px;color:var(--navy)}.hero-score span{font-size:11px;color:var(--muted)}
    .metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}.metric,.panel{background:var(--panel);border:1px solid rgba(220,229,235,.9);border-radius:18px;box-shadow:0 10px 30px rgba(11,53,88,.05)}.metric{padding:18px}.metric span{font-size:12px;color:var(--muted)}.metric strong{display:block;margin:8px 0 4px;font-size:28px;color:var(--navy)}.metric small{color:var(--muted)}
    .grid{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(320px,.65fr);gap:16px}.panel{padding:18px}.title{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:14px}.title h2{margin:5px 0 0;font-size:18px}.count{padding:7px 10px;border-radius:9px;background:#eef4f7;color:var(--blue);font-size:11px}
    .store-list{display:grid;gap:8px}.store-row{width:100%;display:grid;grid-template-columns:40px minmax(120px,1fr) minmax(100px,.8fr) 70px 80px;align-items:center;gap:10px;padding:11px;border:1px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);cursor:pointer;text-align:right}.store-row:hover,.store-row.active{border-color:#8eb9d2;background:#f1f8fc}.rank{width:30px;height:30px;display:grid;place-items:center;border-radius:9px;background:#e8f1f6;color:var(--navy)}.store-name strong,.store-name small{display:block}.store-name small{margin-top:4px;color:var(--muted);font-size:10px}.track{height:8px;background:#e7eef2;border-radius:99px;overflow:hidden}.track i{display:block;height:100%;background:linear-gradient(90deg,var(--blue),var(--green))}.pill{padding:6px;border-radius:8px;background:#e9f5ef;color:#237657;font-size:10px;text-align:center}.pill.bad{background:#fff0f0;color:#a64343}
    .selectors{display:grid;grid-template-columns:1fr 1fr;gap:10px}.field{display:grid;gap:6px}.field span{font-size:11px;color:var(--muted)}select{min-height:43px;padding:0 10px;border:1px solid var(--line);border-radius:10px;background:#f7fafb;color:var(--navy);font-weight:700}.facts{display:grid;grid-template-columns:repeat(2,1fr);gap:9px;margin:14px 0}.fact{padding:12px;border-radius:11px;background:#f6fafc}.fact span,.fact strong{display:block}.fact span{font-size:10px;color:var(--muted)}.fact strong{margin-top:5px;font-size:12px}
    .details{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px}.group-list{display:grid;gap:12px}.group-row{display:grid;grid-template-columns:minmax(130px,1fr) minmax(100px,1.3fr) 48px;gap:10px;align-items:center}.group-row small{display:block;color:var(--muted);margin-top:3px}.issues{display:grid;gap:8px;max-height:430px;overflow:auto}.issue{padding:12px;border-right:4px solid var(--red);border-radius:10px;background:#fff7f7}.issue strong,.issue p,.issue small{display:block;margin:0}.issue p{margin:6px 0;color:#6d3a3a;font-size:12px;line-height:1.8}.issue small{color:var(--muted)}
    .table-panel{margin-top:16px;overflow:hidden}.table-wrap{overflow:auto;max-height:640px;border:1px solid var(--line);border-radius:12px}table{width:100%;border-collapse:collapse;min-width:920px;font-size:12px}th{position:sticky;top:0;z-index:1;background:var(--purple);color:#fff;padding:11px;text-align:right}td{padding:10px;border-bottom:1px solid var(--line);vertical-align:top}tr.fail{background:#fff5f5}.score{display:inline-grid;width:28px;height:28px;place-items:center;border-radius:8px;background:#dff2e9;color:#237657;font-weight:800}.fail .score{background:#f8dede;color:#a64343}.empty{padding:34px;text-align:center;color:var(--muted);line-height:2}.warning-box{margin:0 0 16px;padding:13px 16px;border-radius:12px;background:#fff8dc;color:#755e00;line-height:1.9;font-size:12px}footer{display:flex;justify-content:space-between;gap:16px;color:var(--muted);font-size:11px;padding:20px 4px 0}
    @media(max-width:900px){.hero{grid-template-columns:1fr}.hero-score{display:none}.metrics,.grid,.details{grid-template-columns:1fr}.store-row{grid-template-columns:35px 1fr 65px 72px}.store-row .track{display:none}.topbar,footer{align-items:flex-start;flex-direction:column}}
    @media print{body{background:#fff}.shell{width:100%;padding:0}.topbar,.panel,.metric{box-shadow:none}.store-list{display:none}.grid{grid-template-columns:1fr}.panel{break-inside:avoid}.table-wrap{max-height:none;overflow:visible}}
  </style>
</head>
<body>
<main class="shell">
  <header class="topbar"><div class="brand"><span class="mark">د</span><div><strong>دیدبان فروشگاه</strong><small>داشبورد آفلاین آدیت دوره‌ای</small></div></div><span class="stamp" id="generatedAt"></span></header>
  <section class="hero"><div><span class="eyebrow">نمای مدیریتی شبکه فروشگاه‌ها</span><h1>هر فروشگاه یک‌بار؛ همه دوره‌ها زیر همان فروشگاه.</h1><p>این فایل بدون سرور و اینترنت اجرا می‌شود و در هر بار ساخت، تمام فایل‌های اکسل پوشه data را پردازش می‌کند.</p></div><div class="hero-score" id="heroScore"><div><strong id="heroScoreValue">۰٪</strong><span>میانگین آمادگی</span></div></div></section>
  <div id="warnings"></div>
  <section class="metrics"><article class="metric"><span>فروشگاه یکتا</span><strong id="storeCount">۰</strong><small>بدون تکرار در دوره‌ها</small></article><article class="metric"><span>میانگین آمادگی</span><strong id="averageReadiness">۰٪</strong><small>آخرین دوره هر فروشگاه</small></article><article class="metric"><span>مغایرت باز</span><strong id="issueCount">۰</strong><small>آخرین دوره هر فروشگاه</small></article></section>
  <section class="grid"><article class="panel"><div class="title"><div><span class="eyebrow">انتخاب فروشگاه</span><h2>فروشگاه‌ها و آخرین نتیجه</h2></div><span class="count" id="listCount"></span></div><div class="store-list" id="storeList"></div></article><aside class="panel" id="overview"></aside></section>
  <section class="details" id="details"></section>
  <section class="panel table-panel" id="evaluationPanel"></section>
  <footer><span>دیدبان فروشگاه — نسخه آفلاین HTML</span><span id="sourceSummary"></span></footer>
</main>
<script>
const DATA=__DATA__;
const nf=new Intl.NumberFormat('fa-IR');
const pf=new Intl.NumberFormat('fa-IR',{maximumFractionDigits:1});
const esc=(v)=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let selectedStoreId=DATA.stores[0]?.id||'';
let selectedPeriodId=DATA.stores[0]?.periods[0]?.id||'';
const selectedStore=()=>DATA.stores.find(s=>s.id===selectedStoreId);
const selectedPeriod=()=>selectedStore()?.periods.find(p=>p.id===selectedPeriodId)||selectedStore()?.periods[0];
function metrics(){const latest=DATA.stores.map(s=>s.periods[0]).filter(Boolean);const avg=latest.length?latest.reduce((a,p)=>a+p.readiness,0)/latest.length:0;const issues=latest.reduce((a,p)=>a+p.issues.length,0);document.getElementById('storeCount').textContent=nf.format(DATA.stores.length);document.getElementById('averageReadiness').textContent=pf.format(avg)+'٪';document.getElementById('issueCount').textContent=nf.format(issues);document.getElementById('heroScoreValue').textContent=pf.format(avg)+'٪';document.getElementById('heroScore').style.setProperty('--score',avg);document.getElementById('listCount').textContent=nf.format(DATA.stores.length)+' فروشگاه';}
function renderStores(){const el=document.getElementById('storeList');if(!DATA.stores.length){el.innerHTML='<div class="empty">هنوز فایل اکسل کاملی داخل پوشه data قرار نگرفته است.</div>';return;}el.innerHTML=DATA.stores.map((s,i)=>{const p=s.periods[0];return `<button class="store-row ${s.id===selectedStoreId?'active':''}" onclick="chooseStore('${s.id}')"><span class="rank">${nf.format(i+1)}</span><span class="store-name"><strong>${esc(s.name)}</strong><small>${nf.format(s.periods.length)} دوره ثبت‌شده</small></span><span class="track"><i style="width:${p.readiness}%"></i></span><strong>${pf.format(p.readiness)}٪</strong><span class="pill ${p.issues.length?'bad':''}">${nf.format(p.issues.length)} مغایرت</span></button>`}).join('');}
function chooseStore(id){selectedStoreId=id;selectedPeriodId=selectedStore()?.periods[0]?.id||'';render();}
function choosePeriod(id){selectedPeriodId=id;renderOverview();renderDetails();renderTable();}
function renderOverview(){const s=selectedStore(),p=selectedPeriod(),el=document.getElementById('overview');if(!s||!p){el.innerHTML='<div class="empty">داده‌ای برای نمایش وجود ندارد.</div>';return;}el.innerHTML=`<div class="title"><div><span class="eyebrow">فروشگاه منتخب</span><h2>${esc(s.name)}</h2></div><span class="count">${pf.format(p.readiness)}٪ آمادگی</span></div><div class="selectors"><label class="field"><span>فروشگاه</span><select onchange="chooseStore(this.value)">${DATA.stores.map(x=>`<option value="${x.id}" ${x.id===s.id?'selected':''}>${esc(x.name)}</option>`).join('')}</select></label><label class="field"><span>تاریخ / دوره</span><select onchange="choosePeriod(this.value)">${s.periods.map(x=>`<option value="${x.id}" ${x.id===p.id?'selected':''}>${esc(x.auditDate)}</option>`).join('')}</select></label></div><div class="facts"><div class="fact"><span>ارزیاب</span><strong>${esc(p.evaluator||'ثبت نشده')}</strong></div><div class="fact"><span>سرپرست</span><strong>${esc(p.supervisor||'ثبت نشده')}</strong></div><div class="fact"><span>امتیاز</span><strong>${nf.format(p.passed)} از ${nf.format(p.total)}</strong></div><div class="fact"><span>فایل منبع</span><strong title="${esc(p.sourceFile)}">${esc(p.sourceFile)}</strong></div></div>`;}
function renderDetails(){const p=selectedPeriod(),el=document.getElementById('details');if(!p){el.innerHTML='';return;}const groups=p.groups.map(g=>{const ratio=g.total?g.passed/g.total*100:0;return `<div class="group-row"><div><strong>${esc(g.name)}</strong><small>${nf.format(g.passed)} از ${nf.format(g.total)}</small></div><div class="track"><i style="width:${ratio}%"></i></div><strong>${pf.format(ratio)}٪</strong></div>`}).join('');const issues=p.issues.length?p.issues.map((x,i)=>`<article class="issue"><strong>${nf.format(i+1)}. ${esc(x.indicator)}</strong><p>${esc(x.comment)}</p><small>${esc(x.group)}${x.status?' — '+esc(x.status):''}</small></article>`).join(''):'<div class="empty">در این دوره مغایرتی ثبت نشده است.</div>';el.innerHTML=`<article class="panel"><div class="title"><div><span class="eyebrow">نتیجه دوره</span><h2>امتیاز گروه‌های شاخص</h2></div></div><div class="group-list">${groups}</div></article><article class="panel"><div class="title"><div><span class="eyebrow">اقدامات اصلاحی</span><h2>مغایرت‌های این دوره</h2></div><span class="count">${nf.format(p.issues.length)} مورد</span></div><div class="issues">${issues}</div></article>`;}
function renderTable(){const p=selectedPeriod(),el=document.getElementById('evaluationPanel');if(!p){el.innerHTML='';return;}const rows=p.evaluations.map((x,i)=>`<tr class="${x.score===0?'fail':''}"><td>${nf.format(i+1)}</td><td>${esc(x.code)}</td><td>${esc(x.group)}</td><td>${esc(x.indicator)}</td><td><span class="score">${nf.format(x.score)}</span></td><td>${esc(x.status||'—')}</td><td>${esc(x.comment||'—')}</td></tr>`).join('');el.innerHTML=`<div class="title"><div><span class="eyebrow">جزئیات کامل</span><h2>تمام ۵۵ ردیف ارزیابی — ${esc(p.auditDate)}</h2></div><span class="count">${nf.format(p.passed)} امتیاز</span></div><div class="table-wrap"><table><thead><tr><th>ردیف</th><th>کد</th><th>گروه</th><th>شاخص</th><th>امتیاز</th><th>وضعیت</th><th>توضیح</th></tr></thead><tbody>${rows}</tbody></table></div>`;}
function renderWarnings(){const el=document.getElementById('warnings');el.innerHTML=DATA.warnings.length?`<div class="warning-box"><strong>توجه:</strong><br>${DATA.warnings.map(esc).join('<br>')}</div>`:'';}
function render(){metrics();renderWarnings();renderStores();renderOverview();renderDetails();renderTable();document.getElementById('generatedAt').textContent='ساخته‌شده در '+DATA.generatedAt;const files=DATA.stores.reduce((n,s)=>n+s.periods.length,0);document.getElementById('sourceSummary').textContent=nf.format(files)+' دوره از پوشه data';}
render();
</script>
</body>
</html>'''
    return template.replace("__DATA__", payload)


def main() -> int:
    stores, warnings = collect_data()
    html_output = build_html(stores, warnings)
    OUTPUT_FILE.write_text(html_output, encoding="utf-8")
    INDEX_FILE.write_text(html_output, encoding="utf-8")
    ERROR_FILE.write_text("\n".join(warnings) if warnings else "خطایی وجود ندارد.", encoding="utf-8")
    total_periods = sum(len(store["periods"]) for store in stores)
    print(f"داشبورد ساخته شد: {len(stores)} فروشگاه و {total_periods} دوره")
    if warnings:
        print(f"تعداد هشدارها: {len(warnings)} — فایل گزارش خطاها را بررسی کنید.")
    print(OUTPUT_FILE)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        ERROR_FILE.write_text(str(exc), encoding="utf-8")
        print(f"خطا: {exc}", file=sys.stderr)
        raise
