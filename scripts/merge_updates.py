#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automated data-update merger for فيوتشر كوزماتكس (Future Cosmetics) PWA.

Reads whichever of the 4 standard update files are present in updates/,
merges them into index.html (same rules Claude has been applying manually:
stock refresh, customer-balance refresh-and-zero-absent, collections/sales
dedupe-by-key), and writes a plain-text review report for anything that
needs a human decision (new products with no known price, unexpected file
structure such as multiple warehouses).

This script deliberately does NOT guess prices for brand-new product codes
or silently handle structural surprises -- those always go to the review
report so Waleed or Claude can look at them.
"""
import json
import re
import sys
import datetime
from pathlib import Path
from collections import defaultdict

try:
    import xlrd
except ImportError:
    print("xlrd is required: pip install xlrd", file=sys.stderr)
    sys.exit(1)

REPO_ROOT = Path(__file__).resolve().parent.parent
INDEX_PATH = REPO_ROOT / "index.html"
UPDATES_DIR = REPO_ROOT / "updates"
REPORT_PATH = REPO_ROOT / "تقرير_المراجعة.txt"

STOCK_FILE = "تحديث_أرصدة_المخازن.xls"
BALANCE_FILE = "تحديث_كشف_حساب_العملاء.xls"
COLLECTIONS_FILE = "تحديث_المقبوضات.xls"
SALES_FILE = "تحديث__تفصيلي_فواتير_المبيعات.xls"
CUSTOMERS_FILE = "تحديث_بـيـانات_العملاء.xls"

MAIN_WAREHOUSE_NAME = "مخزن القاهره"


# ---------------------------------------------------------------- helpers --

def extract_span(content, varname, is_obj=False):
    """Locate the `const NAME = ...;` or `let NAME = ...;` JSON literal in
    index.html so it can be sliced out and replaced."""
    marker_const = f"const {varname} = "
    marker_let = f"let {varname} = "
    if marker_const in content:
        idx = content.index(marker_const)
        start = idx + len(marker_const)
    elif marker_let in content:
        idx = content.index(marker_let)
        start = idx + len(marker_let)
    else:
        raise ValueError(f"Could not find variable {varname} in index.html")
    open_ch = '{' if (is_obj or content[start] == '{') else '['
    close_ch = '}' if open_ch == '{' else ']'
    depth = 0
    in_str = False
    esc = False
    i = start
    while i < len(content):
        ch = content[i]
        if in_str:
            if esc:
                esc = False
            elif ch == '\\':
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == open_ch:
                depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    return start, i + 1
        i += 1
    raise ValueError(f"Could not find matching close for {varname}")


def load_var(content, varname, is_obj=False):
    s, e = extract_span(content, varname, is_obj=is_obj)
    return json.loads(content[s:e]), s, e


def save_var(content, varname, obj, is_obj=False):
    s, e = extract_span(content, varname, is_obj=is_obj)
    new_json = json.dumps(obj, ensure_ascii=False, separators=(',', ':'))
    return content[:s] + new_json + content[e:]


def excel_to_yyyymmdd(serial):
    d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(serial))
    return int(d.strftime('%Y%m%d'))


class Report:
    """Collects everything a human needs to look at."""
    def __init__(self):
        self.sections = []

    def add(self, title, lines):
        if lines:
            self.sections.append((title, lines))

    def is_empty(self):
        return not self.sections

    def render(self, run_time):
        out = [f"تقرير مراجعة التحديث التلقائي — {run_time}", "=" * 50, ""]
        if self.is_empty():
            out.append("لا يوجد أي شيء يحتاج مراجعة — كل التحديثات طُبّقت تلقائيًا بنجاح.")
        else:
            for title, lines in self.sections:
                out.append(f"### {title}")
                out.extend(f"- {line}" for line in lines)
                out.append("")
        return "\n".join(out)


# ------------------------------------------------------------ stock merge --

def merge_stock(content, report, summary):
    path = UPDATES_DIR / STOCK_FILE
    if not path.exists():
        return content

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)

    # Detect warehouse-header rows (only col0 populated, rest blank).
    header_rows = []
    for r in range(sh.nrows):
        vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if vals[0] and not any(vals[1:]):
            header_rows.append((r, str(vals[0]).strip()))

    if len(header_rows) > 1:
        report.add(
            "ملف أرصدة المخازن — أكتر من مخزن في نفس الملف",
            [f"لقيت {len(header_rows)} أقسام مخازن: " +
             ", ".join(name for _, name in header_rows) +
             " — اتعامل بس مع '" + MAIN_WAREHOUSE_NAME + "' وسبت الباقي، تأكد إن ده الصح."]
        )

    # Row range for the main warehouse section only.
    start_row = None
    end_row = sh.nrows
    for i, (r, name) in enumerate(header_rows):
        if MAIN_WAREHOUSE_NAME in name:
            start_row = r + 1
            if i + 1 < len(header_rows):
                end_row = header_rows[i + 1][0]
            break
    if start_row is None:
        report.add(
            "ملف أرصدة المخازن — تعذر إيجاد المخزن الرئيسي",
            [f"مفيش قسم اسمه '{MAIN_WAREHOUSE_NAME}' في الملف — الملف اتجوهش خالص."]
        )
        return content

    PRODUCTS, sP, eP = load_var(content, 'PRODUCTS')
    by_code = {}
    for p in PRODUCTS:
        by_code[str(p['code'])] = p
        by_code[str(p['code']).zfill(6)] = p

    updated = 0
    unknown_products = []
    for r in range(start_row + 1, end_row):  # +1 skips the column-label row
        code_cell = sh.cell_value(r, 4)
        if not code_cell:
            continue
        try:
            code = int(float(code_cell))
        except (TypeError, ValueError):
            continue
        stock = sh.cell_value(r, 1)
        key = str(code)
        padded = key.zfill(6)
        prod = by_code.get(key) or by_code.get(padded)
        if prod:
            prod['stock'] = stock
            updated += 1
        elif stock and float(stock) != 0:
            name = sh.cell_value(r, 3)
            unknown_products.append((code, name, stock))

    if unknown_products:
        report.add(
            "أصناف جديدة (غير موجودة في الكتالوج) — محتاجة سعر قبل الإضافة",
            [f"كود {code} — {name} — رصيد {stock}" for code, name, stock in unknown_products]
        )

    summary.append(f"📦 أرصدة المخازن: تحديث رصيد {updated} صنف")
    if unknown_products:
        summary.append(f"   ⚠️ {len(unknown_products)} صنف جديد محتاج مراجعة (شوف تقرير المراجعة)")

    new_json = json.dumps(PRODUCTS, ensure_ascii=False, separators=(',', ':'))
    return content[:sP] + new_json + content[eP:]


# --------------------------------------------------------- balances merge --

def merge_balances(content, report, summary):
    path = UPDATES_DIR / BALANCE_FILE
    if not path.exists():
        return content

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)

    CUST_BALANCES, s, e = load_var(content, 'CUST_BALANCES', is_obj=True)

    updated = 0
    added = 0
    seen_codes = set()
    for r in range(1, sh.nrows):
        name_idx = None
        for c in range(sh.ncols):
            if sh.cell_type(r, c) == 1:  # text cell -> customer name column
                name_idx = c
                break
        if name_idx is None or name_idx == 0:
            continue
        balance = sh.cell_value(r, 0)
        prev = sh.cell_value(r, name_idx - 1)
        code_cell = sh.cell_value(r, name_idx + 1) if name_idx + 1 < sh.ncols else None
        if code_cell in (None, ''):
            continue
        try:
            code = str(int(float(code_cell)))
        except (TypeError, ValueError):
            continue
        seen_codes.add(code)
        if code in CUST_BALANCES:
            updated += 1
        else:
            added += 1
        CUST_BALANCES[code] = [prev, balance]

    zeroed = []
    for code, val in CUST_BALANCES.items():
        if code not in seen_codes and val[1] != 0:
            zeroed.append(code)
            CUST_BALANCES[code] = [0, 0]

    summary.append(
        f"📋 كشف حساب العملاء: تحديث رصيد {updated} عميل، إضافة {added} عميل جديد، "
        f"تصفير رصيد {len(zeroed)} عميل سددوا بالكامل"
    )

    return save_var(content, 'CUST_BALANCES', CUST_BALANCES, is_obj=True)


# ------------------------------------------------------- collections merge --

def merge_collections(content, report, summary):
    path = UPDATES_DIR / COLLECTIONS_FILE
    if not path.exists():
        return content

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)

    COLL_TX, s, e = load_var(content, 'COLL_TX')
    existing_receipts = set(row[1] for row in COLL_TX)

    added = 0
    unassigned = []
    for r in range(1, sh.nrows):
        repcode_cell = sh.cell_value(r, 2)
        debtor = sh.cell_value(r, 3)
        amount = sh.cell_value(r, 4)
        date_serial = sh.cell_value(r, 5)
        receipt_cell = sh.cell_value(r, 7)
        if not receipt_cell or not debtor:
            continue
        try:
            repcode = int(float(repcode_cell)) if repcode_cell != '' else 0
            receipt_num = int(float(receipt_cell))
            date_int = excel_to_yyyymmdd(date_serial)
        except (TypeError, ValueError):
            continue
        if receipt_num in existing_receipts:
            continue
        COLL_TX.append([date_int, receipt_num, repcode, str(debtor), amount])
        existing_receipts.add(receipt_num)
        added += 1
        if repcode == 0:
            unassigned.append(f"إيصال {receipt_num} — {debtor} — {amount} ج.م (بدون مندوب)")

    if unassigned:
        report.add("مقبوضات بدون كود مندوب", unassigned)

    summary.append(f"💰 المقبوضات: إضافة {added} إيصال تحصيل جديد")
    if unassigned:
        summary.append(f"   ⚠️ {len(unassigned)} إيصال بدون مندوب محتاج مراجعة")

    return save_var(content, 'COLL_TX', COLL_TX)


# -------------------------------------------------------------- sales merge --

def merge_sales(content, report, summary):
    path = UPDATES_DIR / SALES_FILE
    if not path.exists():
        return content

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)

    INVOICE_ITEMS, sI, eI = load_var(content, 'INVOICE_ITEMS', is_obj=True)
    SALES_TX, sS, eS = load_var(content, 'SALES_TX')
    ITEM_SALES, sIS, eIS = load_var(content, 'ITEM_SALES')
    SALES_COMPANY_TX, sSC, eSC = load_var(content, 'SALES_COMPANY_TX')
    PRODUCTS, _, _ = load_var(content, 'PRODUCTS')

    code_to_company = {}
    for p in PRODUCTS:
        comp = p.get('company', 'متنوع')
        code_to_company[str(p['code'])] = comp
        code_to_company[str(p['code']).zfill(6)] = comp

    def lookup_company(item_code):
        return (code_to_company.get(str(item_code))
                or code_to_company.get(str(item_code).zfill(6))
                or 'متنوع')

    existing_keys = set(INVOICE_ITEMS.keys())
    groups = defaultdict(list)
    for r in range(2, sh.nrows):
        try:
            invoice_num = int(float(sh.cell_value(r, 13)))
        except (TypeError, ValueError):
            continue
        tx_type_str = sh.cell_value(r, 12)
        if not tx_type_str:
            continue
        groups[(invoice_num, tx_type_str)].append(r)

    new_invoices = 0
    unassigned = []
    item_sales_delta = defaultdict(float)
    company_tx_delta = defaultdict(float)

    for (invoice_num, tx_type_str), rows in groups.items():
        typ = 1 if tx_type_str == 'مبيعات' else (-1 if tx_type_str == 'مرتد بيع' else None)
        if typ is None:
            continue
        key = f"{invoice_num}_{typ}"
        if key in existing_keys:
            continue

        first = rows[0]
        date_int = excel_to_yyyymmdd(sh.cell_value(first, 15))
        rep_code = int(float(sh.cell_value(first, 11)))
        cust_code = int(float(sh.cell_value(first, 18)))

        items = []
        invoice_total = 0
        for r in rows:
            item_code_raw = sh.cell_value(r, 9)
            try:
                item_code = int(item_code_raw)
            except (TypeError, ValueError):
                try:
                    item_code = int(float(item_code_raw))
                except (TypeError, ValueError):
                    continue
            unit = sh.cell_value(r, 6)
            qty = sh.cell_value(r, 7)
            sales_value = sh.cell_value(r, 3)
            price = round(sales_value / qty, 4) if qty else sh.cell_value(r, 5)
            items.append([item_code, unit, qty, price, sales_value])
            invoice_total += sales_value

            item_sales_delta[(cust_code, item_code)] += qty * typ
            company = lookup_company(item_code)
            company_tx_delta[(date_int, rep_code, company, typ)] += sales_value

        INVOICE_ITEMS[key] = [date_int, cust_code, rep_code, invoice_num, typ, items]
        SALES_TX.append([date_int, invoice_num, cust_code, rep_code, typ, round(invoice_total, 4)])
        existing_keys.add(key)
        new_invoices += 1
        if rep_code == 0:
            unassigned.append(f"فاتورة {invoice_num} — عميل {cust_code} — {invoice_total} ج.م (بدون مندوب)")

    if unassigned:
        report.add("فواتير مبيعات بدون كود مندوب", unassigned)

    item_sales_idx = {(row[0], row[1]): i for i, row in enumerate(ITEM_SALES)}
    for (cust_code, item_code), delta in item_sales_delta.items():
        idx_key = (cust_code, item_code)
        if idx_key in item_sales_idx:
            i = item_sales_idx[idx_key]
            ITEM_SALES[i][2] = round(ITEM_SALES[i][2] + delta, 4)
        else:
            ITEM_SALES.append([cust_code, item_code, round(delta, 4)])

    company_tx_idx = {(row[0], row[1], row[2], row[3]): i for i, row in enumerate(SALES_COMPANY_TX)}
    for (date_int, rep_code, company, typ), delta in company_tx_delta.items():
        idx_key = (date_int, rep_code, company, typ)
        if idx_key in company_tx_idx:
            i = company_tx_idx[idx_key]
            SALES_COMPANY_TX[i][4] = round(SALES_COMPANY_TX[i][4] + delta, 4)
        else:
            SALES_COMPANY_TX.append([date_int, rep_code, company, typ, round(delta, 4)])

    summary.append(f"🧾 فواتير المبيعات: إضافة {new_invoices} فاتورة مبيعات جديدة")
    if unassigned:
        summary.append(f"   ⚠️ {len(unassigned)} فاتورة بدون مندوب محتاجة مراجعة")

    content = save_var(content, 'INVOICE_ITEMS', INVOICE_ITEMS, is_obj=True)
    content = save_var(content, 'SALES_TX', SALES_TX)
    content = save_var(content, 'ITEM_SALES', ITEM_SALES)
    content = save_var(content, 'SALES_COMPANY_TX', SALES_COMPANY_TX)
    return content


# --------------------------------------------------------- customers merge --

def merge_customers(content, report, summary):
    path = UPDATES_DIR / CUSTOMERS_FILE
    if not path.exists():
        return content

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)

    CUSTOMERS, s, e = load_var(content, 'CUSTOMERS')
    by_code = {c['code']: c for c in CUSTOMERS}

    updated = 0
    changed = 0
    added = 0
    for r in range(1, sh.nrows):
        code_cell = sh.cell_value(r, 7)
        if code_cell == '':
            continue
        try:
            code = int(float(code_cell))
        except (TypeError, ValueError):
            continue
        tel1 = str(sh.cell_value(r, 0)).strip()
        tel2 = str(sh.cell_value(r, 1)).strip()
        phone = tel1 if tel1 else tel2
        try:
            rep_code = int(float(sh.cell_value(r, 2))) if sh.cell_value(r, 2) != '' else 0
        except (TypeError, ValueError):
            rep_code = 0
        contact = str(sh.cell_value(r, 3)).strip()
        region = str(sh.cell_value(r, 4)).strip()
        addr = str(sh.cell_value(r, 5)).strip()
        name = str(sh.cell_value(r, 6)).strip()

        new_vals = {'name': name, 'phone': phone, 'addr': addr, 'repCode': rep_code,
                    'region': region, 'contact': contact}
        if code in by_code:
            c = by_code[code]
            if any(c.get(k) != v for k, v in new_vals.items()):
                changed += 1
            c.update(new_vals)
            updated += 1
        else:
            newc = {'code': code, **new_vals}
            CUSTOMERS.append(newc)
            by_code[code] = newc
            added += 1

    summary.append(
        f"👤 بيانات العملاء: تحديث بيانات {changed} عميل (من أصل {updated} اتفحص)، "
        f"إضافة {added} عميل جديد"
    )

    return save_var(content, 'CUSTOMERS', CUSTOMERS)


# ------------------------------------------------------------------- main --

def main():
    if not INDEX_PATH.exists():
        print("index.html not found at repo root", file=sys.stderr)
        sys.exit(1)

    with open(INDEX_PATH, encoding='utf-8') as f:
        content = f.read()

    report = Report()
    summary = []

    content = merge_stock(content, report, summary)
    content = merge_balances(content, report, summary)
    content = merge_collections(content, report, summary)
    content = merge_sales(content, report, summary)
    content = merge_customers(content, report, summary)

    with open(INDEX_PATH, 'w', encoding='utf-8') as f:
        f.write(content)

    run_time = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=2)  # Cairo (UTC+2, no DST)
    run_time_str = run_time.strftime('%Y-%m-%d %I:%M %p') + " (توقيت القاهرة)"

    with open(REPORT_PATH, 'w', encoding='utf-8') as f:
        f.write(report.render(run_time_str))

    # Clear out the processed update files so the next push only contains
    # genuinely new data (and doesn't reprocess old files).
    for name in (STOCK_FILE, BALANCE_FILE, COLLECTIONS_FILE, SALES_FILE, CUSTOMERS_FILE):
        p = UPDATES_DIR / name
        if p.exists():
            p.unlink()

    # Print the summary for the GitHub Actions log too.
    print("\n".join(summary) if summary else "لا توجد ملفات تحديث في مجلد updates/ هذه المرة.")
    if not report.is_empty():
        print("\n⚠️ في حاجات محتاجة مراجعة — شوف تقرير_المراجعة.txt")


if __name__ == "__main__":
    main()
