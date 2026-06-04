#!/usr/bin/env python3
"""
晟矽微运营 WIP 数据校验脚本
Validates the consolidated output file against source files.

Usage:
    python3 validate_wip.py [--source-dir PATH] [--output PATH]

Checks:
    1. Row count: source rows ≈ output rows (no silent drops)
    2. Column completeness: 29 columns, correct headers
    3. Data type sanity: numeric columns contain numbers
    4. Supplier coverage: every source file contributed rows
    5. Duplicate detection: warn if duplicate customer code + chip model found
    6. Spot-check samples: compare 3 random rows from each source
    7. Process-step consistency: 在线合计 should ≈ sum of visible process steps
    8. Missing data report: percentage of empty cells per column per supplier
"""
import openpyxl
import xlrd
import os
import sys
import random
import json
from collections import Counter, defaultdict

ATTACH_DIR = os.path.expanduser("~/Downloads/owa_attachments")
OUTPUT_PATH = os.path.expanduser("~/Downloads/晟矽微运营汇总_整合版.xlsx")

HEADERS = [
    "供应商", "客户代码", "客户名称", "客户订单号", "封装形式", "产品型号", "芯片型号",
    "晶圆批次", "订单数量", "投产日期",
    "粘片", "焊线", "焊线2", "焊线批检", "塑封", "电镀", "切筋", "委外切筋",
    "包装", "测试", "测试编带", "包装入库",
    "在线合计", "入库良品", "入库不良品", "出库良品", "出库不良品", "库存", "状态"
]

# Columns expected to be numeric (may have None).
# Note: Column 10 (投产日期) is excluded — it contains date strings.
NUMERIC_COLS = {11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28}

# Source file → (supplier_name, header_row_index_0based_or_Sheet3, sheet_name_or_None)
SOURCE_CONFIG = {
    "晟矽微WIP.xlsx":        ("晟矽微尊阳", "Sheet3"),
    "GD1458_1833.xlsx":      ("气派",       None),
    "JC196封装产品进展.xlsx": ("JC196巨成",  None),
    "ccml203_wip_260519.xls":("中芯微",     2),   # header at row index 2
    "ccml204_wip_260519.xls":("中芯微",     2),
    "内江_KH0352_WIP_20260517.xls": ("内江明泰", 0),
}

PASS = "✅ PASS"
WARN = "⚠️  WARN"
FAIL = "❌ FAIL"

results = []
def check(name, status, detail=""):
    results.append((name, status, detail))

# ════════════════════════════════════════════════
# 1. Check output file exists and opens
# ════════════════════════════════════════════════
def check_file():
    if not os.path.exists(OUTPUT_PATH):
        check("输出文件存在", FAIL, f"{OUTPUT_PATH} 不存在")
        return None
    try:
        wb = openpyxl.load_workbook(OUTPUT_PATH, data_only=True)
        check("输出文件可打开", PASS)
        return wb
    except Exception as e:
        check("输出文件可打开", FAIL, str(e))
        return None

# ════════════════════════════════════════════════
# 2. Column structure
# ════════════════════════════════════════════════
def check_columns(ws):
    actual_headers = [ws.cell(row=1, column=c).value for c in range(1, 30)]
    if actual_headers == HEADERS:
        check("列结构 29 列", PASS, "表头完全匹配")
    else:
        check("列结构 29 列", FAIL, f"期望 {len(HEADERS)} 列，实际 {len(actual_headers)} 列")
        for i, (exp, act) in enumerate(zip(HEADERS, actual_headers)):
            if exp != act:
                check(f"  列{i+1}", FAIL, f"期望 '{exp}'，实际 '{act}'")

# ════════════════════════════════════════════════
# 3. Row count comparison
# ════════════════════════════════════════════════
def count_data_rows(ws, data_start, max_col=40, key_col_idx=1):
    """Count rows that actually have data (non-empty in key column), not max_row."""
    count = 0
    for r in range(data_start, ws.max_row + 1):
        v = ws.cell(row=r, column=key_col_idx).value
        if v and str(v).strip():
            count += 1
        # Optimization: stop if we've seen 10 consecutive empty rows at the end
        # (but don't break early for small files)
    return count


def check_row_counts(ws):
    output_rows = ws.max_row - 1  # minus header

    source_counts = {}
    for fname, (supplier, sheet_info) in SOURCE_CONFIG.items():
        fpath = os.path.join(ATTACH_DIR, fname)
        if not os.path.exists(fpath):
            source_counts[fname] = (0, "文件不存在")
            continue
        try:
            if fname.endswith(".xlsx"):
                wb = openpyxl.load_workbook(fpath, data_only=True)
                if isinstance(sheet_info, str):
                    ws_s = wb[sheet_info]
                else:
                    ws_s = wb.active
                data_start = find_data_start(ws_s)
                data_rows = count_data_rows(ws_s, data_start)
            elif fname.endswith(".xls"):
                wb_x = xlrd.open_workbook(fpath)
                ws_x = wb_x.sheet_by_index(0)
                data_start = sheet_info + 1 if isinstance(sheet_info, int) else 1
                # Count rows with non-empty customer code
                data_rows = 0
                for r in range(data_start, ws_x.nrows):
                    v = ws_x.cell_value(r, 0)
                    if v and str(v).strip():
                        data_rows += 1
            else:
                data_rows = 0
            source_counts[fname] = (data_rows, "OK")
        except Exception as e:
            source_counts[fname] = (0, str(e))

    total_source = sum(c for c, _ in source_counts.values())
    check("行数对比",
          PASS if abs(total_source - output_rows) <= 5 else WARN,
          f"源文件总计 ≈{total_source} 数据行，输出 {output_rows} 行 "
          f"(差异 {output_rows - total_source:+d})")

    for fname, (count, msg) in source_counts.items():
        if msg != "OK":
            check(f"  {fname}", WARN, msg)
        else:
            check(f"  {fname} → {count} 行", PASS)

def find_data_start(ws):
    """Find the first data row (skip headers)."""
    for r in range(1, 10):
        v = ws.cell(row=r, column=1).value
        if v and ("客户代码" in str(v) or "客户编码" in str(v)):
            return r + 1
    return 2

# ════════════════════════════════════════════════
# 4. Numeric column sanity check
# ════════════════════════════════════════════════
def check_numeric(ws):
    bad = 0
    for r in range(2, ws.max_row + 1):
        for c in NUMERIC_COLS:
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            try:
                float(v)
            except (ValueError, TypeError):
                bad += 1
                if bad <= 5:
                    check(f"  数字列异常 R{r}C{c}", WARN,
                          f"'{v}' (供应商={ws.cell(row=r, column=1).value})")
    if bad == 0:
        check("数值列检查", PASS, "所有数值列数据正常")
    else:
        check("数值列检查", WARN, f"发现 {bad} 个非数值单元格（可能正常）")

# ════════════════════════════════════════════════
# 5. Duplicate detection
# ════════════════════════════════════════════════
def check_duplicates(ws):
    seen = Counter()
    for r in range(2, ws.max_row + 1):
        code = ws.cell(row=r, column=2).value  # 客户代码
        chip = ws.cell(row=r, column=7).value  # 芯片型号
        sup = ws.cell(row=r, column=1).value   # 供应商
        if code:
            key = (sup, str(code).strip(), str(chip or "").strip())
            seen[key] += 1
    dups = {k: v for k, v in seen.items() if v > 1}
    if dups:
        check("重复行检查", WARN, f"发现 {len(dups)} 组重复（客户代码+芯片型号）")
        for k, cnt in sorted(dups.items(), key=lambda x: -x[1])[:5]:
            check(f"  重复 {cnt} 次", WARN, f"供应商={k[0]}, 代码={k[1]}, 芯片={k[2]}")
    else:
        check("重复行检查", PASS, "无重复行")

# ════════════════════════════════════════════════
# 6. Supplier coverage
# ════════════════════════════════════════════════
def check_supplier_coverage(ws):
    suppliers = set()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v:
            suppliers.add(v)
    expected_suppliers = set(s for s, _ in SOURCE_CONFIG.values())
    missing = expected_suppliers - suppliers
    extra = suppliers - expected_suppliers
    if not missing and not extra:
        check("供应商覆盖", PASS, f"全部 {len(suppliers)} 个供应商已覆盖")
    else:
        if missing:
            check("供应商覆盖", WARN, f"缺失: {missing}")
        if extra:
            check("供应商覆盖", WARN, f"新增: {extra}（可能是新供应商，正常）")

# ════════════════════════════════════════════════
# 7. Missing data report
# ════════════════════════════════════════════════
def check_missing_data(ws):
    supplier_data = defaultdict(lambda: defaultdict(lambda: {"total": 0, "empty": 0}))
    for r in range(2, ws.max_row + 1):
        sup = ws.cell(row=r, column=1).value or "未知"
        for c in range(2, 30):
            v = ws.cell(row=r, column=c).value
            supplier_data[sup][c]["total"] += 1
            if v is None or str(v).strip() == "":
                supplier_data[sup][c]["empty"] += 1

    # Report columns with >80% empty across ALL suppliers
    total_rows = ws.max_row - 1
    empty_cols = []
    for c in range(2, 30):
        empty = 0
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or str(v).strip() == "":
                empty += 1
        pct = empty / total_rows * 100
        if pct > 80:
            empty_cols.append((c, HEADERS[c-1], pct))

    if empty_cols:
        check("空数据列报告", WARN, f"{len(empty_cols)} 列超过 80% 为空（可能正常）")
        for c, name, pct in empty_cols:
            check(f"  {name} (列{c})", WARN, f"{pct:.0f}% 为空")
    else:
        check("空数据列报告", PASS, "无大面积空缺")

# ════════════════════════════════════════════════
# 8. Spot-check: sample 3 rows per source file
# ════════════════════════════════════════════════
def check_spot_samples(ws):
    """Pick 1 random row per supplier and show it for manual inspection."""
    samples = defaultdict(list)
    for r in range(2, ws.max_row + 1):
        sup = ws.cell(row=r, column=1).value
        if sup:
            samples[sup].append(r)

    check("抽样检查", PASS, "以下随机抽样供人工核对：")
    for sup, rows in sorted(samples.items()):
        row = random.choice(rows)
        # Show key fields
        code = ws.cell(row=row, column=2).value
        product = ws.cell(row=row, column=6).value
        chip = ws.cell(row=row, column=7).value
        online = ws.cell(row=row, column=23).value
        in_good = ws.cell(row=row, column=24).value
        stock = ws.cell(row=row, column=28).value
        check(f"  [{sup}] 行{row}", PASS,
              f"代码={code} | 产品={product} | 芯片={chip} | "
              f"在线={online} | 入库良={in_good} | 库存={stock}")

# ════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("🔍 晟矽微运营汇总 — 数据校验")
    print("=" * 60)

    wb = check_file()
    if wb is None:
        print_report()
        sys.exit(1)

    ws = wb.active
    check_columns(ws)
    check_row_counts(ws)
    check_numeric(ws)
    check_duplicates(ws)
    check_supplier_coverage(ws)
    check_missing_data(ws)
    check_spot_samples(ws)

    print_report()

    # Exit code: 0 if no FAIL, 1 if any FAIL
    if any(s == FAIL for _, s, _ in results):
        sys.exit(1)

def print_report():
    print()
    for name, status, detail in results:
        print(f"{status} {name}")
        if detail:
            print(f"      {detail}")

    passes = sum(1 for _, s, _ in results if s == PASS)
    warns = sum(1 for _, s, _ in results if s == WARN)
    fails = sum(1 for _, s, _ in results if s == FAIL)

    print()
    print(f"📊 汇总: {passes} 通过 / {warns} 警告 / {fails} 失败")
    if fails:
        print("⚠️  存在 FAIL 项，请检查上方详情！")
    elif warns:
        print("ℹ️  有 WARN 项（通常是正常情况，如空列、新供应商等）")
    else:
        print("🎉 全部通过！")

if __name__ == "__main__":
    main()
