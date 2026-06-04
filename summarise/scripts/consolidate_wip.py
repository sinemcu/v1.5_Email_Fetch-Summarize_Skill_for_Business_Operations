#!/usr/bin/env python3
"""
晟矽微运营 WIP 数据整合脚本
Consolidate multiple WIP source files into the standard 29-column 晟矽微运营汇总 format.

Usage:
    python3 consolidate_wip.py [input_dir]

  input_dir  Directory containing WIP attachment files (default: ~/Downloads/owa_attachments/)
             Also supports date subdirectories like ~/Downloads/owa_attachments/2026-05-29/

Dependencies:
    pip install openpyxl xlrd

Output: ~/Downloads/晟矽微运营汇总_<YYYYMMDD>.xlsx  (when date subdir is used)
        ~/Downloads/晟矽微运营汇总_整合版.xlsx      (default)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xlrd
import os
import sys
import glob
from datetime import datetime

# ── Input directory: from argv or default ──
if len(sys.argv) > 1:
    ATTACH_DIR = os.path.expanduser(sys.argv[1])
else:
    ATTACH_DIR = os.path.expanduser("~/Downloads/owa_attachments")

if not os.path.isdir(ATTACH_DIR):
    print(f"❌ Input directory not found: {ATTACH_DIR}")
    sys.exit(1)

# ── Output: include date in filename if input dir has a date ──
dir_basename = os.path.basename(ATTACH_DIR.rstrip('/'))
if dir_basename and dir_basename[:4].isdigit() and '-' in dir_basename[:10]:
    date_part = dir_basename[:10].replace('-', '')
    OUTPUT_PATH = os.path.expanduser(f"~/Downloads/晟矽微运营汇总_{date_part}.xlsx")
else:
    OUTPUT_PATH = os.path.expanduser("~/Downloads/晟矽微运营汇总_整合版.xlsx")

print(f"📂 Input:  {ATTACH_DIR}")
print(f"📤 Output: {OUTPUT_PATH}")
print()

# ── Helper: find file by keyword in ATTACH_DIR ──
def _find_file(keyword):
    """Find first file whose name contains the keyword (case-insensitive)."""
    for f in sorted(os.listdir(ATTACH_DIR)):
        fp = os.path.join(ATTACH_DIR, f)
        if os.path.isfile(fp) and keyword.lower() in f.lower():
            return fp
    return None


# ── Target 29-column headers ──
HEADERS = [
    "供应商", "客户代码", "客户名称", "客户订单号", "封装形式", "产品型号", "芯片型号",
    "晶圆批次", "订单数量", "投产日期",
    "粘片", "焊线", "焊线2", "焊线批检", "塑封", "电镀", "切筋", "委外切筋",
    "包装", "测试", "测试编带", "包装入库",
    "在线合计", "入库良品", "入库不良品", "出库良品", "出库不良品", "库存", "状态"
]

# ── Styles ──
hdr_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='2E75B6')
hdr_align = Alignment(horizontal='center', vertical='center')
data_font = Font(name='Calibri', size=11, color='FFFFFF')
data_align = Alignment(horizontal='center', vertical='center')
thin = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin')
)

# ── Color palette for auto-assignment ──
COLOR_PALETTE = [
    "9C27B0", "03A9F4", "4CAF50", "607D8B", "795548", "E91E63",
    "FF9800", "FFEB3B", "8BC34A", "3F51B5", "CDDC39", "00BCD4",
    "FF5722", "9E9E9E", "673AB7", "2196F3",
]

# ── Known supplier name aliases (aliases → canonical name) ──
SUPPLIER_ALIASES = {
    "气派":       "气派",
    "气派科技":   "气派",
    "代工-晟矽微": "晟矽微尊阳",
}


def build_supplier_color_map(all_rows):
    """Auto-detect suppliers, merge aliases, assign colors by data volume."""
    from collections import Counter

    canonical_rows = []
    for row in all_rows:
        sup = row[0]
        canonical_sup = SUPPLIER_ALIASES.get(sup, sup)
        canonical_rows.append(canonical_sup)

    counts = Counter(canonical_rows)
    sorted_suppliers = sorted(counts.keys(), key=lambda s: (-counts[s], s))

    known_colors = {
        "JC196巨成": "9C27B0", "晟矽微尊阳": "03A9F4",
        "中芯微": "607D8B", "气派": "FFEB3B", "内江明泰": "4CAF50",
    }
    cmap = {}
    used_colors = set()
    for sup in sorted_suppliers:
        if sup in known_colors:
            cmap[sup] = known_colors[sup]
            used_colors.add(known_colors[sup])

    for sup in sorted_suppliers:
        if sup not in cmap:
            for c in COLOR_PALETTE:
                if c not in used_colors:
                    cmap[sup] = c
                    used_colors.add(c)
                    break
            else:
                cmap[sup] = COLOR_PALETTE[hash(sup) % len(COLOR_PALETTE)]

    i = 0
    for row in all_rows:
        canonical_sup = SUPPLIER_ALIASES.get(row[0], row[0])
        row[0] = canonical_sup
        i += 1

    return cmap


def clean(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v == int(v):
            return int(v)
        return v
    s = str(v).strip()
    return s if s else None


def sf(v):
    try:
        return float(v)
    except Exception:
        return None


def supplier_fill(name, color_map):
    hexc = color_map.get(name, None)
    if hexc:
        return PatternFill('solid', fgColor=hexc)
    idx = hash(name) % len(COLOR_PALETTE)
    return PatternFill('solid', fgColor=COLOR_PALETTE[idx])


all_rows = []  # list of 29-element lists

# ════════════════════════════════════════════════
# 1. 晟矽微尊阳 — 晟矽微WIP*.xlsx (Sheet3) or 晟矽微尊阳WIP*.xls
# ════════════════════════════════════════════════
f_zunyang = _find_file("晟矽微WIP") or _find_file("晟矽微尊阳WIP")
if not f_zunyang:
    print("[1] 晟矽微尊阳: skipped (no file found)")
elif f_zunyang.endswith(".xlsx"):
    # Format: Sheet3, column-name mapped
    print(f"[1] 晟矽微尊阳: reading {os.path.basename(f_zunyang)}")
    wb = openpyxl.load_workbook(f_zunyang, data_only=True)
    ws = wb["Sheet3"]
    hr = None
    for r in range(1, 10):
        c1 = ws.cell(row=r, column=1).value
        if c1 and "客户代码" in str(c1):
            hr = r
            break

    if hr:
        smap = {}
        for c in range(1, 45):
            v = ws.cell(row=hr, column=c).value
            if v:
                smap[str(v).strip()] = c

        def g(r, name):
            return ws.cell(row=r, column=smap.get(name, 0)).value if name in smap else None

        for r in range(hr + 1, ws.max_row + 1):
            if not g(r, "客户代码"):
                continue
            all_rows.append([
                "晟矽微尊阳", clean(g(r, "客户代码")), clean(g(r, "客户名称")),
                clean(g(r, "客户订单号")), clean(g(r, "工艺路线（封装形式）")),
                clean(g(r, "产品型号")), clean(g(r, "芯片型号")),
                clean(g(r, "晶圆批次")), clean(g(r, "订单数量(下单数量)")),
                clean(g(r, "投产日期")),
                clean(g(r, "Die_Bonding粘片")), clean(g(r, "Wire_Bonding键合")),
                None, None,
                clean(g(r, "Mol_ding模封")), clean(g(r, "Plating电镀")),
                clean(g(r, "Trim_Form切筋打弯")), None,
                clean(g(r, "Packing包装")), clean(g(r, "Testtu_bassembly测试管装")),
                None, None,
                clean(g(r, "在线合计")), clean(g(r, "入库合计（良品）")),
                clean(g(r, "入库合计（不良品）")), clean(g(r, "出库（良品）")),
                clean(g(r, "出库（不良品）")), clean(g(r, "库存")),
                clean(g(r, "是否结批")),
            ])
    print(f"    → {sum(1 for r in all_rows if r[0] == '晟矽微尊阳')} rows")
else:
    # Format: 晟矽微尊阳WIP*.xls (xlrd, Sheet1, fixed-column positions)
    print(f"[1] 晟矽微尊阳: reading {os.path.basename(f_zunyang)} (xls format)")
    wb_x = xlrd.open_workbook(f_zunyang)
    ws_x = wb_x.sheet_by_index(0)
    start = len(all_rows)
    for r in range(1, ws_x.nrows):
        cc = ws_x.cell(r, 2).value  # 客户编码
        if not cc:
            continue
        all_rows.append([
            "晟矽微尊阳", clean(cc), None,
            clean(ws_x.cell(r, 3).value),  # 客户订单号
            clean(ws_x.cell(r, 5).value),  # 客户封装形式
            None,  # 产品型号
            None,  # 芯片型号
            clean(ws_x.cell(r, 9).value),  # 扩散批号 → 晶圆批次
            clean(ws_x.cell(r, 10).value),  # 来料数 → 订单数量
            None,
            clean(ws_x.cell(r, 13).value),  # 贴片数量 → 粘片
            clean(ws_x.cell(r, 29).value),  # 焊线1数量 → 焊线
            clean(ws_x.cell(r, 31).value),  # 焊线2数量
            None,
            clean(ws_x.cell(r, 39).value),  # 塑封数量
            clean(ws_x.cell(r, 47).value),  # 电镀数量
            clean(ws_x.cell(r, 52).value),  # 切筋数量
            None,
            clean(ws_x.cell(r, 65).value),  # 包装数量
            clean(ws_x.cell(r, 62).value),  # 测试数量
            None, None,
            None,
            clean(ws_x.cell(r, 66).value),  # ERP库存数量 → 入库良品
            None,
            clean(ws_x.cell(r, 67).value),  # ERP出库数量
            None, None, None,
        ])
    print(f"    → {len(all_rows) - start} rows")

# ════════════════════════════════════════════════
# 2. 气派 — GD*.xlsx
# ════════════════════════════════════════════════
f = _find_file("GD1458") or _find_file("GD")
if not f:
    print("[2] 气派: skipped (no file found)")
else:
    print(f"[2] 气派: reading {os.path.basename(f)}")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active
    smap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v:
            smap[str(v).strip()] = c

    def g(r, name):
        return ws.cell(row=r, column=smap.get(name, 0)).value if name in smap else None

    start = len(all_rows)
    for r in range(2, ws.max_row + 1):
        if not g(r, "客户代码"):
            continue
        cc = str(g(r, "客户代码")).strip()
        sup = "气派科技" if "1458" in cc else "气派"

        steps = ["装片", "固化", "等离子清洗1", "键合", "键合QC", "键合全检",
                 "塑封", "去浇口", "后固化", "激光去胶", "回流焊", "打印1",
                 "打印2", "打印QC", "打印全检", "电镀", "划不良", "电镀后外观检",
                 "切筋", "成型", "切割", "烘烤", "测试待投", "测试", "编带",
                 "挑打叉", "三次固化", "测试后全检", "外观", "FQC", "包装"]
        online = sum(int(sf(g(r, s)) or 0) for s in steps)

        all_rows.append([
            sup, clean(g(r, "客户代码")), None,
            clean(g(r, "客户订单号")), clean(g(r, "封装形式")),
            clean(g(r, "电路名")), clean(g(r, "芯片名")),
            clean(g(r, "芯片批号")), clean(g(r, "订单数量")),
            clean(g(r, "投料时间")),
            clean(g(r, "装片")), clean(g(r, "键合")),
            None, None,
            clean(g(r, "塑封")), clean(g(r, "电镀")),
            clean(g(r, "切筋")), None,
            clean(g(r, "包装")), clean(g(r, "测试")),
            clean(g(r, "编带")), None,
            online if online else None,
            clean(g(r, "入库数量")), None,
            clean(g(r, "出库数量")), None,
            clean(g(r, "库存数量")),
            clean(g(r, "订单状态")),
        ])
    print(f"    → {len(all_rows) - start} rows")

# ════════════════════════════════════════════════
# 3. JC196巨成 — JC196*.xlsx
# ════════════════════════════════════════════════
f = _find_file("JC196")
if not f:
    print("[3] JC196巨成: skipped (no file found)")
else:
    print(f"[3] JC196巨成: reading {os.path.basename(f)}")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active
    smap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v:
            smap[str(v).strip()] = c

    def g(r, name):
        return ws.cell(row=r, column=smap.get(name, 0)).value if name in smap else None

    start = len(all_rows)
    for r in range(2, ws.max_row + 1):
        if not g(r, "客户代码"):
            continue
        all_rows.append([
            "JC196巨成", clean(g(r, "客户代码")), None,
            clean(g(r, "客户订单号")), clean(g(r, "封装形式")),
            clean(g(r, "产品名称")), clean(g(r, "芯片名称")),
            clean(g(r, "芯片批号")), clean(g(r, "工单数量")),
            clean(g(r, "上线日期")),
            clean(g(r, "粘片1")), clean(g(r, "焊线1")),
            clean(g(r, "焊线2")), None,
            clean(g(r, "注塑")), None,
            clean(g(r, "切筋")), None,
            clean(g(r, "包装入库")), clean(g(r, "测试")),
            None, None,
            None, clean(g(r, "成品入库")),
            None, clean(g(r, "发货数量")),
            None, clean(g(r, "工序总数")),
            None,
        ])
    print(f"    → {len(all_rows) - start} rows")

# ════════════════════════════════════════════════
# 4. 中芯微 — ccml*.xls
# ════════════════════════════════════════════════
ccml_files = sorted(glob.glob(os.path.join(ATTACH_DIR, "ccml*.xls")))
if not ccml_files:
    print("[4] 中芯微: skipped (no ccml*.xls found)")
else:
    total_ccml = 0
    for fp in ccml_files:
        wb_x = xlrd.open_workbook(fp)
        ws_x = wb_x.sheet_by_index(0)
        smap = {}
        for c in range(ws_x.ncols):
            v = ws_x.cell_value(2, c)
            if v:
                smap[str(v).strip()] = c

        def gx(r, name):
            if name not in smap:
                return None
            cell = ws_x.cell(r, smap[name])
            return cell.value if cell.ctype != xlrd.XL_CELL_EMPTY else None

        start = len(all_rows)
        for r in range(3, ws_x.nrows):
            if not gx(r, "客户代码"):
                continue
            all_rows.append([
                "中芯微", clean(gx(r, "客户代码")), None,
                clean(gx(r, "客户订单号")), clean(gx(r, "封装形式")),
                clean(gx(r, "产品型号")), clean(gx(r, "芯片型号")),
                clean(gx(r, "芯片批次")), clean(gx(r, "工单数量")),
                clean(gx(r, "工单日期")),
                clean(gx(r, "装片")), clean(gx(r, "键合")),
                None, None,
                clean(gx(r, "塑封")), clean(gx(r, "电镀")),
                clean(gx(r, "切筋")), None,
                clean(gx(r, "包装")), clean(gx(r, "测试")),
                None, None,
                None, clean(gx(r, "完工数量")),
                None, None, None, None,
                clean(gx(r, "订单类型")),
            ])
        cnt = len(all_rows) - start
        total_ccml += cnt
        print(f"    {os.path.basename(fp)} → {cnt} rows")
    print(f"    → {total_ccml} rows total")

# ════════════════════════════════════════════════
# 5. 内江明泰 — 内江*.xls
# ════════════════════════════════════════════════
nj_files = sorted(glob.glob(os.path.join(ATTACH_DIR, "内江*.xls")))
if not nj_files:
    print("[5] 内江明泰: skipped (no 内江*.xls found)")
else:
    total_nj = 0
    for fp in nj_files:
        print(f"[5] 内江明泰: reading {os.path.basename(fp)}")
        wb_x = xlrd.open_workbook(fp)
        ws_x = wb_x.sheet_by_index(0)
        smap = {}
        for c in range(ws_x.ncols):
            v = ws_x.cell_value(0, c)
            if v:
                smap[str(v).strip()] = c

        def gn(r, name):
            if name not in smap:
                return None
            cell = ws_x.cell(r, smap[name])
            return cell.value if cell.ctype != xlrd.XL_CELL_EMPTY else None

        start = len(all_rows)
        for r in range(1, ws_x.nrows):
            if not gn(r, "客户编码"):
                continue
            all_rows.append([
                "内江明泰", clean(gn(r, "客户编码")), None,
                clean(gn(r, "客户PO")), clean(gn(r, "封装形式")),
                clean(gn(r, "产品代码")), clean(gn(r, "产品名称")),
                clean(gn(r, "A芯片晶圆批号")), clean(gn(r, "订单数量")),
                clean(gn(r, "计划开工日期")),
                None, None, None, None, None, None, None, None,
                None, None, None, None,
                clean(gn(r, "在制品数量")), None, None, None, None, None, None,
            ])
        cnt = len(all_rows) - start
        total_nj += cnt
        print(f"    → {cnt} rows")
    print(f"    → {total_nj} rows total")

# ════════════════════════════════════════════════
# 6. 新工序 — 新工序在制WIP*.xls / 新工序在制WIP*.xlsx
#     (供应商名称取 "产品归属" 列, e.g. "代工-晟矽微")
# ════════════════════════════════════════════════
xgx_files = sorted(glob.glob(os.path.join(ATTACH_DIR, "新工序在制WIP*")))
if not xgx_files:
    print("[6] 新工序: skipped (no file found)")
else:
    total_xgx = 0
    for fp in xgx_files:
        print(f"[6] 新工序: reading {os.path.basename(fp)}")
        if fp.endswith(".xlsx"):
            wb_x = openpyxl.load_workbook(fp, data_only=True)
            ws_x = wb_x.active
            # Header in row 1, data from row 2
            hmap = {}
            for c in range(1, ws_x.max_column + 1):
                v = ws_x.cell(row=1, column=c).value
                if v:
                    hmap[str(v).strip()] = c

            def gxg(r, name):
                if name not in hmap:
                    return None
                return ws_x.cell(row=r, column=hmap[name]).value

            start = len(all_rows)
            for r in range(2, ws_x.max_row + 1):
                order_no = gxg(r, "客户订单号")
                if not order_no:
                    continue
                sup = clean(gxg(r, "产品归属")) or "新工序"
                all_rows.append([
                    sup, None, None,
                    clean(order_no), clean(gxg(r, "规格明细")),
                    clean(gxg(r, "物料名称")), clean(gxg(r, "芯片型号")),
                    clean(gxg(r, "芯片批号")), clean(gxg(r, "订单数量")),
                    clean(gxg(r, "单据日期")),
                    clean(gxg(r, "粘片")), clean(gxg(r, "焊线")),
                    clean(gxg(r, "焊线2")), clean(gxg(r, "焊线批检")),
                    clean(gxg(r, "塑封")), clean(gxg(r, "电镀")),
                    clean(gxg(r, "切筋")), clean(gxg(r, "委外切筋")),
                    clean(gxg(r, "包装")), clean(gxg(r, "测试")),
                    clean(gxg(r, "测试编带")), clean(gxg(r, "包装入库")),
                    None,
                    clean(gxg(r, "未入库数量")),
                    None, None, None, None, None,
                ])
            cnt = len(all_rows) - start
            total_xgx += cnt
            print(f"    → {cnt} rows")
    print(f"    → {total_xgx} rows total")

# ════════════════════════════════════════════════
# 7. 锐骏 — 174-WIP*.xls (by process-step aggregation)
# ════════════════════════════════════════════════
rj_files = sorted(glob.glob(os.path.join(ATTACH_DIR, "174-WIP*.xls")))
if not rj_files:
    print("[7] 锐骏: skipped (no file found)")
else:
    total_rj = 0
    for fp in rj_files:
        print(f"[7] 锐骏: reading {os.path.basename(fp)}")
        wb_x = xlrd.open_workbook(fp)
        ws_x = wb_x.sheet_by_index(0)
        start = len(all_rows)
        # Build col map from header row
        hmap = {}
        for c in range(ws_x.ncols):
            v = ws_x.cell_value(0, c)
            if v:
                hmap[str(v).strip()] = c

        def grj(r, name):
            if name not in hmap:
                return None
            cell = ws_x.cell(r, hmap[name])
            return cell.value if cell.ctype != xlrd.XL_CELL_EMPTY else None

        for r in range(1, ws_x.nrows):
            cc = grj(r, "客户代码")
            if not cc:
                continue
            all_rows.append([
                "锐骏", clean(cc), None,
                clean(grj(r, "客户订单")), clean(grj(r, "封装形式")),
                clean(grj(r, "产品型号")), None,
                clean(grj(r, "晶圆批号")), clean(grj(r, "数量")),
                None,
                None, None, None, None, None, None, None, None,
                None, None, None, None,
                None, None, None, None, None, None, None,
            ])
        cnt = len(all_rows) - start
        total_rj += cnt
        print(f"    → {cnt} rows")
    print(f"    → {total_rj} rows total")

# ════════════════════════════════════════════════
# 8. 芯丰 — 132*.xlsx / 102*.xlsx (生产订单汇总格式)
# ════════════════════════════════════════════════
xf_files = sorted(glob.glob(os.path.join(ATTACH_DIR, "132*.xlsx")) +
                  glob.glob(os.path.join(ATTACH_DIR, "102*.xlsx")))
if not xf_files:
    print("[8] 芯丰: skipped (no file found)")
else:
    total_xf = 0
    for fp in xf_files:
        print(f"[8] 芯丰: reading {os.path.basename(fp)}")
        wb_x = openpyxl.load_workbook(fp, data_only=True)
        ws_x = wb_x.active
        hmap = {}
        for c in range(1, ws_x.max_column + 1):
            v = ws_x.cell(row=1, column=c).value
            if v:
                hmap[str(v).strip()] = c

        def gxf(r, name):
            if name not in hmap:
                return None
            return ws_x.cell(row=r, column=hmap[name]).value

        # Extract client code from filename (132 or 102)
        basename = os.path.basename(fp)
        client_code = basename.split(".")[0].strip()

        start = len(all_rows)
        for r in range(2, ws_x.max_row + 1):
            order_no = gxf(r, "生产订单号")
            if not order_no:
                continue
            all_rows.append([
                "芯丰", clean(client_code), None,
                clean(gxf(r, "客户订单编码")), clean(gxf(r, "封装形式")),
                clean(gxf(r, "产品型号")), None,
                None, clean(gxf(r, "生产订单数量")),
                clean(gxf(r, "订单开始日期")),
                None, None, None, None, None, None, None, None,
                None, None, None, None,
                None, None, None, None, None, None, None,
            ])
        cnt = len(all_rows) - start
        total_xf += cnt
        print(f"    → {cnt} rows")
    print(f"    → {total_xf} rows total")

# ════════════════════════════════════════════════
# Write output workbook
# ════════════════════════════════════════════════
if not all_rows:
    print("\n❌ No data rows collected. Check that input files exist and are supported.")
    sys.exit(1)

supplier_color_map = build_supplier_color_map(all_rows)
print(f"\n📊 供应商分布（按数据量排序，自动分配配色）:")
for sup, hexc in supplier_color_map.items():
    count = sum(1 for r in all_rows if r[0] == sup)
    print(f"  {sup}: {count} 行 → #{hexc}")

wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "Sheet"

for c, h in enumerate(HEADERS, 1):
    cell = ws_out.cell(row=1, column=c, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = hdr_align
    cell.border = thin

for i, row in enumerate(all_rows):
    r = i + 2
    fill = supplier_fill(row[0], supplier_color_map)
    for c, val in enumerate(row, 1):
        cell = ws_out.cell(row=r, column=c, value=val)
        cell.font = data_font
        cell.fill = fill
        cell.alignment = data_align
        cell.border = thin

for c in range(1, 30):
    ws_out.column_dimensions[get_column_letter(c)].width = 14

ws_out.freeze_panes = "A2"

wb_out.save(OUTPUT_PATH)
print(f"\n✅ Saved {len(all_rows)} rows → {OUTPUT_PATH}")
