# -*- coding: utf-8 -*-
"""
销售预测 捆绑SKU拆解 工具 (v6 美化闭环版)
================================================================
读取销售预测数据（运营表 + 捆绑BOM + 单品产品资料），将捆绑SKU按 BOM 关系
递归拆解为单品SKU，原单品保留。

若同时提供 --inv-file（库存文件），则额外生成 Sheet1「销售需求汇总」
（按拆解需求主SKU+店铺+发货模式+运营汇总，含按月从老到新分配+表格美化）。
--feishu-url 可选，默认内置飞书整套替代关系表URL。

输出:
  有 --inv-file:  3 sheets → 销售需求汇总 + 未拆解整套版销售需求 + 异常提醒
  无 --inv-file:  2 sheets → 未拆解整套版销售需求 + 异常提醒（向后兼容）

用法:
  python decompose.py <源文件.xlsx> [--out 输出.xlsx]
                       [--op-sheet 运营表名] [--bom-sheet 捆绑表名]
                       [--bom-csv 同款套装库存.csv] [--product-csv product.csv]
                       [--enc 编码]
                       [--feishu-url 飞书表格URL] [--inv-file 库存文件.xlsx]

依赖: pywin32, openpyxl  （仅 Windows，且本机已安装 Excel）
      lark-cli（仅在使用 --feishu-url 时需要）

业务规则（与已确认的最终版本一致）:
- 捆绑判定 = BOM 匹配制：运营表 SKU 能在捆绑表 A 列(捆绑SKU)匹配到即拆解，不看 -KB 后缀。
- 拆解键 = 按运营表的 (SKU + 店铺 + 发货模式 + 运营) 逐行拆解。
- 拆解出行填充:
    1) 一级分类/产品定级/渠道/国家/销售状态 按 (组件SKU + 店铺 + 运营 + 发货模式) 匹配
    2) MSKU = /
    3) 品名描述/型号SPU 用捆绑表 J-L 列（产品名称/产品款式）
- 取整: 非整数预测 -> 四舍五入(ROUND_HALF_UP); 安全库存(O列) 也取整。
- 汇总(N列,逐行) = 取整(安全库存) + 12个月销售预测求和。
- 异常: 后缀 -KB 但 BOM 无对应捆绑关系 -> 保留原行不拆解并标红提醒;
        单品未匹配产品资料(J-L) -> 提醒。
- 输出模板: 原始 27 列(A~AA) 严格按原顺序 + 原两行表头(销量预测合并 P:AA + 月份)，
  其后追加 6 列溯源/警告(来源/父捆绑SKU/关联数量/产品名称/产品款式/警告)。

Sheet1 销售需求汇总 逻辑:
  - 从飞书「整套替代关系确认」+「混用SKU」读取SKU映射
  - 从库存文件「海外仓库存」读取归一化映射（店铺→主店铺, 运营→运营归一）
  - 按拆解需求主SKU+主店铺+发货模式+归一运营汇总，输出含月份的汇总表
  - 相似SKU优先从非下单套装（老套装）获取
"""
import argparse
import csv
import os
import sys
import json
import io
import re
import math
import subprocess
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- 常量 ----
MAX_DEPTH = 20
FLOAT_TOL = 1e-9
DEFAULT_FSTART = 15
ROW_WIDTH = 40
HELPER = ["来源", "父捆绑SKU", "关联数量", "产品名称", "产品款式", "警告"]

# 飞书默认URL（整套替代关系+混用SKU表）
FEISHU_URL_DEFAULT = "https://ccn2tbrt0geg.feishu.cn/sheets/DqBNsJNNghEkgkt6cjMcbCWqnVb"

# 库存组件
INV_COMPONENTS = ["海外在途", "海外在仓", "借入", "借出", "待交付合同", "国内在仓", "国内待下合同"]


# ==================== 取整 / 单元格工具 ====================

def to_int(cell):
    """取整（四舍五入 ROUND_HALF_UP），空值返回 0。"""
    if isinstance(cell, (int, float)):
        s = str(cell)
    else:
        s = (cell or "").strip()
    if s == "":
        return 0
    try:
        return int(Decimal(s).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except Exception:
        try:
            return int(round(float(s)))
        except Exception:
            return 0


def to_int_tracked(cell):
    """取整并返回是否为非整数（发生了四舍五入）。返回 (int, was_rounded)。"""
    if isinstance(cell, (int, float)):
        s = str(cell)
    else:
        s = (cell or "").strip()
    if s == "":
        return 0, False
    try:
        d = Decimal(s)
        result = int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        was_rounded = abs(float(s) - result) > FLOAT_TOL
        return result, was_rounded
    except Exception:
        try:
            f = float(s)
            result = int(round(f))
            was_rounded = abs(f - result) > FLOAT_TOL
            return result, was_rounded
        except Exception:
            return 0, False


def cell(row, idx):
    if idx is None or not (0 <= idx < len(row)):
        return ""
    return row[idx]


def _norm_cell(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return "%d年%d月" % (v.year, v.month)
    return v


def safe_str(v):
    if v is None:
        return ""
    return str(v).strip()


def to_num(v):
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return 0


# ==================== Excel COM 读取 ====================

def read_workbook(path, op_sheet=None, bom_sheet=None, read_bom=True):
    """用 pywin32 启动 Excel，单次打开工作簿读取运营/捆绑表。
    返回 (names, op_name, bom_name, op_rows, bom_rows)。
    """
    import win32com.client as win32
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    wb = None
    try:
        wb = xl.Workbooks.Open(path, False, True)
        names = [ws.Name for ws in wb.Worksheets]
        op_name = op_sheet or next((n for n in names if "运营" in n), names[0])
        bom_name = bom_sheet or next((n for n in names if "捆绑" in n), names[-1])

        def grab(sheet_name):
            val = wb.Worksheets(sheet_name).UsedRange.Value
            if val is None:
                return []
            if isinstance(val, tuple):
                rows = [list(r) if isinstance(r, tuple) else [r] for r in val]
            else:
                rows = [[val]]
            mx = max((len(r) for r in rows), default=0)
            return [[_norm_cell(v) for v in (list(r) + [""] * (mx - len(r)))] for r in rows]

        op_rows = grab(op_name)
        bom_rows = grab(bom_name) if read_bom else []
    finally:
        if wb is not None:
            wb.Close(False)
        xl.Quit()
    return names, op_name, bom_name, op_rows, bom_rows


class COMWb:
    """Context manager for reading DRM-encrypted xlsx via Excel COM."""
    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.xl = None
        self.wb = None
    def __enter__(self):
        import win32com.client as win32
        self.xl = win32.DispatchEx("Excel.Application")
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.wb = self.xl.Workbooks.Open(self.path, ReadOnly=True)
        return self.wb
    def __exit__(self, *a):
        try:
            if self.wb:
                self.wb.Close(False)
        except Exception:
            pass
        try:
            if self.xl:
                self.xl.Quit()
        except Exception:
            pass


def com_read_sheet(wb, sheet_name, start_row=1):
    """Read a sheet as 2D list using CurrentRegion + Value2 batch read."""
    ws = wb.Sheets(sheet_name)
    used = ws.Range("A1").CurrentRegion
    nrows = used.Rows.Count
    ncols = used.Columns.Count
    data = ws.Range(ws.Cells(start_row, 1), ws.Cells(nrows, ncols)).Value2
    if data is None:
        return []
    if isinstance(data, (int, float, str)):
        return [[data]]
    return [list(r) for r in data]


def com_read_sheet_used(wb, sheet_name, start_row=1):
    """Read a sheet using UsedRange (handles sheets with empty spacer rows)."""
    ws = wb.Sheets(sheet_name)
    used = ws.UsedRange
    nrows = used.Rows.Count
    ncols = used.Columns.Count
    data = ws.Range(ws.Cells(start_row, 1), ws.Cells(nrows, ncols)).Value2
    if data is None:
        return []
    if isinstance(data, (int, float, str)):
        return [[data]]
    return [list(r) for r in data]


# ==================== 列定位 ====================

def col_index(header, kw, exclude=None):
    for i, h in enumerate(header):
        if h and kw in str(h) and (exclude is None or exclude not in str(h)):
            return i
    return None


def find_header_row(rows, kw, col=None):
    for i, r in enumerate(rows):
        if col is not None:
            if str(cell(r, col)).strip() == kw:
                return i
        else:
            if any(kw in str(c) for c in r if c):
                return i
    return 0


# ==================== 递归拆解 ====================

def make_expand(bom_map):
    def expand(sku, mult, visited, root_sku, depth=0):
        if depth > MAX_DEPTH:
            return [(sku, mult, root_sku, True)]
        su = sku.upper()
        if su in bom_map:
            res = []
            new_visited = visited | {su}
            for c, q in bom_map[su]:
                cu = c.upper()
                if cu in new_visited:
                    res.append((c, mult * q, root_sku, True))
                    continue
                res += expand(c, mult * q, new_visited, root_sku, depth + 1)
            return res
        return [(sku, mult, root_sku, False)]
    return expand


# ==================== CSV 读取 ====================

def _open_csv(path, enc=None):
    if enc:
        return open(path, encoding=enc, errors="ignore", newline="")
    for e in ("gbk", "utf-8-sig", "utf-8"):
        try:
            with open(path, encoding=e, errors="strict", newline="") as t:
                t.read(4096)
            return open(path, encoding=e, errors="ignore", newline="")
        except (UnicodeDecodeError, LookupError):
            continue
    return open(path, encoding="latin-1", errors="ignore", newline="")


def build_bom_product(bom_csv, product_csv, enc=None):
    prod_master = {}
    with _open_csv(product_csv, enc) as f:
        r = csv.reader(f)
        hdr = next(r)
        s_c = col_index(hdr, "产品SKU")
        n_c = col_index(hdr, "产品名称")
        st_c = col_index(hdr, "产品款式")
        if s_c is None:
            raise RuntimeError("product_csv 未找到 产品SKU 列")
        for row in r:
            if s_c >= len(row):
                continue
            j = row[s_c].strip()
            if j:
                prod_master[j.upper()] = (
                    row[n_c].strip() if (n_c is not None and n_c < len(row)) else "",
                    row[st_c].strip() if (st_c is not None and st_c < len(row)) else "",
                )

    bom_map = defaultdict(list)
    seen = set()
    with _open_csv(bom_csv, enc) as f:
        r = csv.reader(f)
        hdr = next(r)
        sku_c = col_index(hdr, "仓库SKU")
        parent_c = col_index(hdr, "父仓库SKU")
        qty_c = col_index(hdr, "套装数量")
        if parent_c is None or sku_c is None:
            raise RuntimeError("bom_csv 未找到 仓库SKU / 父仓库SKU 列")
        for row in r:
            p = row[parent_c].strip() if parent_c < len(row) else ""
            c = row[sku_c].strip() if sku_c < len(row) else ""
            q = row[qty_c].strip() if (qty_c is not None and qty_c < len(row)) else "1"
            if p and c:
                try:
                    qq = float(q)
                except Exception:
                    qq = 1.0
                key = (p.upper(), c.upper())
                if key not in seen:
                    bom_map[p.upper()].append((c, qq))
                    seen.add(key)
    return prod_master, bom_map


# ==================== 明细处理 ====================

def process_rows(data, bom_map, prod_master, op_attr, cc):
    expand = make_expand(bom_map)

    def get_prod(sku):
        if sku in prod_master:
            return prod_master[sku]
        if sku.upper() in prod_master:
            return prod_master[sku.upper()]
        return None

    detail = []
    missing_bundle = []
    missing_prod = []
    rounded_cells = 0
    n_bundle = n_single = n_missbundle = 0

    for row in data:
        row = list(row) + [""] * max(0, ROW_WIDTH - len(row))
        sku = str(cell(row, cc["sku_c"])).strip()
        ship = str(cell(row, cc["ship_c"])).strip()
        shop = str(cell(row, cc["shop_c"])).strip()
        opg = str(cell(row, cc["opg_c"])).strip()

        orig_monthly = []
        for i in cc["fc"]:
            val, was_rounded = to_int_tracked(cell(row, i))
            orig_monthly.append(val)
            if was_rounded:
                rounded_cells += 1

        if sku.upper() in bom_map:
            results = expand(sku, 1.0, set(), sku)
            for (leaf, mult, root_sku, cyc) in results:
                leaf_monthly = [to_int(mv * mult) for mv in orig_monthly]
                pm = get_prod(leaf)
                pname = pm[0] if pm else ""
                pstyle = pm[1] if pm else ""
                key = (leaf.upper(), shop, opg, ship)
                if key in op_attr:
                    cat, grade, channel, country, status = op_attr[key]
                else:
                    cat = str(cell(row, cc["cat_c"])).strip()
                    grade = str(cell(row, cc["grade_c"])).strip()
                    channel = str(cell(row, cc["chan_c"])).strip()
                    country = str(cell(row, cc["coun_c"])).strip()
                    status = str(cell(row, cc["stat_c"])).strip()
                warn = []
                if cyc:
                    warn.append("⚠检测到循环依赖，未进一步拆解")
                if pm is None:
                    warn.append("⚠未匹配到产品资料(J-L)")
                    missing_prod.append((leaf, "捆绑拆解", root_sku))
                ss = to_int(cell(row, cc["ss_col"]))
                total = ss + sum(leaf_monthly)
                out = [leaf, pname, "/", pstyle, cat, grade, channel, country,
                       shop, ship, opg, str(cell(row, cc["list_c"])).strip(), status, total, ss
                       ] + leaf_monthly + [
                    "捆绑拆解", root_sku, (mult if mult != 1 else 1),
                    pname, pstyle, "；".join(warn)]
                detail.append(out)
                n_bundle += 1
        else:
            pm = get_prod(sku)
            pname = pm[0] if pm else ""
            pstyle = pm[1] if pm else ""
            is_kb = sku.upper().endswith("-KB")
            warn = []
            if is_kb:
                warn.append("⚠后缀-KB但BOM中无对应捆绑关系，未拆解")
                missing_bundle.append((sku, str(cell(row, cc["pname_c"])).strip(), str(cell(row, cc["chan_c"])).strip(),
                                       str(cell(row, cc["coun_c"])).strip(), str(cell(row, cc["shop_c"])).strip(), ship))
                n_missbundle += 1
                src_label = "原捆绑(缺失关系)"
            else:
                src_label = "原单品"
            if pm is None and not is_kb:
                warn.append("⚠未匹配到产品资料(J-L)")
                missing_prod.append((sku, src_label, ""))
            ss = to_int(cell(row, cc["ss_col"]))
            total = ss + sum(orig_monthly)
            base = list(row[:cc["fstart"]])
            base[cc["tot_col"]] = total
            base[cc["ss_col"]] = ss
            out = base + orig_monthly + [src_label, "", 1, pname, pstyle, "；".join(warn)]
            detail.append(out)
            n_single += 1

    return detail, missing_bundle, missing_prod, n_bundle, n_single, n_missbundle, rounded_cells


def check_formulas(detail):
    formula_bad = []
    for i, d in enumerate(detail, start=3):
        ss, total = d[14], d[13]
        monthly = d[15:27]
        if not (isinstance(total, (int, float)) and isinstance(ss, (int, float))
                and total == ss + sum(monthly)):
            formula_bad.append((i, d[0]))
    return formula_bad


# ==================== 飞书映射读取（Sheet1用） ====================

def lark_csv(sheet_name, rng, feishu_url):
    """Read Feishu sheet as CSV via lark-cli."""
    r = subprocess.run(
        ["lark-cli", "sheets", "+csv-get", "--url", feishu_url,
         "--sheet-name", sheet_name, "--range", rng],
        capture_output=True, text=True, timeout=120)
    data = json.loads(r.stdout)
    raw = data.get("data", {}).get("annotated_csv", "")
    lines = [re.sub(r'^\[row=\d+\] ', '', l) for l in raw.split('\n') if l.strip()]
    return list(csv.reader(io.StringIO('\n'.join(lines))))


def read_feishu_mapping(feishu_url):
    """Build mapping from Feishu tables.
    Returns: (src_to_map, zhu_to_sets, zhu_to_xiangsi, zhu_to_order)
    """
    src_to_map = {}
    zhu_to_sets = defaultdict(list)
    zhu_to_xiangsi = {}
    zhu_to_order = {}
    zhu_xiangsi_candidates = defaultdict(list)

    rows = lark_csv("整套替代关系确认", "A1:M911", feishu_url)
    for row in rows[1:]:
        if len(row) < 5:
            continue
        zhu = safe_str(row[0])
        xiangsi = safe_str(row[1])
        hun = safe_str(row[2])
        yuan = safe_str(row[3])
        order = safe_str(row[4]) if len(row) > 4 else ""

        if yuan:
            src_to_map[yuan] = {"zhu": zhu, "hun": hun, "xiangsi": xiangsi, "order": order}
        zhu_xiangsi_candidates[zhu].append((xiangsi, hun, order))
        zhu_to_order[zhu] = order

        sets = zhu_to_sets[zhu]
        found = False
        for s in sets:
            if s["hun"] == hun:
                if yuan and yuan not in s["sources"]:
                    s["sources"].append(yuan)
                found = True
                break
        if not found:
            sets.append({"hun": hun, "is_order": False, "sources": [yuan] if yuan else []})

    for zhu, sets in zhu_to_sets.items():
        order = zhu_to_order.get(zhu, "")
        for s in sets:
            s["is_order"] = (s["hun"] == order)

    # 相似SKU优先从非下单套装（老套装）获取
    for zhu, candidates in zhu_xiangsi_candidates.items():
        order = zhu_to_order.get(zhu, "")
        non_order = [c[0] for c in candidates if c[1] != order and c[0]]
        if non_order:
            zhu_to_xiangsi[zhu] = non_order[0]
        else:
            zhu_to_xiangsi[zhu] = candidates[0][0] if candidates else ""

    for yuan, m in src_to_map.items():
        zhu = m["zhu"]
        if zhu in zhu_to_xiangsi:
            m["xiangsi"] = zhu_to_xiangsi[zhu]

    rows2 = lark_csv("混用SKU", "A1:C2415", feishu_url)
    for row in rows2[2:]:
        if len(row) < 3:
            continue
        hun = safe_str(row[0])
        yuan = safe_str(row[1])
        xiangsi = safe_str(row[2])
        if yuan and yuan not in src_to_map:
            src_to_map[yuan] = {"zhu": yuan, "hun": hun, "xiangsi": xiangsi, "order": ""}

    return src_to_map, dict(zhu_to_sets), zhu_to_xiangsi, zhu_to_order


# ==================== 归一化映射读取（Sheet1用） ====================

def read_normalization(inv_file):
    """Build normalization mapping from 海外仓库存 sheet.
    Returns: (shop_to_main, person_to_norm)
    """
    shop_to_main = {}
    person_to_norm = {}

    with COMWb(inv_file) as wb:
        data = com_read_sheet(wb, "海外仓库存", start_row=2)
        for row in data:
            if not row or not row[0]:
                continue
            main_shop = safe_str(row[2]) if len(row) > 2 else ""
            norm_person = safe_str(row[4]) if len(row) > 4 else ""
            shop = safe_str(row[5]) if len(row) > 5 else ""
            person = safe_str(row[6]) if len(row) > 6 else ""

            if shop and main_shop:
                shop_to_main[shop] = main_shop
            if person and norm_person:
                person_to_norm[person] = norm_person

    return shop_to_main, person_to_norm


def normalize_shop(shop, shop_to_main):
    return shop_to_main.get(shop, shop)


def normalize_person(person, person_to_norm):
    return person_to_norm.get(person, person)


# ==================== 库存读取（捆绑拆分用） ====================

def read_inventory(inv_file):
    """Read inventory from 5 sheets, build key -> component dict.
    Key format: 混用主SKU&主店铺&发货模式&运营归一
    """
    inv_map = defaultdict(lambda: dict(zip(INV_COMPONENTS, [0]*7)))

    with COMWb(inv_file) as wb:
        # 1. 海外仓库存: col1=key, col8=库存
        data = com_read_sheet(wb, "海外仓库存", start_row=2)
        for row in data:
            if not row or not row[0]:
                continue
            key = safe_str(row[0])
            qty = to_num(row[7]) if len(row) > 7 else 0
            inv_map[key]["海外在仓"] += qty

        # 2. 在途&配柜: col1=key, col13=出货数量
        data = com_read_sheet(wb, "在途&配柜", start_row=2)
        for row in data:
            if not row or not row[0]:
                continue
            key = safe_str(row[0])
            qty = to_num(row[12]) if len(row) > 12 else 0
            inv_map[key]["海外在途"] += qty

        # 3. 合同待交付: col1=key, col20=浦江可用, col26=工厂可用
        data = com_read_sheet(wb, "合同待交付", start_row=2)
        for row in data:
            if not row or not row[0]:
                continue
            key = safe_str(row[0])
            qty = (to_num(row[19]) if len(row) > 19 else 0) + \
                  (to_num(row[25]) if len(row) > 25 else 0)
            inv_map[key]["待交付合同"] += qty

        # 4. 采购单待下单: col1=key, col15=采购量
        data = com_read_sheet(wb, "采购单待下单", start_row=2)
        for row in data:
            if not row or not row[0]:
                continue
            key = safe_str(row[0])
            qty = to_num(row[14]) if len(row) > 14 else 0
            inv_map[key]["国内待下合同"] += qty

        # 5. 海外仓调拨: 借入key=col1, 借出key=col3+col7+col10+col9, qty=col15
        data = com_read_sheet(wb, "海外仓调拨", start_row=3)
        for row in data:
            if not row:
                continue
            key_in = safe_str(row[0]) if len(row) > 0 else ""
            out_sku = safe_str(row[2]) if len(row) > 2 else ""
            out_shop = safe_str(row[6]) if len(row) > 6 else ""
            out_person = safe_str(row[8]) if len(row) > 8 else ""
            out_mode = safe_str(row[9]) if len(row) > 9 else ""
            qty = to_num(row[14]) if len(row) > 14 else 0

            if key_in:
                inv_map[key_in]["借入"] += qty
            if out_sku and out_shop and out_mode and out_person:
                key_out = "%s&%s&%s&%s" % (out_sku, out_shop, out_mode, out_person)
                inv_map[key_out]["借出"] += qty

    return dict(inv_map)


def get_inv_total(inv_map, key):
    inv = inv_map.get(key)
    if not inv:
        return 0, dict(zip(INV_COMPONENTS, [0]*7))
    total = (inv["海外在途"] + inv["海外在仓"] + inv["借入"] - inv["借出"]
             + inv["待交付合同"] + inv["国内在仓"] + inv["国内待下合同"])
    return total, inv


def build_inv_key(hun, shop, mode, person, shop_to_main, person_to_norm):
    main_shop = normalize_shop(shop, shop_to_main)
    norm_person = normalize_person(person, person_to_norm)
    return "%s&%s&%s&%s" % (hun, main_shop, mode, norm_person)


# ==================== Sheet2 捆绑拆分生成 ====================

def build_sheet2_bundle(records, months, src_to_map, zhu_to_sets, zhu_to_xiangsi,
                        zhu_to_order, inv_map, shop_to_main, person_to_norm):
    """捆绑拆分：套装需求+库存拆解，从老到新分配。

    核心公式:
      V(套装总需求) = SUMIFS(P, 店铺, 发货模式, 相似SKU)
      W(需求占比) = IF(V=0, 1, P/V)
      AB(库存套数) = ROUNDDOWN(总库存 / W)
      X(拆解1) = IF(AND(I="",J=""), P, MIN(ROUNDDOWN(MINIFS(AB)*W), ROUNDDOWN(V*W)))
      Y(拆解2) = IF(X=P, 0, IF(J="", P-X, MIN(ROUNDDOWN(MINIFS(AC)*W), ROUNDDOWN(V*W)-X)))
      Z(拆解3) = IF(J="", 0, P-SUM(X:Y))
      AE(总库存) = 海外在途+海外在仓+借入-借出+待交付合同+国内在仓+国内待下合同

    MINIFS key: (店铺, 相似SKU) — NOT 发货模式!
    """

    # Step 1: Group records by (拆解需求主SKU, 主店铺, 发货模式, 归一运营)
    groups = defaultdict(lambda: {"months": defaultdict(float), "text": {}, "count": 0, "hun": "",
                                   "shop": "", "mode": "", "person": ""})
    for rec in records:
        m = src_to_map.get(rec["SKU"])
        zhu = m["zhu"] if m else rec["SKU"]
        hun = m["hun"] if m else rec["SKU"]
        main_shop = normalize_shop(rec["店铺"], shop_to_main)
        norm_person = normalize_person(rec["listing负责人"], person_to_norm)
        key = (zhu, main_shop, rec["发货模式"], norm_person)
        g = groups[key]
        g["count"] += 1
        g["hun"] = hun
        g["shop"] = main_shop
        g["mode"] = rec["发货模式"]
        g["person"] = norm_person
        for mn in months:
            g["months"][mn] += rec[mn]
        for f in ["品名描述", "MSKU", "型号SPU", "一级分类", "产品定级",
                   "渠道", "国家", "运营组"]:
            if f not in g["text"] or not g["text"][f]:
                g["text"][f] = rec[f]

    # Step 2: Build xiangsi lookup
    def get_xiangsi(zhu):
        if zhu in zhu_to_xiangsi:
            return zhu_to_xiangsi[zhu]
        for rec in records:
            m = src_to_map.get(rec["SKU"])
            if m and m["zhu"] == zhu:
                return m.get("xiangsi", "")
        return ""

    # Step 3: Calculate 套装总需求 (V) — grouped by (shop, mode, xiangsi)
    v_map = defaultdict(float)
    for (zhu, shop, mode, person), g in groups.items():
        xiangsi = get_xiangsi(zhu)
        total = sum(g["months"].values())
        v_map[(shop, mode, xiangsi)] += total

    # Step 4: Build pass1 rows (calculate AB values)
    # Only include rows where 拆解需求主SKU contains "新" (整套替代)
    pass1_rows = []

    for (zhu, shop, mode, person), g in sorted(groups.items()):
        if "新" not in zhu:
            continue

        xiangsi = get_xiangsi(zhu)
        order_sku = zhu_to_order.get(zhu, "")
        is_zhengtao = "新" in zhu

        total_demand = sum(g["months"].values())
        v_val = v_map.get((shop, mode, xiangsi), 0)
        w_val = 1.0 if v_val == 0 else total_demand / v_val

        sets = zhu_to_sets.get(zhu, [{"hun": g["hun"], "is_order": True, "sources": []}])
        sets_sorted = sorted(sets, key=lambda s: (s["is_order"], s["hun"]))
        sets_3 = sets_sorted[:3]
        while len(sets_3) < 3:
            sets_3.append({"hun": "", "is_order": False, "sources": []})

        # Build inventory keys using normalized values
        set_keys = []
        set_skus = []
        for s in sets_3:
            if s["hun"]:
                full_key = build_inv_key(s["hun"], shop, mode, person,
                                        shop_to_main, person_to_norm)
                set_keys.append(full_key)
                set_skus.append(s["hun"])
            else:
                set_keys.append("")
                set_skus.append("")

        # Get inventory per set
        set_inv = []
        set_total_inv = []
        for i, sk in enumerate(set_keys):
            if sk:
                total, comps = get_inv_total(inv_map, sk)
                set_inv.append(comps)
                set_total_inv.append(total)
            else:
                set_inv.append(dict(zip(INV_COMPONENTS, [0]*7)))
                set_total_inv.append(0)

        # 库存套数 (AB/AC/AD) = ROUNDDOWN(总库存 / 需求占比)
        set_ab = []
        for i in range(3):
            if set_skus[i] and w_val > 0:
                ab = math.floor(set_total_inv[i] / w_val)
            else:
                ab = 0
            set_ab.append(ab)

        row = {
            "拆解需求主SKU+店铺+发货模式+运营": "%s&%s&%s&%s" % (zhu, shop, mode, person),
            "拆解需求主SKU": zhu,
            "相似SKU": xiangsi,
            "下单SKU": order_sku,
            "set1_full": set_keys[0], "set2_full": set_keys[1], "set3_full": set_keys[2],
            "set1_sku": set_skus[0], "set2_sku": set_skus[1], "set3_sku": set_skus[2],
            "运营": person,
            "需求总计": total_demand,
            "店铺": shop,
            "发货模式": mode,
            "套装总需求": v_val,
            "需求占比": w_val,
            "set_ab": set_ab,
            "set_total_inv": set_total_inv,
            "set_inv": set_inv,
            "is_zhengtao": is_zhengtao,
        }
        pass1_rows.append((row, (shop, xiangsi)))

    # Step 5: Compute MINIFS per (shop, xiangsi) — NO 发货模式!
    minifs_map = defaultdict(lambda: [float('inf'), float('inf'), float('inf')])
    for row, (shop, xiangsi) in pass1_rows:
        key = (shop, xiangsi)
        for i in range(3):
            if row["set_ab"][i] < minifs_map[key][i]:
                minifs_map[key][i] = row["set_ab"][i]

    for key in minifs_map:
        for i in range(3):
            if minifs_map[key][i] == float('inf'):
                minifs_map[key][i] = 0

    # Step 6: Calculate demand split (X, Y, Z)
    sheet2_rows = []
    for row, (shop, xiangsi) in pass1_rows:
        P = row["需求总计"]
        V = row["套装总需求"]
        W = row["需求占比"]
        I_empty = (row["set2_sku"] == "")
        J_empty = (row["set3_sku"] == "")

        minifs = minifs_map[(shop, xiangsi)]

        if I_empty and J_empty:
            X = P
        else:
            val1 = math.floor(minifs[0] * W) if W > 0 else 0
            val2 = math.floor(V * W)
            X = min(val1, val2)

        if X == P:
            Y = 0
        elif J_empty:
            Y = P - X
        else:
            val1 = math.floor(minifs[1] * W) if W > 0 else 0
            val2 = math.floor(V * W) - X
            Y = min(val1, val2)

        if J_empty:
            Z = 0
        else:
            Z = P - X - Y

        X = max(0, X)
        Y = max(0, Y)
        Z = max(0, Z)

        row["X"] = X
        row["Y"] = Y
        row["Z"] = Z
        row["汇总"] = X + Y + Z
        row["结果"] = "Y" if (X + Y + Z) == P else "N"

        row["冗余1"] = max(0, row["set_total_inv"][0] - X)
        row["冗余2"] = max(0, row["set_total_inv"][1] - Y)
        row["冗余3"] = max(0, row["set_total_inv"][2] - Z)

        sheet2_rows.append(row)

    return sheet2_rows


# ==================== Sheet1 生成（销售需求汇总） ====================

def detail_to_records(detail, month_row, fstart):
    """Convert in-memory detail list to records format for build_sheet1.
    detail row: [SKU, 品名描述, MSKU, 型号SPU, 一级分类, 产品定级, 渠道, 国家,
                 店铺, 发货模式, 运营组, listing负责人, 销售状态, 汇总, 安全库存,
                 month1..month12, 来源, 父捆绑SKU, 关联数量, 产品名称, 产品款式, 警告]
    """
    month_cols = []
    for i in range(fstart, fstart + 12):
        if i < len(month_row) and month_row[i]:
            month_cols.append(str(month_row[i]))
        else:
            month_cols.append("月%d" % (i - fstart + 1))
    months = month_cols

    records = []
    for row in detail:
        rec = {
            "SKU": safe_str(row[0]),
            "品名描述": safe_str(row[1]),
            "MSKU": safe_str(row[2]),
            "型号SPU": safe_str(row[3]),
            "一级分类": safe_str(row[4]),
            "产品定级": safe_str(row[5]),
            "渠道": safe_str(row[6]),
            "国家": safe_str(row[7]),
            "店铺": safe_str(row[8]),
            "发货模式": safe_str(row[9]),
            "运营组": safe_str(row[10]),
            "listing负责人": safe_str(row[11]),
        }
        for i, mn in enumerate(months):
            rec[mn] = to_num(row[15 + i])
        records.append(rec)

    return records, months


def build_sheet1(records, months, src_to_map, zhu_to_sets, zhu_to_order,
                 shop_to_main, person_to_norm, sheet2_data=None):
    """Sheet1: 销售需求汇总（含按月从老到新分配）。

    对于整套替代行（在sheet2_data中）：拆分为老套装/新套装多行，
    按月从老到新分配需求（老套装先满足前面月份，不够的递减到新套装）。
    对于非整套行：保持单行汇总。
    """
    s2_lookup = {}
    if sheet2_data:
        for s2row in sheet2_data:
            key = (s2row["拆解需求主SKU"], s2row["店铺"], s2row["发货模式"], s2row["运营"])
            s2_lookup[key] = s2row

    groups = defaultdict(lambda: {"months": defaultdict(float), "text": {}, "count": 0, "hun": ""})
    for rec in records:
        m = src_to_map.get(rec["SKU"])
        zhu = m["zhu"] if m else rec["SKU"]
        hun = m["hun"] if m else rec["SKU"]
        main_shop = normalize_shop(rec["店铺"], shop_to_main)
        norm_person = normalize_person(rec["listing负责人"], person_to_norm)
        key = (zhu, main_shop, rec["发货模式"], norm_person)
        g = groups[key]
        g["count"] += 1
        g["hun"] = hun
        for mn in months:
            g["months"][mn] += rec[mn]
        for f in ["品名描述", "MSKU", "型号SPU", "一级分类", "产品定级",
                   "渠道", "国家", "运营组"]:
            if f not in g["text"] or not g["text"][f]:
                g["text"][f] = rec[f]

    output = []
    for (zhu, shop, mode, person), g in sorted(groups.items()):
        order = zhu_to_order.get(zhu, "")
        is_zhengtao = "新" in zhu
        total_demand = sum(g["months"].values())

        text_fields = {
            "品名描述": g["text"].get("品名描述", ""),
            "MSKU": g["text"].get("MSKU", ""),
            "型号/SPU": g["text"].get("型号SPU", ""),
            "一级分类": g["text"].get("一级分类", ""),
            "产品定级": g["text"].get("产品定级", ""),
            "渠道": g["text"].get("渠道", ""),
            "国家": g["text"].get("国家", ""),
            "运营组": g["text"].get("运营组", ""),
        }

        s2key = (zhu, shop, mode, person)

        if is_zhengtao and s2key in s2_lookup:
            s2 = s2_lookup[s2key]
            X = s2["X"]
            Y = s2["Y"]
            Z_val = s2["Z"]

            sets_sorted = sorted(
                zhu_to_sets.get(zhu, [{"hun": g["hun"], "is_order": True, "sources": []}]),
                key=lambda s: (s["is_order"], s["hun"])
            )
            set_huns = [s["hun"] for s in sets_sorted[:3]]
            while len(set_huns) < 3:
                set_huns.append("")

            # 按月从老到新分配
            remaining = [X, Y, Z_val]
            set_months = [{}, {}, {}]

            for mn in months:
                month_demand = round(g["months"][mn])
                for si in range(3):
                    if month_demand <= 0:
                        set_months[si][mn] = 0
                    elif remaining[si] >= month_demand:
                        set_months[si][mn] = month_demand
                        remaining[si] -= month_demand
                        month_demand = 0
                    elif remaining[si] > 0:
                        set_months[si][mn] = remaining[si]
                        month_demand -= remaining[si]
                        remaining[si] = 0
                    else:
                        set_months[si][mn] = 0
                if month_demand > 0:
                    set_months[2][mn] = set_months[2].get(mn, 0) + month_demand

            allocs = [X, Y, Z_val]
            for si in range(3):
                if allocs[si] <= 0 and not any(v > 0 for v in set_months[si].values()):
                    continue

                row = {
                    "拆解需求主SKU+店铺+发货模式+运营": "%s&%s&%s&%s" % (zhu, shop, mode, person),
                    "拆解需求主SKU": zhu,
                    "混用主SKU+店铺+发货模式+运营": "%s&%s&%s&%s" % (set_huns[si], shop, mode, person) if set_huns[si] else "",
                    "混用主SKU": set_huns[si],
                    "理论-下单SKU（不考虑补配套）": order,
                    "拆解需求数": allocs[si],
                    "是否整套": "是",
                    **text_fields,
                    "店铺": shop,
                    "发货模式": mode,
                    "listing负责人": person,
                    "需求总计": total_demand,
                    "安全库存": 0,
                }
                for mn in months:
                    row[mn] = set_months[si][mn]
                output.append(row)
        else:
            row = {
                "拆解需求主SKU+店铺+发货模式+运营": "%s&%s&%s&%s" % (zhu, shop, mode, person),
                "拆解需求主SKU": zhu,
                "混用主SKU+店铺+发货模式+运营": "%s&%s&%s&%s" % (g["hun"], shop, mode, person),
                "混用主SKU": g["hun"],
                "理论-下单SKU（不考虑补配套）": order,
                "拆解需求数": total_demand,
                "是否整套": "是" if is_zhengtao else "否",
                **text_fields,
                "店铺": shop,
                "发货模式": mode,
                "listing负责人": person,
                "需求总计": total_demand,
                "安全库存": 0,
            }
            for mn in months:
                row[mn] = round(g["months"][mn])
            output.append(row)
    return output


# ==================== 写出工作簿 ====================

def write_output(out_path, detail, op_hdr, month_row, fstart,
                 missing_bundle, missing_prod, formula_bad,
                 sheet1_data=None, months=None, sheet2_data=None):
    """写出 xlsx：
       有 sheet1: 3 sheets → 销售需求汇总 + 未拆解整套版销售需求 + 异常提醒
       无 sheet1: 2 sheets → 未拆解整套版销售需求 + 异常提醒
       （捆绑拆分为内部计算，不单独输出Sheet）
    """
    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF")
    warn_fill = PatternFill("solid", fgColor="FFC7CE")
    warn_font = Font(color="9C0006")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    month_align = Alignment(horizontal="center")

    # === Sheet1: 销售需求汇总（可选） ===
    if sheet1_data is not None and months:
        ws1 = wb.active
        ws1.title = "销售需求汇总"
        h1 = ["拆解需求主SKU+店铺+发货模式+运营", "拆解需求主SKU",
              "混用主SKU+店铺+发货模式+运营", "混用主SKU", "理论-下单SKU（不考虑补配套）",
              "拆解需求数", "是否整套", "品名描述", "MSKU", "型号/SPU", "一级分类",
              "产品定级", "渠道", "国家", "店铺", "发货模式", "运营组", "listing负责人",
              "需求总计", "安全库存"] + months
        for c, h in enumerate(h1, 1):
            cell_ = ws1.cell(1, c, h)
            cell_.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=10)
            cell_.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            cell_.alignment = center
            cell_.border = border
        s1_font = Font(name="Microsoft YaHei", size=10)
        s1_center = Alignment(horizontal="center", vertical="center")
        s1_left = Alignment(vertical="center", wrap_text=True)
        s1_right = Alignment(horizontal="right", vertical="center")
        zebra = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")
        zhengtao_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        old_set_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        new_set_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        month_set = set(months)
        num_cols_1 = {6, 19, 20} | set(range(21, 21 + len(months)))
        center_cols_1 = {6, 7, 10, 11, 13, 14, 15, 16, 19, 20} | set(range(21, 21 + len(months)))
        col_w_1 = {1: 38, 2: 26, 3: 38, 4: 24, 5: 28, 6: 10, 7: 10, 8: 36,
                   9: 24, 10: 16, 11: 12, 12: 30, 13: 10, 14: 8, 15: 12,
                   16: 10, 17: 16, 18: 14, 19: 12, 20: 12}
        for c, w in col_w_1.items():
            ws1.column_dimensions[get_column_letter(c)].width = w
        for i in range(len(col_w_1) + 1, len(h1) + 1):
            ws1.column_dimensions[get_column_letter(i)].width = 12
        for r, row in enumerate(sheet1_data, 2):
            is_zhengtao = row.get("是否整套", "") == "是"
            hun_val = str(row.get("混用主SKU", ""))
            is_old_set = is_zhengtao and "-V1" not in hun_val and hun_val != ""
            is_new_set = is_zhengtao and "-V1" in hun_val
            for c, h in enumerate(h1, 1):
                val = row.get(h, "")
                cell_ = ws1.cell(r, c, val)
                cell_.font = s1_font
                cell_.border = border
                if c in num_cols_1:
                    cell_.number_format = '#,##0'
                    cell_.alignment = s1_right
                elif c in center_cols_1:
                    cell_.alignment = s1_center
                else:
                    cell_.alignment = s1_left
                if r % 2 == 0 and not is_zhengtao:
                    cell_.fill = zebra
                if is_old_set:
                    cell_.fill = old_set_fill
                elif is_new_set:
                    cell_.fill = new_set_fill
        ws1.freeze_panes = "A2"
        ws1.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(h1)), len(sheet1_data) + 1)
        ws1.row_dimensions[1].height = 36
        ws1 = wb.create_sheet()
    else:
        ws1 = wb.active

    # === Sheet2: 捆绑拆分（可选） ===
    if sheet2_data is not None and months:
        ws2 = ws1
        ws2.title = "捆绑拆分"
        h2 = [
            "拆解需求主SKU+店铺+发货模式+运营", "拆解需求主SKU", "相似SKU", "下单SKU",
            "拆解1-包含SKU&key", "拆解2-包含SKU&key", "拆解3-包含SKU&key",
            "拆解1-混用主SKU", "拆解2-混用主SKU", "拆解3-混用主SKU",
            "运营", "需求总计-待拆分", "店铺", "发货模式",
            "套装总需求", "需求占比",
            "拆解1需求", "拆解2需求", "拆解3需求",
            "拆解1总库存", "拆解2总库存", "拆解3总库存",
            "汇总", "结果",
            "拆解1冗余", "拆解2冗余", "拆解3冗余",
            "拆解1海外在途", "拆解1海外在仓", "拆解1借入", "拆解1借出",
            "拆解1待交付合同", "拆解1国内在仓", "拆解1国内待下合同",
            "拆解2海外在途", "拆解2海外在仓", "拆解2借入", "拆解2借出",
            "拆解2待交付合同", "拆解2国内在仓", "拆解2国内待下合同",
            "拆解3海外在途", "拆解3海外在仓", "拆解3借入", "拆解3借出",
            "拆解3待交付合同", "拆解3国内在仓", "拆解3国内待下合同",
        ]
        s2_font = Font(name="Microsoft YaHei", size=10)
        s2_center = Alignment(horizontal="center", vertical="center")
        s2_left = Alignment(vertical="center", wrap_text=True)
        s2_right = Alignment(horizontal="right", vertical="center")
        s2_zebra = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")
        w2 = {1: 35, 2: 25, 3: 22, 4: 25, 5: 35, 6: 35, 7: 35, 8: 22, 9: 22, 10: 22,
              11: 12, 12: 16, 13: 12, 14: 10, 15: 14, 16: 12,
              17: 12, 18: 12, 19: 12, 20: 14, 21: 14, 22: 14, 23: 12, 24: 8,
              25: 12, 26: 12, 27: 12}
        for c, h in enumerate(h2, 1):
            cell_ = ws2.cell(1, c, h)
            cell_.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=10)
            cell_.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            cell_.alignment = center
            cell_.border = border
        for c, w in w2.items():
            ws2.column_dimensions[get_column_letter(c)].width = w
        for i in range(len(w2) + 1, len(h2) + 1):
            ws2.column_dimensions[get_column_letter(i)].width = 12
        num_cols_2 = set(range(12, 28)) | set(range(29, len(h2) + 1))
        for r, row_data in enumerate(sheet2_data, 2):
            is_zebra = (r % 2 == 0)
            for c, h in enumerate(h2, 1):
                if h == "需求占比":
                    val = "%.4f" % row_data["需求占比"] if isinstance(row_data.get("需求占比"), float) else row_data.get("需求占比", "")
                elif h.startswith("拆解") and len(h) > 2 and h[2].isdigit() and ("需求" in h or "总库存" in h or "冗余" in h):
                    idx = int(h[2]) - 1
                    if "冗余" in h:
                        val = row_data.get("冗余%d" % (idx + 1), "")
                    elif "总库存" in h:
                        val = row_data.get("set_total_inv", [0, 0, 0])[idx] if idx < len(row_data.get("set_total_inv", [])) else ""
                    elif "需求" in h:
                        val = row_data.get(["X", "Y", "Z"][idx], "")
                    else:
                        val = row_data.get(h, "")
                elif h.startswith("拆解") and len(h) > 2 and h[2].isdigit() and any(x in h for x in INV_COMPONENTS):
                    set_idx = int(h[2]) - 1
                    comp_name = h[4:]
                    if set_idx < len(row_data.get("set_inv", [])):
                        val = row_data["set_inv"][set_idx].get(comp_name, 0)
                    else:
                        val = 0
                else:
                    val = row_data.get(h, "")
                cell_ = ws2.cell(r, c, val)
                cell_.font = s2_font
                cell_.border = border
                if c in num_cols_2:
                    cell_.number_format = '#,##0'
                    cell_.alignment = s2_right
                elif c in {3, 4, 11, 14, 24}:
                    cell_.alignment = s2_center
                else:
                    cell_.alignment = s2_left
                if is_zebra:
                    cell_.fill = s2_zebra
        ws2.freeze_panes = "C2"
        ws2.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(h2)), len(sheet2_data) + 1)
        ws2.row_dimensions[1].height = 30
        ws2 = wb.create_sheet()
    else:
        ws2 = ws1

    # === Sheet2: 未拆解整套版销售需求 ===
    ws = ws2
    ws.title = "未拆解整套版销售需求"
    HDR1 = list(op_hdr[:fstart + 12]) + HELPER
    HDR2 = list(month_row[:fstart + 12]) + [""] * len(HELPER)
    ws.append(HDR1)
    ws.append(HDR2)
    ncols = len(HDR1)
    s2_hdr_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    s2_hdr_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=10)
    s2_data_font = Font(name="Microsoft YaHei", size=10)
    s2_zebra = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")
    s2_center = Alignment(horizontal="center", vertical="center")
    s2_right = Alignment(horizontal="right", vertical="center")
    for r in (1, 2):
        for c in range(1, ncols + 1):
            cell_ = ws.cell(row=r, column=c)
            cell_.fill = s2_hdr_fill
            cell_.font = s2_hdr_font
            cell_.alignment = center
            cell_.border = border
    ws.merge_cells(start_row=1, start_column=fstart + 1, end_row=1, end_column=fstart + 12)
    for d in detail:
        ws.append(d)

    warn_idx = ncols - 1
    month_cols = set(range(fstart + 1, fstart + 13))
    num_cols_s2 = month_cols | {fstart}
    for row_cells in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=ncols):
        is_warn = False
        for c in row_cells:
            c.border = border
            c.font = s2_data_font
        for ci in range(fstart, fstart + 12):
            row_cells[ci].alignment = s2_center
            if ci in num_cols_s2 and isinstance(row_cells[ci].value, (int, float)):
                row_cells[ci].number_format = '#,##0'
        wcell = row_cells[warn_idx]
        if wcell.value:
            is_warn = True
            for c in row_cells:
                c.fill = warn_fill
            wcell.font = warn_font
        elif row_cells[0].row % 2 == 0:
            for c in row_cells:
                c.fill = s2_zebra

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 32
    for c in range(3, fstart + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12
    for i in range(12):
        ws.column_dimensions[get_column_letter(fstart + 1 + i)].width = 10
    for nm, w in [("来源", 16), ("父捆绑SKU", 22), ("关联数量", 10),
                   ("产品名称", 28), ("产品款式", 14), ("警告", 36)]:
        ws.column_dimensions[get_column_letter(HDR1.index(nm) + 1)].width = w
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = "A2:%s%d" % (get_column_letter(ncols), ws.max_row)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20

    # === Sheet3: 异常提醒 ===
    ws3 = wb.create_sheet("异常提醒")
    title1_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
    title2_fill = PatternFill(start_color="D35400", end_color="D35400", fill_type="solid")
    hdr3_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    hdr3_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=10)
    data3_font = Font(name="Microsoft YaHei", size=10)
    zebra3 = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")
    s3_center = Alignment(horizontal="center", vertical="center")

    row1 = 1
    ws3.cell(row1, 1, "【异常类型一】后缀为 -KB 但在「捆绑-BOM-发货模式」中找不到对应捆绑关系（共 %d 个）" % len(missing_bundle))
    ws3.merge_cells(start_row=row1, start_column=1, end_row=row1, end_column=6)
    for c in range(1, 7):
        cc = ws3.cell(row1, c)
        cc.fill = title1_fill
        cc.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
        cc.alignment = s3_center
        cc.border = border

    row2 = 2
    hdrs1 = ["销售预测SKU", "品名描述", "渠道", "国家", "店铺", "发货模式"]
    for ci, htext in enumerate(hdrs1, 1):
        cc = ws3.cell(row2, ci, htext)
        cc.fill = hdr3_fill
        cc.font = hdr3_font
        cc.alignment = s3_center
        cc.border = border

    for ri, m in enumerate(missing_bundle):
        r = row2 + 1 + ri
        for ci, val in enumerate(m, 1):
            cc = ws3.cell(r, ci, val)
            cc.font = data3_font
            cc.border = border
            if r % 2 == 0:
                cc.fill = zebra3

    blank_row = row2 + 1 + len(missing_bundle)
    ws3.cell(blank_row, 1, "")
    title2_row = blank_row + 1
    ws3.cell(title2_row, 1, "【异常类型二】单品SKU 未匹配到产品资料(J-L)（共 %d 个）" % len(missing_prod))
    ws3.merge_cells(start_row=title2_row, start_column=1, end_row=title2_row, end_column=6)
    for c in range(1, 7):
        cc = ws3.cell(title2_row, c)
        cc.fill = title2_fill
        cc.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
        cc.alignment = s3_center
        cc.border = border

    hdr2_row = title2_row + 1
    hdrs2 = ["单品SKU", "来源", "父捆绑SKU"]
    for ci, htext in enumerate(hdrs2, 1):
        cc = ws3.cell(hdr2_row, ci, htext)
        cc.fill = hdr3_fill
        cc.font = hdr3_font
        cc.alignment = s3_center
        cc.border = border

    for ri, m in enumerate(missing_prod):
        r = hdr2_row + 1 + ri
        for ci, val in enumerate(m, 1):
            cc = ws3.cell(r, ci, val)
            cc.font = data3_font
            cc.border = border
            if r % 2 == 0:
                cc.fill = zebra3

    ws3.column_dimensions["A"].width = 26
    ws3.column_dimensions["B"].width = 34
    for col in ["C", "D", "E", "F"]:
        ws3.column_dimensions[col].width = 14
    ws3.row_dimensions[row1].height = 28
    ws3.row_dimensions[title2_row].height = 28

    wb.save(out_path)
    return wb.sheetnames


# ==================== 运行摘要 ====================

def print_summary(out_path, detail, n_bundle, n_single, n_missbundle,
                  missing_bundle, missing_prod, formula_bad, rounded_cells, sheetnames,
                  sheet1_count=0, sheet2_count=0):
    print("=" * 56)
    print("运行摘要")
    print("=" * 56)
    print("  输出文件 :", out_path)
    print("  明细总行数 :", len(detail))
    print("    - 捆绑拆解组件 :", n_bundle)
    print("    - 原单品(含缺失关系) :", n_single)
    print("  缺失捆绑关系(保留原行) :", n_missbundle, [m[0] for m in missing_bundle])
    print("  未匹配产品资料 :", len(missing_prod), [m[0] for m in missing_prod][:10])
    print("  汇总公式自检 :",
          "全部通过" if not formula_bad else "存在 %d 行不符: %s" % (len(formula_bad), formula_bad[:10]))
    print("  非整数预测取整单元格 :", rounded_cells)
    if sheet1_count:
        print("  Sheet1 销售需求汇总 :", sheet1_count, "行")
    if sheet2_count:
        print("  Sheet2 捆绑拆分 :", sheet2_count, "行")
    print("  工作表 :", sheetnames)
    print("=" * 56)
    if formula_bad:
        print("警告：发现汇总公式不符的行，请检查源数据或脚本逻辑！")


# ==================== 主流程 ====================

def main():
    ap = argparse.ArgumentParser(description="销售预测 捆绑SKU拆解 (v5 集成版)")
    ap.add_argument("src", help="源 xlsx（含运营表与捆绑表，可能 DRM 加密）")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径")
    ap.add_argument("--op-sheet", default=None, help="运营表 sheet 名")
    ap.add_argument("--bom-sheet", default=None, help="捆绑表 sheet 名")
    ap.add_argument("--bom-csv", default=None, help="同款套装库存 CSV")
    ap.add_argument("--product-csv", default=None, help="product 单品产品资料 CSV")
    ap.add_argument("--enc", default=None, help="CSV 编码")
    ap.add_argument("--feishu-url", default=FEISHU_URL_DEFAULT,
                    help="飞书整套替代关系表URL（默认内置URL）")
    ap.add_argument("--inv-file", default=None,
                    help="库存文件 xlsx（含海外仓库存等sheet），用于归一化映射+Sheet1生成")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        print("源文件不存在:", src)
        sys.exit(1)

    if bool(args.bom_csv) ^ bool(args.product_csv):
        print("错误：--bom-csv 与 --product-csv 必须同时提供")
        sys.exit(2)

    # ---- 读取运营表/捆绑表 ----
    names, op_name, bom_name, op, bom = read_workbook(
        src, args.op_sheet, args.bom_sheet, read_bom=(args.bom_csv is None))
    if args.bom_csv and args.product_csv:
        print("运营表:", op_name, "| 捆绑关系/产品资料: CSV 模式")
    else:
        print("运营表:", op_name, "| 捆绑表:", bom_name)

    # ---- 运营表 header / month / 数据 ----
    hidx = find_header_row(op, "SKU", col=0)
    if hidx == 0 and not str(cell(op[0], 0)).strip() == "SKU":
        print("警告：未在运营表定位到表头行，已回退到第 1 行")
    op_hdr = op[hidx]
    month_row = op[hidx + 1] if hidx + 1 < len(op) else []
    data = [r for r in op[hidx + 2:] if r and str(cell(r, 0)).strip()]

    # ---- 列索引 ----
    fstart = col_index(op_hdr, "销量预测")
    if fstart is None:
        fstart = DEFAULT_FSTART
    ss_col = fstart - 1
    tot_col = fstart - 2
    fc = list(range(fstart, fstart + 12))
    cc = {
        "fstart": fstart, "ss_col": ss_col, "tot_col": tot_col, "fc": fc,
        "sku_c": col_index(op_hdr, "SKU", exclude="产品") or 0,
        "shop_c": col_index(op_hdr, "店铺"),
        "opg_c": col_index(op_hdr, "运营"),
        "ship_c": col_index(op_hdr, "发货模式"),
        "cat_c": col_index(op_hdr, "一级分类"),
        "grade_c": col_index(op_hdr, "产品定级"),
        "chan_c": col_index(op_hdr, "渠道"),
        "coun_c": col_index(op_hdr, "国家"),
        "stat_c": col_index(op_hdr, "销售状态"),
        "list_c": col_index(op_hdr, "listing") or col_index(op_hdr, "负责人"),
        "msku_c": col_index(op_hdr, "MSKU"),
        "pname_c": col_index(op_hdr, "品名描述"),
        "spu_c": col_index(op_hdr, "型号SPU"),
    }

    # ---- 属性查找表 ----
    op_attr = {}
    for row in data:
        sku = str(cell(row, cc["sku_c"])).strip()
        shop = str(cell(row, cc["shop_c"])).strip()
        opg = str(cell(row, cc["opg_c"])).strip()
        ship = str(cell(row, cc["ship_c"])).strip()
        vals = (str(cell(row, cc["cat_c"])).strip(), str(cell(row, cc["grade_c"])).strip(),
                str(cell(row, cc["chan_c"])).strip(), str(cell(row, cc["coun_c"])).strip(),
                str(cell(row, cc["stat_c"])).strip())
        key = (sku.upper(), shop, opg, ship)
        if key not in op_attr:
            op_attr[key] = vals

    # ---- 捆绑表 / 产品资料 ----
    if args.bom_csv and args.product_csv:
        prod_master, bom_map = build_bom_product(args.bom_csv, args.product_csv, args.enc)
        print("（CSV 模式）prod_master:", len(prod_master), "| bom_map:", len(bom_map))
    else:
        bidx = find_header_row(bom, "包含产品SKU")
        if bidx is None:
            bidx = find_header_row(bom, "产品SKU")
        if bidx is None:
            bidx = 0
        bh = bom[bidx]
        b_sku = col_index(bh, "SKU", exclude="产品")
        b_comp = col_index(bh, "包含产品SKU")
        b_qty = col_index(bh, "关联数量")
        b_prod = col_index(bh, "产品SKU", exclude="包含")
        b_pname = col_index(bh, "产品名称")
        b_pstyle = col_index(bh, "产品款式")

        prod_master = {}
        bom_map = defaultdict(list)
        for row in bom[bidx + 1:]:
            j = str(cell(row, b_prod)).strip()
            if j:
                prod_master[j] = (str(cell(row, b_pname)).strip(), str(cell(row, b_pstyle)).strip())
            a = str(cell(row, b_sku)).strip()
            c = str(cell(row, b_comp)).strip()
            e = str(cell(row, b_qty)).strip()
            if a and c:
                try:
                    q = float(e)
                except Exception:
                    q = 1.0
                bom_map[a.upper()].append((c, q))

    # ---- 构建明细 ----
    detail, missing_bundle, missing_prod, n_bundle, n_single, n_missbundle, rounded_cells = \
        process_rows(data, bom_map, prod_master, op_attr, cc)

    # ---- 自检 ----
    formula_bad = check_formulas(detail)

    # ---- Sheet1+Sheet2 生成（可选，需要库存文件） ----
    sheet1_data = None
    sheet2_data = None
    months = None
    if args.inv_file:
        inv_path = os.path.abspath(args.inv_file)
        if not os.path.exists(inv_path):
            print("警告：库存文件不存在:", inv_path, "→ 跳过Sheet1/Sheet2生成")
        else:
            try:
                print("[Sheet1] 读取飞书映射表...")
                src_to_map, zhu_to_sets, zhu_to_xiangsi, zhu_to_order = read_feishu_mapping(args.feishu_url)
                print("  源SKU映射:", len(src_to_map), "| 整套替代:", len(zhu_to_sets))

                print("[Sheet1] 读取归一化映射...")
                shop_to_main, person_to_norm = read_normalization(inv_path)
                print("  店铺→主店铺:", len(shop_to_main), "| 运营→归一:", len(person_to_norm))

                print("[Sheet1] 转换明细为records...")
                records, months = detail_to_records(detail, month_row, fstart)
                print("  records:", len(records), "| months:", len(months))

                print("[Sheet2] 读取库存文件(5个sheet)...")
                inv_map = read_inventory(inv_path)
                print("  库存key:", len(inv_map))

                print("[Sheet2] 生成捆绑拆分...")
                sheet2_data = build_sheet2_bundle(records, months, src_to_map, zhu_to_sets,
                                                  zhu_to_xiangsi, zhu_to_order, inv_map,
                                                  shop_to_main, person_to_norm)
                zhengtao = sum(1 for r in sheet2_data if r["is_zhengtao"])
                results_ok = sum(1 for r in sheet2_data if r["结果"] == "Y")
                matched = sum(1 for r in sheet2_data if any(r["set_total_inv"]))
                print("  Sheet2行数:", len(sheet2_data), "| 整套替代:", zhengtao,
                      "| 结果Y:", results_ok, "N:", len(sheet2_data)-results_ok,
                      "| 有库存:", matched)

                print("[Sheet1] 生成销售需求汇总（含按月从老到新分配）...")
                sheet1_data = build_sheet1(records, months, src_to_map, zhu_to_sets,
                                          zhu_to_order, shop_to_main, person_to_norm,
                                          sheet2_data)
                print("  Sheet1行数:", len(sheet1_data))
            except Exception as e:
                import traceback
                traceback.print_exc()
                print("警告：Sheet1/Sheet2生成失败:", e, "→ 仅输出Sheet3+Sheet4")
                sheet1_data = None
                sheet2_data = None
                months = None

    # ---- 写出 ----
    if sheet1_data is not None:
        out_path = args.out or os.path.join(
            os.path.dirname(src),
            "销售需求汇总_整套分配_%s.xlsx" % datetime.now().strftime("%Y%m%d"))
    else:
        out_path = args.out or os.path.join(
            os.path.dirname(src),
            "销售预测拆解汇总_明细_%s.xlsx" % datetime.now().strftime("%Y%m%d"))
    out_path = os.path.abspath(out_path)

    sheetnames = write_output(out_path, detail, op_hdr, month_row, fstart,
                              missing_bundle, missing_prod, formula_bad,
                              sheet1_data=sheet1_data, months=months,
                              sheet2_data=None)

    # ---- 摘要 ----
    print_summary(out_path, detail, n_bundle, n_single, n_missbundle,
                  missing_bundle, missing_prod, formula_bad, rounded_cells, sheetnames,
                  sheet1_count=len(sheet1_data) if sheet1_data else 0,
                  sheet2_count=0)


if __name__ == "__main__":
    main()
