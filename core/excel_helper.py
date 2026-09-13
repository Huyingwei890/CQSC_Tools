# -*- coding: utf-8 -*-
"""Excel 操作公共函数"""

import re
import datetime

# Excel 常量
XL_UP        = -4162
XL_MANUAL    = -4135
XL_AUTOMATIC = -4105


def col_letter(n):
    """1=A, 2=B, 27=AA ..."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def normalize_text(s):
    """文本规范化: 去所有空白、标点符号、转小写"""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u00a0", "").replace("\u3000", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^\w\u4e00-\u9fff]", "", s, flags=re.UNICODE)
    return s.lower()


def fuzzy_match(actual, expected_val=None, keywords=None):
    """
    模糊匹配:
      1. 规范化后完全相等 -> 命中
      2. 规范化后包含 expected -> 命中
      3. 规范化后包含所有 keywords -> 命中
    """
    if actual is None:
        return False
    actual_norm = normalize_text(actual)
    if not actual_norm:
        return False
    if expected_val is not None:
        exp_norm = normalize_text(expected_val)
        if actual_norm == exp_norm:
            return True
        if exp_norm and exp_norm in actual_norm:
            return True
    if keywords:
        ok = True
        for kw in keywords:
            if kw.lower() not in actual_norm:
                ok = False
                break
        if ok:
            return True
    return False


def read_cell_state(ws, addr):
    """返回 (Type, Value, NumberFormat)
       Type 取值: Formula / Value / Empty
    """
    rng = ws.Range(addr)
    fmt = rng.NumberFormat
    if rng.HasFormula:
        return ("Formula", rng.Formula, fmt)
    v = rng.Value
    if v is None:
        return ("Empty", "", fmt)
    return ("Value", v, fmt)


def _smart_parse(s):
    """把字符串智能解析回 Python 原生类型 (数字 / 日期 / 字符串)
    用于回滚时把 OldValue 字符串还原成正确的数据类型"""
    if not isinstance(s, str):
        return s
    s_stripped = s.strip()
    if not s_stripped:
        return s

    # 尝试 int (不含小数点/科学计数法)
    try:
        if "." not in s_stripped and "e" not in s_stripped.lower():
            return int(s_stripped)
    except (ValueError, TypeError):
        pass

    # 尝试 float
    try:
        return float(s_stripped)
    except (ValueError, TypeError):
        pass

    # 尝试 datetime (各种常见格式)
    # 先去时区后缀 (例如 2026-06-01 00:00:00+00:00)
    s_clean = s_stripped.split("+")[0].strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
                "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(s_clean, fmt)
        except ValueError:
            continue

    # 都不是, 原样返回字符串
    return s


def write_cell_by_type(ws, addr, vtype, val, fmt=None):
    """按类型写入单元格
    Value 类型: 字符串会智能转回数字/日期/字符串 (避免日期被当文本写入)
    """
    rng = ws.Range(addr)
    if vtype == "Formula":
        rng.Formula = val
    elif vtype == "Value":
        if val == "" or val is None:
            rng.ClearContents()
        else:
            rng.Value = _smart_parse(val)
    elif vtype == "Empty":
        rng.ClearContents()
    if fmt:
        try:
            rng.NumberFormat = fmt
        except Exception:
            pass


def find_sheet_by_name(wb, name):
    """按名查找 sheet，找不到返回 None"""
    for s in wb.Sheets:
        if s.Name == name:
            return s
    return None


def make_unique_sheet_name(wb, base):
    """生成不重名的 sheet 名（自动加 (2) (3) ...），截断到 31 字符"""
    existing = set()
    for s in wb.Sheets:
        existing.add(s.Name)
    name = base[:31]
    if name not in existing:
        return name
    i = 2
    while True:
        suffix = " (" + str(i) + ")"
        cand = base[:31 - len(suffix)] + suffix
        if cand not in existing:
            return cand
        i += 1


def wrap_formula_times(formula_str, factor):
    """把公式 =xxx 包成 =(xxx)*factor"""
    s = str(formula_str)
    if s.startswith("="):
        s = s[1:]
    return "=(" + s + ")*" + str(factor)


def to_2d_array(data, nr, nc):
    """把 win32 返回的各种形状统一成 2D 元组 (nr × nc)"""
    if data is None:
        return tuple(tuple(None for _ in range(nc)) for _ in range(nr))
    if nr == 1 and nc == 1:
        return ((data,),)
    if nr == 1:
        if isinstance(data, tuple) and len(data) > 0 and isinstance(data[0], tuple):
            return data
        return (tuple(data),)
    if nc == 1:
        if isinstance(data, tuple) and len(data) > 0 and isinstance(data[0], tuple):
            return data
        return tuple((v,) for v in data)
    return data


def calc_full(excel, logw=None):
    """强制全工作簿重算 (Rebuild = 重建依赖关系)"""
    try:
        excel.Calculation = XL_AUTOMATIC
        excel.CalculateFullRebuild()
        if logw:
            logw.log("  [重算] CalculateFullRebuild 完成")
    except Exception as e:
        try:
            excel.CalculateFull()
            if logw:
                logw.log("  [重算] CalculateFull 完成")
        except Exception as e2:
            if logw:
                logw.log("  [重算] 失败: " + str(e2))


class FastExcel:
    """临时关闭刷屏/计算/事件，结束自动还原"""
    def __init__(self, excel):
        self.excel = excel

    def __enter__(self):
        self.scr  = self.excel.ScreenUpdating
        self.calc = self.excel.Calculation
        self.evt  = self.excel.EnableEvents
        self.excel.ScreenUpdating = False
        self.excel.Calculation = XL_MANUAL
        self.excel.EnableEvents = False
        return self

    def __exit__(self, *a):
        try:
            self.excel.Calculation = self.calc
        except Exception:
            pass
        self.excel.ScreenUpdating = True
        self.excel.EnableEvents = self.evt