# -*- coding: utf-8 -*-
"""
[1] Tooling 批量修改
对 CQSC 文件中的 Tooling* 工作簿做批量修改:
  - W125 = =-S118
  - Y125 = =S64
  - U126 = 今天 (yyyy-m-1)
  - U132 = CBD 的年月 (yyyy-m-1)
  - X132 = 100%
  - W135 / Y135 清空
修改记录写入 _ChangeLog, 直接修改原文件 (不另存)
"""

import os
import datetime

from core.excel_helper import (
    FastExcel, read_cell_state, write_cell_by_type, calc_full
)
from core.change_log import (
    ensure_log_sheet, flush_log_rows, make_log_row, make_batch_id
)


MODULE_INFO = {
    "id": "tooling_modify",
    "name": "[1] Tooling 批量修改",
    "description": "对 CQSC 文件中的 Tooling* 工作簿做批量修改",
    "needs_files": True,
    "file_dialog_title": "选择需要批量修改 Tooling 的 Excel (可多选)",
    "settings_section": "tooling_modify",
}


def today_yyyymd1():
    """今天日期格式化为 yyyy-m-1 (月份不补零)"""
    d = datetime.date.today()
    return str(d.year) + "-" + str(d.month) + "-1"


def fmt_yyyymd1(year, month):
    """指定年月格式化为 yyyy-m-1"""
    return str(int(year)) + "-" + str(int(month)) + "-1"


def process_one_file(excel, file_path, settings, logw):
    """处理单个文件"""
    fn = os.path.basename(file_path)
    logw.log("===== [Tooling 修改] " + fn + " =====", status="修改中: " + fn)

    # 取参数
    common = settings.get("_common", {})
    password = common.get("open_password", "")

    tooling_prefix    = settings.get("tooling_prefix", "Tooling")
    tooling_flag_cell = settings.get("tooling_flag_cell", "R121")
    tooling_flag_val  = settings.get("tooling_flag_val",
                                     "Cash Flow Projection for Tooling")

    cbd_prefix     = settings.get("cbd_prefix", "CBD")
    cbd_flag_cell  = settings.get("cbd_flag_cell", "C8")
    cbd_flag_val   = settings.get("cbd_flag_val", "SOP Year")
    cbd_year_cell  = settings.get("cbd_year_cell", "D8")
    cbd_month_cell = settings.get("cbd_month_cell", "F8")

    log_sheet = settings.get("log_sheet_name", "_ChangeLog")

    # 打开文件
    try:
        wb = excel.Workbooks.Open(file_path, False, False,
                                  None, password, password)
    except Exception as e:
        logw.log("  [错误] 打开失败: " + str(e))
        return False

    try:
        with FastExcel(excel):
            tooling_sheets = []
            cbd_sheet = None
            logw.log("  扫描 sheet ...")
            for sh in wb.Sheets:
                name = sh.Name
                # 跳过日志 sheet
                if name == log_sheet:
                    continue

                # CBD 命中（取第一个）
                if cbd_sheet is None and name.startswith(cbd_prefix):
                    try:
                        v = sh.Range(cbd_flag_cell).Value
                    except Exception:
                        v = None
                    if v is not None and str(v).strip() == cbd_flag_val:
                        cbd_sheet = sh
                        logw.log("  [CBD ] 命中: " + name)

                # Tooling 命中（全部）
                if name.startswith(tooling_prefix):
                    try:
                        v = sh.Range(tooling_flag_cell).Value
                    except Exception:
                        v = None
                    if v is not None and str(v).strip() == tooling_flag_val:
                        tooling_sheets.append(sh)
                        logw.log("  [Tool] 命中: " + name)

            if not tooling_sheets:
                logw.log("  [跳过] 未找到符合条件的 Tooling* sheet")
                wb.Close(SaveChanges=False)
                return False

            # 取 CBD 年月
            cbd_year = None
            cbd_month = None
            if cbd_sheet is not None:
                try:
                    cbd_year  = cbd_sheet.Range(cbd_year_cell).Value
                    cbd_month = cbd_sheet.Range(cbd_month_cell).Value
                    logw.log("  [CBD ] " + cbd_year_cell + "=" + str(cbd_year) +
                             ", " + cbd_month_cell + "=" + str(cbd_month))
                except Exception as e:
                    logw.log("  [警告] 读 CBD 年月失败: " + str(e))
            else:
                logw.log("  [警告] 未找到 CBD* sheet, U132 留空")

            today_s = today_yyyymd1()
            if cbd_year and cbd_month:
                u132_val = fmt_yyyymd1(cbd_year, cbd_month)
            else:
                u132_val = ""

            log_ws, next_row = ensure_log_sheet(wb, log_sheet)
            batch_id = make_batch_id()

            # 修改清单: (cell, NewType, NewValue, NewFormat)
            changes = [
                ("W125", "Formula", "=-S118", None),
                ("Y125", "Formula", "=S64",   None),
                ("U126", "Value",   today_s,  None),
                ("U132", "Value",   u132_val, None),
                ("X132", "Value",   1,        "0%"),    # 100%
                ("W135", "Empty",   "",       None),
                ("Y135", "Empty",   "",       None),
            ]

            log_rows = []
            for sh in tooling_sheets:
                logw.log("  写入 " + sh.Name + " ...")
                try:
                    for addr, ntype, nval, nfmt in changes:
                        old_type, old_val, old_fmt = read_cell_state(sh, addr)
                        write_cell_by_type(sh, addr, ntype, nval, nfmt)
                        log_rows.append(make_log_row(
                            batch_id, "Modify", sh.Name, addr,
                            old_type, old_val, old_fmt,
                            ntype, nval, (nfmt or old_fmt),
                            "Tooling 批量修改"))
                    logw.log("  [完成] " + sh.Name)
                except Exception as e:
                    logw.log("  [错误] 写入 " + sh.Name + " 失败: " + str(e))

            # 写日志
            logw.log("  写入日志 (" + str(len(log_rows)) + " 行) ...")
            flush_log_rows(log_ws, next_row, log_rows)

            # 重算 + 保存
            calc_full(excel, logw)
            logw.log("  保存 ...")
            wb.Save()

        wb.Close(SaveChanges=False)
        logw.log("  [完成] BatchID=" + batch_id)
        return True
    except Exception as e:
        logw.log("  [错误] " + str(e))
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        return False


def run(excel, files, settings, logw):
    """模块统一入口"""
    ok = 0
    fail = 0
    for idx, f in enumerate(files, 1):
        f = os.path.abspath(f)
        logw.log("\n>>>>> 文件 " + str(idx) + "/" + str(len(files)) +
                 ": " + os.path.basename(f),
                 status="处理 " + str(idx) + "/" + str(len(files)))
        if process_one_file(excel, f, settings, logw):
            ok += 1
        else:
            fail += 1
    return ok, fail