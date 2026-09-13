# -*- coding: utf-8 -*-
"""
[4] To CM Version
对 CQSC 文件的 GP*/CBD* 工作簿做 CM 版本修改:
  - GP*  : I26 = 0.5%, I27 清空
  - CBD* : M321:M420 *0.7, V321:V420 *0.3
复制原文件 -> +CM 新文件, 修改记录写入 _CM_ChangeLog
"""

import os
import shutil

from core.excel_helper import (
    FastExcel, col_letter, fuzzy_match,
    read_cell_state, write_cell_by_type,
    wrap_formula_times, to_2d_array, calc_full
)
from core.change_log import (
    ensure_log_sheet, flush_log_rows, make_log_row, make_batch_id
)


MODULE_INFO = {
    "id": "cm_version",
    "name": "[4] To CM Version",
    "description": "复制 CQSC 文件 -> +CM 版本，自动修改 GP*/CBD* 工作簿",
    "needs_files": True,
    "file_dialog_title": "选择需要 To CM Version 的 CQSC 文件 (可多选)",
    "settings_section": "cm_version",
}


def build_cm_output_path(src_path, suffix):
    """生成 +CM 文件路径，已存在自动加 (2)/(3)..."""
    folder = os.path.dirname(src_path)
    base, ext = os.path.splitext(os.path.basename(src_path))
    p = os.path.join(folder, base + suffix + ext)
    i = 2
    while os.path.exists(p):
        p = os.path.join(folder, base + suffix + "(" + str(i) + ")" + ext)
        i += 1
    return p


def process_cm_range(ws, range_addr, factor, sheet_name, batch_id,
                     log_rows, logw):
    """对一个区域内的非空单元格执行 *factor 修改"""
    rng = ws.Range(range_addr)
    n_rows = rng.Rows.Count
    n_cols = rng.Columns.Count
    first_row = rng.Row
    first_col = rng.Column

    vals = to_2d_array(rng.Value, n_rows, n_cols)
    fmls = to_2d_array(rng.Formula, n_rows, n_cols)

    cnt = 0
    skipped_text = 0
    for r in range(n_rows):
        for c in range(n_cols):
            v = vals[r][c]
            f = fmls[r][c]
            if v is None or v == "":
                continue
            addr = col_letter(first_col + c) + str(first_row + r)
            cell = ws.Range(addr)
            if isinstance(f, str) and f.startswith("="):
                old_type = "Formula"
                old_val = f
                new_val = wrap_formula_times(f, factor)
                new_type = "Formula"
                cell.Formula = new_val
            else:
                try:
                    num = float(v)
                except Exception:
                    skipped_text += 1
                    continue
                old_type = "Value"
                old_val = num
                new_val = num * factor
                new_type = "Value"
                cell.Value = new_val
            log_rows.append(make_log_row(
                batch_id, "CM_Modify", sheet_name, addr,
                old_type, old_val, "",
                new_type, new_val, "",
                "区域 " + range_addr + " *" + str(factor)))
            cnt += 1

    msg = "    [CBD ] " + range_addr + " *" + str(factor) + " 修改 " + str(cnt) + " 个单元格"
    if skipped_text > 0:
        msg += "  (跳过 " + str(skipped_text) + " 个文本)"
    logw.log(msg)
    return cnt


def process_one_file(excel, src_path, settings, logw):
    """处理单个文件"""
    fn = os.path.basename(src_path)
    logw.log("===== [To CM] " + fn + " =====", status="To CM: " + fn)

    # 取参数
    common = settings.get("_common", {})
    password = common.get("open_password", "")

    gp_prefix     = settings.get("gp_prefix", "GP")
    gp_flag_cell  = settings.get("gp_flag_cell", "C4")
    gp_flag_val   = settings.get("gp_flag_val", "Volume (pcs)")

    cbd_prefix    = settings.get("cbd_prefix", "CBD")
    cbd_flag_cell = settings.get("cbd_flag_cell", "T320")
    cbd_flag_val  = settings.get("cbd_flag_val", "No.of DL Required per part")
    cbd_keywords  = settings.get("cbd_flag_keywords", ["dl", "required", "part"])

    m_range  = settings.get("cbd_m_range", "M321:M420")
    m_factor = settings.get("cbd_m_factor", 0.7)
    v_range  = settings.get("cbd_v_range", "V321:V420")
    v_factor = settings.get("cbd_v_factor", 0.3)

    i26_val   = settings.get("gp_i26_value", 0.005)
    i26_fmt   = settings.get("gp_i26_format", "0.0%")
    i27_clear = settings.get("gp_i27_clear", True)

    suffix    = settings.get("file_suffix", "+CM")
    log_sheet = settings.get("log_sheet_name", "_CM_ChangeLog")

    # 1) 复制文件
    out_path = build_cm_output_path(src_path, suffix)
    try:
        shutil.copyfile(src_path, out_path)
        logw.log("  [复制] -> " + os.path.basename(out_path))
    except Exception as e:
        logw.log("  [错误] 复制失败: " + str(e))
        return False

    # 2) 打开新文件
    try:
        wb = excel.Workbooks.Open(out_path, False, False,
                                  None, password, password)
    except Exception as e:
        logw.log("  [错误] 打开 +CM 文件失败: " + str(e))
        return False

    try:
        with FastExcel(excel):
            gp_sheets = []
            cbd_sheets = []
            logw.log("  扫描 sheet ...")
            for sh in wb.Sheets:
                name = sh.Name
                if name == log_sheet:
                    continue

                # GP 命中
                if name.startswith(gp_prefix):
                    try:
                        v = sh.Range(gp_flag_cell).Value
                    except Exception:
                        v = None
                    if v is not None and str(v).strip() == gp_flag_val:
                        gp_sheets.append(sh)
                        logw.log("  [GP  ] 命中: " + name)
                    elif v is not None:
                        if fuzzy_match(v, gp_flag_val,
                                       keywords=["volume", "pcs"]):
                            gp_sheets.append(sh)
                            logw.log("  [GP  ] 命中(模糊): " + name +
                                     "  原值=" + repr(v))

                # CBD 命中(模糊匹配)
                if name.startswith(cbd_prefix):
                    try:
                        v = sh.Range(cbd_flag_cell).Value
                    except Exception:
                        v = None
                    if v is not None:
                        if fuzzy_match(v, cbd_flag_val,
                                       keywords=cbd_keywords):
                            cbd_sheets.append(sh)
                            logw.log("  [CBD ] 命中: " + name +
                                     "  原值=" + repr(v))
                        else:
                            logw.log("  [CBD ] 未命中: " + name +
                                     "  " + cbd_flag_cell + "=" + repr(v))

            if not gp_sheets and not cbd_sheets:
                logw.log("  [跳过] 未找到符合条件的 GP* / CBD* sheet")
                wb.Close(SaveChanges=False)
                try: os.remove(out_path)
                except Exception: pass
                return False

            log_ws, next_row = ensure_log_sheet(wb, log_sheet)
            batch_id = make_batch_id()
            log_rows = []

            # 3) GP 修改
            for sh in gp_sheets:
                logw.log("  写入 GP: " + sh.Name)
                try:
                    addr = "I26"
                    old_t, old_v, old_f = read_cell_state(sh, addr)
                    write_cell_by_type(sh, addr, "Value", i26_val, i26_fmt)
                    log_rows.append(make_log_row(
                        batch_id, "CM_Modify", sh.Name, addr,
                        old_t, old_v, old_f,
                        "Value", i26_val, i26_fmt,
                        "GP I26 设为 " + str(i26_val)))
                    logw.log("    [GP  ] I26 = " + str(i26_val))

                    if i27_clear:
                        addr = "I27"
                        old_t, old_v, old_f = read_cell_state(sh, addr)
                        write_cell_by_type(sh, addr, "Empty", "", None)
                        log_rows.append(make_log_row(
                            batch_id, "CM_Modify", sh.Name, addr,
                            old_t, old_v, old_f,
                            "Empty", "", old_f,
                            "GP I27 清空"))
                        logw.log("    [GP  ] I27 清空")
                except Exception as e:
                    logw.log("    [错误] GP " + sh.Name + " 失败: " + str(e))

            # 4) CBD 修改
            for sh in cbd_sheets:
                logw.log("  写入 CBD: " + sh.Name)
                try:
                    process_cm_range(sh, m_range, m_factor,
                                     sh.Name, batch_id, log_rows, logw)
                    process_cm_range(sh, v_range, v_factor,
                                     sh.Name, batch_id, log_rows, logw)
                except Exception as e:
                    logw.log("    [错误] CBD " + sh.Name + " 失败: " + str(e))

            # 5) 写日志
            logw.log("  写入日志 (" + str(len(log_rows)) + " 行) ...")
            flush_log_rows(log_ws, next_row, log_rows)

            # 6) 重算 + 保存
            calc_full(excel, logw)
            logw.log("  保存 ...")
            wb.Save()

        wb.Close(SaveChanges=False)
        logw.log("  [完成] 输出: " + out_path + "  BatchID=" + batch_id)
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