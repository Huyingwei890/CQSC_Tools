# -*- coding: utf-8 -*-
"""
[3] ROI 合并工具
将多个 CQSC 文件中的 SFP254v1.5_CQSC-PreCER* 工作簿合并到 Template
"""

import os
import shutil
import datetime

from core.excel_helper import (
    FastExcel, col_letter, find_sheet_by_name,
    make_unique_sheet_name, calc_full
)


MODULE_INFO = {
    "id": "roi_combine",
    "name": "[3] ROI 合并工具",
    "description": "将多个 CQSC 文件合并到 ROI Template",
    "needs_files": True,
    "file_dialog_title": "选择需要合并的 CQSC Excel 文件 (可多选)",
    "settings_section": "roi_combine",
}


def build_roi_output_path(output_folder, template_path):
    """生成 output_folder/ROI Combine YYYY-MM-DD.xlsx 路径"""
    if output_folder and output_folder.strip():
        folder = output_folder.strip()
    else:
        folder = os.path.dirname(template_path)

    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass

    date_s = datetime.date.today().strftime("%Y-%m-%d")
    base = "ROI Combine " + date_s
    p = os.path.join(folder, base + ".xlsx")
    i = 2
    while os.path.exists(p):
        p = os.path.join(folder, base + "-" + str(i) + ".xlsx")
        i += 1
    return p


def build_external_link(file_path, sheet_name, cell):
    folder = os.path.dirname(file_path)
    fname = os.path.basename(file_path)
    return "='" + folder + "\\[" + fname + "]" + sheet_name + "'!" + cell


def _safe_close(wb):
    """安全关闭 workbook, 忽略所有错误"""
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def run(excel, files, settings, logw):
    """模块统一入口"""
    logw.log("===== [ROI 合并] 启动 =====", status="ROI 合并启动")

    common = settings.get("_common", {})
    password = common.get("open_password", "")

    template_path = settings.get("template_path", "")
    output_folder = settings.get("output_folder", "")

    precer_prefix    = settings.get("precer_prefix", "SFP254v1.5_CQSC-PreCER")
    precer_flag_cell = settings.get("precer_flag_cell", "I7")
    precer_flag_val  = settings.get("precer_flag_val",
                                    "Pre-CER / CER Financial Evaluation")

    pi_prefix    = settings.get("pi_prefix", "PI")
    pi_flag_cell = settings.get("pi_flag_cell", "B3")
    pi_flag_val  = settings.get("pi_flag_val", "Part Description")

    pi_src_header_row     = settings.get("pi_src_header_row", 3)
    pi_src_data_row_start = settings.get("pi_src_data_row_start", 4)
    pi_src_col_start      = settings.get("pi_src_col_start", 2)
    pi_src_col_end        = settings.get("pi_src_col_end", 15)

    pi_dst_header_row     = settings.get("pi_dst_header_row", 2)
    pi_dst_data_row_start = settings.get("pi_dst_data_row_start", 3)
    pi_dst_col_start      = settings.get("pi_dst_col_start", 1)
    pi_data_rows          = settings.get("pi_data_rows", 2)
    pi_summary_sheet_name = settings.get("pi_summary_sheet", "PI Summary")
    pi_clear_rows         = 2000

    cons_sheet_name = settings.get("cons_sheet", "SFP254v1.5_CONSOLIDATION")
    cons_pic_range  = settings.get("cons_pic_range", "I352:Y392")
    cons_fill_cells = settings.get("cons_fill_cells", ["L11", "L12", "N365"])

    ppt_sheet_name   = settings.get("ppt_sheet", "PPT1")
    ppt_anchor_cell  = settings.get("ppt_anchor_cell", "A1")

    end_sheet_name = settings.get("end_sheet", "<<End")

    # 1) 检查模板
    if not template_path or not os.path.exists(template_path):
        logw.log("  [错误] 找不到模板文件: " + str(template_path))
        logw.log("  请到 [设置] 参数设置 -> ROI 合并 里设置正确的 Template 路径")
        return 0, 1

    # 2) 生成输出路径并复制模板
    try:
        out_path = build_roi_output_path(output_folder, template_path)
    except Exception as e:
        logw.log("  [错误] 准备输出路径失败: " + str(e))
        return 0, 1

    logw.log("  [输出] 目标: " + out_path)

    try:
        shutil.copyfile(template_path, out_path)
        logw.log("  [模板] 已复制 -> " + os.path.basename(out_path))
    except Exception as e:
        logw.log("  [错误] 复制模板失败: " + str(e))
        return 0, 1

    # 3) 打开目标文件
    try:
        out_wb = excel.Workbooks.Open(out_path, 0)
    except Exception as e:
        logw.log("  [错误] 打开目标失败: " + str(e))
        return 0, 1

    open_cqsc_wbs = []

    try:
        with FastExcel(excel):
            end_sheet        = find_sheet_by_name(out_wb, end_sheet_name)
            cons_sheet       = find_sheet_by_name(out_wb, cons_sheet_name)
            ppt_sheet        = find_sheet_by_name(out_wb, ppt_sheet_name)
            pi_summary_sheet = find_sheet_by_name(out_wb, pi_summary_sheet_name)

            if end_sheet is None:
                logw.log("  [错误] 模板中找不到 " + end_sheet_name)
                _safe_close(out_wb)
                return 0, 1
            if cons_sheet is None:
                logw.log("  [错误] 模板中找不到 " + cons_sheet_name)
                _safe_close(out_wb)
                return 0, 1
            if ppt_sheet is None:
                logw.log("  [警告] 模板中找不到 " + ppt_sheet_name)
            if pi_summary_sheet is None:
                logw.log("  [警告] 模板中找不到 " + pi_summary_sheet_name)

            cons_values = None
            pi_links = []
            copied_precer_total = 0

            # 4) 遍历每个 CQSC
            for idx, cqsc_path in enumerate(files, 1):
                cqsc_path = os.path.abspath(cqsc_path)
                fn = os.path.basename(cqsc_path)
                logw.log("\n  >>> 处理 CQSC " + str(idx) + "/" +
                         str(len(files)) + ": " + fn,
                         status="处理 " + str(idx) + "/" + str(len(files)))
                try:
                    src_wb = excel.Workbooks.Open(
                        cqsc_path, False, True,
                        None, password, password)
                    open_cqsc_wbs.append(src_wb)
                except Exception as e:
                    logw.log("    [错误] 打开失败: " + str(e))
                    continue

                try:
                    precer_sheets = []
                    pi_sheets_local = []
                    for sh in src_wb.Sheets:
                        nm = sh.Name
                        if nm.startswith(precer_prefix):
                            try:
                                v = sh.Range(precer_flag_cell).Value
                            except Exception:
                                v = None
                            if v is not None and str(v).strip() == precer_flag_val:
                                precer_sheets.append(sh)
                                logw.log("    [PreCER] 命中: " + nm)
                        if nm.startswith(pi_prefix):
                            try:
                                v = sh.Range(pi_flag_cell).Value
                            except Exception:
                                v = None
                            if v is not None and str(v).strip() == pi_flag_val:
                                pi_sheets_local.append(sh)
                                logw.log("    [PI    ] 命中: " + nm)

                    if cons_values is None and precer_sheets:
                        sh0 = precer_sheets[0]
                        cons_values = {}
                        for c in cons_fill_cells:
                            try:
                                cons_values[c] = sh0.Range(c).Value
                            except Exception:
                                cons_values[c] = None
                        msg = "    [CONS ] 取值 "
                        msg += ", ".join(c + "=" + str(cons_values.get(c))
                                         for c in cons_fill_cells)
                        msg += "  (来源: " + sh0.Name + ")"
                        logw.log(msg)

                    for sh in precer_sheets:
                        original_name = sh.Name
                        try:
                            sh.Copy(Before=out_wb.Sheets(end_sheet_name))
                            new_sh = excel.ActiveSheet
                            unique = make_unique_sheet_name(out_wb, original_name)
                            if unique != original_name:
                                new_sh.Name = unique
                                logw.log("    [拷贝 ] " + original_name + " -> " + unique)
                            else:
                                logw.log("    [拷贝 ] " + original_name)
                            copied_precer_total += 1
                        except Exception as e:
                            logw.log("    [错误] 拷贝 " + original_name +
                                     " 失败: " + str(e))

                    for sh in pi_sheets_local:
                        pi_links.append((cqsc_path, sh.Name))
                except Exception as e:
                    logw.log("    [错误] 处理失败: " + str(e))

            logw.log("\n  ----- 共拷贝 PreCER " + str(copied_precer_total) +
                     " 个, 识别 PI " + str(len(pi_links)) + " 个 -----")

            # 5) 填 CONSOLIDATION
            if cons_values is not None:
                try:
                    for c in cons_fill_cells:
                        cons_sheet.Range(c).Value = cons_values.get(c)
                    logw.log("  [CONS ] 已写入 " +
                             " / ".join(cons_fill_cells))
                except Exception as e:
                    logw.log("  [错误] 写 CONSOLIDATION 失败: " + str(e))
            else:
                logw.log("  [警告] 未能从任何 CQSC 取到 CONSOLIDATION 值")

            # 6) PPT1 链接图片
            if ppt_sheet is not None:
                try:
                    excel.ScreenUpdating = True
                    cons_sheet.Range(cons_pic_range).Copy()
                    ppt_sheet.Activate()
                    ppt_sheet.Range(ppt_anchor_cell).Select()
                    ppt_sheet.Pictures().Paste(Link=True)
                    excel.CutCopyMode = False
                    excel.ScreenUpdating = False
                    logw.log("  [PPT1 ] 已粘贴链接图片 (" + cons_pic_range +
                             " -> " + ppt_sheet_name + "!" + ppt_anchor_cell + ")")
                except Exception as e:
                    logw.log("  [错误] PPT1 截图失败: " + str(e))
                    try:
                        excel.CutCopyMode = False
                    except Exception:
                        pass

            # 7) PI Summary
            if pi_summary_sheet is not None and pi_links:
                try:
                    n_cols = pi_src_col_end - pi_src_col_start + 1
                    dst_c_start = pi_dst_col_start
                    dst_c_end   = pi_dst_col_start + n_cols - 1
                    c_dst_start = col_letter(dst_c_start)
                    c_dst_end   = col_letter(dst_c_end)

                    clear_addr = (c_dst_start + str(pi_dst_header_row) + ":" +
                                  c_dst_end + str(pi_dst_header_row + pi_clear_rows))
                    pi_summary_sheet.Range(clear_addr).ClearContents()
                    logw.log("  [PI   ] 清空旧数据: " + clear_addr)

                    logw.log("  [PI   ] 写入 " + str(len(pi_links)) +
                             " 个来源的链接 ...")

                    head_src_path, head_pi_name = pi_links[0]
                    head_formulas = []
                    for c in range(pi_src_col_start, pi_src_col_end + 1):
                        cl = col_letter(c)
                        head_formulas.append(
                            build_external_link(head_src_path, head_pi_name,
                                                cl + str(pi_src_header_row)))
                    head_addr = (c_dst_start + str(pi_dst_header_row) + ":" +
                                 c_dst_end + str(pi_dst_header_row))
                    pi_summary_sheet.Range(head_addr).Formula = [head_formulas]
                    logw.log("    [PI 标题] " + os.path.basename(head_src_path) +
                             " | " + head_pi_name + " -> " + head_addr)

                    for i, (src_path, pi_name) in enumerate(pi_links):
                        target_row_start = pi_dst_data_row_start + i * pi_data_rows
                        formulas = []
                        for r_off in range(pi_data_rows):
                            src_row = pi_src_data_row_start + r_off
                            row_formulas = []
                            for c in range(pi_src_col_start, pi_src_col_end + 1):
                                cl = col_letter(c)
                                f = build_external_link(src_path, pi_name,
                                                        cl + str(src_row))
                                row_formulas.append(f)
                            formulas.append(row_formulas)
                        addr = (c_dst_start + str(target_row_start) + ":" +
                                c_dst_end + str(target_row_start + pi_data_rows - 1))
                        pi_summary_sheet.Range(addr).Formula = formulas
                        logw.log("    [PI #" + str(i + 1) + "] " +
                                 os.path.basename(src_path) + " | " + pi_name +
                                 " -> " + addr)
                except Exception as e:
                    logw.log("  [错误] 写 PI Summary 失败: " + str(e))
            elif pi_summary_sheet is not None:
                logw.log("  [PI   ] 无可链接的 PI sheet, 已跳过")

            # 8) 保存
            try:
                out_wb.UpdateLinks = 0
            except Exception:
                pass

            calc_full(excel, logw)

            logw.log("  保存 -> " + os.path.basename(out_path))
            out_wb.Save()

        _safe_close(out_wb)
        logw.log("\n  [完成] 输出文件: " + out_path)
        return 1, 0

    except Exception as e:
        logw.log("  [错误] " + str(e))
        _safe_close(out_wb)
        return 0, 1

    finally:
        for src_wb in open_cqsc_wbs:
            _safe_close(src_wb)
        logw.log("  [清理] 已关闭 " + str(len(open_cqsc_wbs)) + " 个源 CQSC 文件")