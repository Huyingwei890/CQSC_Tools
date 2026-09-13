# -*- coding: utf-8 -*-
"""
[2] Tooling 按批次回滚
读取文件中的 _ChangeLog, 按 BatchID 选择回滚
"""

import os

from core.change_log import rollback_workbook


MODULE_INFO = {
    "id": "tooling_rollback",
    "name": "[2] Tooling 按批次回滚",
    "description": "在文件中按批次回滚 Tooling 修改",
    "needs_files": True,
    "file_dialog_title": "选择需要 Tooling 回滚的 Excel 文件 (可多选)",
    "settings_section": "tooling_rollback",
}


def process_one_file(excel, file_path, settings, logw):
    """处理单个文件回滚"""
    fn = os.path.basename(file_path)
    logw.log("===== [Tooling 回滚] " + fn + " =====", status="Tooling 回滚: " + fn)

    common = settings.get("_common", {})
    password = common.get("open_password", "")
    log_sheet = settings.get("log_sheet_name", "_ChangeLog")

    try:
        wb = excel.Workbooks.Open(file_path, False, False,
                                  None, password, password)
    except Exception as e:
        logw.log("  [错误] 打开失败: " + str(e))
        return False

    try:
        result = rollback_workbook(excel, wb, log_sheet, "Rollback", logw)
        if result is None:
            # 用户取消或日志不存在
            wb.Close(SaveChanges=False)
            return False
        ok_n, err_n = result

        logw.log("  保存 ...")
        wb.Save()
        wb.Close(SaveChanges=False)
        logw.log("  [完成] 成功 " + str(ok_n) + ", 失败 " + str(err_n))
        return err_n == 0
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