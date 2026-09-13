# -*- coding: utf-8 -*-
"""修改日志的读写和回滚通用逻辑"""

import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from core.excel_helper import (
    XL_UP, FastExcel, read_cell_state, write_cell_by_type, calc_full
)


LOG_HEADERS = ["BatchID", "Timestamp", "Operation", "SheetName", "Cell",
               "OldType", "OldValue", "OldFormat",
               "NewType", "NewValue", "NewFormat", "Note"]


def make_batch_id():
    return "B" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _safe_str(v):
    """把值安全地转成字符串。
    关键: 公式或以 = + @ 开头的字符串前加单引号 (Excel 文本前缀),
    防止 Excel 写入时把它当公式二次求值。"""
    if v is None:
        return ""
    s = str(v)
    if s and s[0] in ("=", "+", "@"):
        return "'" + s
    return s


def make_log_row(batch_id, op, sheet_name, cell,
                 old_type, old_val, old_fmt,
                 new_type, new_val, new_fmt, note=""):
    """生成一行日志记录"""
    return [batch_id, now_str(), op, sheet_name, cell,
            old_type, _safe_str(old_val), (old_fmt or ""),
            new_type, _safe_str(new_val), (new_fmt or ""),
            note]


def ensure_log_sheet(wb, log_sheet_name):
    """获取或创建日志 sheet，返回 (sheet对象, 下一个空行号)
    关键: 把 OldValue / NewValue 列设置为文本格式，防止公式被二次求值"""
    log_ws = None
    for sh in wb.Sheets:
        if sh.Name == log_sheet_name:
            log_ws = sh
            break

    if log_ws is None:
        log_ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        log_ws.Name = log_sheet_name
        for i, h in enumerate(LOG_HEADERS, start=1):
            log_ws.Cells(1, i).Value = h
        log_ws.Rows(1).Font.Bold = True
        log_ws.Columns("A:L").ColumnWidth = 16

    # 关键保险: 不管是新建还是已有，都强制把 G(OldValue)/J(NewValue) 列设文本
    try:
        log_ws.Columns("G:G").NumberFormat = "@"
        log_ws.Columns("J:J").NumberFormat = "@"
    except Exception:
        pass

    last_row = log_ws.Cells(log_ws.Rows.Count, 1).End(XL_UP).Row
    if last_row < 1:
        last_row = 1
    return log_ws, last_row + 1


def flush_log_rows(log_ws, start_row, rows):
    """批量写入多行日志（数组赋值，比逐格快几十倍）"""
    if not rows:
        return
    n_rows = len(rows)
    n_cols = len(LOG_HEADERS)
    rng = log_ws.Range(log_ws.Cells(start_row, 1),
                       log_ws.Cells(start_row + n_rows - 1, n_cols))
    rng.Value = rows


def _clean_log_value(s):
    """读取日志值时，去除前导单引号（写入时为防 Excel 求值加的）"""
    if s is None:
        return s
    if isinstance(s, str) and s.startswith("'"):
        return s[1:]
    return s


def pick_batch_dialog(batches, file_name, title_prefix="选择要回滚的批次"):
    """显示批次选择对话框，返回选中的 BatchID 或 None"""
    sel = {"val": None}
    win = tk.Toplevel()
    win.title(title_prefix + " - " + file_name)
    win.geometry("760x420")
    win.grab_set()

    tk.Label(
        win,
        text="文件: " + file_name + "\n选择批次后，该批次涉及的所有单元格将恢复到改之前的状态:",
        anchor="w", justify="left"
    ).pack(fill="x", padx=10, pady=8)

    cols = ("BatchID", "Timestamp", "Operation", "Changes")
    tv = ttk.Treeview(win, columns=cols, show="headings", height=14)
    for c, w in zip(cols, (180, 170, 180, 90)):
        tv.heading(c, text=c)
        tv.column(c, width=w, anchor="w")
    tv.pack(fill="both", expand=True, padx=10)

    for b in batches:
        tv.insert("", "end", values=(b["batch_id"], b["ts"], b["op"], b["count"]))

    def ok():
        item = tv.focus()
        if not item:
            messagebox.showwarning("提示", "请先选择一个批次")
            return
        sel["val"] = tv.item(item)["values"][0]
        win.destroy()

    def cancel():
        win.destroy()

    frm = tk.Frame(win)
    frm.pack(pady=8)
    tk.Button(frm, text="回滚此批次", width=14, command=ok).pack(side="left", padx=10)
    tk.Button(frm, text="取消", width=14, command=cancel).pack(side="left", padx=10)
    win.wait_window()
    return sel["val"]


def rollback_workbook(excel, wb, log_sheet_name, op_label, logw):
    """对一个已打开的工作簿执行回滚操作
    返回:
      (ok_count, err_count) -- 成功完成
      None                  -- 用户取消或日志不存在
    """
    log_ws = None
    for sh in wb.Sheets:
        if sh.Name == log_sheet_name:
            log_ws = sh
            break

    if log_ws is None:
        logw.log("  [跳过] 无 " + log_sheet_name)
        return None

    last_row = log_ws.Cells(log_ws.Rows.Count, 1).End(XL_UP).Row
    if last_row < 2:
        logw.log("  [跳过] 日志为空")
        return None

    n_cols = len(LOG_HEADERS)
    logw.log("  读取日志 " + str(last_row - 1) + " 行 ...")
    data = log_ws.Range(log_ws.Cells(2, 1),
                        log_ws.Cells(last_row, n_cols)).Value
    if last_row - 1 == 1:
        data = (data,)

    entries = []
    for row in data:
        entries.append({
            "batch_id": row[0], "ts": row[1], "op": row[2],
            "sheet": row[3], "cell": row[4],
            "old_type": row[5],
            "old_val":  _clean_log_value(row[6]),   # 去掉前导'
            "old_fmt":  row[7],
            "new_type": row[8],
            "new_val":  _clean_log_value(row[9]),   # 去掉前导'
            "new_fmt":  row[10],
        })

    # 聚合批次
    batch_map = {}
    for e in entries:
        bid = e["batch_id"]
        if not bid:
            continue
        if bid not in batch_map:
            batch_map[bid] = {"batch_id": bid, "ts": e["ts"],
                              "op": e["op"], "count": 0}
        batch_map[bid]["count"] += 1
    batches = sorted(batch_map.values(),
                     key=lambda x: x["batch_id"], reverse=True)

    selected = pick_batch_dialog(batches, wb.Name)
    if not selected:
        logw.log("  [取消] 用户未选择批次")
        return None

    target = [e for e in entries if str(e["batch_id"]) == str(selected)]
    logw.log("  目标批次: " + str(selected) + ", 涉及 " + str(len(target)) + " 条")

    with FastExcel(excel):
        new_batch = make_batch_id()
        log_rows = []
        ok_n = 0
        err_n = 0

        for e in target:
            try:
                ws = wb.Sheets(e["sheet"])
                cur_type, cur_val, cur_fmt = read_cell_state(ws, e["cell"])

                old_type = e["old_type"] or "Empty"
                old_val = e["old_val"]
                old_fmt = e["old_fmt"] or None
                write_cell_by_type(ws, e["cell"], old_type, old_val, old_fmt)

                op_lbl = op_label + "(" + str(selected) + ")"
                log_rows.append(make_log_row(
                    new_batch, op_lbl,
                    e["sheet"], e["cell"],
                    cur_type, cur_val, cur_fmt,
                    old_type, old_val, old_fmt,
                    "回滚批次 " + str(selected)))
                ok_n += 1
                logw.log("    [还原] " + str(e["sheet"]) + "!" + str(e["cell"]) +
                         "  <- " + str(old_type) + ":" + str(old_val))
            except Exception as ex:
                err_n += 1
                logw.log("    [错误] " + str(e["sheet"]) + "!" + str(e["cell"]) +
                         " 失败: " + str(ex))

        logw.log("  写入回滚日志 (" + str(len(log_rows)) + " 行) ...")
        flush_log_rows(log_ws, last_row + 1, log_rows)

        calc_full(excel, logw)

    logw.log("  [完成] 成功 " + str(ok_n) + ", 失败 " + str(err_n) +
             ", 新批次 " + new_batch)
    return ok_n, err_n