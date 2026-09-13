# -*- coding: utf-8 -*-
"""配置加载、保存、可视化编辑"""

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from settings_default import DEFAULT_SETTINGS, SETTINGS_META


class SettingsManager:
    """配置管理: 加载 / 保存 / 合并默认值"""

    def __init__(self, root_dir):
        self.root_dir = root_dir
        self.path = os.path.join(root_dir, "settings.json")

    def load(self):
        """加载配置，文件不存在则用默认值初始化"""
        if not os.path.exists(self.path):
            self.save(DEFAULT_SETTINGS)
            return json.loads(json.dumps(DEFAULT_SETTINGS))
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                user = json.load(f)
        except Exception:
            user = {}
        # 合并默认值（默认值有但用户没有的键，自动补上）
        merged = json.loads(json.dumps(DEFAULT_SETTINGS))
        for sec_key, sec_val in user.items():
            if sec_key in merged and isinstance(merged[sec_key], dict):
                merged[sec_key].update(sec_val)
            else:
                merged[sec_key] = sec_val
        return merged

    def save(self, settings):
        """保存配置到 settings.json"""
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)


def open_settings_dialog(parent, mgr):
    """打开参数设置对话框，返回 True 表示有修改"""
    settings = mgr.load()
    win = tk.Toplevel(parent)
    win.title("参数设置")
    win.geometry("780x620")
    win.grab_set()

    # Notebook 多 Tab
    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True, padx=10, pady=10)

    entries = {}   # {(section, key): (Entry, type_str)}

    # 每个区段一个 Tab
    for section, meta in SETTINGS_META.items():
        if section not in settings:
            settings[section] = {}

        frm = tk.Frame(nb, padx=12, pady=12)
        nb.add(frm, text=meta.get("_label", section))

        # 表头
        tk.Label(frm, text="参数", font=("微软雅黑", 9, "bold"),
                 width=28, anchor="w").grid(row=0, column=0, sticky="w", pady=(0, 6))
        tk.Label(frm, text="值", font=("微软雅黑", 9, "bold"),
                 anchor="w").grid(row=0, column=1, sticky="w", pady=(0, 6))

        row = 1
        for key, info in meta.items():
            if key.startswith("_"):
                continue

            tk.Label(frm, text=info.get("label", key),
                     anchor="w").grid(row=row, column=0, sticky="w", pady=2)

            cur_val = settings[section].get(key, "")
            if info["type"] == "list_str" and isinstance(cur_val, list):
                cur_val = ",".join(str(x) for x in cur_val)
            else:
                cur_val = str(cur_val) if cur_val is not None else ""

            ent = tk.Entry(frm, width=58)
            ent.insert(0, cur_val)
            ent.grid(row=row, column=1, sticky="w", padx=(8, 4), pady=2)
            entries[(section, key)] = (ent, info["type"])

            # path / folder 类型加 "..." 浏览按钮
            if info["type"] in ("path", "folder"):
                def make_browse(e, vtype):
                    def f():
                        if vtype == "folder":
                            p = filedialog.askdirectory(title="选择文件夹")
                        else:
                            p = filedialog.askopenfilename(
                                title="选择模板文件",
                                filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")])
                        if p:
                            e.delete(0, tk.END)
                            e.insert(0, p)
                    return f
                tk.Button(frm, text="...", width=3,
                          command=make_browse(ent, info["type"])).grid(
                            row=row, column=2, padx=2)
            row += 1

    changed_flag = {"val": False}

    def save():
        try:
            new_settings = json.loads(json.dumps(settings))
            for (section, key), (ent, vtype) in entries.items():
                raw = ent.get().strip()
                if vtype == "float":
                    new_settings[section][key] = float(raw) if raw else 0.0
                elif vtype == "int":
                    new_settings[section][key] = int(raw) if raw else 0
                elif vtype == "bool":
                    new_settings[section][key] = raw.lower() in ("true", "1", "yes")
                elif vtype == "list_str":
                    new_settings[section][key] = [
                        x.strip() for x in raw.split(",") if x.strip()
                    ]
                else:
                    new_settings[section][key] = raw
            mgr.save(new_settings)
            changed_flag["val"] = True
            messagebox.showinfo("提示", "参数已保存。\n部分参数下次运行该功能时生效。")
            win.destroy()
        except ValueError as e:
            messagebox.showerror("错误", "值类型错误: " + str(e))

    def restore_default():
        if messagebox.askyesno("确认", "确定恢复所有参数为默认值吗？"):
            mgr.save(DEFAULT_SETTINGS)
            changed_flag["val"] = True
            messagebox.showinfo("提示", "已恢复默认。")
            win.destroy()

    def cancel():
        win.destroy()

    # 按钮区
    btn_frm = tk.Frame(win)
    btn_frm.pack(fill="x", pady=8, padx=10)
    tk.Button(btn_frm, text="恢复默认", width=12,
              command=restore_default).pack(side="left", padx=4)
    tk.Button(btn_frm, text="保存", width=12,
              command=save).pack(side="right", padx=4)
    tk.Button(btn_frm, text="取消", width=12,
              command=cancel).pack(side="right", padx=4)

    win.wait_window()
    return changed_flag["val"]
