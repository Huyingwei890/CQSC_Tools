# -*- coding: utf-8 -*-
"""
CQSC Batch Tools v2.0 Main
Author: Hu Ying Wei
"""
import os
import sys
import datetime
import importlib
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client as win32
import pythoncom

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from core.log_window import LogWindow
from core.settings_manager import SettingsManager, open_settings_dialog


MODULE_REGISTRY = [
    ("tooling_modify",   "Tooling 批量修改",   "1", "对 Tooling* 工作簿批量修改单元格"),
    ("tooling_rollback", "Tooling 按批次回滚", "2", "按 BatchID 回滚 Tooling 修改"),
    ("roi_combine",      "ROI 合并工具",       "3", "合并多个 CQSC 到 Template"),
    ("cm_version",       "To CM Version",      "4", "生成 +CM 文件并修改 GP/CBD"),
    ("cm_rollback",      "To CM 按批次回滚",   "5", "在 +CM 文件中回滚修改"),
]


# ============ 配色 (深色赛博风) ============
COLOR_BG          = "#0f1729"      # 主背景 深蓝紫
COLOR_CARD        = "#1a2342"      # 卡片背景
COLOR_CARD_HOV    = "#243056"      # 卡片 hover
COLOR_CARD_DIS    = "#161e35"      # 卡片 disabled
COLOR_BORDER      = "#2d3a5f"      # 边框
COLOR_BORDER_HOV  = "#b388ff"      # hover 时的霓虹边框

# 文字
COLOR_TITLE_FG    = "#ffffff"
COLOR_SUB_FG      = "#e1bee7"
COLOR_TEXT_FG     = "#e8eaf6"
COLOR_DESC_FG     = "#8a93b8"
COLOR_DIS_FG      = "#4a5378"

# 强调
COLOR_NEON_PURPLE = "#b388ff"      # 编号霓虹紫
COLOR_NEON_PINK   = "#ff4081"      # 粉
COLOR_NEON_CYAN   = "#00e5ff"      # 青

# 标题渐变三色 (Canvas 画)
GRADIENT_LEFT     = "#7c4dff"      # 紫
GRADIENT_MID      = "#e040fb"      # 品红
GRADIENT_RIGHT    = "#00b0ff"      # 蓝

# 状态栏
COLOR_STATUS_BG   = "#0a1020"
COLOR_STATUS_FG   = "#80deea"

# 字体
FONT_TITLE    = ("Microsoft YaHei UI", 18, "bold")
FONT_SUBTITLE = ("Segoe UI", 9)
FONT_VERSION  = ("Segoe UI", 10, "bold")
FONT_SECTION  = ("Microsoft YaHei UI", 10, "bold")
FONT_HINT     = ("Segoe UI", 20, "bold")
FONT_BTN_TEXT = ("Microsoft YaHei UI", 11, "bold")
FONT_BTN_DESC = ("Microsoft YaHei UI", 8)
FONT_TOOL     = ("Microsoft YaHei UI", 9)
FONT_STATUS   = ("Microsoft YaHei UI", 9)


def load_module(mod_name):
    try:
        return importlib.import_module("modules." + mod_name)
    except ImportError:
        return None


def run_module(mod_name, display_name, settings):
    mod = load_module(mod_name)
    if mod is None:
        messagebox.showinfo("提示",
            "模块 " + mod_name + " 尚未安装。\n请将对应 .py 文件放入 modules/ 文件夹。")
        return

    info = getattr(mod, "MODULE_INFO", {})
    needs_files = info.get("needs_files", True)
    dialog_title = info.get("file_dialog_title", "选择文件 - " + display_name)
    section = info.get("settings_section", mod_name)

    files = ()
    if needs_files:
        root = tk.Tk()
        root.withdraw()
        files = filedialog.askopenfilenames(title=dialog_title,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")])
        root.destroy()
        if not files:
            return

    logw = LogWindow(title="运行日志 - " + display_name)
    logw.log("模块: " + display_name)
    logw.log("文件数: " + str(len(files)), status="启动 Excel ...")

    pythoncom.CoInitialize()
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
    except Exception as e:
        logw.log("[错误] 启动 Excel 失败: " + str(e))
        logw.finish("启动 Excel 失败")
        logw.wait_close()
        pythoncom.CoUninitialize()
        return

    mod_settings = settings.get(section, {})
    mod_settings["_common"] = settings.get("common", {})

    ok = 0
    fail = 0
    try:
        result = mod.run(excel, list(files), mod_settings, logw)
        if isinstance(result, tuple) and len(result) == 2:
            ok, fail = result
    except Exception as e:
        logw.log("[错误] 模块运行异常: " + str(e))
        import traceback
        logw.log(traceback.format_exc())
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    log_dir = os.path.join(ROOT_DIR, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir,
        datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + "_" + mod_name + ".log")
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(logw.lines))
        logw.log("\n日志已保存: " + log_path)
    except Exception as e:
        logw.log("\n日志保存失败: " + str(e))

    logw.log("\n===== 汇总 =====  成功: " + str(ok) + "  失败/跳过: " + str(fail),
             status="完成  成功 " + str(ok) + " / 失败 " + str(fail))
    logw.finish("完成 - 关闭窗口返回主界面")
    logw.wait_close()


def _interpolate_color(c1, c2, t):
    """颜色线性插值, c1/c2 为 #RRGGBB"""
    r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
    r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return "#%02x%02x%02x" % (r, g, b)


class GradientHeader(tk.Canvas):
    """渐变标题栏 (紫 -> 品红 -> 蓝)"""
    def __init__(self, parent, width, height):
        super().__init__(parent, width=width, height=height,
                         highlightthickness=0, bd=0)
        self.w = width
        self.h = height
        self._draw_gradient()
        self._draw_text()

    def _draw_gradient(self):
        # 三色渐变, 横向
        steps = self.w
        for i in range(steps):
            t = i / float(steps - 1)
            if t < 0.5:
                color = _interpolate_color(GRADIENT_LEFT, GRADIENT_MID, t * 2)
            else:
                color = _interpolate_color(GRADIENT_MID, GRADIENT_RIGHT,
                                           (t - 0.5) * 2)
            self.create_line(i, 0, i, self.h, fill=color)

    def _draw_text(self):
        # 主标题
        self.create_text(20, 20, anchor="nw",
                         text="CQSC 批量工具",
                         fill=COLOR_TITLE_FG, font=FONT_TITLE)
        # 副标题
        self.create_text(20, 50, anchor="nw",
                         text="ROI / CM / Tooling 自动化",
                         fill=COLOR_SUB_FG, font=FONT_SUBTITLE)
        # 右上 v2.0 + 作者
        self.create_text(self.w - 20, 22, anchor="ne",
                         text="v2.0",
                         fill=COLOR_TITLE_FG, font=FONT_VERSION)
        self.create_text(self.w - 20, 48, anchor="ne",
                         text="Hu Ying Wei",
                         fill=COLOR_SUB_FG, font=FONT_SUBTITLE)


class ModuleButton(tk.Frame):
    """美化按钮 (卡片)"""
    def __init__(self, parent, hint, title, desc, enabled, command):
        super().__init__(parent, bg=COLOR_CARD if enabled else COLOR_CARD_DIS,
                         highlightbackground=COLOR_BORDER,
                         highlightthickness=1,
                         cursor="hand2" if enabled else "arrow")
        self.enabled = enabled
        self.command = command
        self.normal_bg = COLOR_CARD if enabled else COLOR_CARD_DIS

        # 左侧大编号
        hint_fg = COLOR_NEON_PURPLE if enabled else COLOR_DIS_FG
        self.hint_lbl = tk.Label(self, text=hint, font=FONT_HINT,
                                 fg=hint_fg, bg=self.normal_bg,
                                 width=3)
        self.hint_lbl.pack(side="left", padx=(14, 0), pady=10)

        # 右侧主+副
        right = tk.Frame(self, bg=self.normal_bg)
        right.pack(side="left", fill="both", expand=True, padx=12, pady=10)

        title_fg = COLOR_TEXT_FG if enabled else COLOR_DIS_FG
        title_text = title if enabled else title + "  (未安装)"
        self.title_lbl = tk.Label(right, text=title_text, font=FONT_BTN_TEXT,
                                  fg=title_fg, bg=self.normal_bg, anchor="w")
        self.title_lbl.pack(fill="x")

        desc_fg = COLOR_DESC_FG if enabled else COLOR_DIS_FG
        self.desc_lbl = tk.Label(right, text=desc, font=FONT_BTN_DESC,
                                 fg=desc_fg, bg=self.normal_bg, anchor="w")
        self.desc_lbl.pack(fill="x", pady=(2, 0))

        self.right_frame = right

        if enabled:
            self._bind_events()

    def _bind_events(self):
        for w in (self, self.hint_lbl, self.right_frame,
                  self.title_lbl, self.desc_lbl):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _set_bg(self, bg, border):
        self.configure(bg=bg, highlightbackground=border)
        self.hint_lbl.configure(bg=bg)
        self.right_frame.configure(bg=bg)
        self.title_lbl.configure(bg=bg)
        self.desc_lbl.configure(bg=bg)

    def _on_enter(self, e):
        self._set_bg(COLOR_CARD_HOV, COLOR_BORDER_HOV)

    def _on_leave(self, e):
        self._set_bg(self.normal_bg, COLOR_BORDER)

    def _on_click(self, e):
        if self.command:
            self.command()


class ToolButton(tk.Frame):
    """工具栏按钮"""
    def __init__(self, parent, text, command):
        super().__init__(parent, bg=COLOR_CARD,
                         highlightbackground=COLOR_BORDER,
                         highlightthickness=1,
                         cursor="hand2")
        self.command = command
        self.lbl = tk.Label(self, text=text, font=FONT_TOOL,
                            fg=COLOR_TEXT_FG, bg=COLOR_CARD,
                            padx=10, pady=8)
        self.lbl.pack(fill="both", expand=True)

        for w in (self, self.lbl):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, e):
        self.configure(bg=COLOR_CARD_HOV,
                       highlightbackground=COLOR_NEON_CYAN)
        self.lbl.configure(bg=COLOR_CARD_HOV, fg=COLOR_NEON_CYAN)

    def _on_leave(self, e):
        self.configure(bg=COLOR_CARD, highlightbackground=COLOR_BORDER)
        self.lbl.configure(bg=COLOR_CARD, fg=COLOR_TEXT_FG)

    def _on_click(self, e):
        if self.command:
            self.command()


class MainWindow:
    def __init__(self):
        self.settings_mgr = SettingsManager(ROOT_DIR)
        self.settings = self.settings_mgr.load()

        self.root = tk.Tk()
        self.root.title("CQSC 批量工具 v2.0")
        WIN_W = 500
        WIN_H = 760
        self.root.geometry(str(WIN_W) + "x" + str(WIN_H))
        self.root.resizable(False, False)
        self.root.configure(bg=COLOR_BG)

        # 渐变标题
        header = GradientHeader(self.root, WIN_W, 80)
        header.pack(fill="x")

        # 主体
        body = tk.Frame(self.root, bg=COLOR_BG, padx=18, pady=14)
        body.pack(fill="both", expand=True)

        # ============ 功能区 ============
        section_lbl = tk.Frame(body, bg=COLOR_BG)
        section_lbl.pack(fill="x", pady=(0, 8))
        tk.Label(section_lbl, text="▎",
                 fg=COLOR_NEON_PINK, bg=COLOR_BG,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        tk.Label(section_lbl, text="选择功能",
                 font=FONT_SECTION,
                 fg=COLOR_TEXT_FG, bg=COLOR_BG,
                 anchor="w").pack(side="left", padx=(2, 0))

        for mod_name, display_name, hint, desc in MODULE_REGISTRY:
            installed = load_module(mod_name) is not None
            btn = ModuleButton(
                body, hint=hint, title=display_name, desc=desc,
                enabled=installed,
                command=(lambda m=mod_name, d=display_name:
                         self.run_module(m, d)) if installed else None)
            btn.pack(fill="x", pady=4)

        # ============ 分隔 ============
        sep = tk.Frame(body, height=1, bg=COLOR_BORDER)
        sep.pack(fill="x", pady=(14, 10))

        # ============ 工具区 ============
        tool_lbl = tk.Frame(body, bg=COLOR_BG)
        tool_lbl.pack(fill="x", pady=(0, 6))
        tk.Label(tool_lbl, text="▎",
                 fg=COLOR_NEON_CYAN, bg=COLOR_BG,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        tk.Label(tool_lbl, text="工具",
                 font=FONT_SECTION,
                 fg=COLOR_TEXT_FG, bg=COLOR_BG,
                 anchor="w").pack(side="left", padx=(2, 0))

        tools_frame = tk.Frame(body, bg=COLOR_BG)
        tools_frame.pack(fill="x")

        tools = [
            ("⚙   参数设置", self.open_settings),
            ("📁   打开日志文件夹", self.open_logs_folder),
            ("ℹ    关于", self.show_about),
        ]
        for text, cmd in tools:
            tb = ToolButton(tools_frame, text=text, command=cmd)
            tb.pack(fill="x", pady=3)

        # ============ 状态栏 ============
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Frame(self.root, bg=COLOR_STATUS_BG, height=26)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)
        tk.Label(status_bar, text="●",
                 fg=COLOR_NEON_CYAN, bg=COLOR_STATUS_BG,
                 font=("Segoe UI", 10)).pack(
                 side="left", padx=(12, 4))
        tk.Label(status_bar, textvariable=self.status_var,
                 anchor="w", font=FONT_STATUS,
                 fg=COLOR_STATUS_FG, bg=COLOR_STATUS_BG).pack(
                 side="left", fill="y")

    def run_module(self, mod_name, display_name):
        self.status_var.set("执行中: " + display_name)
        self.root.withdraw()
        try:
            run_module(mod_name, display_name, self.settings)
        finally:
            self.root.deiconify()
            self.status_var.set("就绪")

    def open_settings(self):
        self.status_var.set("打开参数设置 ...")
        changed = open_settings_dialog(self.root, self.settings_mgr)
        if changed:
            self.settings = self.settings_mgr.load()
            self.status_var.set("参数已更新")
        else:
            self.status_var.set("就绪")

    def open_logs_folder(self):
        log_dir = os.path.join(ROOT_DIR, "logs")
        os.makedirs(log_dir, exist_ok=True)
        try:
            os.startfile(log_dir)
        except Exception as e:
            messagebox.showerror("错误", "打开日志文件夹失败: " + str(e))

    def show_about(self):
        msg = ("CQSC 批量工具 v2.0\n\n"
               "作者: Hu Ying Wei\n"
               "用途: ROI / CM / Tooling 批量自动化\n\n"
               "配置文件: settings.json\n"
               "日志文件: logs/\n"
               "模块文件夹: modules/\n\n"
               "如需新增功能, 把对应 .py 放入 modules/ 即可。")
        messagebox.showinfo("关于", msg)

    def run(self):
        self.root.mainloop()


def main():
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()