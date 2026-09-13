# -*- coding: utf-8 -*-
"""实时日志窗口"""

import datetime
import time
import tkinter as tk
from tkinter import scrolledtext


class LogWindow:
    def __init__(self, title="运行日志"):
        # 如果已有 Tk 根窗口就用 Toplevel，否则新建 Tk
        if tk._default_root is not None:
            self.root = tk.Toplevel()
        else:
            self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("880x580")

        top = tk.Frame(self.root)
        top.pack(fill="x", padx=8, pady=4)
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(top, textvariable=self.status_var, anchor="w",
                 font=("微软雅黑", 10, "bold")).pack(side="left")

        self.text = scrolledtext.ScrolledText(
            self.root, wrap="word", font=("Consolas", 10))
        self.text.pack(fill="both", expand=True, padx=8, pady=6)

        bot = tk.Frame(self.root)
        bot.pack(fill="x", padx=8, pady=4)
        self.close_btn = tk.Button(bot, text="关闭并返回主界面", width=20,
                                   command=self._do_close, state="disabled")
        self.close_btn.pack(side="right")

        self.lines = []
        self._closed = False

        # 拦截右上角 X：未完成时禁止关闭
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        self.root.update()

    def log(self, msg, status=None):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = "[" + ts + "] " + str(msg)
        self.lines.append(line)
        try:
            self.text.insert("end", line + "\n")
            self.text.see("end")
            if status:
                self.status_var.set(status)
            self.root.update()
        except tk.TclError:
            pass

    def finish(self, status="完成"):
        try:
            self.status_var.set(status)
            self.close_btn.config(state="normal")
            self.root.update()
        except tk.TclError:
            pass

    def _on_window_close(self):
        # 处理中不允许关
        if str(self.close_btn["state"]) != "normal":
            return
        self._do_close()

    def _do_close(self):
        self._closed = True
        try:
            self.root.destroy()
        except Exception:
            pass

    def wait_close(self):
        try:
            while not self._closed:
                self.root.update()
                time.sleep(0.05)
        except (tk.TclError, KeyboardInterrupt):
            pass