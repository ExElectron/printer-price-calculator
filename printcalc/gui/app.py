# -*- coding: utf-8 -*-
"""应用主窗口。"""

import os
import tkinter as tk
from tkinter import messagebox, ttk

from .. import APP_NAME, VERSION
from ..config import config_path, load_config, save_config
from ..converter import default_temp_dir
from ..defaults import DEFAULT_EXCEL_PATH
from ..excel_import import import_from_excel
from .calc_tab import CalcTab
from .common import TaskRunner, pick_font
from .coverage_tab import CoverageTab
from .info_tab import InfoTab
from .settings_tab import SettingsTab


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        first_run = not os.path.isfile(config_path())
        self.cfg = load_config()
        if first_run:
            self._bootstrap_from_excel()
        self.docs = []
        self.quote = None
        self.current_item = None
        self.temp_dir = default_temp_dir()

        self.title("%s  v%s" % (APP_NAME, VERSION))
        self.geometry("1220x780")
        self.minsize(1060, 660)
        self._setup_style()

        self.runner = TaskRunner(self)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        self.calc_tab = CalcTab(self.notebook, self)
        self.coverage_tab = CoverageTab(self.notebook, self)
        self.settings_tab = SettingsTab(self.notebook, self)
        self.info_tab = InfoTab(self.notebook, self)

        self.notebook.add(self.calc_tab, text="  计价  ")
        self.notebook.add(self.coverage_tab, text="  覆盖率明细  ")
        self.notebook.add(self.settings_tab, text="  设置  ")
        self.notebook.add(self.info_tab, text="  公示信息  ")

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.calc_tab.refresh_options()
        self.calc_tab.recalculate()

    # ------------------------------------------------------------------
    def _bootstrap_from_excel(self):
        """首次运行时，用 Excel 公示文件初始化默认数据。"""
        path = self.cfg.get("excel_path") or DEFAULT_EXCEL_PATH
        if not os.path.isfile(path):
            return
        try:
            data = import_from_excel(path)
        except Exception:  # noqa: BLE001
            return
        if not data:
            return
        for key in ("items", "notes", "materials", "equipment"):
            if key in data:
                self.cfg[key] = data[key]
        self.cfg["excel_path"] = path
        try:
            save_config(self.cfg)
        except OSError:
            pass

    def _setup_style(self):
        font = pick_font(self, 9)
        self.option_add("*Font", font)
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
        style.configure("Treeview", rowheight=24, font=font)
        style.configure("Treeview.Heading", font=(font[0], 9, "bold"))
        style.configure("Total.TLabel", font=(font[0], 14, "bold"), foreground="#1f6feb")
        style.configure("Title.TLabel", font=(font[0], 11, "bold"))
        style.configure("Hint.TLabel", foreground="#666666")

    # ------------------------------------------------------------------
    def on_config_changed(self):
        self.calc_tab.refresh_options()
        self.calc_tab.recalculate()
        self.calc_tab.update_engine_hint()
        self.coverage_tab.refresh()
        self.info_tab.refresh()

    def on_docs_changed(self):
        self.calc_tab.refresh_file_list()
        self.calc_tab.recalculate()
        self.coverage_tab.refresh()

    def save(self):
        return save_config(self.cfg)

    # ------------------------------------------------------------------
    def _on_close(self):
        if self.runner.busy:
            if not messagebox.askyesno("退出", "任务仍在运行，确定要退出吗？"):
                return
            self.runner.cancel()
        try:
            if not self.cfg.get("conversion", {}).get("keep_temp", False):
                from ..converter import clean_temp_dir
                clean_temp_dir(self.temp_dir)
        except Exception:  # noqa: BLE001
            pass
        self.destroy()


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    app = App()
    app.mainloop()
