# -*- coding: utf-8 -*-
"""设置页：所有可调数据都集中在此页。"""

import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .. import config as cfgmod
from .. import converter, pricing
from ..defaults import SIDES_DOUBLE, SIDES_SINGLE
from ..excel_import import ExcelImportError, import_from_excel
from .common import PAD, ScrollableFrame, center_window

ITEM_COLUMNS = ("纸张介质", "规格", "打印方式", "色彩模式", "面数", "单价(元/张)", "覆盖率上限", "状态", "计费说明")
ITEM_WIDTHS = (120, 190, 80, 80, 80, 90, 90, 70, 240)


class ItemDialog(tk.Toplevel):
    def __init__(self, master, item=None):
        super().__init__(master)
        self.title("编辑价目行" if item else "新增价目行")
        self.resizable(False, False)
        self.transient(master)
        self.result = None

        item = item or {}
        self.vars = {
            "paper": tk.StringVar(value=item.get("paper", "")),
            "spec": tk.StringVar(value=item.get("spec", "")),
            "method": tk.StringVar(value=item.get("method", "")),
            "color": tk.StringVar(value=item.get("color", "")),
            "price": tk.StringVar(value="" if item.get("price") in (None, "") else str(item.get("price"))),
            "limit": tk.StringVar(value="" if item.get("limit") in (None, "") else str(float(item["limit"]) * 100)),
            "note": tk.StringVar(value=item.get("note", "")),
        }
        sides = item.get("sides") or [SIDES_SINGLE]
        self.single_var = tk.BooleanVar(value=SIDES_SINGLE in sides)
        self.double_var = tk.BooleanVar(value=SIDES_DOUBLE in sides)
        self.available_var = tk.BooleanVar(value=bool(item.get("available", True)))

        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)
        rows = [
            ("纸张介质", "paper"), ("规格尺寸/克重", "spec"),
            ("打印方式", "method"), ("色彩模式", "color"),
            ("单价（元/张，留空=暂不提供）", "price"),
            ("覆盖率上限（%，留空=不设上限）", "limit"),
            ("计费说明", "note"),
        ]
        for index, (label, key) in enumerate(rows):
            ttk.Label(body, text=label).grid(row=index, column=0, sticky="w", pady=3)
            ttk.Entry(body, textvariable=self.vars[key], width=32).grid(
                row=index, column=1, sticky="ew", pady=3)
        row = len(rows)
        ttk.Label(body, text="打印面数").grid(row=row, column=0, sticky="w", pady=3)
        sides_box = ttk.Frame(body)
        sides_box.grid(row=row, column=1, sticky="w")
        ttk.Checkbutton(sides_box, text="单面", variable=self.single_var).pack(side="left")
        ttk.Checkbutton(sides_box, text="双面", variable=self.double_var).pack(side="left", padx=8)
        ttk.Checkbutton(body, text="提供该项服务", variable=self.available_var).grid(
            row=row + 1, column=1, sticky="w", pady=3)
        body.columnconfigure(1, weight=1)

        buttons = ttk.Frame(self, padding=(12, 0, 12, 12))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="确定", command=self._ok).pack(side="right")
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=6)
        self.bind("<Return>", lambda _e: self._ok())
        self.bind("<Escape>", lambda _e: self.destroy())
        center_window(self, 420, 340, master)
        self.grab_set()

    def _ok(self):
        paper = self.vars["paper"].get().strip()
        if not paper:
            messagebox.showwarning("提示", "请填写纸张介质。", parent=self)
            return
        sides = []
        if self.single_var.get():
            sides.append(SIDES_SINGLE)
        if self.double_var.get():
            sides.append(SIDES_DOUBLE)
        if not sides:
            messagebox.showwarning("提示", "请至少选择一种打印面数。", parent=self)
            return

        price_text = self.vars["price"].get().strip()
        limit_text = self.vars["limit"].get().strip()
        try:
            price = float(price_text) if price_text else None
            limit = (float(limit_text) / 100.0) if limit_text else None
        except ValueError:
            messagebox.showwarning("提示", "单价与覆盖率上限必须是数字。", parent=self)
            return

        self.result = {
            "paper": paper,
            "spec": self.vars["spec"].get().strip(),
            "method": self.vars["method"].get().strip(),
            "color": self.vars["color"].get().strip(),
            "sides": sides,
            "price": price,
            "limit": limit,
            "note": self.vars["note"].get().strip(),
            "available": bool(self.available_var.get()) and price is not None,
        }
        self.destroy()


class SettingsTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.vars = {}
        self.item_rows = []
        self._build()
        self.load_from_config()

    # ------------------------------------------------------------------
    def _build(self):
        self.scroll = ScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True)
        root = self.scroll.inner
        root.columnconfigure(0, weight=1)

        self._build_pricing(root)
        self._build_items(root)
        self._build_conversion(root)
        self._build_defaults(root)
        self._build_data(root)

    # ------------------------------------------------------------------
    def _section(self, parent, title, row):
        box = ttk.LabelFrame(parent, text=title)
        box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(PAD, 0))
        box.columnconfigure(2, weight=1)
        return box

    def _add_row(self, box, row, label, widget, hint=""):
        ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=3)
        widget.grid(row=row, column=1, sticky="w", padx=6, pady=3)
        if hint:
            ttk.Label(box, text=hint, style="Hint.TLabel").grid(
                row=row, column=2, sticky="w", padx=6)

    # ------------------------------------------------------------------
    def _build_pricing(self, root):
        box = self._section(root, "计价参数", 0)

        self.vars["overage_factor"] = tk.StringVar()
        self.vars["limit_mode"] = tk.StringVar()
        self.vars["overage_basis"] = tk.StringVar()
        self.vars["overage_custom_basis"] = tk.StringVar()
        self.vars["round_digits"] = tk.StringVar()

        mode_box = ttk.Combobox(box, textvariable=self.vars["limit_mode"], state="readonly",
                                values=list(pricing.MODE_LABELS.values()))
        basis_box = ttk.Combobox(box, textvariable=self.vars["overage_basis"], state="readonly",
                                 values=list(pricing.BASIS_LABELS.values()))
        self._add_row(box, 0, "超标加收系数", ttk.Entry(box, textvariable=self.vars["overage_factor"]),
                      "总价 + 超出折算页数 × 单价 × 系数（默认 0.6）")
        self._add_row(box, 1, "超标判定方式", mode_box)
        self._add_row(box, 2, "超标覆盖折算基准", basis_box)
        self._add_row(box, 3, "自定义折算基准（%）",
                      ttk.Entry(box, textvariable=self.vars["overage_custom_basis"]),
                      "选择“自定义折算基准”时生效，如 5 表示按 5% 一页折算")
        self._add_row(box, 4, "金额小数位", ttk.Entry(box, textvariable=self.vars["round_digits"]))

    # ------------------------------------------------------------------
    def _build_items(self, root):
        box = ttk.LabelFrame(root, text="价目表（纸张 / 打印方式 / 色彩 / 面数 / 单价 / 覆盖率上限）")
        box.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(PAD, 0))
        box.columnconfigure(0, weight=1)

        holder = ttk.Frame(box)
        holder.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)

        from .common import treeview_with_scroll
        frame, self.item_tree = treeview_with_scroll(
            holder, ITEM_COLUMNS, heights=12, widths=ITEM_WIDTHS,
            anchors=("w", "w", "center", "center", "center", "e", "e", "center", "w"))
        frame.grid(row=0, column=0, sticky="nsew")
        self.item_tree.bind("<Double-1>", lambda _e: self.edit_item())

        buttons = ttk.Frame(box)
        buttons.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))
        ttk.Button(buttons, text="新增", command=self.add_item).pack(side="left")
        ttk.Button(buttons, text="编辑", command=self.edit_item).pack(side="left", padx=4)
        ttk.Button(buttons, text="删除", command=self.delete_item).pack(side="left")
        ttk.Button(buttons, text="上移", command=lambda: self.move_item(-1)).pack(side="left", padx=4)
        ttk.Button(buttons, text="下移", command=lambda: self.move_item(1)).pack(side="left")

    # ------------------------------------------------------------------
    def _build_conversion(self, root):
        box = self._section(root, "文档转 PDF", 2)
        self.vars["engine"] = tk.StringVar()
        self.vars["acrobat_dir"] = tk.StringVar()
        self.vars["office_dir"] = tk.StringVar()
        self.vars["image_dpi"] = tk.StringVar()
        self.vars["timeout_seconds"] = tk.StringVar()
        self.keep_temp_var = tk.BooleanVar()

        engine_box = ttk.Combobox(box, textvariable=self.vars["engine"], state="readonly",
                                  values=[label for _value, label in converter.ENGINES])
        self._add_row(box, 0, "转换引擎", engine_box)
        self._add_row(box, 1, "Acrobat 目录", self._dir_row(box, "acrobat_dir"))
        self._add_row(box, 2, "Office 目录", self._dir_row(box, "office_dir"))
        self._add_row(box, 3, "图片默认 DPI（Pillow 引擎）",
                      ttk.Entry(box, textvariable=self.vars["image_dpi"]))
        self._add_row(box, 4, "转换超时（秒）", ttk.Entry(box, textvariable=self.vars["timeout_seconds"]))
        ttk.Checkbutton(box, text="保留转换产生的临时 PDF", variable=self.keep_temp_var).grid(
            row=5, column=1, sticky="w", padx=6, pady=3)

        self.vars["dpi"] = tk.StringVar()
        dpi_box = ttk.Combobox(box, textvariable=self.vars["dpi"], state="readonly",
                               values=("100", "150", "200", "300", "400", "600"))
        self._add_row(box, 6, "覆盖率计算 DPI", dpi_box,
                      "DPI 越高越精确，速度越慢（推荐 300）")

    def _dir_row(self, parent, key):
        holder = ttk.Frame(parent)
        ttk.Entry(holder, textvariable=self.vars[key], width=64).pack(side="left")
        ttk.Button(holder, text="浏览…", width=8,
                   command=lambda k=key: self._browse_dir(k)).pack(side="left", padx=(6, 0))
        return holder

    def _browse_dir(self, key):
        initial = self.vars[key].get() or os.path.expanduser("~")
        folder = filedialog.askdirectory(title="选择目录", initialdir=initial)
        if folder:
            self.vars[key].set(os.path.normpath(folder))

    # ------------------------------------------------------------------
    def _build_defaults(self, root):
        box = self._section(root, "默认打印选项", 3)
        self.vars["default_paper"] = tk.StringVar()
        self.vars["default_method"] = tk.StringVar()
        self.vars["default_color"] = tk.StringVar()
        self.vars["default_sides"] = tk.StringVar()
        self.vars["default_copies"] = tk.StringVar()
        self.default_combos = {}
        for index, (label, key) in enumerate([
            ("默认纸张介质", "default_paper"), ("默认打印方式", "default_method"),
            ("默认色彩模式", "default_color"), ("默认打印面数", "default_sides"),
        ]):
            combo = ttk.Combobox(box, textvariable=self.vars[key], state="readonly")
            self.default_combos[key] = combo
            self._add_row(box, index, label, combo)
        self._add_row(box, 4, "默认份数", ttk.Entry(box, textvariable=self.vars["default_copies"]))

    # ------------------------------------------------------------------
    def _build_data(self, root):
        box = self._section(root, "数据来源与备份", 4)
        self.vars["excel_path"] = tk.StringVar()
        holder = ttk.Frame(box)
        ttk.Entry(holder, textvariable=self.vars["excel_path"], width=64).pack(side="left")
        ttk.Button(holder, text="浏览…", width=8,
                   command=self._browse_excel).pack(side="left", padx=(6, 0))
        self._add_row(box, 0, "价格公示 Excel", holder)

        buttons = ttk.Frame(box)
        ttk.Button(buttons, text="从 Excel 导入", command=self.import_excel).pack(side="left")
        ttk.Button(buttons, text="恢复内置默认", command=self.reset_defaults).pack(side="left", padx=4)
        ttk.Button(buttons, text="导出设置…", command=self.export_settings).pack(side="left")
        ttk.Button(buttons, text="导入设置…", command=self.import_settings).pack(side="left", padx=4)
        buttons.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        save_box = ttk.Frame(root)
        save_box.grid(row=5, column=0, sticky="ew", padx=PAD, pady=PAD)
        self.save_btn = ttk.Button(save_box, text="保存设置并应用", command=self.save_settings)
        self.save_btn.pack(side="right")
        ttk.Button(save_box, text="放弃修改", command=self.load_from_config).pack(side="right", padx=6)
        self.status_var = tk.StringVar(value="")
        ttk.Label(save_box, textvariable=self.status_var, style="Hint.TLabel").pack(side="left")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择价格公示 Excel", filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.vars["excel_path"].set(path)

    # ------------------------------------------------------------------
    def load_from_config(self):
        cfg = self.app.cfg
        pr = cfg.get("pricing", {})
        conv = cfg.get("conversion", {})
        defaults = cfg.get("defaults", {})
        self.vars["overage_factor"].set(str(pr.get("overage_factor", 0.6)))
        self.vars["limit_mode"].set(pricing.MODE_LABELS.get(pr.get("limit_mode", pricing.MODE_PER_PAGE),
                                                            list(pricing.MODE_LABELS.values())[0]))
        self.vars["overage_basis"].set(pricing.BASIS_LABELS.get(pr.get("overage_basis", pricing.BASIS_LIMIT),
                                                                list(pricing.BASIS_LABELS.values())[0]))
        self.vars["overage_custom_basis"].set(str(pr.get("overage_custom_basis", 0.05) * 100))
        self.vars["round_digits"].set(str(pr.get("round_digits", 2)))
        self.vars["dpi"].set(str(cfg.get("coverage", {}).get("dpi", 300)))
        self.vars["engine"].set(dict((v, l) for v, l in converter.ENGINES).get(
            conv.get("engine", "acrobat"), converter.ENGINES[0][1]))
        self.vars["acrobat_dir"].set(conv.get("acrobat_dir", ""))
        self.vars["office_dir"].set(conv.get("office_dir", ""))
        self.vars["image_dpi"].set(str(conv.get("image_dpi", 96)))
        self.vars["timeout_seconds"].set(str(conv.get("timeout_seconds", 300)))
        self.keep_temp_var.set(bool(conv.get("keep_temp", False)))
        self.vars["default_paper"].set(defaults.get("paper", ""))
        self.vars["default_method"].set(defaults.get("method", ""))
        self.vars["default_color"].set(defaults.get("color", ""))
        self.vars["default_sides"].set(defaults.get("sides", ""))
        self.vars["default_copies"].set(str(defaults.get("copies", 1)))
        self.vars["excel_path"].set(cfg.get("excel_path", ""))
        self.item_rows = [dict(it) for it in cfg.get("items", [])]
        self._refresh_default_options()
        self.refresh_items()
        self.status_var.set("")

    def _refresh_default_options(self):
        items = self.item_rows
        self.default_combos["default_paper"].configure(values=cfgmod.paper_list(items))
        self.default_combos["default_method"].configure(values=sorted({it["method"] for it in items if it["method"]}))
        self.default_combos["default_color"].configure(values=sorted({it["color"] for it in items if it["color"]}))
        self.default_combos["default_sides"].configure(values=["单面", "双面"])

    # ------------------------------------------------------------------
    def refresh_items(self):
        self.item_tree.delete(*self.item_tree.get_children())
        for index, item in enumerate(self.item_rows):
            price = cfgmod.item_price(item)
            limit = cfgmod.item_limit(item)
            self.item_tree.insert("", "end", iid=str(index), values=(
                item.get("paper", ""),
                item.get("spec", ""),
                item.get("method", ""),
                item.get("color", ""),
                "/".join(item.get("sides", [])),
                "%.2f" % price if price is not None else "—",
                "不设上限" if limit is None else "%.1f%%" % (limit * 100),
                "提供" if cfgmod.item_available(item) else "暂不提供",
                item.get("note", ""),
            ))

    def _selected_item_index(self):
        selection = self.item_tree.selection()
        if not selection:
            return None
        try:
            return int(selection[0])
        except (TypeError, ValueError):
            return None

    def add_item(self):
        dialog = ItemDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            self.item_rows.append(dialog.result)
            self.refresh_items()
            self._refresh_default_options()
            self.item_tree.selection_set(str(len(self.item_rows) - 1))

    def edit_item(self):
        index = self._selected_item_index()
        if index is None:
            messagebox.showinfo("提示", "请先选择一行。")
            return
        dialog = ItemDialog(self, self.item_rows[index])
        self.wait_window(dialog)
        if dialog.result:
            self.item_rows[index] = dialog.result
            self.refresh_items()
            self._refresh_default_options()
            self.item_tree.selection_set(str(index))

    def delete_item(self):
        index = self._selected_item_index()
        if index is None:
            messagebox.showinfo("提示", "请先选择一行。")
            return
        if not messagebox.askyesno("删除", "确定删除选中的价目行吗？"):
            return
        self.item_rows.pop(index)
        self.refresh_items()
        self._refresh_default_options()

    def move_item(self, delta):
        index = self._selected_item_index()
        if index is None:
            return
        target = index + delta
        if target < 0 or target >= len(self.item_rows):
            return
        self.item_rows[index], self.item_rows[target] = self.item_rows[target], self.item_rows[index]
        self.refresh_items()
        self.item_tree.selection_set(str(target))

    # ------------------------------------------------------------------
    def _collect(self):
        cfg = self.app.cfg
        try:
            factor = float(self.vars["overage_factor"].get())
            custom = float(self.vars["overage_custom_basis"].get()) / 100.0
            digits = int(self.vars["round_digits"].get())
            image_dpi = int(self.vars["image_dpi"].get())
            timeout = int(self.vars["timeout_seconds"].get())
            dpi = int(self.vars["dpi"].get())
            copies = max(1, int(self.vars["default_copies"].get()))
        except ValueError:
            messagebox.showwarning("提示", "单价系数、DPI、小数位等必须是数字。")
            return False

        mode = {label: value for value, label in pricing.MODE_LABELS.items()}[self.vars["limit_mode"].get()]
        basis = {label: value for value, label in pricing.BASIS_LABELS.items()}[self.vars["overage_basis"].get()]
        engine = {label: value for value, label in converter.ENGINES}[self.vars["engine"].get()]

        cfg["pricing"] = {
            "overage_factor": factor,
            "limit_mode": mode,
            "overage_basis": basis,
            "overage_custom_basis": custom,
            "round_digits": digits,
        }
        cfg["coverage"] = {"dpi": dpi}
        cfg["conversion"] = {
            "engine": engine,
            "acrobat_dir": self.vars["acrobat_dir"].get().strip(),
            "office_dir": self.vars["office_dir"].get().strip(),
            "image_dpi": image_dpi,
            "keep_temp": bool(self.keep_temp_var.get()),
            "timeout_seconds": timeout,
        }
        cfg["defaults"] = {
            "paper": self.vars["default_paper"].get(),
            "method": self.vars["default_method"].get(),
            "color": self.vars["default_color"].get(),
            "sides": self.vars["default_sides"].get(),
            "copies": copies,
        }
        cfg["excel_path"] = self.vars["excel_path"].get().strip()
        for idx, item in enumerate(self.item_rows, start=1):
            item["id"] = idx
        cfg["items"] = [dict(it) for it in self.item_rows]
        return True

    def save_settings(self):
        if not self._collect():
            return
        try:
            path = cfgmod.save_config(self.app.cfg)
        except OSError as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self.status_var.set("已保存：%s" % path)
        self.app.on_config_changed()

    def reset_defaults(self):
        if not messagebox.askyesno("恢复默认", "将丢弃当前所有设置，恢复为内置默认数据。继续？"):
            return
        self.app.cfg = cfgmod.reset_config()
        self.load_from_config()
        self.save_settings()
        self.status_var.set("已恢复内置默认设置")

    def import_excel(self):
        path = self.vars["excel_path"].get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请先选择有效的 Excel 文件。")
            return
        try:
            data = import_from_excel(path)
        except ExcelImportError as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        if not data:
            messagebox.showwarning("提示", "未能从该 Excel 中解析出数据。")
            return
        self.app.cfg["excel_path"] = path
        for key in ("items", "notes", "materials", "equipment"):
            if key in data:
                self.app.cfg[key] = data[key]
        self.load_from_config()
        self.save_settings()
        self.status_var.set("已从 Excel 导入 %d 条价目" % len(data.get("items", [])))

    def export_settings(self):
        path = filedialog.asksaveasfilename(
            title="导出设置", defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")], initialfile="printcalc_settings.json")
        if not path:
            return
        if not self._collect():
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(self.app.cfg, fh, ensure_ascii=False, indent=2)
        except OSError as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.status_var.set("设置已导出：%s" % path)

    def import_settings(self):
        path = filedialog.askopenfilename(
            title="导入设置", filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, ValueError) as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        if not isinstance(data, dict):
            messagebox.showerror("导入失败", "文件格式不正确。")
            return
        from ..config import _deep_merge
        from ..defaults import default_config
        merged = default_config()
        _deep_merge(merged, data)
        if not merged.get("items"):
            merged["items"] = list(default_config()["items"])
        self.app.cfg = merged
        self.load_from_config()
        self.save_settings()
        self.status_var.set("设置已导入")
