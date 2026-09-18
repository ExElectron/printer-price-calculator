# -*- coding: utf-8 -*-
"""公示信息页：服务须知、耗材公示、设备软件。"""

import tkinter as tk
from tkinter import ttk

from .common import PAD, treeview_with_scroll

MATERIAL_COLUMNS = ("序号", "耗材分类", "物品名称", "品牌/厂商", "规格型号/核心参数", "包装规格", "主要用途/适用场景")
MATERIAL_WIDTHS = (50, 90, 170, 130, 230, 100, 240)

EQUIPMENT_COLUMNS = ("序号", "资产类别", "厂商/品牌", "设备/软件名称", "版本型号/规格参数")
EQUIPMENT_WIDTHS = (50, 90, 170, 190, 220)


class InfoTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        notes_frame = ttk.Frame(notebook)
        self.notes_text = tk.Text(notes_frame, wrap="word", relief="flat", padx=12, pady=12,
                                  font=("Microsoft YaHei UI", 10), background="#fbfbfc")
        self.notes_text.pack(fill="both", expand=True)
        self.notes_text.configure(state="disabled")
        notebook.add(notes_frame, text="  服务须知与计价规则  ")

        material_frame = ttk.Frame(notebook)
        frame, self.material_tree = treeview_with_scroll(
            material_frame, MATERIAL_COLUMNS, heights=18, widths=MATERIAL_WIDTHS,
            anchors=("center", "center", "w", "w", "w", "center", "w"))
        frame.pack(fill="both", expand=True, padx=6, pady=6)
        notebook.add(material_frame, text="  耗材公示  ")

        equipment_frame = ttk.Frame(notebook)
        frame2, self.equipment_tree = treeview_with_scroll(
            equipment_frame, EQUIPMENT_COLUMNS, heights=18, widths=EQUIPMENT_WIDTHS,
            anchors=("center", "center", "w", "w", "w"))
        frame2.pack(fill="both", expand=True, padx=6, pady=6)
        notebook.add(equipment_frame, text="  设备与软件公示  ")

        self.refresh()

    def refresh(self):
        cfg = self.app.cfg
        self.notes_text.configure(state="normal")
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", "文印服务须知与计价规则说明\n\n")
        for line in cfg.get("notes", []):
            self.notes_text.insert("end", line + "\n\n")
        self.notes_text.configure(state="disabled")

        self.material_tree.delete(*self.material_tree.get_children())
        for item in cfg.get("materials", []):
            self.material_tree.insert("", "end", values=(
                item.get("no", ""), item.get("category", ""), item.get("name", ""),
                item.get("brand", ""), item.get("spec", ""), item.get("pack", ""),
                item.get("usage", ""),
            ))

        self.equipment_tree.delete(*self.equipment_tree.get_children())
        for item in cfg.get("equipment", []):
            self.equipment_tree.insert("", "end", values=(
                item.get("no", ""), item.get("category", ""), item.get("brand", ""),
                item.get("name", ""), item.get("spec", ""),
            ))
