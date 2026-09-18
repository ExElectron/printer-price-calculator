# -*- coding: utf-8 -*-
"""覆盖率明细页：逐页覆盖率与页面预览。"""

import io
import os
import tkinter as tk
from tkinter import ttk

from .. import config as cfgmod
from .. import coverage
from .common import PAD, treeview_with_scroll

DOC_COLUMNS = ("文件名", "页数", "平均覆盖率", "总覆盖率")
PAGE_COLUMNS = ("页码", "覆盖率", "状态")


class CoverageTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self._photo = None
        self.current_doc = None
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 4))
        ttk.Label(toolbar, text="逐页查看墨水覆盖率（以满版 A4 纯黑为 100%）",
                  style="Title.TLabel").pack(side="left")
        self.dpi_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self.dpi_var, style="Hint.TLabel").pack(side="right")

        pane = ttk.PanedWindow(self, orient="horizontal")
        pane.grid(row=1, column=0, sticky="nsew", padx=PAD, pady=4)

        left = ttk.LabelFrame(pane, text="已分析文档")
        frame, self.doc_tree = treeview_with_scroll(
            left, DOC_COLUMNS, heights=16, widths=(170, 45, 85, 85),
            anchors=("w", "center", "e", "e"))
        frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.doc_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_doc_selected())
        pane.add(left, weight=2)

        middle = ttk.LabelFrame(pane, text="页面覆盖率")
        frame2, self.page_tree = treeview_with_scroll(
            middle, PAGE_COLUMNS, heights=16, widths=(55, 85, 110),
            anchors=("center", "e", "w"))
        frame2.pack(fill="both", expand=True, padx=6, pady=6)
        self.page_tree.tag_configure("over", foreground="#c0392b")
        self.page_tree.tag_configure("ok", foreground="#1e8449")
        self.page_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_page_selected())
        pane.add(middle, weight=2)

        right = ttk.LabelFrame(pane, text="页面预览")
        self.canvas = tk.Canvas(right, width=340, height=520, background="#e9ecef",
                                highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, padx=6, pady=6)
        self.preview_info = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.preview_info, style="Hint.TLabel",
                  wraplength=320, justify="left").pack(fill="x", padx=6, pady=(0, 6))
        pane.add(right, weight=3)

    # ------------------------------------------------------------------
    def refresh(self):
        docs = [d for d in self.app.docs if d.ready]
        selection = self.current_doc.path if self.current_doc else None
        for item in self.doc_tree.get_children():
            self.doc_tree.delete(item)
        for index, doc in enumerate(docs):
            self.doc_tree.insert("", "end", iid=str(index), values=(
                doc.name,
                doc.pages,
                "%.2f%%" % (doc.avg_coverage * 100),
                "%.3f" % doc.total_coverage,
            ))
        self.dpi_var.set("计算 DPI：%s" % self.app.cfg.get("coverage", {}).get("dpi", 300))
        if selection:
            for index, doc in enumerate(docs):
                if doc.path == selection:
                    self.doc_tree.selection_set(str(index))
                    break
        elif docs:
            self.doc_tree.selection_set("0")
        else:
            self.page_tree.delete(*self.page_tree.get_children())
            self.current_doc = None
            self.canvas.delete("all")
            self.preview_info.set("")

    def select_current_doc(self):
        docs = [d for d in self.app.docs if d.ready]
        if not docs:
            return
        selected = self.doc_tree.selection()
        if not selected:
            self.doc_tree.selection_set("0")

    # ------------------------------------------------------------------
    def _docs(self):
        return [d for d in self.app.docs if d.ready]

    def _on_doc_selected(self):
        selection = self.doc_tree.selection()
        docs = self._docs()
        if not selection:
            return
        try:
            self.current_doc = docs[int(selection[0])]
        except (ValueError, IndexError):
            self.current_doc = None
        self._fill_pages()

    def _fill_pages(self):
        self.page_tree.delete(*self.page_tree.get_children())
        doc = self.current_doc
        if doc is None:
            self.canvas.delete("all")
            return
        limit = cfgmod.item_limit(self.app.current_item) if self.app.current_item else None
        for index, value in enumerate(doc.coverages):
            over = limit is not None and value > limit
            status = "超出上限" if over else ("未超限" if limit is not None else "无上限")
            self.page_tree.insert("", "end", iid=str(index), values=(
                index + 1, "%.3f%%" % (value * 100), status,
            ), tags=("over" if over else "ok",))
        if doc.coverages:
            self.page_tree.selection_set("0")

    def _on_page_selected(self):
        doc = self.current_doc
        selection = self.page_tree.selection()
        if doc is None or not selection:
            return
        try:
            page_index = int(selection[0])
        except (TypeError, ValueError):
            return
        self._render(doc, page_index)

    def _render(self, doc, page_index):
        try:
            data = coverage.render_page_png(doc.pdf_path, page_index, dpi=72)
        except Exception as exc:  # noqa: BLE001
            self.preview_info.set("预览失败：%s" % exc)
            return
        self.canvas.delete("all")
        if not data:
            self.preview_info.set("无法渲染该页面")
            return
        try:
            from PIL import Image, ImageTk
        except ImportError:
            self.preview_info.set("缺少 Pillow，无法显示预览")
            return
        image = Image.open(io.BytesIO(data))
        canvas_w = max(self.canvas.winfo_width(), 120)
        canvas_h = max(self.canvas.winfo_height(), 120)
        image.thumbnail((canvas_w - 12, canvas_h - 12))
        self._photo = ImageTk.PhotoImage(image)
        self.canvas.create_image(canvas_w // 2, canvas_h // 2, image=self._photo, anchor="center")
        value = doc.coverages[page_index] if page_index < len(doc.coverages) else 0.0
        limit = cfgmod.item_limit(self.app.current_item) if self.app.current_item else None
        extra = ""
        if limit is not None and value > limit:
            extra = "（超出上限 %.1f%%，超出 %.3f）" % (limit * 100, value - limit)
        elif limit is not None:
            extra = "（上限内）"
        self.preview_info.set("%s 第 %d 页，墨水覆盖率 %.3f%%%s"
                              % (doc.name, page_index + 1, value * 100, extra))
