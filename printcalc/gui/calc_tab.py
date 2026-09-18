# -*- coding: utf-8 -*-
"""计价页。"""

import csv
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .. import config as cfgmod
from .. import converter, coverage, pricing
from .common import PAD, treeview_with_scroll

FILE_COLUMNS = ("文件名", "类型", "页数", "平均覆盖率", "最高页覆盖率", "超标覆盖", "状态")
FILE_WIDTHS = (300, 70, 60, 100, 100, 100, 200)
FILE_ANCHORS = ("w", "center", "center", "e", "e", "e", "w")


class CalcTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.line_by_path = {}
        self._build()

    # ------------------------------------------------------------------
    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 4))
        ttk.Button(toolbar, text="添加文件", command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="添加文件夹", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(toolbar, text="移除所选", command=self.remove_selected).pack(side="left")
        ttk.Button(toolbar, text="清空", command=self.clear_files).pack(side="left", padx=4)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
        self.analyze_btn = ttk.Button(toolbar, text="转换并分析", command=self.start_analysis)
        self.analyze_btn.pack(side="left")
        self.cancel_btn = ttk.Button(toolbar, text="取消", command=self.app.runner.cancel,
                                     state="disabled")
        self.cancel_btn.pack(side="left", padx=4)
        ttk.Button(toolbar, text="重新计算", command=self.recalculate).pack(side="left", padx=4)
        ttk.Button(toolbar, text="导出报价单", command=self.export_csv).pack(side="left")
        ttk.Button(toolbar, text="导出转换后的 PDF", command=self.export_pdfs).pack(side="left", padx=4)

        main = ttk.Frame(self)
        main.grid(row=1, column=0, sticky="nsew", padx=PAD, pady=4)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)

        left = ttk.LabelFrame(main, text="待打印文档")
        left.grid(row=0, column=0, sticky="nsew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        frame, self.file_tree = treeview_with_scroll(
            left, FILE_COLUMNS, heights=14, widths=FILE_WIDTHS, anchors=FILE_ANCHORS)
        frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.file_tree.tag_configure("over", foreground="#c0392b")
        self.file_tree.tag_configure("error", foreground="#b9770e")
        self.file_tree.tag_configure("ok", foreground="#1e8449")
        self.file_tree.bind("<Double-1>", self._open_coverage_tab)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsw", padx=(PAD, 0))
        self._build_options(right)
        self._build_summary(right)

        status = ttk.Frame(self)
        status.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(0, PAD))
        status.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(status, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status, textvariable=self.status_var, style="Hint.TLabel").grid(row=0, column=1, padx=8)

    # ------------------------------------------------------------------
    def _build_options(self, parent):
        box = ttk.LabelFrame(parent, text="打印选项")
        box.pack(fill="x")

        self.paper_var = tk.StringVar()
        self.method_var = tk.StringVar()
        self.color_var = tk.StringVar()
        self.sides_var = tk.StringVar()
        self.copies_var = tk.IntVar(value=1)

        rows = [
            ("纸张介质", self.paper_var, "paper_combo"),
            ("打印方式", self.method_var, "method_combo"),
            ("色彩模式", self.color_var, "color_combo"),
            ("打印面数", self.sides_var, "sides_combo"),
        ]
        for index, (label, var, attr) in enumerate(rows):
            ttk.Label(box, text=label).grid(row=index, column=0, sticky="w", padx=6, pady=3)
            combo = ttk.Combobox(box, textvariable=var, state="readonly", width=20)
            combo.grid(row=index, column=1, sticky="ew", padx=6, pady=3)
            setattr(self, attr, combo)

        ttk.Label(box, text="打印份数").grid(row=4, column=0, sticky="w", padx=6, pady=3)
        spin = ttk.Spinbox(box, from_=1, to=9999, textvariable=self.copies_var, width=8,
                           command=self.recalculate)
        spin.grid(row=4, column=1, sticky="w", padx=6, pady=3)
        spin.bind("<FocusOut>", lambda _e: self.recalculate())
        spin.bind("<Return>", lambda _e: self.recalculate())

        self.paper_combo.bind("<<ComboboxSelected>>", lambda _e: self._cascade("paper"))
        self.method_combo.bind("<<ComboboxSelected>>", lambda _e: self._cascade("method"))
        self.color_combo.bind("<<ComboboxSelected>>", lambda _e: self._cascade("color"))
        self.sides_combo.bind("<<ComboboxSelected>>", lambda _e: self.recalculate())

        ttk.Separator(box).grid(row=5, column=0, columnspan=2, sticky="ew", pady=6)

        self.spec_var = tk.StringVar(value="—")
        self.price_var = tk.StringVar(value="—")
        self.limit_var = tk.StringVar(value="—")
        self.factor_var = tk.StringVar(value="—")
        for index, (label, var) in enumerate([
            ("规格", self.spec_var), ("单价", self.price_var),
            ("覆盖率上限", self.limit_var), ("加收阶梯", self.factor_var),
        ], start=6):
            ttk.Label(box, text=label).grid(row=index, column=0, sticky="w", padx=6, pady=2)
            style = "Tier.TLabel" if label == "加收阶梯" else "Title.TLabel"
            value_label = ttk.Label(box, textvariable=var, style=style,
                                    wraplength=180, justify="left")
            value_label.grid(row=index, column=1, sticky="w", padx=6, pady=2)
        box.columnconfigure(1, weight=1)

        self.engine_hint = ttk.Label(parent, text="", style="Hint.TLabel", wraplength=260,
                                     justify="left")
        self.engine_hint.pack(fill="x", pady=(6, 0))

    # ------------------------------------------------------------------
    def _build_summary(self, parent):
        box = ttk.LabelFrame(parent, text="报价汇总")
        box.pack(fill="x", pady=(PAD, 0))
        self.summary_vars = {
            "files": tk.StringVar(value="0"),
            "pages": tk.StringVar(value="0"),
            "faces": tk.StringVar(value="0"),
            "sheets": tk.StringVar(value="0"),
            "base": tk.StringVar(value="0.00"),
            "surcharge": tk.StringVar(value="0.00"),
            "total": tk.StringVar(value="0.00"),
        }
        labels = [
            ("文件数", "files"), ("总页数", "pages"), ("打印面", "faces"), ("打印张数", "sheets"),
            ("基础价", "base"), ("超标加收", "surcharge"),
        ]
        for index, (label, key) in enumerate(labels):
            ttk.Label(box, text=label).grid(row=index, column=0, sticky="w", padx=6, pady=2)
            ttk.Label(box, textvariable=self.summary_vars[key]).grid(
                row=index, column=1, sticky="e", padx=6, pady=2)
        ttk.Label(box, text="合计").grid(row=len(labels), column=0, sticky="w", padx=6, pady=(6, 4))
        ttk.Label(box, textvariable=self.summary_vars["total"], style="Total.TLabel").grid(
            row=len(labels), column=1, sticky="e", padx=6, pady=(6, 4))
        box.columnconfigure(1, weight=1)

        self.notes_text = tk.Text(parent, height=7, width=34, wrap="word", relief="flat",
                                  background="#f5f6f8", highlightthickness=0)
        self.notes_text.pack(fill="x", pady=(6, 0))
        self.notes_text.configure(state="disabled")

    # ------------------------------------------------------------------
    def update_engine_hint(self):
        conv = self.app.cfg.get("conversion", {})
        engine = conv.get("engine", "acrobat")
        names = dict(converter.ENGINES)
        text = "转换引擎：%s" % names.get(engine, engine)
        if engine in ("acrobat", "auto"):
            acrobat = os.path.join(conv.get("acrobat_dir", ""), "Acrobat.exe")
            text += "\nAcrobat：%s" % (acrobat if os.path.isfile(acrobat) else "未找到（将尝试回退）")
        self.engine_hint.configure(text=text)

    # ------------------------------------------------------------------
    def _set_combo(self, combo, values, preferred=None):
        current = preferred if preferred is not None else combo.get()
        combo.configure(values=values)
        if current in values:
            combo.set(current)
        elif values:
            combo.set(values[0])
        else:
            combo.set("")

    def refresh_options(self):
        items = self.app.cfg["items"]
        defaults = self.app.cfg.get("defaults", {})

        papers = cfgmod.paper_list(items)
        self._set_combo(self.paper_combo, papers, defaults.get("paper"))
        paper = self.paper_var.get()
        self._set_combo(self.method_combo, cfgmod.method_list(items, paper), defaults.get("method"))
        method = self.method_var.get()
        self._set_combo(self.color_combo, cfgmod.color_list(items, paper, method), defaults.get("color"))
        color = self.color_var.get()
        self._set_combo(self.sides_combo, cfgmod.sides_list(items, paper, method, color), defaults.get("sides"))
        self.spec_var.set(cfgmod.spec_of(items, paper) or "—")
        self.update_engine_hint()
        self.recalculate()

    def _cascade(self, level):
        items = self.app.cfg["items"]
        paper, method, color = self.paper_var.get(), self.method_var.get(), self.color_var.get()
        if level == "paper":
            self._set_combo(self.method_combo, cfgmod.method_list(items, paper))
            method = self.method_var.get()
            self.spec_var.set(cfgmod.spec_of(items, paper) or "—")
        if level in ("paper", "method"):
            self._set_combo(self.color_combo, cfgmod.color_list(items, paper, method))
            color = self.color_var.get()
        self._set_combo(self.sides_combo, cfgmod.sides_list(items, paper, method, color))
        self.recalculate()

    # ------------------------------------------------------------------
    def current_selection(self):
        items = self.app.cfg["items"]
        paper = self.paper_var.get()
        method = self.method_var.get()
        color = self.color_var.get()
        sides = self.sides_var.get()
        item = cfgmod.find_item(items, paper, method, color, sides) if paper else None
        return item, sides, max(1, int(self.copies_var.get() or 1))

    # ------------------------------------------------------------------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要打印的文档",
            filetypes=[
                ("支持的文档", " ".join("*" + e for e in converter.SUPPORTED_EXTS)),
                ("所有文件", "*.*"),
            ])
        self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return
        paths = []
        for root, _dirs, files in os.walk(folder):
            for name in files:
                if os.path.splitext(name)[1].lower() in converter.SUPPORTED_EXTS:
                    paths.append(os.path.join(root, name))
        if paths:
            self._add_paths(paths)
        else:
            messagebox.showinfo("提示", "该文件夹中没有找到支持的文档。")

    def _add_paths(self, paths):
        existing = {os.path.abspath(doc.path) for doc in self.app.docs}
        added = 0
        for path in paths:
            full = os.path.abspath(path)
            if full in existing:
                continue
            ext = os.path.splitext(full)[1].lower()
            doc = pricing.DocResult(path=full, name=os.path.basename(full), ext=ext)
            self.app.docs.append(doc)
            existing.add(full)
            added += 1
        if added:
            self.app.on_docs_changed()
            self.status_var.set("已添加 %d 个文档，点击“转换并分析”开始计算" % added)

    def remove_selected(self):
        selected = self._selected_paths()
        if not selected:
            return
        self.app.docs = [d for d in self.app.docs if d.path not in selected]
        self.app.on_docs_changed()

    def clear_files(self):
        if not self.app.docs:
            return
        self.app.docs = []
        self.app.quote = None
        self.app.on_docs_changed()

    def _selected_paths(self):
        paths = []
        for item in self.file_tree.selection():
            try:
                idx = int(item)
            except (TypeError, ValueError):
                continue
            if 0 <= idx < len(self.app.docs):
                paths.append(self.app.docs[idx].path)
        return paths

    def _open_coverage_tab(self, _event=None):
        self.app.notebook.select(self.app.coverage_tab)
        self.app.coverage_tab.select_current_doc()

    # ------------------------------------------------------------------
    def start_analysis(self):
        if not self.app.docs:
            messagebox.showinfo("提示", "请先添加文档。")
            return
        docs = [d for d in self.app.docs if not d.ready]
        if not docs:
            docs = list(self.app.docs)
        self._analyze(docs)

    def _analyze(self, docs):
        cfg = self.app.cfg
        dpi = int(cfg.get("coverage", {}).get("dpi", 300) or 300)
        temp_dir = self.app.temp_dir
        self.analyze_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        def work(emit, cancel):
            os.makedirs(temp_dir, exist_ok=True)
            emit(0, len(docs), "开始转换文档…")
            results = converter.convert_many(
                [d.path for d in docs], temp_dir, cfg,
                progress=lambda done, total, text: emit(done, total, "转换：" + text),
                should_cancel=cancel)
            by_src = {r.src: r for r in results}
            for index, doc in enumerate(docs, 1):
                if cancel():
                    break
                result = by_src.get(doc.path)
                if result and result.ok:
                    doc.pdf_path = result.pdf_path
                    doc.engine = result.engine
                    doc.converted = True
                    emit(index, len(docs), "分析：%s" % doc.name)
                    try:
                        pages, covs = coverage.analyze_pdf(doc.pdf_path, dpi, should_cancel=cancel)
                        doc.pages = pages
                        doc.coverages = covs
                        doc.analyzed = True
                        doc.error = ""
                    except Exception as exc:  # noqa: BLE001
                        doc.analyzed = False
                        doc.error = "覆盖率分析失败：%s" % exc
                else:
                    doc.converted = False
                    doc.analyzed = False
                    doc.error = (result.error if result else "转换失败")
            return docs

        def on_progress(done, total, text):
            self.progress.configure(maximum=max(total, 1), value=done)
            self.status_var.set(text)

        def on_done(_):
            cancelled = self.app.runner.is_cancelled()
            self.analyze_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
            self.progress.configure(value=0)
            self.status_var.set("已取消" if cancelled else "分析完成")
            self.app.on_docs_changed()

        def on_error(exc):
            self.analyze_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
            self.progress.configure(value=0)
            self.status_var.set("任务失败：%s" % exc)
            messagebox.showerror("任务失败", str(exc))

        self.app.runner.submit(work, on_done=on_done, on_error=on_error, on_progress=on_progress)

    # ------------------------------------------------------------------
    def refresh_file_list(self):
        selection = self.file_tree.selection()
        selected_index = self.file_tree.index(selection[0]) if selection else None
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        lines = self.line_by_path
        for index, doc in enumerate(self.app.docs):
            line = lines.get(doc.path)
            if doc.error:
                status = "失败：%s" % doc.error
                tag = "error"
            elif doc.ready:
                engine = doc.engine or "pdf"
                status = "已分析（%s）" % engine
                tag = "ok"
            else:
                status = "待分析"
                tag = ""
            over = ""
            if line and line.excess_coverage > 0:
                over = "%.3f" % line.excess_coverage
                tag = "over"
            self.file_tree.insert("", "end", iid=str(index), values=(
                doc.name,
                doc.ext.lstrip("."),
                doc.pages if doc.pages else "—",
                ("%.2f%%" % (doc.avg_coverage * 100)) if doc.ready else "—",
                ("%.2f%%" % (doc.max_coverage * 100)) if doc.ready else "—",
                over or "—",
                status,
            ), tags=(tag,) if tag else ())
        if selected_index is not None and str(selected_index) in self.file_tree.get_children():
            self.file_tree.selection_set(str(selected_index))

    def recalculate(self):
        item, sides, copies = self.current_selection()
        self.app.current_item = item
        self.line_by_path = {}

        if item is None:
            self.spec_var.set("—")
            self.price_var.set("暂不提供")
            self.limit_var.set("—")
            self.factor_var.set("—")
            quote = None
        else:
            price = cfgmod.item_price(item)
            limit = cfgmod.item_limit(item)
            self.spec_var.set(item.get("spec") or "—")
            self.price_var.set("暂不提供" if price is None else "%.2f 元/张" % price)
            self.limit_var.set("不设上限" if limit is None else "%.1f%%" % (limit * 100))
            tiers = pricing.normalize_tiers(
                self.app.cfg.get("pricing", {}).get("overage_tiers"),
                self.app.cfg.get("pricing", {}).get("overage_factor", 0.6))
            self.factor_var.set(pricing.tier_summary(tiers, limit))

            ready = [d for d in self.app.docs if d.ready]
            if price is None or not ready:
                quote = None
                if price is None:
                    self.price_var.set("暂不提供")
            else:
                quote = pricing.compute_quote(ready, item, sides, copies, self.app.cfg.get("pricing", {}))
                for line in quote.lines:
                    self.line_by_path[line.doc.path] = line

        self.app.quote = quote
        self._update_summary(quote)
        self.refresh_file_list()

    def _update_summary(self, quote):
        digits = int(self.app.cfg.get("pricing", {}).get("round_digits", 2) or 2)
        if quote is None:
            for key in ("files", "pages", "faces", "sheets"):
                self.summary_vars[key].set("0")
            for key in ("base", "surcharge", "total"):
                self.summary_vars[key].set(pricing.format_money(0, digits))
            self._set_notes(["请选择有单价的价目行并完成分析。"])
            return
        self.summary_vars["files"].set(str(len(quote.lines)))
        self.summary_vars["pages"].set(str(quote.pages))
        self.summary_vars["faces"].set(str(quote.faces))
        self.summary_vars["sheets"].set(str(quote.sheets))
        self.summary_vars["base"].set(pricing.format_money(quote.base, digits))
        self.summary_vars["surcharge"].set(pricing.format_money(quote.surcharge, digits))
        self.summary_vars["total"].set(pricing.format_money(quote.total, digits))
        self._set_notes(quote.notes)

    def _set_notes(self, lines):
        self.notes_text.configure(state="normal")
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", "\n".join(lines))
        self.notes_text.configure(state="disabled")

    # ------------------------------------------------------------------
    def export_pdfs(self):
        converted = [d for d in self.app.docs if d.analyzed and d.pdf_path and os.path.isfile(d.pdf_path)]
        if not converted:
            messagebox.showinfo("提示", "还没有转换好的 PDF，请先点击“转换并分析”。")
            return
        folder = filedialog.askdirectory(title="选择导出 PDF 的文件夹")
        if not folder:
            return
        import shutil
        used = set()
        count = 0
        for doc in converted:
            name = os.path.splitext(doc.name)[0] + ".pdf"
            target = os.path.join(folder, name)
            index = 1
            while target in used or os.path.exists(target):
                target = os.path.join(folder, "%s(%d).pdf" % (os.path.splitext(doc.name)[0], index))
                index += 1
            try:
                shutil.copyfile(doc.pdf_path, target)
                used.add(target)
                count += 1
            except OSError as exc:
                messagebox.showerror("导出失败", "%s：%s" % (doc.name, exc))
                return
        self.status_var.set("已导出 %d 个 PDF 到 %s" % (count, folder))

    # ------------------------------------------------------------------
    def export_csv(self):
        if not self.app.quote or not self.app.quote.lines:
            messagebox.showinfo("提示", "还没有可导出的报价结果。")
            return
        path = filedialog.asksaveasfilename(
            title="导出报价单", defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")], initialfile="打印报价单.csv")
        if not path:
            return
        quote = self.app.quote
        pricing_cfg = self.app.cfg.get("pricing", {})
        digits = int(pricing_cfg.get("round_digits", 2) or 2)
        divisor = pricing.divisor_for(
            quote.basis, quote.limit,
            float(pricing_cfg.get("overage_custom_basis", 0.05) or 0.05))
        limit_text = "不设上限" if quote.limit is None else "%.1f%%" % (quote.limit * 100)

        def pct(value, digits_=2):
            return "%.*f%%" % (digits_, (value or 0.0) * 100)

        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(["PrintCalc 打印报价单"])
                writer.writerow(["生成时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                writer.writerow([])
                writer.writerow(["【打印方案】"])
                writer.writerow(["纸张介质", quote.item.get("paper", "")])
                writer.writerow(["规格", quote.item.get("spec", "")])
                writer.writerow(["打印方式", quote.item.get("method", "")])
                writer.writerow(["色彩模式", quote.item.get("color", "")])
                writer.writerow(["打印面数", quote.sides])
                writer.writerow(["打印份数", quote.copies])
                writer.writerow(["单价（元/张）", "%.2f" % quote.unit_price])
                writer.writerow(["覆盖率上限", limit_text])
                writer.writerow(["超标判定", pricing.MODE_LABELS.get(quote.mode, quote.mode)])
                writer.writerow(["超标加收阶梯", pricing.tier_summary(quote.tiers, quote.limit)])
                writer.writerow(["超标覆盖折算基准", pricing.BASIS_LABELS.get(quote.basis, quote.basis)])

                writer.writerow([])
                writer.writerow(["【报价汇总】"])
                writer.writerow(["文件数", len(quote.lines)])
                writer.writerow(["总页数", quote.pages])
                writer.writerow(["打印面", quote.faces])
                writer.writerow(["打印张数", quote.sheets])
                writer.writerow(["总覆盖率", "%.3f" % quote.total_coverage])
                writer.writerow(["平均覆盖率", pct(quote.avg_coverage)])
                writer.writerow(["超出覆盖量", "%.4f" % quote.excess_coverage])
                writer.writerow(["加权覆盖量", "%.4f" % quote.weighted_coverage])
                writer.writerow(["折算页数", "%.2f" % quote.extra_pages])
                writer.writerow(["实际平均加收系数", "%.3f" % quote.effective_factor])
                writer.writerow(["基础价", pricing.format_money(quote.base, digits)])
                writer.writerow(["超标加收", pricing.format_money(quote.surcharge, digits)])
                writer.writerow(["应付合计", pricing.format_money(quote.total, digits)])

                writer.writerow([])
                writer.writerow(["【各文件汇总】"])
                writer.writerow(["文件", "页数", "打印面", "张数", "平均覆盖率", "最高页覆盖率",
                                 "超出覆盖量", "折算页数", "实际平均系数", "基础价", "超标加收", "小计"])
                for line in quote.lines:
                    writer.writerow([
                        line.doc.name, line.pages, line.faces, line.sheets,
                        pct(line.avg_coverage), pct(line.doc.max_coverage),
                        "%.4f" % line.excess_coverage,
                        "%.2f" % line.extra_pages,
                        "%.3f" % line.effective_factor,
                        pricing.format_money(line.base, digits),
                        pricing.format_money(line.surcharge, digits),
                        pricing.format_money(line.total, digits),
                    ])

                writer.writerow([])
                writer.writerow(["【覆盖量明细（逐页）】"])
                if quote.mode == pricing.MODE_AGGREGATE:
                    writer.writerow(["说明：当前为“整体判定”，下表逐页金额仅供参考，实际按总量结算。"])
                writer.writerow(["文件", "页码", "覆盖率", "覆盖率上限", "超出量",
                                 "该页适用系数", "折算页数", "该页加收(元)"])
                for line in quote.lines:
                    rows = pricing.page_details(line.doc, quote.limit, quote.tiers,
                                                divisor, quote.unit_price, quote.copies)
                    for row in rows:
                        writer.writerow([
                            line.doc.name,
                            row["page"],
                            pct(row["coverage"], 3),
                            limit_text,
                            "%.4f" % row["excess"],
                            "—" if row["factor"] is None else "%.3f" % row["factor"],
                            "%.2f" % row["extra_pages"],
                            pricing.format_money(row["surcharge"], digits),
                        ])
        except OSError as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.status_var.set("报价单已导出：%s" % path)
