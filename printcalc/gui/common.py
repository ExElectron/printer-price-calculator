# -*- coding: utf-8 -*-
"""界面通用组件与辅助函数。"""

import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk

PAD = 8


def pick_font(root, size=9):
    from tkinter import font as tkfont
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimHei", "SimSun"):
        if name in families:
            return (name, size)
    return ("Segoe UI", size)


def center_window(win, width, height, parent=None):
    win.update_idletasks()
    if parent is not None:
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    else:
        x = (win.winfo_screenwidth() - width) // 2
        y = (win.winfo_screenheight() - height) // 2
    win.geometry("%dx%d+%d+%d" % (max(width, 1), max(height, 1), max(x, 0), max(y, 0)))


class ScrollableFrame(ttk.Frame):
    """带垂直滚动条的容器，内容放在 .inner。"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        for widget in (self.canvas, self.inner):
            widget.bind("<Enter>", self._bind_wheel)
            widget.bind("<Leave>", self._unbind_wheel)

    def _on_inner_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)

    def _bind_wheel(self, _event):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _unbind_wheel(self, _event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class TaskRunner:
    """在后台线程运行任务，并把结果回调回 Tk 主线程。"""

    def __init__(self, widget):
        self.widget = widget
        self.queue = queue.Queue()
        self._cancel = threading.Event()
        self.busy = False
        self.widget.after(80, self._poll)

    def submit(self, work, on_done=None, on_error=None, on_progress=None):
        if self.busy:
            messagebox.showinfo("请稍候", "已有任务正在运行，请等待完成。")
            return False
        self._cancel.clear()
        self.busy = True

        def emit(done, total, text=""):
            self.queue.put(("progress", on_progress, (done, total, text)))

        def runner():
            try:
                result = work(emit, self._cancel.is_set)
                self.queue.put(("done", on_done, result))
            except Exception as exc:  # noqa: BLE001
                self.queue.put(("error", on_error, exc))

        threading.Thread(target=runner, daemon=True).start()
        return True

    def cancel(self):
        self._cancel.set()

    def is_cancelled(self):
        return self._cancel.is_set()

    def _poll(self):
        try:
            while True:
                kind, callback, payload = self.queue.get_nowait()
                if kind == "progress":
                    if callback:
                        callback(*payload)
                    continue
                self.busy = False
                if kind == "done":
                    if callback:
                        callback(payload)
                else:
                    if callback:
                        callback(payload)
                    else:
                        messagebox.showerror("出错了", str(payload))
        except queue.Empty:
            pass
        self.widget.after(80, self._poll)


def labeled_entry(parent, label, variable, width=18, row=None, column=0, **kwargs):
    box = ttk.Frame(parent)
    ttk.Label(box, text=label).pack(side="left")
    entry = ttk.Entry(box, textvariable=variable, width=width, **kwargs)
    entry.pack(side="left", padx=(6, 0))
    return box, entry


def treeview_with_scroll(parent, columns, heights=10, widths=None, anchors=None):
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=heights)
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    for index, name in enumerate(columns):
        tree.heading(name, text=name)
        width = (widths[index] if widths and index < len(widths) else 110)
        anchor = (anchors[index] if anchors and index < len(anchors) else "center")
        tree.column(name, width=width, anchor=anchor, stretch=True)
    return frame, tree
