# -*- coding: utf-8 -*-
"""墨水覆盖率计算。

算法与 https://github.com/hgjazhgj/PrinterPageCoverage 保持一致：
将页面渲染为位图，取所有通道像素值的平均值，按 A4 面积归一化。

    coverage = (1 - mean(pixels)/255) * 实际像素数 / (A4面积(inch) * dpi^2)

因此满版纯黑 A4 页面覆盖率 = 1.0（100%），纯白 = 0。
"""

import os

import numpy

try:
    import pymupdf
except ImportError:  # pragma: no cover
    import fitz as pymupdf

A4_WIDTH_IN = 8.268
A4_HEIGHT_IN = 11.693
A4_AREA_IN2 = A4_WIDTH_IN * A4_HEIGHT_IN


def page_coverage(page, dpi=300):
    """返回单个 PDF 页面的墨水覆盖率（以 A4 满版为 1.0）。"""
    pix = page.get_pixmap(dpi=dpi)
    buffer = getattr(pix, "samples_mv", None)
    if buffer is None:
        buffer = pix.samples
    array = numpy.frombuffer(buffer, dtype=numpy.uint8)
    if array.size == 0:
        return 0.0
    mean = float(array.mean())
    return (1.0 - mean / 255.0) * pix.height * pix.width / (A4_AREA_IN2 * dpi * dpi)


class Cancelled(Exception):
    pass


def analyze_pdf(path, dpi=300, progress=None, should_cancel=None):
    """分析 PDF，返回 (页数, 每页覆盖率列表)。"""
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    doc = pymupdf.open(path)
    try:
        pages = doc.page_count
        results = []
        for index in range(pages):
            if should_cancel and should_cancel():
                raise Cancelled()
            results.append(page_coverage(doc[index], dpi))
            if progress:
                progress(index + 1, pages)
        return pages, results
    finally:
        doc.close()


def render_page_png(path, page_index, dpi=72):
    """把某一页渲染成 PNG 字节，用于界面预览。"""
    doc = pymupdf.open(path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            return None
        pix = doc[page_index].get_pixmap(dpi=dpi)
        return pix.tobytes("png")
    finally:
        doc.close()
