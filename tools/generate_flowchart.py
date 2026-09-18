# -*- coding: utf-8 -*-
"""生成《打印计价算法》Graph LR 流程图（PDF + PNG + Mermaid 源码）。

流程严格对应程序实现（printcalc/converter.py、coverage.py、pricing.py）。
默认参数取自《打印价格暂定方案和材料公示.xlsx》。

用法：
    python tools/generate_flowchart.py [输出目录]
"""

import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from printcalc import defaults as D  # noqa: E402
from printcalc.excel_import import import_from_excel  # noqa: E402

# --------------------------------------------------------------------------
# 图定义：rank 为列号（0 起），shape ∈ rect / diamond / stadium / terminator
# --------------------------------------------------------------------------

NODES = [
    ("start", 0, "stadium", "开始：添加待打印文档"),
    ("topdf", 1, "rect",
     "① 文档 → PDF\n· 已是 PDF：直接读取原文件，不转换\n· 其他格式：Adobe Acrobat 转换\n"
     "· 失败回退：Office → Pillow → 文本排版"),
    ("cover", 2, "rect",
     "② 逐页渲染位图并计算覆盖率\n默认 300 DPI（可调 100~600）\n"
     "覆盖率 c =(1 − 平均像素值 ÷ 255)\n            × 像素数 ÷(A4面积 × DPI²)\n"
     "A4 = 8.268 × 11.693 in²\n满版 A4 纯黑 = 100%，纯白 = 0%"),
    ("opt", 3, "rect",
     "③ 选择打印方案\n纸张介质 / 打印方式 /\n色彩模式 / 打印面数 / 打印份数"),
    ("d_item", 4, "diamond", "④ 价目表匹配？"),
    ("price", 5, "rect",
     "⑤ 取单价 p（元/张）\n与覆盖率上限 L\n（该介质无上限时 L 为空）"),
    ("na", 5, "rect", "暂不提供\n不计价（结束）"),
    ("sheets", 6, "rect",
     "⑥ 打印张数\n单面 = 页数\n双面 = 页数 ÷ 2\n（向上取整，再乘份数）\n基础价 = 张数 × 单价 p"),
    ("d_mode", 7, "diamond", "⑦ 超标判定方式"),
    ("pp", 8, "rect", "逐页判定（默认）\n每页单独考核\n存在 c > L 的页"),
    ("agg", 8, "rect", "整体判定\n按全单平均覆盖率考核\n平均覆盖率 > L"),
    ("ladder", 9, "rect",
     "⑧ 分段阶梯累进递减\n超出上限的覆盖量按区间拆分：\n· 上限 ~ 25%   × 0.60\n"
     "· 25% ~ 45%    × 0.40\n· 45% 以上      × 0.20\n"
     "加权覆盖量 = Σ(各档超出量 × 系数)"),
    ("fee", 10, "rect",
     "⑨ 超标加收金额\n加收 = 加权覆盖量\n        ÷ 折算基准 × 单价 p\n"
     "折算基准：按覆盖率上限(默认)\n/ 按整页 100% / 自定义\n（L 为空则加收 = 0）"),
    ("settle", 11, "rect",
     "⑩ 应收合计\n应收 = 基础价 + 超标加收\n打印完成后录入实收金额\n"
     "优惠 = 应收合计 − 实收金额"),
    ("finish", 12, "stadium", "结束：导出报价单\n（含逐页覆盖量明细）"),
]

EDGES = [
    ("start", "topdf", ""),
    ("topdf", "cover", ""),
    ("cover", "opt", ""),
    ("opt", "d_item", ""),
    ("d_item", "price", "是"),
    ("d_item", "na", "否"),
    ("price", "sheets", ""),
    ("sheets", "d_mode", ""),
    ("d_mode", "pp", "逐页"),
    ("d_mode", "agg", "整体"),
    ("pp", "ladder", ""),
    ("agg", "ladder", ""),
    ("ladder", "fee", ""),
    ("fee", "settle", ""),
    ("settle", "finish", ""),
]

# 分带：每带内一律自左向右阅读，带间用 U 形回线连接，逻辑顺序清晰
STAGE_BANDS = [
    ("阶段一　文档处理与覆盖率计算", [0, 1, 2, 3]),
    ("阶段二　选型、查表与基础价", [4, 5, 6]),
    ("阶段三　超标判定与阶梯加收", [7, 8, 9, 10]),
    ("阶段四　结算与输出", [11, 12]),
]

STAGE_COLORS = ["#1f6feb", "#0a8f6a", "#d68910", "#8e44ad"]

LEVEL_FONT_SIZE = 8.6
GAP_X = 44
GAP_Y = 10
BAND_GAP = 58
WRAP = 250
DIAMOND_WRAP = 120
MAX_BAND_WIDTH = 1560


class FNode:
    def __init__(self, key, rank, shape, text):
        self.key = key
        self.rank = rank
        self.shape = shape
        self.text = text
        self.lines = []
        self.w = 0.0
        self.h = 0.0
        self.x = 0.0
        self.y = 0.0


def register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for regular, rindex, bold, bindex in [
        (r"C:\Windows\Fonts\msyh.ttc", 0, r"C:\Windows\Fonts\msyhbd.ttc", 0),
        (r"C:\Windows\Fonts\simhei.ttf", None, r"C:\Windows\Fonts\simhei.ttf", None),
    ]:
        try:
            pdfmetrics.registerFont(TTFont("CJK", regular) if rindex is None
                                    else TTFont("CJK", regular, subfontIndex=rindex))
            pdfmetrics.registerFont(TTFont("CJK-Bold", bold) if bindex is None
                                    else TTFont("CJK-Bold", bold, subfontIndex=bindex))
            return "CJK", "CJK-Bold"
        except Exception:  # noqa: BLE001
            continue
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    return "STSong-Light", "STSong-Light"


BREAK_AFTER = set("，。：；！？、（）《》【】“”—…·%×÷≤≥①②③④⑤⑥⑦⑧⑨⑩⌈⌉·/")


def wrap_text(text, font, size, max_width):
    from reportlab.pdfbase import pdfmetrics
    out = []
    for paragraph in str(text).split("\n"):
        if not paragraph:
            out.append("")
            continue
        tokens = []
        token = ""
        for ch in paragraph:
            if ("\u4e00" <= ch <= "\u9fff") or ch in BREAK_AFTER:
                if token:
                    tokens.append(token)
                    token = ""
                tokens.append(ch)
            elif ch == " ":
                if token:
                    tokens.append(token)
                    token = ""
                tokens.append(" ")
            else:
                token += ch
        if token:
            tokens.append(token)

        current = ""
        for item in tokens:
            if item == " " and not current:
                continue
            if pdfmetrics.stringWidth(current + item, font, size) <= max_width or not current:
                current += item
            else:
                out.append(current.rstrip())
                current = "" if item == " " else item
        out.append(current.rstrip())
    return out


def measure(node, font):
    from reportlab.pdfbase import pdfmetrics
    width = DIAMOND_WRAP if node.shape == "diamond" else WRAP
    node.lines = wrap_text(node.text, font, LEVEL_FONT_SIZE, width)
    text_w = max((pdfmetrics.stringWidth(line, font, LEVEL_FONT_SIZE) for line in node.lines),
                 default=0.0)
    leading = LEVEL_FONT_SIZE * 1.34
    pad_x = 14 if node.shape != "diamond" else 26
    pad_y = 11 if node.shape != "diamond" else 14
    if node.shape == "diamond":
        node.w = text_w * 1.35 + pad_x * 2
        node.h = len(node.lines) * leading + pad_y * 2
    else:
        node.w = text_w + pad_x * 2
        node.h = len(node.lines) * leading + pad_y * 2





# --------------------------------------------------------------------------
# 绘制
# --------------------------------------------------------------------------

def shade(hex_color, factor):
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % (int(r + (255 - r) * factor),
                              int(g + (255 - g) * factor),
                              int(b + (255 - b) * factor))


def draw_node(c, node, font, bold):
    from reportlab.lib.colors import HexColor
    style = "term" if node.shape in ("stadium", "terminator") else "process"
    if style == "term":
        fill, stroke, text_color = "#12355b", "#12355b", "#ffffff"
    elif node.shape == "diamond":
        fill, stroke, text_color = "#fff5e6", "#d68910", "#5b3b00"
    else:
        fill, stroke, text_color = "#f2f7fd", "#7aa7d8", "#12263a"

    x, y, w, h = node.x, node.y, node.w, node.h
    c.setFillColor(HexColor(fill))
    c.setStrokeColor(HexColor(stroke))
    c.setLineWidth(1.1)
    if node.shape == "diamond":
        p = c.beginPath()
        p.moveTo(x, y + h / 2)
        p.lineTo(x + w / 2, y + h)
        p.lineTo(x + w, y + h / 2)
        p.lineTo(x + w / 2, y)
        p.close()
        c.drawPath(p, stroke=1, fill=1)
    elif style == "term":
        c.roundRect(x, y, w, h, h / 2, stroke=1, fill=1)
    else:
        c.roundRect(x, y, w, h, 5, stroke=1, fill=1)

    c.setFillColor(HexColor(text_color))
    c.setFont(font, LEVEL_FONT_SIZE)
    leading = LEVEL_FONT_SIZE * 1.34
    top = y + h - (node.h - len(node.lines) * leading) / 2
    baseline = top - LEVEL_FONT_SIZE * 0.82
    for index, line in enumerate(node.lines):
        c.drawString(x + (node.w - c.stringWidth(line, font, LEVEL_FONT_SIZE)) / 2,
                     baseline - index * leading, line)


def arrow_head(c, x, y, direction):
    from reportlab.lib.colors import HexColor
    size = 6.2
    c.setFillColor(HexColor("#5a6b7d"))
    p = c.beginPath()
    p.moveTo(x, y)
    p.lineTo(x - direction * size, y + size * 0.44)
    p.lineTo(x - direction * size, y - size * 0.44)
    p.close()
    c.drawPath(p, stroke=0, fill=1)


def draw_edge(c, src, dst, label, direction, font):
    from reportlab.lib.colors import HexColor
    c.setStrokeColor(HexColor("#5a6b7d"))
    c.setLineWidth(1.15)
    c.setFillColor(HexColor("#5a6b7d"))

    if direction > 0:
        x1, x2 = src.x + src.w, dst.x
    else:
        x1, x2 = src.x, dst.x + dst.w
    y1, y2 = src.y + src.h / 2, dst.y + dst.h / 2
    mid = x1 + (x2 - x1) * 0.5
    c.bezier(x1, y1, mid, y1, mid, y2, x2, y2)
    arrow_head(c, x2, y2, 1 if direction > 0 else -1)
    if label:
        lx = (x1 + x2) / 2
        ly = (y1 + y2) / 2 + (7 if y2 >= y1 else -13)
        c.setFont(font, 7.4)
        c.setFillColor(HexColor("#8a5a00"))
        c.drawCentredString(lx, ly, label)


def draw_wrap_edge(c, src, dst, label, bus_y, font, page_left, page_right):
    """带间 U 形回线：右出 → 下行 → 左行 → 下行 → 右入（始终自左向右阅读）。"""
    from reportlab.lib.colors import HexColor
    c.setStrokeColor(HexColor("#5a6b7d"))
    c.setLineWidth(1.15)
    c.setFillColor(HexColor("#5a6b7d"))

    out_x = min(src.x + src.w + 22, page_right)
    in_x = max(dst.x - 22, page_left)
    sx, sy = src.x + src.w, src.y + src.h / 2
    ex, ey = dst.x, dst.y + dst.h / 2
    c.line(sx, sy, out_x, sy)
    c.line(out_x, sy, out_x, bus_y)
    c.line(out_x, bus_y, in_x, bus_y)
    c.line(in_x, bus_y, in_x, ey)
    c.line(in_x, ey, ex, ey)
    arrow_head(c, ex, ey, 1)
    c.setFont(font, 7.4)
    c.setFillColor(HexColor("#5a6b7d"))
    c.drawString(out_x + 4, sy + 4, "续下行")
    if label:
        c.setFillColor(HexColor("#8a5a00"))
        c.drawCentredString((out_x + in_x) / 2, bus_y + 4, label)


def export_mermaid(path):
    lines = ["graph LR"]
    for key, _rank, shape, text in NODES:
        label = text.replace("\n", "<br/>")
        if shape in ("stadium", "terminator"):
            lines.append("    %s([%s])" % (key, label))
        elif shape == "diamond":
            lines.append("    %s{%s}" % (key, label))
        else:
            lines.append("    %s[%s]" % (key, label))
    for src, dst, label in EDGES:
        if label:
            lines.append("    %s -->|%s| %s" % (src, label, dst))
        else:
            lines.append("    %s --> %s" % (src, dst))
    lines.append("")
    lines.append("    classDef term fill:#12355b,stroke:#12355b,color:#fff,stroke-width:1px;")
    lines.append("    classDef proc fill:#f2f7fd,stroke:#7aa7d8,color:#12263a,stroke-width:1px;")
    lines.append("    classDef dec fill:#fff5e6,stroke:#d68910,color:#5b3b00,stroke-width:1px;")
    lines.append("    class " + ",".join(k for k, _r, s, _t in NODES
                                         if s in ("stadium", "terminator")) + " term;")
    lines.append("    class " + ",".join(k for k, _r, s, _t in NODES
                                         if s == "rect") + " proc;")
    lines.append("    class " + ",".join(k for k, _r, s, _t in NODES
                                         if s == "diamond") + " dec;")
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")


def render(out_pdf, from_excel):
    from reportlab.lib.colors import HexColor
    from reportlab.pdfgen import canvas

    font, bold = register_fonts()
    nodes = {}
    for key, rank, shape, text in NODES:
        node = FNode(key, rank, shape, text)
        measure(node, font)
        nodes[key] = node

    ranks = sorted({n.rank for n in nodes.values()})
    by_rank = {r: [n for n in nodes.values() if n.rank == r] for r in ranks}
    colw = {r: max(n.w for n in by_rank[r]) for r in ranks}
    rank_h = {r: sum(n.h for n in by_rank[r]) + GAP_Y * (len(by_rank[r]) - 1) for r in ranks}

    bands = []
    covered = set()
    for label, band_ranks in STAGE_BANDS:
        present = [r for r in band_ranks if r in by_rank]
        if not present:
            continue
        covered.update(present)
        bands.append((label, present))
    leftover = [r for r in ranks if r not in covered]
    if leftover:
        bands.append(("其余步骤", leftover))
    band_index = {r: i for i, (_l, rs) in enumerate(bands) for r in rs}

    band_w = [sum(colw[r] for r in rs) + GAP_X * (len(rs) - 1) for _l, rs in bands]
    band_h = [max(rank_h[r] for r in rs) for _l, rs in bands]

    margin = 40
    header = 96
    footer = 36
    page_w = max(band_w) + 2 * margin
    total_h = header + sum(band_h) + BAND_GAP * (len(bands) - 1) + footer + 52
    content_w = page_w - 2 * margin

    c = canvas.Canvas(out_pdf, pagesize=(page_w, total_h))
    c.setTitle("打印计价算法流程图")
    c.setFillColor(HexColor("#ffffff"))
    c.rect(0, 0, page_w, total_h, stroke=0, fill=1)

    c.setFillColor(HexColor("#12355b"))
    c.setFont(bold, 21)
    c.drawString(margin, total_h - 46, "打印计价算法流程图")
    c.setFillColor(HexColor("#5a6b7d"))
    c.setFont(font, 9.4)
    c.drawString(margin, total_h - 64,
                 "图形完全对应程序实现：文档转换 → 覆盖率计算 → 查表计价 → 分级加收 → 结算导出。"
                 "默认参数依据《打印价格暂定方案和材料公示》%s"
                 % ("" if from_excel else "（内置默认值）"))
    c.setStrokeColor(HexColor("#d5dbe2"))
    c.setLineWidth(1)
    c.line(margin, total_h - 74, page_w - margin, total_h - 74)

    top = total_h - header
    bus_y = {}
    for bi, (label, band_ranks) in enumerate(bands):
        bottom = top - band_h[bi]
        width = band_w[bi]
        origin = margin + (content_w - width) / 2
        color = STAGE_COLORS[bi % len(STAGE_COLORS)]

        c.setFillColor(HexColor(shade(color, 0.93)))
        c.roundRect(origin - 8, bottom - 8, width + 16, band_h[bi] + 16, 7, stroke=0, fill=1)
        c.setFillColor(HexColor(color))
        c.setFont(bold, 9.6)
        c.drawString(origin - 2, top + 4, label)

        x = origin
        for rank in band_ranks:
            for node in by_rank[rank]:
                node.x = x + (colw[rank] - node.w) / 2
            x += colw[rank] + GAP_X

        center = bottom + band_h[bi] / 2
        for rank in band_ranks:
            y = center + rank_h[rank] / 2
            for node in by_rank[rank]:
                node.y = y - node.h
                y -= node.h + GAP_Y

        bus_y[bi] = bottom - BAND_GAP / 2
        top = bottom - BAND_GAP

    for src_key, dst_key, label in EDGES:
        src, dst = nodes[src_key], nodes[dst_key]
        src_band = band_index[src.rank]
        if src_band == band_index[dst.rank]:
            draw_edge(c, src, dst, label, 1, font)
        else:
            draw_wrap_edge(c, src, dst, label, bus_y[src_band], font, margin, page_w - margin)

    for node in nodes.values():
        draw_node(c, node, font, bold)

    c.setFillColor(HexColor("#7b8794"))
    c.setFont(font, 7.4)
    c.drawString(margin, 22,
                 "说明：超标加收仅对超出覆盖率上限的覆盖量计费；超出量按区间分段累加，"
                 "各档系数见步骤⑧；折算基准默认取该介质覆盖率上限（上限 10% 即每 10% 覆盖量折算 1 页）。")
    c.drawRightString(page_w - margin, 22,
                      "生成时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    c.showPage()
    c.save()
    return page_w, total_h


def export_png(pdf_path, png_path, dpi=160):
    try:
        import pymupdf
    except ImportError:
        import fitz as pymupdf
    doc = pymupdf.open(pdf_path)
    doc[0].get_pixmap(dpi=dpi).save(png_path)
    doc.close()


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.expanduser("~"), "Desktop")
    os.makedirs(out_dir, exist_ok=True)
    try:
        from_excel = bool(import_from_excel(D.DEFAULT_EXCEL_PATH).get("items"))
    except Exception:  # noqa: BLE001
        from_excel = False

    pdf_path = os.path.join(out_dir, "打印计价算法流程图.pdf")
    png_path = os.path.join(out_dir, "打印计价算法流程图.png")
    mmd_path = os.path.join(out_dir, "打印计价算法流程图.mmd")

    width, height = render(pdf_path, from_excel)
    export_png(pdf_path, png_path)
    export_mermaid(mmd_path)
    print("PDF :", pdf_path, "  %.1f x %.1f cm" % (width / 72 * 2.54, height / 72 * 2.54))
    print("PNG :", png_path)
    print("源码:", mmd_path)


if __name__ == "__main__":
    main()
