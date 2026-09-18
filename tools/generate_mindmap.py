# -*- coding: utf-8 -*-
"""生成《打印计价算法》思维导图（PDF + PNG）。

数据取自《打印价格暂定方案和材料公示.xlsx》的默认参数；
若 Excel 不可读，则退回程序内置默认值。

用法：
    python tools/generate_mindmap.py [输出目录]

默认输出到桌面：
    打印计价算法思维导图.pdf
    打印计价算法思维导图.png
"""

import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from printcalc import defaults as D  # noqa: E402
from printcalc import pricing  # noqa: E402
from printcalc.excel_import import import_from_excel  # noqa: E402

# --------------------------------------------------------------------------
# 数据
# --------------------------------------------------------------------------

PAPER_ORDER = [
    "普通打印纸", "不干胶纸", "柯达高厚白卡纸", "柯达双面铜版纸",
    "柯达高光相片纸", "富士RC绒面相纸", "自带纸张",
]


def load_items():
    try:
        data = import_from_excel(D.DEFAULT_EXCEL_PATH)
        if data.get("items"):
            return data["items"], True
    except Exception:  # noqa: BLE001
        pass
    return [dict(it) for it in D.DEFAULT_ITEMS], False


def money(value):
    if value is None:
        return "暂不提供"
    return "%.2f" % value


def sides_text(sides):
    return "/".join(sides)


def build_price_groups(items):
    """按纸张分组，返回 [(标题, [价格行, ...]), ...]。"""
    grouped = {}
    for item in items:
        grouped.setdefault(item.get("paper", ""), []).append(item)

    papers = [p for p in PAPER_ORDER if p in grouped] + \
             [p for p in grouped if p not in PAPER_ORDER]

    groups = []
    for paper in papers:
        rows = grouped[paper]
        spec = rows[0].get("spec", "")
        head = "%s  %s" % (paper, spec)
        sub = []
        buckets = {}
        order = []
        for row in rows:
            key = (row.get("method", ""), row.get("color", ""))
            if key not in buckets:
                buckets[key] = {}
                order.append(key)
            for side in row.get("sides", []):
                buckets[key][side] = ("暂不提供" if row.get("price") is None
                                      else "%.2f 元" % float(row["price"]))
        for key in order:
            method, color = key
            price_map = buckets[key]
            if "单面" in price_map and "双面" in price_map:
                price_text = "单面 %s / 双面 %s" % (price_map["单面"], price_map["双面"])
            elif "单面" in price_map:
                price_text = "单面 %s" % price_map["单面"]
            elif "双面" in price_map:
                price_text = "双面 %s" % price_map["双面"]
            else:
                price_text = "—"
            prefix = method if color in ("", "黑白") else "%s·%s" % (method, color)
            sub.append("%s：%s" % (prefix, price_text))
        groups.append((head, sub))
    return groups


def paper_limits(items):
    out = []
    seen = {}
    for item in items:
        paper = item.get("paper", "")
        if paper in seen:
            continue
        limit = None
        for row in items:
            if row.get("paper") != paper:
                continue
            raw = row.get("limit")
            if raw in (None, ""):
                continue
            value = float(raw)
            limit = value if limit is None else max(limit, value)
        seen[paper] = limit
    for paper in PAPER_ORDER:
        if paper in seen:
            out.append((paper, seen[paper]))
    for paper, limit in seen.items():
        if paper not in PAPER_ORDER:
            out.append((paper, limit))
    return out


def build_tree():
    items, from_excel = load_items()

    root = Node("打印计价算法\n（默认参数公示）", level=0)

    # ① 输入
    n = root.add("① 计费输入")
    n.add("纸张介质：%s" % "、".join(p for p in PAPER_ORDER))
    n.add("打印方式：喷墨打印（激光打印暂不提供）")
    n.add("色彩模式：黑白 / 灰度 / 彩色 / 彩色高精")
    n.add("打印面数：单面 / 双面")
    n.add("打印份数：1 ~ 9999 份")
    n.add("文档：PDF / Word / Excel / PPT / 图片 / 网页 / 纯文本，可批量添加")

    # ② 转 PDF
    n = root.add("② 文档 → PDF")
    n.add("输入已是 PDF → 直接读取原文件，不做任何转换")
    n.add("其他格式 → 调用 Adobe Acrobat 转换")
    n.add("回退链：Acrobat → Microsoft Office → Pillow(图片) → 内置文本排版")
    n.add("转换产物存放于临时目录，退出可自动清理")

    # ③ 覆盖率
    n = root.add("③ 墨水覆盖率计算")
    n.add("页面渲染为位图（默认 300 DPI，可调 100~600）")
    n.add("公式：(1 − 平均像素值/255) × 像素数 ÷ (A4面积 × DPI²)")
    n.add("A4 面积 = 8.268 × 11.693 平方英寸")
    n.add("基准：满版 A4 纯黑 = 100%；纯白 = 0%")
    n.add("逐页统计；文档平均覆盖率 = 总覆盖量 ÷ 打印面数")
    n.add("示例：50% 灰整页 = 50%；纯红整页 ≈ 66.7%")

    # ④ 基础价
    n = root.add("④ 基础价")
    n.add("打印张数：单面 = 页数；双面 = ⌈页数 ÷ 2⌉；再乘份数")
    n.add("基础价 = 打印张数 × 单价（元/张）")
    n.add("各文件分别取整后求和（奇数页末张只算一面）")

    # ⑤ 覆盖率上限
    n = root.add("⑤ 覆盖率上限（按介质）")
    for paper, limit in paper_limits(items):
        n.add("%s：%s" % (paper, "不设上限" if limit is None else "≤ %.0f%%" % (limit * 100)))

    # ⑥ 判定方式
    n = root.add("⑥ 超标判定方式")
    n.add("逐页判定（默认）：每张单独考核，超标页各自计费")
    n.add("整体判定：按全单平均覆盖率考核，低覆盖页可抵扣高覆盖页")

    # ⑦ 阶梯
    tiers = pricing.normalize_tiers(D.DEFAULT_TIERS)
    n = root.add("⑦ 超标加收阶梯（分段累进递减）")
    n.add("档 1 适度：覆盖率上限 ~ 25%%   系数 %.2f" % tiers[0]["factor"])
    n.add("档 2 重度：25%% ~ 45%%   系数 %.2f" % tiers[1]["factor"])
    n.add("档 3 极限：> 45%%   系数 %.2f" % tiers[2]["factor"])
    n.add("超出上限的覆盖量按上述区间拆分，各档分别乘以对应系数")
    n.add("加权覆盖量 = Σ（各档超出量 × 该档系数）")
    n.add("阶梯档位与系数可在“设置”页自由调整")
    sub = n.add("累进示例（上限 10%%，某页覆盖 60%%）")
    sub.add("10%~25% 段：0.15 × 0.60 = 0.090")
    sub.add("25%~45% 段：0.20 × 0.40 = 0.080")
    sub.add(">45% 段：0.15 × 0.20 = 0.030")
    sub.add("超出 0.50，加权 0.200，实际平均系数 0.40")

    # ⑧ 折算基准
    n = root.add("⑧ 覆盖折算基准")
    n.add("按覆盖率上限折算（默认）：10% 上限 → 每 10% 覆盖量记 1 页")
    n.add("按整页折算：每 100% 覆盖量记 1 页")
    n.add("自定义基准：如按 5% 折算 1 页")

    # ⑨ 加收金额
    n = root.add("⑨ 超标加收金额")
    n.add("折算页数 = 超出覆盖量 ÷ 折算基准")
    n.add("加收金额 = 加权覆盖量 ÷ 折算基准 × 单价")
    n.add("覆盖率未超限，或该介质不设上限 → 加收 = 0")
    n.add("加收按全单汇总；逐页明细仅作参考时以汇总为准")

    # ⑩ 结算
    n = root.add("⑩ 结算与输出")
    n.add("应收合计 = 基础价 + 超标加收")
    n.add("实收金额：打印完成后录入")
    n.add("优惠 = 应收合计 − 实收金额（实收高于应收则为负数）")
    n.add("报价单 CSV：打印方案 / 报价汇总 / 各文件汇总 / 逐页覆盖量明细")
    n.add("覆盖率明细页：逐页覆盖率、是否超限、适用系数、页面预览")

    # ⑪ 价目表
    n = root.add("⑪ 默认价目表（元/张）")
    for head, sub_lines in build_price_groups(items):
        n.add("\n".join([head] + sub_lines))
    n.add("数据来源：%s" % ("《打印价格暂定方案和材料公示》Excel" if from_excel else "程序内置默认值"))

    # ⑫ 示例
    n = root.add("⑫ 示例计算")
    n.add("场景：146 页文档，普通纸 80g，喷墨·灰度，双面，0.30 元/张，1 份")
    n.add("张数 = ⌈146 ÷ 2⌉ = 73 张\n基础价 = 73 × 0.30 = 21.90 元")
    n.add("实测：平均覆盖率 9.37%，最高页 61.45%；逐页超标\n超出覆盖量 3.915，加权覆盖量 1.778")
    n.add("折算页数 = 3.915 ÷ 0.10 = 39.15 页\n加收 = 1.778 ÷ 0.10 × 0.30 = 5.33 元")
    n.add("应收合计 = 21.90 + 5.33 = 27.23 元\n若实收 25.00 元 → 优惠 2.23 元")

    return root, from_excel


# --------------------------------------------------------------------------
# 节点与排版
# --------------------------------------------------------------------------

LEVELS = [
    dict(font="CJK-Bold", size=15.5, wrap=210, pad_x=10, pad_y=7, radius=8),
    dict(font="CJK-Bold", size=11.5, wrap=225, pad_x=8, pad_y=5, radius=6),
    dict(font="CJK", size=9.0, wrap=300, pad_x=7, pad_y=4, radius=5),
    dict(font="CJK", size=8.2, wrap=320, pad_x=6, pad_y=3, radius=4),
]

PALETTE = [
    "#1f6feb", "#0a8f6a", "#c0392b", "#8e44ad", "#d68910", "#0e7490",
    "#b3541e", "#2f6f3e", "#7b3fa0", "#1a5276", "#a93226", "#117a65",
]


class Node:
    def __init__(self, text, level=0, children=None):
        self.text = text
        self.level = level
        self.children = children or []
        self.w = 0
        self.h = 0
        self.x = 0
        self.y = 0
        self.lines = []
        self.branch = 0
        self.color = "#333333"

    def add(self, text):
        child = Node(text, self.level + 1)
        self.children.append(child)
        return child


def wrap_text(text, font, size, max_width):
    from reportlab.pdfbase import pdfmetrics
    lines = []
    for paragraph in str(text).split("\n"):
        if not paragraph:
            lines.append("")
            continue
        token = ""

        def flush():
            nonlocal token
            if token:
                tokens.append(token)
                token = ""

        tokens = []
        for ch in paragraph:
            if "\u4e00" <= ch <= "\u9fff" or ch in "，。：；！？、（）《》【】“”—…·%×÷≤≥①②③④⑤⑥⑦⑧⑨⑩⑪⑫⌈⌉":
                flush()
                tokens.append(ch)
            elif ch == " ":
                flush()
                tokens.append(" ")
            else:
                token += ch
        flush()

        current = ""
        for item in tokens:
            if item == " " and not current:
                continue
            candidate = current + item
            if pdfmetrics.stringWidth(candidate, font, size) <= max_width or not current:
                current = candidate
            else:
                lines.append(current.rstrip())
                current = "" if item == " " else item
        lines.append(current.rstrip())
    return lines


def measure(node):
    style = LEVELS[min(node.level, len(LEVELS) - 1)]
    node.lines = wrap_text(node.text, style["font"], style["size"], style["wrap"])
    from reportlab.pdfbase import pdfmetrics
    width = max((pdfmetrics.stringWidth(line, style["font"], style["size"])
                 for line in node.lines), default=0)
    leading = style["size"] * 1.32
    node.w = width + 2 * style["pad_x"]
    node.h = len(node.lines) * leading + 2 * style["pad_y"]
    for child in node.children:
        measure(child)


def assign_colors(node, index=0):
    if node.level == 1:
        node.color = PALETTE[index % len(PALETTE)]
    for i, child in enumerate(node.children):
        if node.level == 0:
            assign_colors(child, i)
        else:
            child.color = node.color
            assign_colors(child)


def column_widths(root):
    widths = {}
    def walk(node):
        widths[node.level] = max(widths.get(node.level, 0), node.w)
        for child in node.children:
            walk(child)
    walk(root)
    return widths


def assign_x(root, widths, gap):
    positions = {}
    x = 0
    for level in sorted(widths):
        positions[level] = x
        x += widths[level] + gap
    def walk(node):
        node.x = positions[node.level]
        for child in node.children:
            walk(child)
    walk(root)
    return x - gap


def assign_y(node, top, gap):
    if not node.children:
        node.y = top + node.h / 2
        return top + node.h
    cursor = top
    for child in node.children:
        cursor = assign_y(child, cursor, gap) + gap
    cursor -= gap
    node.y = (node.children[0].y + node.children[-1].y) / 2
    return max(cursor, top + node.h)


def flip_y(node, total):
    """PDF 坐标自下而上，翻转后分支顺序即自上而下阅读。"""
    node.y = total - node.y
    for child in node.children:
        flip_y(child, total)


def span(node):
    """子树占用的垂直高度（需先完成布局）。"""
    values = []

    def walk(n):
        values.append(n.y - n.h / 2)
        values.append(n.y + n.h / 2)
        for child in n.children:
            walk(child)

    walk(node)
    return max(values) - min(values)


def balance_split(branches):
    """按子树高度把分支均分成两组，避免某页大量留白。"""
    spans = [span(b) for b in branches]
    total = sum(spans)
    best_index, best_diff = 1, None
    for index in range(1, len(branches)):
        left = sum(spans[:index])
        diff = abs(left - (total - left))
        if best_diff is None or diff < best_diff:
            best_diff, best_index = diff, index
    return branches[:best_index], branches[best_index:]


# --------------------------------------------------------------------------
# 绘制
# --------------------------------------------------------------------------

def register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    pairs = [
        (r"C:\Windows\Fonts\msyh.ttc", 0, r"C:\Windows\Fonts\msyhbd.ttc", 0),
        (r"C:\Windows\Fonts\simhei.ttf", None, r"C:\Windows\Fonts\simhei.ttf", None),
    ]
    for regular, rindex, bold, bindex in pairs:
        try:
            if rindex is None:
                pdfmetrics.registerFont(TTFont("CJK", regular))
            else:
                pdfmetrics.registerFont(TTFont("CJK", regular, subfontIndex=rindex))
            if bindex is None:
                pdfmetrics.registerFont(TTFont("CJK-Bold", bold))
            else:
                pdfmetrics.registerFont(TTFont("CJK-Bold", bold, subfontIndex=bindex))
            return "CJK", "CJK-Bold"
        except Exception:  # noqa: BLE001
            continue
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    return "STSong-Light", "STSong-Light"


def shade(hex_color, factor):
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return "#%02x%02x%02x" % (r, g, b)


def draw_node(c, node):
    from reportlab.lib.colors import HexColor
    style = LEVELS[min(node.level, len(LEVELS) - 1)]
    x = node.x
    y = node.y - node.h / 2
    if node.level == 0:
        fill = stroke = "#12355b"
        text_color = "#ffffff"
    elif node.level == 1:
        fill = shade(node.color, 0.82)
        stroke = node.color
        text_color = "#12263a"
    else:
        fill = shade(node.color, 0.92) if node.level == 2 else "#f7f8fa"
        stroke = shade(node.color, 0.35)
        text_color = "#2c3e50"
    c.setFillColor(HexColor(fill))
    c.setStrokeColor(HexColor(stroke))
    c.setLineWidth(1.1 if node.level <= 1 else 0.6)
    c.roundRect(x, y, node.w, node.h, style["radius"], stroke=1, fill=1)

    c.setFillColor(HexColor(text_color))
    c.setFont(style["font"], style["size"])
    leading = style["size"] * 1.32
    top = node.y + node.h / 2
    baseline = top - style["pad_y"] - style["size"] * 0.82
    for i, line in enumerate(node.lines):
        c.drawString(x + style["pad_x"], baseline - i * leading, line)


def draw_edges(c, node):
    from reportlab.lib.colors import HexColor
    for child in node.children:
        x1 = node.x + node.w
        y1 = node.y
        x2 = child.x
        y2 = child.y
        c.setStrokeColor(HexColor(shade(child.color, 0.35)))
        c.setLineWidth(max(0.6, 2.2 - child.level * 0.55))
        mid = x1 + (x2 - x1) * 0.55
        c.bezier(x1, y1, mid, y1, mid, y2, x2, y2)
        draw_edges(c, child)


MARGIN = 34
HEADER = 74
FOOTER = 26


def layout(root, gap=26, gap_y=4.5):
    measure(root)
    assign_colors(root)
    widths = column_widths(root)
    tree_width = assign_x(root, widths, gap)
    tree_height = assign_y(root, 0, gap_y)
    flip_y(root, tree_height)
    return tree_width, tree_height


def draw_tree(c, root):
    draw_edges(c, root)

    def walk(node):
        draw_node(c, node)
        for child in node.children:
            walk(child)
    walk(root)


def draw_page(c, root, tree_width, page_width, page_height, title, subtitle):
    from reportlab.lib.colors import HexColor

    c.setFillColor(HexColor("#ffffff"))
    c.rect(0, 0, page_width, page_height, stroke=0, fill=1)

    left = MARGIN
    top = page_height - MARGIN
    c.setFillColor(HexColor("#12355b"))
    c.setFont(LEVELS[0]["font"], 22)
    c.drawString(left, top - 24, title)
    c.setFillColor(HexColor("#5a6b7d"))
    c.setFont(LEVELS[2]["font"], 10)
    c.drawString(left, top - 42, subtitle)
    c.setStrokeColor(HexColor("#d5dbe2"))
    c.setLineWidth(1)
    c.line(left, top - 54, page_width - MARGIN, top - 54)

    offset_x = MARGIN + max(0.0, (page_width - 2 * MARGIN - tree_width) / 2.0)
    c.saveState()
    c.translate(offset_x, MARGIN + FOOTER)
    draw_tree(c, root)
    c.restoreState()

    c.setFillColor(HexColor("#7b8794"))
    c.setFont(LEVELS[2]["font"], 8)
    c.drawString(MARGIN, MARGIN - 4,
                 "本图依据程序默认参数生成，参数可在程序“设置”页调整。")
    c.drawRightString(page_width - MARGIN, MARGIN - 4,
                      "生成时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))


def render(root, out_pdf, title, subtitle):
    """单页思维导图，页面尺寸自适应内容。"""
    from reportlab.pdfgen import canvas

    tree_width, tree_height = layout(root)
    page_width = tree_width + 2 * MARGIN
    page_height = tree_height + HEADER + FOOTER + 2 * MARGIN
    c = canvas.Canvas(out_pdf, pagesize=(page_width, page_height))
    c.setTitle(title)
    draw_page(c, root, tree_width, page_width, page_height, title, subtitle)
    c.showPage()
    c.save()
    return page_width, page_height


def render_paged(roots, out_pdf, title, subtitle):
    """把若干子树排版成统一尺寸的多页思维导图（便于 A3/A4 分页打印）。"""
    from reportlab.pdfgen import canvas

    layouts = [layout(root) for root in roots]
    tree_width = max(w for w, _h in layouts)
    tree_height = max(h for _w, h in layouts)
    page_width = tree_width + 2 * MARGIN
    page_height = tree_height + HEADER + FOOTER + 2 * MARGIN
    c = canvas.Canvas(out_pdf, pagesize=(page_width, page_height))
    c.setTitle(title)
    for index, root in enumerate(roots):
        page_subtitle = "%s（第 %d / %d 页）" % (subtitle, index + 1, len(roots))
        draw_page(c, root, tree_width, page_width, page_height, title, page_subtitle)
        c.showPage()
    c.save()
    return page_width, page_height, len(roots)


def export_png(pdf_path, png_path, dpi=170):
    try:
        import pymupdf
    except ImportError:
        import fitz as pymupdf
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi)
    pix.save(png_path)
    doc.close()


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.expanduser("~"), "Desktop")
    os.makedirs(out_dir, exist_ok=True)
    regular, bold = register_fonts()
    LEVELS[0]["font"] = bold
    LEVELS[1]["font"] = bold
    LEVELS[2]["font"] = regular
    LEVELS[3]["font"] = regular
    root, from_excel = build_tree()
    subtitle = ("打印计价算法说明 · 依据《打印价格暂定方案和材料公示》默认参数%s"
                % ("" if from_excel else "（内置默认值）"))

    pdf_path = os.path.join(out_dir, "打印计价算法思维导图.pdf")
    png_path = os.path.join(out_dir, "打印计价算法思维导图.png")
    size = render(root, pdf_path, "打印计价算法公示", subtitle)
    export_png(pdf_path, png_path)
    print("PDF :", pdf_path, "  %.1f x %.1f cm" % (size[0] / 72 * 2.54, size[1] / 72 * 2.54))
    print("PNG :", png_path)

    branches = list(root.children)
    if len(branches) > 1:
        groups = list(balance_split(branches))
        roots = [Node("打印计价算法\n（%d/%d）" % (i, len(groups)), level=0, children=list(g))
                 for i, g in enumerate(groups, 1)]
        paged = os.path.join(out_dir, "打印计价算法思维导图_A3分页.pdf")
        pw, ph, pages = render_paged(roots, paged, "打印计价算法公示", subtitle)
        print("分页:", paged, "  %d 页  %.1f x %.1f cm" % (pages, pw / 72 * 2.54, ph / 72 * 2.54))


if __name__ == "__main__":
    main()
