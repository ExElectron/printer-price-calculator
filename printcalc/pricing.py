# -*- coding: utf-8 -*-
"""计价引擎。

基础价 = 打印张数 × 单价。
若覆盖率超过价目表规定上限，则加收：
    加收 = 超出的覆盖折算页数 × 单价 × 加收系数(默认 0.6)
"""

import math
from dataclasses import dataclass, field

BASIS_LIMIT = "limit"
BASIS_FULL = "full"
BASIS_CUSTOM = "custom"

BASIS_LABELS = {
    BASIS_LIMIT: "按覆盖率上限折算",
    BASIS_FULL: "按整页(100%)折算",
    BASIS_CUSTOM: "自定义折算基准",
}

MODE_PER_PAGE = "per_page"
MODE_AGGREGATE = "aggregate"

MODE_LABELS = {
    MODE_PER_PAGE: "逐页判定（每张单独考核）",
    MODE_AGGREGATE: "整体判定（平均覆盖率考核）",
}


@dataclass
class DocResult:
    path: str
    name: str = ""
    ext: str = ""
    pdf_path: str = ""
    engine: str = ""
    pages: int = 0
    coverages: list = field(default_factory=list)
    error: str = ""
    converted: bool = False
    analyzed: bool = False

    @property
    def total_coverage(self):
        return float(sum(self.coverages))

    @property
    def avg_coverage(self):
        return self.total_coverage / self.pages if self.pages else 0.0

    @property
    def max_coverage(self):
        return max(self.coverages) if self.coverages else 0.0

    @property
    def ready(self):
        return self.analyzed and not self.error and self.pages > 0


@dataclass
class QuoteLine:
    doc: DocResult
    pages: int = 0
    faces: int = 0
    sheets: int = 0
    avg_coverage: float = 0.0
    excess_coverage: float = 0.0
    extra_pages: float = 0.0
    base: float = 0.0
    surcharge: float = 0.0

    @property
    def total(self):
        return self.base + self.surcharge


@dataclass
class Quote:
    item: dict = None
    sides: str = ""
    copies: int = 1
    unit_price: float = 0.0
    limit: float = None
    mode: str = MODE_PER_PAGE
    basis: str = BASIS_LIMIT
    factor: float = 0.6
    lines: list = field(default_factory=list)
    pages: int = 0
    faces: int = 0
    sheets: int = 0
    total_coverage: float = 0.0
    avg_coverage: float = 0.0
    excess_coverage: float = 0.0
    extra_pages: float = 0.0
    base: float = 0.0
    surcharge: float = 0.0
    total: float = 0.0
    over_limit: bool = False
    notes: list = field(default_factory=list)

    @property
    def has_surcharge(self):
        return self.surcharge > 0.0000001


def _num(cfg, key, default):
    try:
        return float(cfg.get(key, default))
    except (TypeError, ValueError):
        return default


def divisor_for(basis, limit, custom):
    if basis == BASIS_FULL:
        return 1.0
    if basis == BASIS_CUSTOM:
        value = custom if custom and custom > 0 else None
        return value if value is not None else (limit or 1.0)
    return limit or 1.0


def sheets_for(pages, sides):
    if pages <= 0:
        return 0
    if sides == "双面":
        return math.ceil(pages / 2)
    return pages


def compute_quote(docs, item, sides, copies, pricing_cfg):
    """根据分析结果与价目行计算报价。"""
    pricing_cfg = pricing_cfg or {}
    factor = _num(pricing_cfg, "overage_factor", 0.6)
    mode = pricing_cfg.get("limit_mode", MODE_PER_PAGE)
    basis = pricing_cfg.get("overage_basis", BASIS_LIMIT)
    custom = _num(pricing_cfg, "overage_custom_basis", 0.05)
    digits = int(_num(pricing_cfg, "round_digits", 2))

    copies = max(1, int(copies or 1))
    unit_price = 0.0
    limit = None
    if item:
        try:
            unit_price = float(item.get("price") or 0.0)
        except (TypeError, ValueError):
            unit_price = 0.0
        try:
            raw = item.get("limit")
            limit = float(raw) if raw not in (None, "") else None
        except (TypeError, ValueError):
            limit = None
        if limit is not None and limit <= 0:
            limit = None

    quote = Quote(item=item, sides=sides, copies=copies, unit_price=unit_price,
                  limit=limit, mode=mode, basis=basis, factor=factor)

    divisor = divisor_for(basis, limit, custom)
    lines = []
    total_pages = total_faces = total_sheets = 0
    total_cov = 0.0
    sum_doc_excess = 0.0

    for doc in docs:
        if not doc.ready:
            continue
        pages = doc.pages
        faces = pages * copies
        sheets = sheets_for(pages, sides) * copies
        doc_total = doc.total_coverage * copies

        if limit is None:
            doc_excess = 0.0
        elif mode == MODE_PER_PAGE:
            doc_excess = sum(max(0.0, c - limit) for c in doc.coverages) * copies
        else:
            doc_excess = max(0.0, doc_total - faces * limit)

        line = QuoteLine(
            doc=doc,
            pages=pages,
            faces=faces,
            sheets=sheets,
            avg_coverage=doc.avg_coverage,
            excess_coverage=doc_excess,
            base=sheets * unit_price,
        )
        lines.append(line)

        total_pages += pages
        total_faces += faces
        total_sheets += sheets
        total_cov += doc_total
        sum_doc_excess += doc_excess

    if limit is None:
        excess = 0.0
    elif mode == MODE_PER_PAGE:
        excess = sum_doc_excess
    else:
        excess = max(0.0, total_cov - total_faces * limit)

    quote.lines = lines
    quote.pages = total_pages
    quote.faces = total_faces
    quote.sheets = total_sheets
    quote.total_coverage = total_cov
    quote.avg_coverage = (total_cov / total_faces) if total_faces else 0.0
    quote.excess_coverage = excess
    quote.extra_pages = excess / divisor if divisor else 0.0
    quote.base = total_sheets * unit_price
    quote.surcharge = quote.extra_pages * unit_price * factor
    quote.over_limit = excess > 0.0000001

    if mode == MODE_AGGREGATE and sum_doc_excess > 0 and quote.surcharge > 0:
        for line in lines:
            line.excess_coverage = line.excess_coverage
            line.extra_pages = line.excess_coverage / divisor if divisor else 0.0
            line.surcharge = quote.surcharge * (line.excess_coverage / sum_doc_excess)
    else:
        for line in lines:
            line.extra_pages = line.excess_coverage / divisor if divisor else 0.0
            line.surcharge = line.extra_pages * unit_price * factor

    quote.base = round(quote.base, digits)
    quote.surcharge = round(quote.surcharge, digits)
    quote.total = round(quote.base + quote.surcharge, digits)
    for line in lines:
        line.base = round(line.base, digits)
        line.surcharge = round(line.surcharge, digits)

    quote.notes = build_notes(quote)
    return quote


def build_notes(quote):
    notes = []
    if not quote.item:
        return notes
    if not quote.lines:
        return notes
    price_text = "%.2f" % quote.unit_price
    notes.append("单价 %s 元/张（%s）" % (price_text, quote.item.get("note") or "—"))
    notes.append("共 %d 页 / %d 个打印面 / %d 张（%s，%d 份）"
                 % (quote.pages, quote.faces, quote.sheets, quote.sides, quote.copies))
    if quote.limit is None:
        notes.append("该介质未设置覆盖率上限，不计超标加收")
    else:
        notes.append("覆盖率上限 %.1f%%，实际平均覆盖率 %.2f%%"
                     % (quote.limit * 100, quote.avg_coverage * 100))
        if quote.over_limit:
            notes.append("超出覆盖 %.3f（折算 %.2f 页），按 %.0f%% 加收"
                         % (quote.excess_coverage, quote.extra_pages, quote.factor * 100))
        else:
            notes.append("未超出覆盖率要求，无超标加收")
    return notes


def format_money(value, digits=2):
    return ("%%.%df" % digits) % (value or 0.0)
