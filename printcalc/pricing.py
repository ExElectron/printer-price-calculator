# -*- coding: utf-8 -*-
"""计价引擎。

基础价 = 打印张数 × 单价。

若覆盖率超过价目表规定上限，超出部分按“覆盖率区间”分段累进递减加收：

    适度超标  上限 → 25%   系数 0.6
    重度超标  25% → 45%    系数 0.4
    极限超标  > 45%        系数 0.2

即把每页超出上限的覆盖量拆进各档区间，分别乘以该档系数，得到加权覆盖量：

    加收 = 加权覆盖量 ÷ 折算基准 × 单价

阶梯档位（区间上限、系数）均可在设置页自由调整。
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
    weighted_coverage: float = 0.0
    extra_pages: float = 0.0
    base: float = 0.0
    surcharge: float = 0.0

    @property
    def total(self):
        return self.base + self.surcharge

    @property
    def effective_factor(self):
        if self.excess_coverage > 0.0000001:
            return self.weighted_coverage / self.excess_coverage
        return 0.0


@dataclass
class Quote:
    item: dict = None
    sides: str = ""
    copies: int = 1
    unit_price: float = 0.0
    limit: float = None
    mode: str = MODE_PER_PAGE
    basis: str = BASIS_LIMIT
    tiers: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    pages: int = 0
    faces: int = 0
    sheets: int = 0
    total_coverage: float = 0.0
    avg_coverage: float = 0.0
    excess_coverage: float = 0.0
    weighted_coverage: float = 0.0
    extra_pages: float = 0.0
    base: float = 0.0
    surcharge: float = 0.0
    total: float = 0.0
    over_limit: bool = False
    notes: list = field(default_factory=list)

    @property
    def has_surcharge(self):
        return self.surcharge > 0.0000001

    @property
    def effective_factor(self):
        if self.excess_coverage > 0.0000001:
            return self.weighted_coverage / self.excess_coverage
        return 0.0


# --------------------------------------------------------------------------
# 超标阶梯
# --------------------------------------------------------------------------

def normalize_tiers(tiers, fallback_factor=0.6):
    """整理阶梯：系数有效的保留，按上限升序排列，不设上限的档位放最后。"""
    cleaned = []
    for tier in tiers or []:
        if not isinstance(tier, dict):
            continue
        try:
            factor = float(tier.get("factor"))
        except (TypeError, ValueError):
            continue
        raw = tier.get("upper")
        upper = None
        if raw not in (None, ""):
            try:
                upper = float(raw)
            except (TypeError, ValueError):
                upper = None
            if upper is not None and upper <= 0:
                upper = None
        cleaned.append({"upper": upper, "factor": factor})
    if not cleaned:
        cleaned = [{"upper": None, "factor": fallback_factor}]

    bounded = sorted((t for t in cleaned if t["upper"] is not None), key=lambda t: t["upper"])
    unbounded = [t for t in cleaned if t["upper"] is None]
    if not unbounded:
        unbounded = [{"upper": None, "factor": bounded[-1]["factor"] if bounded else fallback_factor}]
    return bounded + unbounded[:1]


def tier_split(coverage, limit, tiers):
    """把 coverage 超出 limit 的部分拆进各档，返回 (超出量, 加权量)。

    加权量 = Σ 各档超出量 × 该档系数。
    """
    if limit is None or coverage <= limit:
        return 0.0, 0.0
    excess = 0.0
    weighted = 0.0
    lower = limit
    last = len(tiers) - 1
    for index, tier in enumerate(tiers):
        upper = tier["upper"]
        if upper is None or index == last:
            amount = coverage - lower
            if amount > 0:
                excess += amount
                weighted += amount * tier["factor"]
            break
        if upper <= lower:
            continue
        if coverage <= lower:
            break
        amount = min(coverage, upper) - lower
        excess += amount
        weighted += amount * tier["factor"]
        lower = upper
        if coverage <= upper:
            break
    return excess, weighted


def tier_factor_for(coverage, limit, tiers):
    """某一页超出量对应的加权平均系数；未超出返回 None。"""
    excess, weighted = tier_split(coverage, limit, tiers)
    if excess <= 0:
        return None
    return weighted / excess


def _pct(value):
    percent = value * 100
    if abs(percent - round(percent)) < 0.05:
        return "%d%%" % round(percent)
    return "%.1f%%" % percent


def tier_summary(tiers, limit=None):
    """把阶梯渲染成一行说明，如 “10%–25%×0.60  25%–45%×0.40  >45%×0.20”。"""
    parts = []
    lower = limit
    for index, tier in enumerate(tiers):
        upper = tier["upper"]
        if upper is not None and lower is not None and upper <= lower:
            continue
        if upper is None:
            band = ">%s" % _pct(lower) if lower is not None else "全部超标"
        elif lower is None:
            band = "≤%s" % _pct(upper)
        else:
            band = "%s–%s" % (_pct(lower), _pct(upper))
        parts.append("%s×%.2f" % (band, tier["factor"]))
        if upper is None:
            break
        lower = upper
    return "  ".join(parts) if parts else "—"


# --------------------------------------------------------------------------
# 计算
# --------------------------------------------------------------------------

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
    mode = pricing_cfg.get("limit_mode", MODE_PER_PAGE)
    basis = pricing_cfg.get("overage_basis", BASIS_LIMIT)
    custom = _num(pricing_cfg, "overage_custom_basis", 0.05)
    digits = int(_num(pricing_cfg, "round_digits", 2))
    tiers = normalize_tiers(pricing_cfg.get("overage_tiers"),
                            _num(pricing_cfg, "overage_factor", 0.6))

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
                  limit=limit, mode=mode, basis=basis, tiers=tiers)

    divisor = divisor_for(basis, limit, custom)
    lines = []
    total_pages = total_faces = total_sheets = 0
    total_cov = 0.0
    sum_excess = sum_weighted = 0.0

    for doc in docs:
        if not doc.ready:
            continue
        pages = doc.pages
        faces = pages * copies
        sheets = sheets_for(pages, sides) * copies
        doc_total = doc.total_coverage * copies

        if limit is None:
            excess = weighted = 0.0
        elif mode == MODE_PER_PAGE:
            excess = weighted = 0.0
            for value in doc.coverages:
                band_excess, band_weighted = tier_split(value, limit, tiers)
                excess += band_excess
                weighted += band_weighted
            excess *= copies
            weighted *= copies
        else:
            representative = (doc.total_coverage / pages) if pages else 0.0
            band_excess, band_weighted = tier_split(representative, limit, tiers)
            excess = band_excess * faces
            weighted = band_weighted * faces

        line = QuoteLine(
            doc=doc,
            pages=pages,
            faces=faces,
            sheets=sheets,
            avg_coverage=doc.avg_coverage,
            excess_coverage=excess,
            weighted_coverage=weighted,
            base=sheets * unit_price,
        )
        lines.append(line)

        total_pages += pages
        total_faces += faces
        total_sheets += sheets
        total_cov += doc_total
        sum_excess += excess
        sum_weighted += weighted

    if limit is None:
        excess = weighted = 0.0
    elif mode == MODE_PER_PAGE:
        excess = sum_excess
        weighted = sum_weighted
    else:
        representative = (total_cov / total_faces) if total_faces else 0.0
        band_excess, band_weighted = tier_split(representative, limit, tiers)
        excess = band_excess * total_faces
        weighted = band_weighted * total_faces

    quote.lines = lines
    quote.pages = total_pages
    quote.faces = total_faces
    quote.sheets = total_sheets
    quote.total_coverage = total_cov
    quote.avg_coverage = (total_cov / total_faces) if total_faces else 0.0
    quote.excess_coverage = excess
    quote.weighted_coverage = weighted
    quote.extra_pages = (excess / divisor) if divisor else 0.0
    quote.base = total_sheets * unit_price
    quote.surcharge = (weighted / divisor) * unit_price if divisor else 0.0
    quote.over_limit = excess > 0.0000001

    if mode == MODE_AGGREGATE and sum_weighted > 0 and quote.surcharge > 0:
        for line in lines:
            line.extra_pages = (line.excess_coverage / divisor) if divisor else 0.0
            line.surcharge = quote.surcharge * (line.weighted_coverage / sum_weighted)
    else:
        for line in lines:
            line.extra_pages = (line.excess_coverage / divisor) if divisor else 0.0
            line.surcharge = ((line.weighted_coverage / divisor) * unit_price) if divisor else 0.0

    quote.base = round(quote.base, digits)
    quote.surcharge = round(quote.surcharge, digits)
    quote.total = round(quote.base + quote.surcharge, digits)
    for line in lines:
        line.base = round(line.base, digits)
        line.surcharge = round(line.surcharge, digits)

    quote.notes = build_notes(quote)
    return quote


def page_details(doc, limit, tiers, divisor, unit_price, copies):
    """逐页覆盖量明细，供报价单导出使用。"""
    rows = []
    for index, value in enumerate(doc.coverages):
        if limit is None:
            excess = weighted = 0.0
        else:
            excess, weighted = tier_split(value, limit, tiers)
        factor = (weighted / excess) if excess > 0 else None
        rows.append({
            "page": index + 1,
            "coverage": value,
            "limit": limit,
            "excess": excess,
            "weighted": weighted,
            "factor": factor,
            "extra_pages": (excess / divisor) if divisor else 0.0,
            "surcharge": ((weighted / divisor) * unit_price * copies) if divisor else 0.0,
        })
    return rows


def build_notes(quote):
    notes = []
    if not quote.item:
        return notes
    if not quote.lines:
        return notes
    notes.append("单价 %.2f 元/张（%s）" % (quote.unit_price, quote.item.get("note") or "—"))
    notes.append("共 %d 页 / %d 个打印面 / %d 张（%s，%d 份）"
                 % (quote.pages, quote.faces, quote.sheets, quote.sides, quote.copies))
    if quote.limit is None:
        notes.append("该介质未设置覆盖率上限，不计超标加收")
    else:
        notes.append("覆盖率上限 %.1f%%，实际平均覆盖率 %.2f%%"
                     % (quote.limit * 100, quote.avg_coverage * 100))
        notes.append("超标阶梯：%s" % tier_summary(quote.tiers, quote.limit))
        if quote.over_limit:
            notes.append("超出覆盖 %.3f（折算 %.2f 页），加权覆盖 %.3f，实际平均系数 %.2f"
                         % (quote.excess_coverage, quote.extra_pages,
                            quote.weighted_coverage, quote.effective_factor))
        else:
            notes.append("未超出覆盖率要求，无超标加收")
    return notes


def format_money(value, digits=2):
    return ("%%.%df" % digits) % (value or 0.0)
