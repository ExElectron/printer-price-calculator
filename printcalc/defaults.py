# -*- coding: utf-8 -*-
"""内置默认数据。

全部数据来源于《打印价格暂定方案和材料公示.xlsx》，同时作为
程序在无法读取 Excel 时的兜底数据。设置页可对任意字段进行修改。
"""

DEFAULT_EXCEL_PATH = r"C:\Users\17151\Desktop\打印价格暂定方案和材料公示.xlsx"
DEFAULT_ACROBAT_DIR = r"C:\Program Files\Adobe\Acrobat DC\Acrobat"
DEFAULT_OFFICE_DIR = r"C:\Program Files\Microsoft Office\root\Office16"

SIDES_SINGLE = "单面"
SIDES_DOUBLE = "双面"
SIDES_BOTH = "单面/双面"

METHOD_LASER = "激光打印"
METHOD_INKJET = "喷墨打印"

COLOR_MONO = "黑白"
COLOR_GRAY = "灰度"
COLOR_COLOR = "彩色"
COLOR_PHOTO = "彩色高精"

_PAPER_PLAIN = "普通打印纸"
_PAPER_LABEL = "不干胶纸"
_PAPER_CARD = "柯达高厚白卡纸"
_PAPER_COATED = "柯达双面铜版纸"
_PAPER_GLOSSY = "柯达高光相片纸"
_PAPER_RC = "富士RC绒面相纸"
_PAPER_OWN = "自带纸张"

_SPEC_PLAIN = "A4 (210×297mm) / 80g"
_SPEC_LABEL = "A4 (210×297mm)"
_SPEC_CARD = "A4 (210×297mm) / 230g"
_SPEC_COATED = "A4 (210×297mm) / 120g"
_SPEC_6IN = "6寸 (4×6英寸 约100×150mm)"
_SPEC_OWN = "自备纸张（需确认设备兼容）"


def _split_sides(text):
    return [p.strip() for p in text.split("/") if p.strip()]


def _item(idx, paper, spec, method, color, sides, price, limit, note="", available=True):
    return {
        "id": idx,
        "paper": paper,
        "spec": spec,
        "method": method,
        "color": color,
        "sides": _split_sides(sides),
        "price": price,
        "limit": limit,
        "note": note,
        "available": available,
    }


DEFAULT_ITEMS = [
    _item(1, _PAPER_PLAIN, _SPEC_PLAIN, METHOD_LASER, COLOR_MONO, SIDES_BOTH, None, None, "暂不提供", False),
    _item(2, _PAPER_PLAIN, _SPEC_PLAIN, METHOD_INKJET, COLOR_GRAY, SIDES_SINGLE, 0.20, 0.10, "墨水覆盖率 ≤ 10%"),
    _item(3, _PAPER_PLAIN, _SPEC_PLAIN, METHOD_INKJET, COLOR_GRAY, SIDES_DOUBLE, 0.30, 0.10, "墨水覆盖率 ≤ 10%"),
    _item(4, _PAPER_PLAIN, _SPEC_PLAIN, METHOD_INKJET, COLOR_COLOR, SIDES_SINGLE, 0.20, 0.10, "墨水覆盖率 ≤ 10%"),
    _item(5, _PAPER_PLAIN, _SPEC_PLAIN, METHOD_INKJET, COLOR_COLOR, SIDES_DOUBLE, 0.30, 0.10, "墨水覆盖率 ≤ 10%"),
    _item(6, _PAPER_LABEL, _SPEC_LABEL, METHOD_LASER, COLOR_MONO, SIDES_SINGLE, None, None, "暂不提供", False),
    _item(7, _PAPER_LABEL, _SPEC_LABEL, METHOD_INKJET, COLOR_GRAY, SIDES_SINGLE, 0.50, 0.10, "墨水覆盖率 ≤ 10%，标贴制作"),
    _item(8, _PAPER_LABEL, _SPEC_LABEL, METHOD_INKJET, COLOR_COLOR, SIDES_SINGLE, 0.50, 0.10, "墨水覆盖率 ≤ 10%，标贴制作"),
    _item(9, _PAPER_CARD, _SPEC_CARD, METHOD_LASER, COLOR_MONO, SIDES_BOTH, None, None, "暂不提供", False),
    _item(10, _PAPER_CARD, _SPEC_CARD, METHOD_INKJET, COLOR_GRAY, SIDES_SINGLE, 0.50, 0.35, "墨水覆盖率 ≤ 35%，封面/卡片/证书输出"),
    _item(11, _PAPER_CARD, _SPEC_CARD, METHOD_INKJET, COLOR_GRAY, SIDES_DOUBLE, 0.80, 0.35, "墨水覆盖率 ≤ 35%"),
    _item(12, _PAPER_CARD, _SPEC_CARD, METHOD_INKJET, COLOR_PHOTO, SIDES_SINGLE, 0.50, 0.35, "墨水覆盖率 ≤ 35%"),
    _item(13, _PAPER_CARD, _SPEC_CARD, METHOD_INKJET, COLOR_PHOTO, SIDES_DOUBLE, 0.80, 0.35, "墨水覆盖率 ≤ 35%"),
    _item(14, _PAPER_COATED, _SPEC_COATED, METHOD_INKJET, COLOR_GRAY, SIDES_SINGLE, 1.00, 0.45, "墨水覆盖率 ≤ 45%，双面涂层喷墨专用"),
    _item(15, _PAPER_COATED, _SPEC_COATED, METHOD_INKJET, COLOR_GRAY, SIDES_DOUBLE, 1.50, 0.45, "墨水覆盖率 ≤ 45%"),
    _item(16, _PAPER_COATED, _SPEC_COATED, METHOD_INKJET, COLOR_PHOTO, SIDES_SINGLE, 1.00, 0.45, "墨水覆盖率 ≤ 45%"),
    _item(17, _PAPER_COATED, _SPEC_COATED, METHOD_INKJET, COLOR_PHOTO, SIDES_DOUBLE, 1.50, 0.45, "墨水覆盖率 ≤ 45%"),
    _item(18, _PAPER_GLOSSY, _SPEC_6IN, METHOD_INKJET, COLOR_PHOTO, SIDES_SINGLE, 1.00, None, "270g 专业超厚高光相纸，相片成品输出"),
    _item(19, _PAPER_RC, _SPEC_6IN, METHOD_INKJET, COLOR_PHOTO, SIDES_SINGLE, 1.00, None, "260g 专业RC绒面相纸，亚光防反光防指纹"),
    _item(20, _PAPER_OWN, _SPEC_OWN, METHOD_INKJET, COLOR_GRAY, SIDES_SINGLE, 0.20, 0.10, "自备纸张纯耗墨服务费"),
    _item(21, _PAPER_OWN, _SPEC_OWN, METHOD_INKJET, COLOR_GRAY, SIDES_DOUBLE, 0.30, 0.10, "自备纸张纯耗墨服务费"),
    _item(22, _PAPER_OWN, _SPEC_OWN, METHOD_INKJET, COLOR_COLOR, SIDES_SINGLE, 0.20, 0.10, "自备纸张纯耗墨服务费"),
    _item(23, _PAPER_OWN, _SPEC_OWN, METHOD_INKJET, COLOR_COLOR, SIDES_DOUBLE, 0.30, 0.10, "自备纸张纯耗墨服务费"),
]

DEFAULT_NOTES = [
    "1. 计费基准：普通纸单双面按标准墨水覆盖率 ≤ 10% 计费；卡纸覆盖率放宽至 ≤ 35%（0.80元/张），铜版纸覆盖率放宽至 ≤ 45%（1.50元/张）。",
    "2. 大面积色块加收：若打印超高密度深色背景、满版照片级渲染图等重度超标耗墨作业，将视实际耗墨情况按 1.5 ~ 2.0 倍协商计费。",
    "3. 双面打印说明：A4 普通复印纸支持硬件自动双面；高厚白卡纸与双面铜版纸支持手动双面打印；不干胶标签纸及相纸仅提供单面打印。",
    "4. 自带纸张说明：若使用自备纸张，需经确认设备兼容性后上机，按单面 0.20 元/张、双面 0.30 元/张收取纯耗墨服务费。",
    "5. 设备服务说明：当前所有文印作业由 Brother DCP-T725DW 彩色喷墨多功能一体机提供，激光打印服务目前处于暂缓开放状态。",
]

DEFAULT_MATERIALS = [
    {"no": 1, "category": "打印墨水", "name": "原装彩色喷墨墨水（青色/蓝）", "brand": "兄弟（Brother）",
     "spec": "BT5009C，48.8ml 原装彩色染料墨水", "pack": "1 盒/瓶", "usage": "喷墨一体机原装彩色墨水，蓝色彩色输出"},
    {"no": 2, "category": "打印墨水", "name": "原装彩色喷墨墨水（品红/红）", "brand": "兄弟（Brother）",
     "spec": "BT5009M，48.8ml 原装彩色染料墨水", "pack": "1 盒/瓶", "usage": "喷墨一体机原装彩色墨水，红色彩色输出"},
    {"no": 3, "category": "打印墨水", "name": "原装彩色喷墨墨水（黄色/黄）", "brand": "兄弟（Brother）",
     "spec": "BT5009Y，48.8ml 原装彩色染料墨水", "pack": "1 盒/瓶", "usage": "喷墨一体机原装彩色墨水，黄色彩色输出"},
    {"no": 4, "category": "打印墨水", "name": "原装大容量黑色喷墨墨水", "brand": "兄弟（Brother）",
     "spec": "BTD60BK，108ml 超大容量黑色墨水", "pack": "1 盒/瓶", "usage": "喷墨一体机原装黑色墨水，文本与黑白高负荷输出"},
    {"no": 5, "category": "文档纸张", "name": "橙安妮 多功能复印打印纸", "brand": "亚太森博（Asia Symbol）",
     "spec": "80g/m²，A4 标准规格（210×297mm）", "pack": "5 包（500张/包）", "usage": "日常黑白/彩色文档、讲义、论文资料双面打印"},
    {"no": 6, "category": "特种介质", "name": "双面彩色喷墨铜版纸", "brand": "柯达（Kodak）",
     "spec": "120g/m²，A4 规格，双面涂层喷墨专用", "pack": "1 包（50张）", "usage": "高清彩色图文宣传单、企业简报、图册封面打印"},
    {"no": 7, "category": "特种介质", "name": "柯达高厚白卡纸", "brand": "柯达（Kodak）",
     "spec": "230g/m²，A4 规格，高挺度加厚卡纸", "pack": "1 包（50张）", "usage": "精装封面、画册卡片、厚卡纸彩色打印与手工制作"},
    {"no": 8, "category": "特种介质", "name": "A4 哑光不干胶打印纸", "brand": "真彩（TrueColor）",
     "spec": "A4 规格，高粘性黄色底纸，全张背胶", "pack": "1 包（100张）", "usage": "标贴、封口贴、分类标识及个性化标签制作"},
    {"no": 9, "category": "照片相纸", "name": "6寸 背胶高光喷墨相纸", "brand": "柯达（Kodak）",
     "spec": "120g/m²，4×6 英寸（约100×150mm），自带高粘背胶", "pack": "1 包（20张）", "usage": "照片贴纸、DIY手账、便携证件照片快速粘贴输出"},
    {"no": 10, "category": "照片相纸", "name": "6寸 RC喷墨照片相纸（绒面）", "brand": "富士胶片（Fujifilm）",
     "spec": "260g/m²，4×6 英寸（约100×150mm），RC微孔防水", "pack": "1 包（100张）", "usage": "高品质人像/艺术照打印，细腻绒面质感防眩光防指纹"},
    {"no": 11, "category": "照片相纸", "name": "6寸 Ultra Premium 高光相纸", "brand": "柯达（Kodak）",
     "spec": "270g/m²，4×6 英寸（约100×150mm），超厚专业级", "pack": "1 包（100张）", "usage": "专业级高光相片打印，色彩饱和鲜艳，光泽度高"},
    {"no": 12, "category": "后期装裱", "name": "A4 冷裱保护膜（高光/光面）", "brand": "古德（GOOD）",
     "spec": "A4 规格，高透PVC冷裱保护膜，防刮防紫外线", "pack": "1 包（50张）", "usage": "打印照片/重要图文冷裱覆膜，长效防水防潮防刮花"},
]

DEFAULT_EQUIPMENT = [
    {"no": 1, "category": "硬件设备", "brand": "兄弟（Brother）", "name": "Brother DCP-T725DW", "spec": "彩色喷墨多功能一体机"},
    {"no": 2, "category": "办公套件", "brand": "微软（Microsoft）", "name": "Microsoft Word 2024", "spec": "Office LTSC 专业增强版 2024"},
    {"no": 3, "category": "办公套件", "brand": "微软（Microsoft）", "name": "Microsoft Excel 2024", "spec": "Office LTSC 专业增强版 2024"},
    {"no": 4, "category": "办公套件", "brand": "微软（Microsoft）", "name": "Microsoft PowerPoint 2024", "spec": "Office LTSC 专业增强版 2024"},
    {"no": 5, "category": "图像设计", "brand": "Adobe", "name": "Adobe Photoshop 2025", "spec": "版本 26.2.0 (64-bit)"},
    {"no": 6, "category": "矢量设计", "brand": "Adobe", "name": "Adobe Illustrator 2025", "spec": "版本 29.1 (64-bit)"},
    {"no": 7, "category": "摄影调色", "brand": "Adobe", "name": "Adobe Lightroom Classic 2025", "spec": "版本 14.1.1 (Camera Raw 17.1)"},
    {"no": 8, "category": "文档管理", "brand": "Adobe", "name": "Adobe Acrobat Pro 2024", "spec": "Continuous Release (2024.005)"},
    {"no": 9, "category": "工程设计", "brand": "欧特克（Autodesk）", "name": "Autodesk AutoCAD 2027", "spec": "2027"},
    {"no": 10, "category": "摄影调色", "brand": "Blackmagic Design", "name": "DaVinci Resolve Studio 21", "spec": "Studio 21.0.2 BUILD 4"},
]


def default_config():
    """返回一份完整的默认配置。"""
    return {
        "version": 1,
        "excel_path": DEFAULT_EXCEL_PATH,
        "pricing": {
            "overage_factor": 0.6,
            "limit_mode": "per_page",
            "overage_basis": "limit",
            "overage_custom_basis": 0.05,
            "round_digits": 2,
        },
        "coverage": {
            "dpi": 300,
        },
        "conversion": {
            "engine": "acrobat",
            "acrobat_dir": DEFAULT_ACROBAT_DIR,
            "office_dir": DEFAULT_OFFICE_DIR,
            "image_dpi": 96,
            "keep_temp": False,
            "timeout_seconds": 300,
        },
        "defaults": {
            "paper": _PAPER_PLAIN,
            "method": METHOD_INKJET,
            "color": COLOR_GRAY,
            "sides": SIDES_SINGLE,
            "copies": 1,
        },
        "items": list(DEFAULT_ITEMS),
        "notes": list(DEFAULT_NOTES),
        "materials": list(DEFAULT_MATERIALS),
        "equipment": list(DEFAULT_EQUIPMENT),
    }
