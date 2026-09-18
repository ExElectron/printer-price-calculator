# -*- coding: utf-8 -*-
"""配置读写与选项辅助函数。"""

import copy
import json
import os

from . import APP_ID
from . import defaults as D

CONFIG_FILENAME = "config.json"


def app_root():
    """程序目录（main.py 所在目录）。"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def config_path():
    return os.path.join(app_root(), CONFIG_FILENAME)


def _deep_merge(base, override):
    """把 override 合并进 base（base 被修改），返回 base。"""
    for key, value in (override or {}).items():
        if key in base and isinstance(base[key], dict) and isinstance(value, dict):
            _deep_merge(base[key], value)
        else:
            base[key] = value
    return base


def load_config(path=None):
    path = path or config_path()
    cfg = D.default_config()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                saved = json.load(fh)
            if isinstance(saved, dict):
                _deep_merge(cfg, saved)
                if not cfg.get("items"):
                    cfg["items"] = copy.deepcopy(D.DEFAULT_ITEMS)
        except (OSError, ValueError):
            pass
    return cfg


def save_config(cfg, path=None):
    path = path or config_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return path


def reset_config():
    return D.default_config()


# --------------------------------------------------------------------------
# 选项辅助
# --------------------------------------------------------------------------

def item_price(item):
    try:
        return float(item["price"])
    except (TypeError, ValueError):
        return None


def item_limit(item):
    try:
        value = float(item["limit"])
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def item_available(item):
    if not item.get("available", True):
        return False
    return item_price(item) is not None


def paper_list(items):
    seen = []
    for it in items:
        if it["paper"] not in seen:
            seen.append(it["paper"])
    return seen


def method_list(items, paper=None):
    seen = []
    for it in items:
        if paper and it["paper"] != paper:
            continue
        if it["method"] not in seen:
            seen.append(it["method"])
    return seen


def color_list(items, paper=None, method=None):
    seen = []
    for it in items:
        if paper and it["paper"] != paper:
            continue
        if method and it["method"] != method:
            continue
        if it["color"] not in seen:
            seen.append(it["color"])
    return seen


def sides_list(items, paper=None, method=None, color=None):
    seen = []
    for it in items:
        if paper and it["paper"] != paper:
            continue
        if method and it["method"] != method:
            continue
        if color and it["color"] != color:
            continue
        for side in it["sides"]:
            if side not in seen:
                seen.append(side)
    order = {D.SIDES_SINGLE: 0, D.SIDES_DOUBLE: 1}
    seen.sort(key=lambda s: order.get(s, 9))
    return seen


def spec_of(items, paper):
    for it in items:
        if it["paper"] == paper and it.get("spec"):
            return it["spec"]
    return ""


def find_item(items, paper, method, color, sides):
    """查找匹配的价目行；优先返回可提供的行。"""
    candidates = []
    for it in items:
        if it["paper"] != paper or it["method"] != method or it["color"] != color:
            continue
        if sides not in it["sides"]:
            continue
        candidates.append(it)
    for it in candidates:
        if item_available(it):
            return it
    return candidates[0] if candidates else None
