# -*- coding: utf-8 -*-
"""文档转 PDF。

优先调用 Adobe Acrobat（通过 PowerShell COM：AcroExch.AVDoc / PDDoc），
失败时按配置回退到 Microsoft Office、Pillow（图片）或内置文本排版。
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
from dataclasses import dataclass, field

RESOURCE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources")

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".jp2"}
OFFICE_EXTS = {".doc", ".docx", ".rtf", ".odt", ".xls", ".xlsx", ".xlsm", ".csv", ".ods", ".ppt", ".pptx", ".odp"}
TEXT_EXTS = {".txt", ".md", ".log", ".ini", ".json", ".xml"}

SUPPORTED_EXTS = sorted({".pdf"} | IMAGE_EXTS | OFFICE_EXTS | TEXT_EXTS)

ENGINES = [
    ("acrobat", "Adobe Acrobat"),
    ("auto", "自动（推荐）"),
    ("office", "Microsoft Office"),
    ("pillow", "Pillow（图片）"),
]

_CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


@dataclass
class ConvertResult:
    src: str
    pdf_path: str = ""
    ok: bool = False
    engine: str = ""
    error: str = ""
    pages: int = 0
    extra: dict = field(default_factory=dict)


def _safe_stem(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", stem).strip() or "document"
    digest = hashlib.md5(os.path.abspath(path).encode("utf-8", "ignore")).hexdigest()[:8]
    return "%s_%s" % (stem, digest)


def output_path(src, out_dir):
    return os.path.join(out_dir, _safe_stem(src) + ".pdf")


def default_temp_dir():
    return os.path.join(tempfile.gettempdir(), "PrintCalc")


def _engine_order(ext, engine):
    if engine == "acrobat":
        return ["acrobat", "office", "pillow", "text"]
    if engine == "office":
        return ["office", "acrobat", "pillow", "text"]
    if engine == "pillow":
        return ["pillow", "acrobat", "office", "text"]
    if ext in IMAGE_EXTS:
        return ["pillow", "acrobat", "office"]
    if ext in OFFICE_EXTS:
        return ["acrobat", "office", "pillow", "text"]
    if ext in TEXT_EXTS:
        return ["acrobat", "text", "office"]
    return ["acrobat", "office", "pillow", "text"]


# --------------------------------------------------------------------------
# PowerShell 引擎
# --------------------------------------------------------------------------

def _write_jsonl(path, rows):
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")


def _read_jsonl(path):
    rows = []
    if not os.path.isfile(path):
        return rows
    with open(path, "r", encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if line:
                try:
                    rows.append(json.loads(line))
                except ValueError:
                    continue
    return rows


def _powershell_exe():
    for candidate in (
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        shutil.which("powershell"),
        shutil.which("pwsh"),
    ):
        if candidate and os.path.isfile(candidate):
            return candidate
    return "powershell"


def _run_powershell(script_name, jobs, timeout):
    """运行 PowerShell 转换脚本。返回 {src: (ok, pages, error)}。"""
    script = os.path.join(RESOURCE_DIR, script_name)
    if not os.path.isfile(script):
        return {job["src"]: (False, 0, "缺少转换脚本：%s" % script_name) for job in jobs}

    with tempfile.TemporaryDirectory(prefix="printcalc_ps_") as tmp:
        jobs_file = os.path.join(tmp, "jobs.jsonl")
        result_file = os.path.join(tmp, "results.jsonl")
        _write_jsonl(jobs_file, jobs)
        cmd = [
            _powershell_exe(), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", script, "-JobsFile", jobs_file, "-ResultFile", result_file,
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=timeout,
                           creationflags=_CREATE_NO_WINDOW, check=False)
        except subprocess.TimeoutExpired:
            return {job["src"]: (False, 0, "转换超时（超过 %s 秒）" % timeout) for job in jobs}
        except OSError as exc:
            return {job["src"]: (False, 0, "无法启动 PowerShell：%s" % exc) for job in jobs}

        results = {}
        for row in _read_jsonl(result_file):
            results[row.get("src", "")] = (
                bool(row.get("ok")),
                int(row.get("pages") or 0),
                row.get("error") or "",
            )
        for job in jobs:
            results.setdefault(job["src"], (False, 0, "转换脚本未返回结果"))
        return results


# --------------------------------------------------------------------------
# Pillow / 文本 引擎
# --------------------------------------------------------------------------

def _convert_image(src, dst, dpi):
    from PIL import Image, ImageSequence

    with Image.open(src) as im:
        frames = []
        for frame in ImageSequence.Iterator(im):
            frames.append(frame.convert("RGB"))
        if not frames:
            raise ValueError("图片没有可用的帧")
        frames[0].save(dst, "PDF", resolution=float(dpi), save_all=len(frames) > 1,
                       append_images=frames[1:])
        return len(frames)


def _read_text_file(path):
    for encoding in ("utf-8-sig", "utf-8", "gbk", "big5", "latin-1"):
        try:
            with open(path, "r", encoding=encoding) as fh:
                return fh.read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        return fh.read()


def _convert_text(src, dst):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfgen import canvas

    text = _read_text_file(src)
    width, height = A4
    margin = 40
    line_height = 14
    font_size = 10
    font = "Helvetica"
    for candidate in ("MicrosoftYaHei", "SimSun", "SimHei"):
        try:
            pdfmetrics.getFont(candidate)
            font = candidate
            break
        except Exception:  # noqa: BLE001
            continue

    def wrap(line, max_width):
        chunks = []
        current = ""
        for ch in line:
            if pdfmetrics.stringWidth(current + ch, font, font_size) <= max_width:
                current += ch
            else:
                chunks.append(current)
                current = ch
        chunks.append(current)
        return chunks

    doc = canvas.Canvas(dst, pagesize=A4)
    doc.setFont(font, font_size)
    y = height - margin
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        for chunk in wrap(raw_line, width - 2 * margin):
            if y < margin:
                doc.showPage()
                doc.setFont(font, font_size)
                y = height - margin
            doc.drawString(margin, y, chunk)
            y -= line_height
    doc.save()
    return 1


# --------------------------------------------------------------------------
# 对外接口
# --------------------------------------------------------------------------

def convert_one(src, out_dir, cfg, progress=None, should_cancel=None):
    results = convert_many([src], out_dir, cfg, progress=progress, should_cancel=should_cancel)
    return results[0]


def convert_many(paths, out_dir, cfg, progress=None, should_cancel=None):
    """把一批文件转换成 PDF，返回与 paths 等长的 ConvertResult 列表。"""
    conv = cfg.get("conversion", {})
    engine = conv.get("engine", "acrobat")
    timeout = int(conv.get("timeout_seconds", 300) or 300)
    image_dpi = conv.get("image_dpi", 96) or 96

    os.makedirs(out_dir, exist_ok=True)
    total = len(paths)
    if total == 0:
        return []

    final = {}
    pending = {}
    processed = 0
    for src in paths:
        ext = os.path.splitext(src)[1].lower()
        if ext == ".pdf":
            # 已经是 PDF：直接使用原文件，不做任何转换
            final[src] = ConvertResult(src=src, pdf_path=src, ok=True, engine="pdf")
            processed += 1
            if progress:
                progress(processed, total, "%s（PDF 直接分析，不转换）" % os.path.basename(src))
            continue
        dst = output_path(src, out_dir)
        pending[src] = {"dst": dst, "ext": ext, "order": _engine_order(ext, engine), "idx": 0}
        final[src] = ConvertResult(src=src, pdf_path=dst)

    def report(src, note):
        nonlocal processed
        processed += 1
        if progress:
            progress(processed, total, "%s（%s）" % (os.path.basename(src), note))

    while pending:
        buckets = {}
        for src, state in pending.items():
            order = state["order"]
            if state["idx"] >= len(order):
                continue
            buckets.setdefault(order[state["idx"]], []).append(src)
        if not buckets:
            break

        for name in ("acrobat", "office"):
            if name not in buckets:
                continue
            jobs = [{"src": src, "dst": pending[src]["dst"]} for src in buckets[name]]
            outcomes = _run_powershell(
                "acrobat_convert.ps1" if name == "acrobat" else "office_convert.ps1", jobs, timeout)
            for src in buckets[name]:
                ok, pages, error = outcomes.get(src, (False, 0, "未知错误"))
                if ok and os.path.isfile(pending[src]["dst"]):
                    result = final[src]
                    result.ok = True
                    result.engine = name
                    result.pages = pages
                    report(src, "Acrobat" if name == "acrobat" else "Office")
                    del pending[src]
                else:
                    state = pending[src]
                    state["last_error"] = error or "转换失败"
                    state["idx"] += 1

        for name in ("pillow", "text"):
            if name not in buckets:
                continue
            for src in list(buckets[name]):
                if should_cancel and should_cancel():
                    break
                state = pending.get(src)
                if state is None:
                    continue
                try:
                    if name == "pillow":
                        if state["ext"] not in IMAGE_EXTS:
                            raise ValueError("Pillow 仅支持图片格式")
                        pages = _convert_image(src, state["dst"], image_dpi)
                    else:
                        if state["ext"] not in TEXT_EXTS:
                            raise ValueError("文本引擎仅支持纯文本")
                        pages = _convert_text(src, state["dst"])
                    result = final[src]
                    result.ok = True
                    result.engine = "pillow" if name == "pillow" else "text"
                    result.pages = pages
                    report(src, "Pillow" if name == "pillow" else "文本排版")
                    del pending[src]
                except Exception as exc:  # noqa: BLE001
                    state["last_error"] = str(exc)
                    state["idx"] += 1

        if should_cancel and should_cancel():
            break

    for src, state in pending.items():
        result = final[src]
        result.ok = False
        result.error = state.get("last_error", "没有可用的转换引擎")
        report(src, "失败")

    order_map = {src: index for index, src in enumerate(paths)}
    return sorted(final.values(), key=lambda r: order_map.get(r.src, 0))


def clean_temp_dir(path):
    if path and os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)


class CancellationToken:
    def __init__(self):
        self._event = threading.Event()

    def cancel(self):
        self._event.set()

    def is_cancelled(self):
        return self._event.is_set()
