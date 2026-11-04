#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""filters.py —— 装饰素材 / 二维码 / 文末 END 的通用过滤器。

单独成模块的原因：`wx2docx.py`（生成时过滤）和 `check_output.py`（交付前自检）
必须用**同一套判据**，否则生成端删了、自检端却看不出来，或者反过来误报。

四类判据：
    1. `is_decorative()`  —— 低清小图标（尺寸阈值）
    2. `is_gif()`         —— 按文件头判定动图
    3. `is_qr_code()`     —— 二维码（含中间带 logo 的公众号二维码）
    4. `has_transparency()` —— 带透明通道的贴图素材（横幅、礼花、分割线）
    5. `is_end_marker()`  —— 文末 END 标记（END / -END- / （完）/ THE END）

二维码为什么不能只看尺寸：公众号文末二维码常见 258×258、260×260，
既过得了 `--min-px 200`，像素面积也远超 `--min-area 40000`，
按尺寸判永远删不掉。只能看**结构**——近正方形 + 高黑白纯度 + 高横向跳变。
"""

import os
import re

# --------------------------------------------------------------------------
# 尺寸：低清装饰图
# --------------------------------------------------------------------------
MIN_PX = 200
MIN_AREA = 40000


def is_decorative(w, h, min_px=MIN_PX, min_area=MIN_AREA):
    """低分辨率装饰图（小图标、细分割线、扫尾小图）判定。"""
    if not w or not h:
        return True
    return max(w, h) < min_px or (w * h) < min_area


def is_gif(path):
    """按文件头判定 gif，避免服务端把 gif 存成 .jpg 时漏判。"""
    try:
        with open(path, "rb") as fh:
            return fh.read(4) == b"GIF8"
    except Exception:
        return False


# --------------------------------------------------------------------------
# 二维码
# --------------------------------------------------------------------------
# 二维码结构特征（实测标定，样本含带 logo 的公众号二维码）
QR_ASPECT_MIN, QR_ASPECT_MAX = 0.80, 1.25   # 近正方形
QR_DARK_MIN, QR_DARK_MAX = 0.15, 0.60       # 深色模块占比
QR_PURITY_MIN = 0.85                        # 近黑/近白像素合计占比（照片远低于此）
QR_TRANS_MIN = 0.06                         # 横向相邻跳变比例（照片通常 <0.02）


def _band(im, mode):
    """取单通道 raw bytes：L 模式灰度，A 模式透明度。"""
    im = im.convert(mode)
    return im, im.tobytes()


def qr_like_image(im):
    """二维码判定核心：近正方形 + 黑白二值化纯度高 + 模块跳变密集。

    三条同时满足才判定，单看任何一条都会误伤：
      - 正方形      → 排除横幅、海报、竖图
      - 黑白纯度高  → 排除实景照片（肤色/天空会让灰度散开）
      - 跳变密集    → 排除纯色块、纯白背景的表格图
    中心压了 logo 的二维码依然满足：logo 只占少数像素。
    """
    try:
        w, h = im.size
        if not w or not h:
            return False
        if not (QR_ASPECT_MIN <= w / float(h) <= QR_ASPECT_MAX):
            return False
        im.thumbnail((140, 140))  # 归一化，避免大图拖慢
        g = im.convert("L")
        data = g.tobytes()
        n = len(data)
        if not n:
            return False
        dark = sum(1 for v in data if v < 90) / float(n)
        light = sum(1 for v in data if v > 165) / float(n)
        if not (QR_DARK_MIN <= dark <= QR_DARK_MAX):
            return False
        if dark + light < QR_PURITY_MIN:
            return False
        gw, gh = g.size
        px = g.load()
        trans = 0
        for y in range(gh):
            for x in range(1, gw):
                if abs(px[x, y] - px[x - 1, y]) > 80:
                    trans += 1
        return trans / float(gw * gh) >= QR_TRANS_MIN
    except Exception:
        return False


def is_qr_code(path):
    """按文件路径判定二维码。"""
    try:
        from PIL import Image
    except ImportError:
        return False
    try:
        with Image.open(path) as im:
            return qr_like_image(im.copy())
    except Exception:
        return False


# --------------------------------------------------------------------------
# 透明贴图素材
# --------------------------------------------------------------------------
def translucent_image(im, min_frac=0.02):
    """已打开的图片是否带透明通道且确实透明（礼花、红绸、华表、分割条贴图）。

    实拍照片几乎不会是 RGBA 且真透明；公众号装饰素材则几乎必然是。
    阈值 2%：整图有 2% 以上像素透出来，就不像照片。
    """
    try:
        if im.mode != "RGBA" and "transparency" not in im.info:
            return False
        alpha = im.convert("RGBA").getchannel("A")
        alpha.thumbnail((120, 120))
        data = alpha.tobytes()
        if not data:
            return False
        return sum(1 for v in data if v < 250) / float(len(data)) > min_frac
    except Exception:
        return False


def has_transparency(path, min_frac=0.02):
    """按文件路径判断是否是透明贴图素材。"""
    try:
        from PIL import Image
    except ImportError:
        return False
    try:
        with Image.open(path) as im:
            return translucent_image(im, min_frac)
    except Exception:
        return False


def is_wide_banner(path, min_ratio=2.2):
    """超宽横条：头条横幅、页边装饰条。只在开头/结尾位置参考，不作全局判据。"""
    try:
        from PIL import Image
        with Image.open(path) as im:
            w, h = im.size
        return bool(w and h) and w / float(h) >= min_ratio
    except Exception:
        return False


# --------------------------------------------------------------------------
# 文末 END
# --------------------------------------------------------------------------
# 去掉空白与全套中英文标点/分隔符号后再比对核心词
_END_NOISE = re.compile(
    r"[\s\-—–－·•*~〜～_=|/\\\[\]【】〔〕（）()<>《》〈〉\"'“”‘’,，.。!！?？:：;；、]+")
_END_WORDS = set(["END", "THEEND", "ENDTHE", "本文END", "全文END",
                  "本文完", "全文完", "完", "ENDTHANKS"])


def is_end_marker(txt):
    """文末结束标记：END / -END- / （完）/ THE END / 全文完 …

    注意「完」单独成段也算——公众号常见，但不能整段只有标点。
    """
    if not txt:
        return False
    core = _END_NOISE.sub("", txt)
    if not core:
        return False
    return core.upper() in _END_WORDS


def strip_tail_end(blocks):
    """删掉文末的 END 标记，以及它之后仅剩的装饰图/二维码。

    只在「文末」生效：END 之后若还有正文，说明它不是结束标记（正文里
    偶尔会出现 END 字样），原样保留，宁可不删也不误删正文。

    blocks 形如 [("p"|"h1"|"h2"|"img", value), ...]，返回新列表。
    """
    last_end = None
    for i, (kind, val) in enumerate(blocks):
        if kind != "img" and is_end_marker(val):
            last_end = i
    if last_end is None:
        return blocks
    # END 之后只允许是图片或另一个 END 标记
    for kind, val in blocks[last_end + 1:]:
        if kind == "img":
            continue
        if is_end_marker(val):
            continue
        return blocks
    return blocks[:last_end]
