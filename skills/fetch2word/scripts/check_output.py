#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_output.py —— 交付前自检：逐篇校验生成的 docx 是否合规。

用法：
    python check_output.py --out "产物/10月"
    python check_output.py --out "产物/10月" --min-px 200 --min-area 40000

校验项（任一不通过则退出码非 0）：
    1. gif = 0                  —— 装饰 gif 已清干净
    2. empty = 0                —— 无完全空段落
    3. dangling = 0             —— 无悬空图片引用（Word 里会显示红叉）
    4. low-res = 0              —— 无残留低清装饰图
    5. 文件名形如 yyyy.mm.dd(Title).docx
    6. 首段严格为 yyyy.mm.dd（发出日期）
    7. 次段非空                  —— 标题行存在
    8. 第 3 段起无居中段落        —— 小标题必须居左
    9. 正文段落均有首行缩进      —— 小标题除外
   10. 体积 < 8MB（默认软阈值）  —— 提示残留部件
"""

import argparse
import glob
import os
import re
import sys

BLIP = ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
# 日期行固定：yyyy.mm.dd（发出日期），一个字都不能少
DATE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2}（发出日期）$")
# 文件名固定规则：yyyy.mm.dd(Title).docx
NAME_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2}\(.+\)\.docx$")
SIZE_LIMIT_KB = 8 * 1024


def is_decorative(w, h, min_px, min_area):
    if not w or not h:
        return True
    return max(w, h) < min_px or (w * h) < min_area


def check_one(path, min_px, min_area):
    from docx import Document
    from docx.image.image import Image as DocxImage
    from docx.oxml.ns import qn

    doc = Document(path)
    body_xml = doc.element.body.xml
    refs = set(re.findall(r'r:embed="(rId\d+)"', body_xml))
    dangling = refs - set(doc.part.rels.keys())

    gif = 0
    lowres = 0
    imgs = 0
    for para in doc.paragraphs:
        for blip in para._p.findall(BLIP):
            rid = blip.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rid or rid not in doc.part.rels:
                continue
            part = doc.part.rels[rid].target_part
            imgs += 1
            if part.content_type == "image/gif":
                gif += 1
            try:
                info = DocxImage.from_blob(part.blob)
                if is_decorative(info.px_width, info.px_height, min_px, min_area):
                    lowres += 1
            except Exception:
                pass

    empty = 0
    for para in doc.paragraphs:
        if para.text.strip():
            continue
        if para._p.findall(BLIP) or para._p.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"):
            continue
        empty += 1

    paras = doc.paragraphs
    head = paras[0].text.strip() if paras else ""
    title = paras[1].text.strip() if len(paras) > 1 else ""
    size_kb = os.path.getsize(path) // 1024

    problems = []
    if not NAME_RE.match(os.path.basename(path)):
        problems.append("文件名不符 yyyy.mm.dd(Title).docx")
    if gif:
        problems.append("gif=%d" % gif)
    if empty:
        problems.append("empty=%d" % empty)
    if dangling:
        problems.append("dangling=%s" % sorted(dangling))
    if lowres:
        problems.append("low-res=%d" % lowres)
    if not DATE_RE.match(head):
        problems.append("日期行不符(yyyy.mm.dd（发出日期）):%r" % head[:24])

    # 小标题一律居左：第 3 段起不允许出现居中段落
    centered = [p.text[:12] for p in paras[2:] if str(p.alignment).startswith("CENTER")]
    if centered:
        problems.append("存在居中段落:%s" % centered[:2])

    # 正文必须有首行缩进（小标题不加缩进）
    no_indent = 0
    for p in paras[2:]:
        if not p.text.strip():
            continue
        first = p.runs[0] if p.runs else None
        if first is not None and first.bold:
            continue
        ppr = p._p.pPr
        ind = ppr.find(qn("w:ind")) if ppr is not None else None
        ok = False
        if ind is not None:
            ok = bool(ind.get(qn("w:firstLineChars")) or ind.get(qn("w:firstLine")))
        if not ok:
            no_indent += 1
    if no_indent:
        problems.append("缺首行缩进:%d段" % no_indent)
    if not title:
        problems.append("缺标题行")
    if size_kb > SIZE_LIMIT_KB:
        problems.append("体积偏大:%dKB" % size_kb)

    return {"file": os.path.basename(path), "imgs": imgs, "size": size_kb,
            "head": head, "problems": problems}


def main():
    ap = argparse.ArgumentParser(description="校验生成的 docx 是否合规")
    ap.add_argument("--out", required=True, help="产出目录")
    ap.add_argument("--min-px", type=int, default=200, help="装饰图最长边阈值")
    ap.add_argument("--min-area", type=int, default=40000, help="装饰图像素面积阈值")
    args = ap.parse_args()

    try:
        import docx  # noqa: F401
    except ImportError:
        sys.exit("缺少依赖：python -m pip install python-docx")

    files = [f for f in sorted(glob.glob(os.path.join(args.out, "*.docx")))
             if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit("目录下没有 docx：%s" % args.out)

    bad = 0
    total_img = 0
    for f in files:
        r = check_one(f, args.min_px, args.min_area)
        total_img += r["imgs"]
        flag = "OK  " if not r["problems"] else "FAIL"
        if r["problems"]:
            bad += 1
        print("%s %-46s 图%3d %5dKB  %s" % (
            flag, r["file"][:44], r["imgs"], r["size"],
            "；".join(r["problems"]) if r["problems"] else r["head"]))

    print("\n共 %d 篇，图片 %d 张，不合规 %d 篇" % (len(files), total_img, bad))
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
