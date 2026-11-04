#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wx2docx.py —— 微信公众号推文批量转 Word（同一模板、图片内嵌）

用法示例（在工作区根目录执行，全部使用相对路径）：
    python .workbuddy/skills/wx-article-to-docx/scripts/wx2docx.py \
        --links "链接/10月.txt" \
        --template "模板/2025.11.26(团员代表).docx" \
        --out "产物/10月"

依赖：python-docx、beautifulsoup4、lxml
    python -m pip install python-docx beautifulsoup4 lxml

设计原则：
1. 不出现任何绝对路径，输入输出全部由命令行参数决定，默认按工作区相对目录推断。
2. 版式（字体字号、页边距、图片宽度）从模板文档本身读取，不硬编码。
3. 网络下载优先 urllib，失败自动回退 curl（兼容 SSL/证书异常的机器）。
"""

import argparse
import csv
import datetime
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from html import unescape
from urllib.parse import parse_qs, urlparse
from urllib.request import Request, urlopen

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
REFERER = "https://mp.weixin.qq.com/"

# 内置模板（已脱敏，版式与原始模板完全一致），随 skill 一起分发
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TEMPLATE = os.path.join(SKILL_DIR, "assets", "template.docx")

# 文件名规则固定为：文件名：yyyy.mm.dd(Title).docx
# Title = 标题去掉栏目前缀，剔除 Windows 非法字符，过长截断
NAME_MAX = 60

# 装饰图判定：最长边不足 MIN_PX，或像素面积不足 MIN_AREA，视为低清装饰图标
MIN_PX = 200
MIN_AREA = 40000
# 图片最小可用体积，低于此判定为下载失败的空包
MIN_BYTES = 500


def is_decorative(w, h, min_px=MIN_PX, min_area=MIN_AREA):
    """低分辨率装饰图（小图标、细分割线、扫尾小图）判定。"""
    if not w or not h:
        return True
    return max(w, h) < min_px or (w * h) < min_area


def strip_empty_paragraphs(doc):
    """删除完全没有内容的空段落（无文字、无图片）。"""
    from docx.oxml.ns import qn
    removed = 0
    for p in list(doc.paragraphs):
        if p.text.strip():
            continue
        el = p._p
        if el.findall(".//" + qn("a:blip")) or el.findall(".//" + qn("w:drawing")) \
                or el.findall(".//" + qn("w:pict")):
            continue
        el.getparent().remove(el)
        removed += 1
    return removed


# --------------------------------------------------------------------------
# 网络
# --------------------------------------------------------------------------
def alt_url(url):
    """`/640?` 是压缩版，`/0?` 是原图；压缩版被拒时换原图重试。"""
    return url.replace("/640?", "/0?").replace("/640#", "/0#")


def download_image(url, dst):
    """带重试 + 原图回退的图片下载，返回 (ok, bytes)。"""
    for attempt in range(3):
        target = url if attempt < 2 else alt_url(url)
        data = fetch(target, referer=REFERER, timeout=45)
        if len(data) > MIN_BYTES:
            with open(dst, "wb") as fh:
                fh.write(data)
            return True, len(data)
    with open(dst, "wb") as fh:
        fh.write(b"")
    return False, 0


def is_gif(path):
    """按文件头判定 gif，避免服务端把 gif 存成 .jpg 时漏判。"""
    try:
        with open(path, "rb") as fh:
            return fh.read(6)[:4] in (b"GIF8",)
    except Exception:
        return False


def ensure_supported(path):
    """python-docx 解析不了的格式（如 webp）转成 jpg，返回可用路径或 None。"""
    from docx.image.image import Image as DocxImage
    try:
        DocxImage.from_file(path)
        return path
    except Exception:
        pass
    try:
        from PIL import Image as PILImage
    except ImportError:
        return None
    try:
        im = PILImage.open(path)
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        new_path = os.path.splitext(path)[0] + ".conv.jpg"
        im.save(new_path, "JPEG", quality=90)
        return new_path
    except Exception:
        return None


def fetch(url, referer=None, timeout=60):
    """返回 bytes；urllib 失败时回退 curl。"""
    headers = {"User-Agent": UA, "Accept": "*/*"}
    if referer:
        headers["Referer"] = referer
    try:
        req = Request(url, headers=headers)
        with urlopen(req, timeout=timeout) as resp:
            return resp.read()
    except Exception:
        pass
    cmd = ["curl", "-sL", "--max-time", str(timeout), "-A", UA]
    if referer:
        cmd += ["-e", referer]
    cmd += [url]
    try:
        out = subprocess.run(cmd, capture_output=True, check=False).stdout
        return out or b""
    except Exception:
        return b""


# --------------------------------------------------------------------------
# 链接清单
# --------------------------------------------------------------------------
def read_links(path):
    """读取链接清单：按出现顺序去重；识别文件内的“排序：从旧到新/从新到旧”提示。"""
    hint = None
    urls = []
    with open(path, encoding="utf-8") as fh:
        for raw in fh:
            line = raw.strip()
            if not line:
                continue
            if line.startswith("http"):
                if line not in urls:
                    urls.append(line)
            elif "排序" in line:
                if "从旧到新" in line or "旧→新" in line:
                    hint = "asc"
                elif "从新到旧" in line or "新→旧" in line:
                    hint = "desc"
    return urls, hint


# --------------------------------------------------------------------------
# 正文解析
# --------------------------------------------------------------------------
def parse_article(url, raw_html):
    from bs4 import BeautifulSoup

    html_text = raw_html.decode("utf-8", errors="ignore")
    soup = BeautifulSoup(html_text, "lxml")

    title = ""
    m = re.search(r'<meta property="og:title" content="([^"]*)"', html_text)
    if m:
        title = unescape(m.group(1))
    if not title:
        m = re.search(r"var msg_title = '([^']*)'", html_text)
        if m:
            title = m.group(1)
    if not title:
        node = soup.select_one("#activity-name")
        title = node.get_text(strip=True) if node else ""
    title = re.sub(r"\s+", " ", title).strip()

    date = ""
    m = re.search(r"var createTime = '(\d{4}-\d{2}-\d{2})", html_text)
    if m:
        date = m.group(1)
    if not date:
        m = re.search(r'var ct = "(\d{10})"', html_text)
        if m:
            date = datetime.datetime.fromtimestamp(int(m.group(1))).strftime("%Y-%m-%d")

    content = soup.select_one("#js_content")
    blocks = []

    if content is not None:
        def walk(el):
            for child in el.children:
                name = getattr(child, "name", None)
                if name is None:
                    txt = str(child).replace("\u200b", "").replace("\xa0", " ").strip()
                    if txt:
                        blocks.append(("p", txt))
                    continue
                if name in ("script", "style", "svg"):
                    continue
                if name == "img":
                    src = child.get("data-src") or child.get("src") or ""
                    if src.startswith("//"):
                        src = "https:" + src
                    if src.startswith("http"):
                        blocks.append(("img", src))
                    continue
                has_block = child.find(["p", "section", "div", "img", "table", "tr",
                                        "td", "li", "h1", "h2", "h3", "h4", "h5", "h6",
                                        "blockquote"])
                if has_block:
                    walk(child)
                else:
                    txt = child.get_text().replace("\u200b", "").replace("\xa0", " ").strip()
                    if txt:
                        blocks.append(("p", re.sub(r"[ \t]+", " ", txt)))

        walk(content)

    return {"url": url, "title": title, "date": date, "blocks": blocks}


def normalize_title(title, column):
    """栏目与标题之间统一用 “||” 分隔（沿用模板写法）。"""
    t = re.sub(r"\s+", " ", title).strip()
    t = re.sub(r"\|+", "||", t)
    if "||" not in t and column:
        t = column + "||" + t
    return t


def safe_name(date, title):
    y, m, d = date.split("-")
    base = title.split("||", 1)[1] if "||" in title else title
    base = re.sub(r'[\\/:*?"<>|\r\n\t]', "", base)
    base = re.sub(r"\s+", " ", base).strip()
    if len(base) > NAME_MAX:
        base = base[:NAME_MAX]
    return "%s.%s.%s(%s).docx" % (y, m, d, base)


# --------------------------------------------------------------------------
# 模板解析：版式全部从模板读取
# --------------------------------------------------------------------------
def read_template_spec(template_path):
    from docx import Document
    from docx.shared import Emu

    doc = Document(template_path)
    sec = doc.sections[0]
    text_w = int(sec.page_width - sec.left_margin - sec.right_margin)

    size_pt, bold = 14.0, True
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip():
                if r.font.size:
                    size_pt = r.font.size.pt
                bold = bool(r.bold)
                break
        else:
            continue
        break

    img_w = text_w
    for shp in doc.inline_shapes:
        w = int(shp.width)
        if w > img_w:
            img_w = w
    img_w = min(img_w, text_w)

    return {"text_width": text_w, "img_width": img_w,
            "size_pt": size_pt, "bold": bold}


def clear_body(doc):
    from docx.oxml.ns import qn
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def build_doc(template_path, spec, art, img_paths, out_path,
              date_note, keep_gif, min_px=MIN_PX, min_area=MIN_AREA):
    from docx import Document
    from docx.shared import Emu, Pt
    from docx.image.image import Image as DocxImage
    from docx.oxml.ns import qn

    doc = Document(template_path)
    clear_body(doc)

    def styled(text):
        p = doc.add_paragraph()
        r = p.add_run(text)
        if spec["bold"]:
            r.bold = True
        r.font.size = Pt(spec["size_pt"])
        rpr = r._element.get_or_add_rPr()
        rpr.get_or_add_rFonts().set(qn("w:hint"), "eastAsia")
        return p

    y, m, d = art["date"].split("-")
    styled("%s.%s.%s%s" % (y, m, d, date_note))
    styled(art["title"])

    n_img = 0
    skips = []  # (原因, 图片URL, 尺寸)
    for kind, val in art["blocks"]:
        if kind == "p":
            txt = val.strip()
            if not txt or txt.upper() == "END":
                continue
            doc.add_paragraph().add_run(txt)
            continue

        path = img_paths.get(val, "")
        if not path or not os.path.exists(path) or os.path.getsize(path) < MIN_BYTES:
            skips.append(("下载失败", val, ""))
            continue
        if (not keep_gif) and is_gif(path):
            skips.append(("gif装饰", val, ""))
            continue

        usable = ensure_supported(path)
        if not usable:
            skips.append(("格式不支持", val, ""))
            continue
        try:
            info = DocxImage.from_file(usable)
        except Exception:
            skips.append(("无法解析", val, ""))
            continue
        if is_decorative(info.px_width, info.px_height, min_px, min_area):
            skips.append(("低清装饰图", val, "%dx%d" % (info.px_width, info.px_height)))
            continue

        width = min(Emu(info.px_width * 9525), Emu(spec["img_width"]))
        doc.add_paragraph().add_run().add_picture(usable, width=width)
        n_img += 1

    strip_empty_paragraphs(doc)

    # 清掉清空正文后残留的模板图片关系（否则体积白涨）
    used = set(re.findall(r'r:(?:embed|id)="(rId\d+)"', doc.element.body.xml))
    for rid, rel in list(doc.part.rels.items()):
        if "image" in rel.reltype and rid not in used:
            try:
                doc.part.drop_rel(rid)
            except Exception:
                pass

    doc.save(out_path)
    return n_img, skips


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="微信公众号推文批量转 Word（模板化、图片内嵌）")
    ap.add_argument("--links", required=True, help="链接清单 txt，一行一个 URL")
    ap.add_argument("--template", default=None,
                    help="可选：覆盖内置模板（默认用 skill 自带的脱敏模板）")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--column", default="学院动态", help="标题无栏目前缀时补的前缀，留空则不补")
    ap.add_argument("--date-note", default="（发出日期）", help="日期后紧跟的标注，留空则不加")
    ap.add_argument("--sort", choices=["auto", "asc", "desc", "none"], default="auto",
                    help="按发布日期排序；auto 表示识别清单里的“排序：”提示")
    ap.add_argument("--keep-gif", action="store_true", help="保留 gif 图片（默认删除装饰性 gif）")
    ap.add_argument("--min-px", type=int, default=MIN_PX,
                    help="图片最长边小于该像素视为装饰图标，默认 %d" % MIN_PX)
    ap.add_argument("--min-area", type=int, default=MIN_AREA,
                    help="图片像素面积小于该值视为低清装饰图，默认 %d" % MIN_AREA)
    ap.add_argument("--workers", type=int, default=8, help="并发数")
    ap.add_argument("--cache", default=None, help="中间缓存目录（缺省为系统临时目录）")
    args = ap.parse_args()

    if not args.links or not os.path.isfile(args.links):
        sys.exit("找不到链接清单：%s" % args.links)

    template = args.template or DEFAULT_TEMPLATE
    if not os.path.isfile(template):
        sys.exit("找不到模板文档：%s" % template)

    try:
        import docx  # noqa: F401
        import bs4  # noqa: F401
    except ImportError:
        sys.exit("缺少依赖，请先执行：python -m pip install python-docx beautifulsoup4 lxml")

    os.makedirs(args.out, exist_ok=True)
    tmp_root = args.cache or tempfile.mkdtemp(prefix="wx2docx_")
    os.makedirs(tmp_root, exist_ok=True)

    urls, hint = read_links(args.links)
    if not urls:
        sys.exit("清单中没有可用链接")
    print("链接 %d 条（已去重），模板：%s" % (len(urls), os.path.basename(template)))

    # 1) 抓页面
    def grab(i_url):
        i, u = i_url
        p = os.path.join(tmp_root, "page_%03d.html" % i)
        if os.path.exists(p) and os.path.getsize(p) > 20000:
            return i, p
        data = fetch(u)
        with open(p, "wb") as fh:
            fh.write(data)
        return i, p

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        pages = dict(ex.map(grab, list(enumerate(urls))))

    # 2) 解析
    arts, failed = [], []
    for i, u in enumerate(urls):
        p = pages.get(i)
        if not p or not os.path.exists(p) or os.path.getsize(p) < 20000:
            failed.append((u, "页面抓取失败"))
            continue
        art = parse_article(u, open(p, "rb").read())
        if not art["title"] or not art["date"] or not art["blocks"]:
            failed.append((u, "解析不完整（可能被拦截或需要验证）"))
            continue
        art["title"] = normalize_title(art["title"], args.column)
        arts.append(art)

    order = args.sort
    if order == "auto":
        order = hint or "none"
    if order in ("asc", "desc"):
        arts.sort(key=lambda a: a["date"], reverse=(order == "desc"))

    # 3) 下载图片
    img_urls = []
    for a in arts:
        for k, v in a["blocks"]:
            if k == "img" and v not in img_urls:
                img_urls.append(v)

    img_dir = os.path.join(tmp_root, "img")
    os.makedirs(img_dir, exist_ok=True)

    def ext_of(u):
        fmt = parse_qs(urlparse(u).query).get("wx_fmt", ["jpg"])[0].lower()
        return "jpg" if fmt in ("jpeg", "jpg") else fmt

    def dl(item):
        idx, u = item
        p = os.path.join(img_dir, "%03d.%s" % (idx, ext_of(u)))
        if os.path.exists(p) and os.path.getsize(p) > MIN_BYTES:
            return u, p
        download_image(u, p)
        return u, p

    img_paths = {}
    with ThreadPoolExecutor(max_workers=max(args.workers, 8)) as ex:
        for u, p in ex.map(dl, list(enumerate(img_urls))):
            img_paths[u] = p
    bad_img = sum(1 for u, p in img_paths.items()
                  if not os.path.exists(p) or os.path.getsize(p) < MIN_BYTES)
    print("图片 %d 张，下载失败 %d 张" % (len(img_urls), bad_img))

    # 4) 生成
    spec = read_template_spec(template)
    rows = []
    lost = []  # 真正丢内容的图片（下载失败/格式不支持）
    for a in arts:
        name = safe_name(a["date"], a["title"])
        dst = os.path.join(args.out, name)
        n_img, skips = build_doc(template, spec, a, img_paths, dst,
                                 args.date_note, args.keep_gif,
                                 args.min_px, args.min_area)
        tally = {}
        for reason, url, dims in skips:
            tally[reason] = tally.get(reason, 0) + 1
            if reason in ("下载失败", "格式不支持", "无法解析"):
                lost.append((name, "%s %s" % (dims, url)))
        detail = " ".join("%s%d" % (k, v) for k, v in sorted(tally.items()))
        rows.append([name, a["date"], a["title"], n_img, len(skips), detail, a["url"]])
        print("OK  %s  (图 %d%s)" % (name, n_img, ("；跳过 " + detail) if detail else ""))

    # 5) 清单
    manifest = os.path.join(args.out, "_manifest.csv")
    with open(manifest, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["文件名", "日期", "标题", "内嵌图片数", "跳过图片数", "跳过明细", "原文链接"])
        w.writerows(rows)

    print("\n完成：%d 篇 → %s" % (len(rows), args.out))
    if lost:
        print("⚠ 有 %d 张正文图片未能嵌入（需人工补图）：" % len(lost))
        for name, url in lost[:20]:
            print("  - [%s] %s" % (name[:24], url))
    if failed:
        print("失败 %d 条：" % len(failed))
        for u, why in failed:
            print("  - %s (%s)" % (u, why))
    print("清单：%s" % manifest)

    if not args.cache:
        shutil.rmtree(tmp_root, ignore_errors=True)


if __name__ == "__main__":
    main()
