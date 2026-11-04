#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
shorten_titles.py —— 文件名标题的「二次简化」工具（不重抓、不重排版）

两条用法：

1) 列出现有产物里的标题，供 AI 只看标题就写简化映射（不读正文、不重抓页面）：
       python shorten_titles.py --out "产物/10月" --list

2) 按映射表改名（同时更新 _manifest.csv）：
       python shorten_titles.py --out "产物/10月" --apply "命名.tsv"
   映射表两列，Tab 或逗号分隔，第一列可以写「原文标题」「URL」或「当前文件名」：
       情暖重阳，爱满新才！机电与建筑工程学院…传递敬老温情	青山路街道敬老
       我为同学办实事|清洗空调滤网…	清洗空调滤网

改名只动文件名，docx 内容一个字节都不碰；新名仍走 yyyy.mm.dd(简化标题).docx 规则。
"""

import argparse
import csv
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from title_naming import (NAME_TITLE_MAX, load_name_map, lookup_name,  # noqa: E402
                          safe_name, short_title)


def read_manifest(out_dir):
    path = os.path.join(out_dir, "_manifest.csv")
    if not os.path.isfile(path):
        sys.exit("找不到清单：%s（先跑 wx2docx.py）" % path)
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        sys.exit("清单为空：%s" % path)
    head, body = rows[0], rows[1:]
    idx = {name: i for i, name in enumerate(head)}
    return path, head, body, idx


def get(row, idx, key, default=0):
    i = idx.get(key)
    return row[i] if i is not None and i < len(row) else default


def cmd_list(out_dir, name_max):
    _, _, body, idx = read_manifest(out_dir)
    print("序号\t日期\t原文标题\t→ 规则简化\t当前文件名")
    for i, row in enumerate(body, 1):
        title = get(row, idx, "标题", "")
        cur = get(row, idx, "文件名", "")
        print("%02d\t%s\t%s\t→ %s\t%s" % (
            i, get(row, idx, "日期", ""), title,
            short_title(title, name_max), cur))
    print("\n共 %d 篇。只根据上面「原文标题」写简化名即可，不需要读正文。" % len(body))
    print("写法：新建 命名.tsv，一行「原文标题<TAB>简化名」，然后 --apply 命名.tsv。")


def cmd_apply(out_dir, map_path, name_max):
    mapping = load_name_map(map_path)
    if not mapping:
        sys.exit("映射表为空：%s" % map_path)
    path, head, body, idx = read_manifest(out_dir)
    if "简化标题" not in idx:  # 兼容旧清单
        head.insert(idx.get("标题", 0) + 1, "简化标题")
        for row in body:
            row.insert(idx.get("标题", 0) + 1, "")
        idx = {n: i for i, n in enumerate(head)}

    renamed, locked, missed = 0, [], []
    for row in body:
        date = get(row, idx, "日期", "")
        title = get(row, idx, "标题", "")
        cur_name = get(row, idx, "文件名", "")
        cur_short = get(row, idx, "简化标题", "")
        url = get(row, idx, "原文链接", "")

        new_short = (lookup_name(mapping, url, title)
                     or mapping.get(cur_name) or mapping.get(cur_short) or "")
        if not new_short:
            # 没给映射就维持现状，顺手把清单里的「简化标题」补齐
            keep = cur_short or (cur_name[11:-5] if cur_name.endswith(".docx")
                                 and len(cur_name) > 16 else short_title(title, name_max))
            row[idx["简化标题"]] = keep
            missed.append(cur_name)
            continue
        new_short = new_short.strip()
        new_name = safe_name(date, title, new_short, name_max)
        if new_name == cur_name:
            row[idx["简化标题"]] = new_short
            continue

        src = os.path.join(out_dir, cur_name)
        dst = os.path.join(out_dir, new_name)
        if not os.path.isfile(src):
            missed.append(cur_name)
            continue
        if os.path.exists(dst):
            print("SKIP 目标已存在：%s" % new_name)
            continue
        try:
            fh = open(src, "r+b")  # 占用探测：被 Word/WPS 打开就改名失败
            fh.close()
            os.rename(src, dst)
        except Exception as exc:
            locked.append((cur_name, str(exc)))
            continue
        row[idx["文件名"]] = new_name
        row[idx["简化标题"]] = new_short
        renamed += 1
        print("RENAME %s → %s" % (cur_name, new_name))

    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(head)
        w.writerows(body)

    print("\n改名 %d 篇；未命中映射 %d 篇" % (renamed, len(missed)))
    if missed:
        print("未命中（映射表里没这一篇，标题或 URL 要能对上）：")
        for n in missed[:20]:
            print("  - %s" % n)
    if locked:
        print("⚠ 被占用/改名失败（先在 Word/WPS 里关闭）：")
        for n, why in locked[:20]:
            print("  - %s (%s)" % (n, why))
    extra = [f for f in glob.glob(os.path.join(out_dir, "*.docx"))
             if os.path.basename(f).startswith("~$")]
    if extra:
        print("提示：目录里有 %d 个 ~$ 锁文件，属于 Word 临时文件，别删也别当产物。" % len(extra))
    print("清单已更新：%s；建议再跑一次 check_output.py" % path)


def main():
    ap = argparse.ArgumentParser(description="文件名标题的二次简化（只看标题，不重抓）")
    ap.add_argument("--out", required=True, help="产物目录（含 _manifest.csv）")
    ap.add_argument("--list", action="store_true", help="列出标题，供写简化映射")
    ap.add_argument("--apply", default=None, help="按映射表改名（两列，Tab/逗号分隔）")
    ap.add_argument("--name-max", type=int, default=NAME_TITLE_MAX,
                    help="标题字数上限，默认 %d" % NAME_TITLE_MAX)
    args = ap.parse_args()

    if args.apply:
        cmd_apply(args.out, args.apply, args.name_max)
    else:
        cmd_list(args.out, args.name_max)


if __name__ == "__main__":
    main()
